"""文档构建与部署任务。

提供命令行接口来管理文档构建和部署。
"""

import shlex
import shutil
import subprocess
import sys
from pathlib import Path

from invoke.collection import Collection
from invoke.context import Context
from invoke.tasks import task

from .config import DeployConfig

_PROJECT_ROOT = Path(__file__).resolve().parents[2]

# 加载配置
config = DeployConfig.from_env()


@task
def serve(c: Context) -> None:
    """启动本地开发服务器（热重载）。"""
    print("🚀 启动 MkDocs 开发服务器 (http://127.0.0.1:8000)")
    c.run("mkdocs serve", pty=True)


@task
def build(c: Context) -> None:
    """构建静态文档。"""
    print("🔨 构建文档...")
    c.run("mkdocs build")
    print("✅ 文档构建完成")


@task
def serve_versioned(c: Context) -> None:
    """启动版本化文档本地服务器。"""
    print("🚀 启动 Mike 版本化文档服务器")
    c.run("mike serve", pty=True)


def sync_gh_pages(c: Context) -> None:
    """同步远程 gh-pages 分支到本地。

    在多人协作场景下，先同步远程分支可避免推送时的 non-fast-forward 冲突。
    Mike 会在同步后的分支基础上增量更新，确保不丢失其他人部署的版本。
    """
    print("🔄 同步远程 gh-pages 分支...")

    # 检查远程 gh-pages 分支是否存在
    result = c.run("git ls-remote --heads origin gh-pages", warn=True, hide=True)
    if not result.stdout.strip():
        print("   远程 gh-pages 分支不存在，跳过同步（首次部署）")
        return

    # 获取远程分支最新状态
    c.run("git fetch origin gh-pages:gh-pages", warn=True)
    print("   ✅ 同步完成")


def _ssh_connect_kwargs() -> dict:
    """构造 paramiko.connect 的 kwargs。优先密钥、其次密码。"""
    cfg = config.server
    kwargs: dict = {"hostname": cfg.host, "port": cfg.port, "username": cfg.user}
    if cfg.key_filename:
        kwargs["key_filename"] = cfg.key_filename
    elif cfg.password:
        kwargs["password"] = cfg.password
    else:
        raise RuntimeError("未配置 SSH 密码或密钥")
    return kwargs


def sync_via_tar(c: Context) -> None:
    """把本地 gh-pages 通过 tar + SSH 流式推送到文档服务器（绕过 GitHub）。

    流程：本地 `git archive gh-pages` | paramiko ssh | 服务器 `tar -xpf -`
    解压到 `deploy_path.new`，成功后用 mv 做原子替换。失败则清理 staging
    并尝试回滚。

    前置：本地 gh-pages 分支必须存在且内容是最新（通常 mike deploy 后即满足）。
    """
    import paramiko

    errors = config.validate()
    if errors:
        print("❌ 配置错误:")
        for error in errors:
            print(f"   - {error}")
        sys.exit(1)

    result = c.run("git rev-parse --verify gh-pages", warn=True, hide=True)
    if not result.ok:
        print("❌ 本地 gh-pages 分支不存在。请先运行 mike deploy 或 "
              "`git fetch origin gh-pages:gh-pages`。")
        sys.exit(1)

    path = config.server.deploy_path
    quoted = shlex.quote(path)
    new_path = f"{quoted}.new"
    old_path = f"{quoted}.old"

    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(**_ssh_connect_kwargs())

    try:
        print(f"   📦 准备 staging: {path}.new")
        _, stdout, stderr = ssh.exec_command(
            f"rm -rf {new_path} && mkdir -p {new_path}"
        )
        rc = stdout.channel.recv_exit_status()
        if rc != 0:
            raise RuntimeError(f"staging 准备失败: {stderr.read().decode()}")

        print("   📤 git archive gh-pages → 服务器 tar 解压")
        stdin, stdout, stderr = ssh.exec_command(f"tar -xpf - -C {new_path}")
        tar_proc = subprocess.Popen(
            ["git", "archive", "--format=tar", "gh-pages"],
            stdout=subprocess.PIPE,
            cwd=str(_PROJECT_ROOT),
        )
        try:
            assert tar_proc.stdout is not None
            shutil.copyfileobj(tar_proc.stdout, stdin)
            stdin.channel.shutdown_write()
            remote_rc = stdout.channel.recv_exit_status()
            local_rc = tar_proc.wait()
            if local_rc != 0:
                raise RuntimeError(f"git archive 失败 (rc={local_rc})")
            if remote_rc != 0:
                err = stderr.read().decode()
                raise RuntimeError(f"远端 tar 解压失败 (rc={remote_rc}): {err}")
        finally:
            if tar_proc.stdout:
                tar_proc.stdout.close()
            if tar_proc.poll() is None:
                tar_proc.terminate()

        print(f"   🔀 原子替换 {path}")
        switch_cmd = (
            "set -e; "
            f"rm -rf {old_path}; "
            f"if [ -d {quoted} ]; then mv {quoted} {old_path}; fi; "
            f"mv {new_path} {quoted}; "
            f"rm -rf {old_path}"
        )
        _, stdout, stderr = ssh.exec_command(switch_cmd)
        rc = stdout.channel.recv_exit_status()
        if rc != 0:
            err = stderr.read().decode()
            ssh.exec_command(
                f"if [ -d {old_path} ] && [ ! -d {quoted} ]; then "
                f"mv {old_path} {quoted}; fi; rm -rf {new_path}"
            )
            raise RuntimeError(f"原子替换失败: {err}")

        _, stdout, _ = ssh.exec_command(
            f"cd {quoted} && "
            "printf 'latest → '; readlink latest 2>/dev/null || echo '(missing)'; "
            "echo '--- versions.json ---'; "
            "cat versions.json 2>/dev/null || echo '(missing)'"
        )
        verify = stdout.read().decode().rstrip()
        if verify:
            print("   🔍 服务器状态：")
            for line in verify.splitlines():
                print(f"      {line}")

    except Exception:
        try:
            ssh.exec_command(f"rm -rf {new_path}")
        except Exception:
            pass
        raise
    finally:
        ssh.close()


def update_server() -> None:
    """触发服务器 Git pull 更新文档（fallback 方案，受 GitHub 网络影响）。

    使用 paramiko SSH 连接到服务器并执行 Git pull。
    """
    import paramiko

    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    try:
        # 连接服务器
        if config.server.password:
            ssh.connect(
                config.server.host,
                port=config.server.port,
                username=config.server.user,
                password=config.server.password,
            )
        elif config.server.key_filename:
            ssh.connect(
                config.server.host,
                port=config.server.port,
                username=config.server.user,
                key_filename=config.server.key_filename,
            )
        else:
            print("⚠️  未配置密码或密钥文件，跳过服务器更新")
            return

        # 执行 Git pull
        cmd = f"cd {config.server.deploy_path} && git fetch origin gh-pages && git reset --hard origin/gh-pages"
        stdin, stdout, stderr = ssh.exec_command(cmd)

        exit_code = stdout.channel.recv_exit_status()
        output = stdout.read().decode()
        error = stderr.read().decode()

        if exit_code == 0:
            print(f"   ✅ 服务器更新成功:\n{output}")
        else:
            print(f"   ⚠️  服务器更新警告:\n{error}")

    except Exception as e:
        print(f"   ❌ 服务器更新失败: {e}")
    finally:
        ssh.close()


@task
def deploy(
    c: Context,
    version: str = "0.1.0",
    alias: str = "latest",
    sync: bool = True,
    sync_via: str = "tar",
) -> None:
    """部署文档到 GitHub Pages 并同步到文档服务器。

    Args:
        version: 版本号（默认: 0.1.0，通常显式指定，如 --version=0.1.9）。
        alias: 版本别名（默认: latest）。
        sync: 是否同步到文档服务器（默认: True）。
        sync_via: 同步方式 "tar"（本地 git archive 直推，默认，绕开 GitHub）
                  或 "git"（服务器端 git pull，受 GitHub 网络影响）。

    Examples:
        inv docs.deploy --version=0.1.9                          # tar 同步
        inv docs.deploy --version=0.1.9 --sync-via=git           # git pull 同步
        inv docs.deploy --version=0.1.9 --no-sync                # 只推 GitHub
    """
    if sync and sync_via not in ("tar", "git"):
        print(f"❌ 未知 sync_via: {sync_via}（可选: tar, git）")
        sys.exit(1)

    print(f"🚀 部署文档 (version={version}, alias={alias})")

    if sync:
        errors = config.validate()
        if errors:
            print("❌ 配置错误:")
            for error in errors:
                print(f"   - {error}")
            sys.exit(1)

    sync_gh_pages(c)

    print(f"📦 部署版本 '{version}' [{alias}] 到 GitHub Pages...")
    c.run(f"mike deploy --push --update-aliases {version} {alias}")

    if sync:
        if sync_via == "tar":
            print(f"🔄 tar-over-SSH 同步到文档服务器 ({config.server.host})...")
            sync_via_tar(c)
        else:
            print(f"🔄 git pull 同步到文档服务器 ({config.server.host})...")
            update_server()

    print("✅ 部署完成")


@task
def push_to_server(c: Context) -> None:
    """仅把本地 gh-pages 通过 tar-over-SSH 推到文档服务器（不 redeploy）。

    用于：部署到 GitHub Pages 之后单独触发服务器同步，或 gh-pages 已是最新、
    仅需同步服务器。整个过程不经过服务器到 GitHub 的网络。
    """
    print(f"🔄 tar-over-SSH 同步到文档服务器 ({config.server.host})...")
    sync_via_tar(c)
    print("✅ 同步完成")


@task
def update_server_task(c: Context) -> None:
    """手动触发服务器更新（Git pull）。"""
    print("🔄 触发服务器更新...")

    errors = config.validate()
    if errors:
        print("❌ 配置错误:")
        for error in errors:
            print(f"   - {error}")
        sys.exit(1)

    update_server()
    print("✅ 更新完成")


@task
def clean(c: Context) -> None:
    """清理构建产物。"""
    c.run("rm -rf site/", warn=True)
    print("✅ 清理完成")


# 创建任务集合
docs_tasks = Collection("docs")
docs_tasks.add_task(serve)
docs_tasks.add_task(build)
docs_tasks.add_task(deploy)
docs_tasks.add_task(push_to_server, name="push-to-server")
docs_tasks.add_task(serve_versioned, name="serve-versioned")
docs_tasks.add_task(update_server_task, name="update-server")
docs_tasks.add_task(clean)
