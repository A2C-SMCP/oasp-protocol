"""文档构建与部署任务。

提供命令行接口来管理文档构建和部署。
"""

import sys

from invoke.collection import Collection
from invoke.context import Context
from invoke.tasks import task

from .config import DeployConfig

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


def update_server() -> None:
    """触发服务器 Git pull 更新文档。

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
) -> None:
    """部署文档到 GitHub Pages 并同步到文档服务器。

    Args:
        version: 版本号（默认: 0.1.0）。
        alias: 版本别名（默认: latest）。
        sync: 是否同步到文档服务器（默认: True）。

    Examples:
        inv docs.deploy                              # 部署 0.1.0 [latest] 并同步
        inv docs.deploy --version=0.2.0              # 部署新版本 0.2.0 [latest]
        inv docs.deploy --version=0.1.0 --no-sync    # 仅部署到 GitHub Pages
    """
    print(f"🚀 部署文档 (version={version}, alias={alias})")

    # 验证配置（如果需要同步到服务器）
    if sync:
        errors = config.validate()
        if errors:
            print("❌ 配置错误:")
            for error in errors:
                print(f"   - {error}")
            sys.exit(1)

    # Step 0: 同步远程 gh-pages 分支（避免多人协作冲突）
    sync_gh_pages(c)

    # Step 1: 使用 mike 部署到 GitHub Pages
    print(f"📦 部署版本 '{version}' [{alias}] 到 GitHub Pages...")
    c.run(f"mike deploy --push --update-aliases {version} {alias}")

    # Step 2: 同步到文档服务器
    if sync:
        print(f"🔄 同步到文档服务器 ({config.server.host})...")
        update_server()

    print("✅ 部署完成")


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
docs_tasks.add_task(serve_versioned, name="serve-versioned")
docs_tasks.add_task(update_server_task, name="update-server")
docs_tasks.add_task(clean)
