"""文档构建与部署任务。

提供命令行接口来管理文档构建和部署。
"""

import re
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
    # 本仓发布走本地 inv docs.deploy（不经 CI），故在此前置对账，
    # 避免错误码漂移绕过 check.yml 直接进入产物
    check_error_codes(c)
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

    # mike deploy 自行构建、不经 docs.build，故在此单独前置对账
    check_error_codes(c)

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


_SPEC_DIR = _PROJECT_ROOT / "docs" / "specification"
_REGISTRY_FILE = _SPEC_DIR / "error-handling.md"

# 注册表行（权威）：| `3000` | `DOCUMENT_ERROR` | 文档操作错误（通用） |
_REGISTRY_ROW = re.compile(r"^\|\s*`(\d{4})`\s*\|\s*`([A-Z_]+)`\s*\|")

# 引用体例。规范内实际存在多种，缺一即留下不被对账的死角：
# (a) 事件表「可能的错误」：| 3000 | `DOCUMENT_ERROR` - 说明 |
_CITE_EVENT_ROW = re.compile(r"^\s*\|\s*(\d{4})\s*\|\s*\**`([A-Z_]+)`")
# (b) 同格写法（conventions.md）：| 条件 | `4001 MISSING_PARAM` | details |
_CITE_INLINE = re.compile(r"`(\d{4})\s+([A-Z_]+)`")
# (c) 名称后紧跟括号编号：小节标题 `### TIMEOUT (1002)`、正文 `` `SELECTION_EMPTY` (3002) ``、
#     `PROTOCOL_VERSION_MISMATCH（2006）`、`HANDSHAKE_FAILED（复用 2003）`
_CITE_NAME_PAREN = re.compile(r"([A-Z_]{4,})`?\s*[(（][^)）]{0,4}?(\d{4})[)）]")
# (e) 锚点：`error-handling.md#protocol_version_mismatch-2006`、`{#element_not_found-3010}`
_CITE_ANCHOR = re.compile(r"#([a-z_]{4,})-(\d{4})\b")
# (d) 规范层映射表（events-excel.md）按「相邻单元格」解析，兼容单码与并列多码：
#     | 触发条件 | `3010` | `ELEMENT_NOT_FOUND` | kind | 事件 |
#     | 参数缺失/非法/越界 | `4001`/`4002`/`4004` | `MISSING_PARAM` / `INVALID_PARAM` / … | — |
_CELL_CODES = re.compile(r"`(\d{4})`")
_CELL_NAMES = re.compile(r"`([A-Z_]+)`")

# 码段数字（1xxx–4xxx）。行内同时出现它与已知错误码名称，才可能是一处「配对」引用；
# 只提名称不带号的散文（如 `"code": "HANDSHAKE_FAILED"`）无配对可对账，不在守护范围。
# 注意：不得排除前置反引号——(b)/(d) 体例的编号本就带反引号，排除即在那里开出盲区。
_CODE_TOKEN = re.compile(r"(?<![\d.])[1-4]\d{3}(?![\d.])")

# 引用处数下限：粗粒度兜底，仅用于捕获「体例整体脱节导致解析面坍塌」。
# 主防线是下方的嗅探器——部分萎缩（如少数体例失配）靠下限拦不住，故此处不写死实际值。
_CITATION_FLOOR = 300

# 可达性豁免码段。1xxx（通用失败）适用于**所有**操作、2xxx 属**连接生命周期**（不在请求-响应
# 事件的错误面内），二者均非事件表可枚举项——要求它们被引用只会逼出噪声引用。
# 对齐注册表自身的分节语义（error-handling.md「1xxx - 通用错误」/「2xxx - 连接与认证错误」）。
# 3xxx/4xxx 描述具体的、事件可触发的条件，无引用即**无人可发出** —— 是缺陷，不是风格问题。
_REACHABILITY_EXEMPT_BANDS = ("1", "2")


def _load_registry() -> set[tuple[str, str]]:
    """解析 error-handling.md 的权威错误码注册表，返回 (编号, 名称) 配对集合。"""
    text = _REGISTRY_FILE.read_text(encoding="utf-8")
    pairs = {
        (m.group(1), m.group(2)) for line in text.splitlines() if (m := _REGISTRY_ROW.match(line))
    }
    if not pairs:
        raise RuntimeError(f"未能从 {_REGISTRY_FILE} 解析出任何注册表条目——正则可能已与文档体例脱节")
    return pairs


def _citations(line: str) -> list[tuple[str, str]]:
    """从一行中提取全部 (编号, 名称) 引用，覆盖规范内各种体例。"""
    found: list[tuple[str, str]] = []

    if m := _CITE_EVENT_ROW.match(line):
        found.append((m.group(1), m.group(2)))
    found += [(m.group(2), m.group(1)) for m in _CITE_NAME_PAREN.finditer(line)]
    found += [(m.group(2), m.group(1).upper()) for m in _CITE_ANCHOR.finditer(line)]
    found += [(m.group(1), m.group(2)) for m in _CITE_INLINE.finditer(line)]

    # 表格单元格按位配对，个数相等才配（覆盖单码、并列多码、码名同格三种）
    if line.lstrip().startswith("|"):
        cells = line.split("|")
        for cell, nxt in zip(cells, cells[1:]):
            codes = _CELL_CODES.findall(cell)
            if not codes:
                continue
            # 码格在前、名格紧随；或码与名同格
            for names in (_CELL_NAMES.findall(nxt), _CELL_NAMES.findall(cell)):
                if len(codes) == len(names):
                    found += list(zip(codes, names))
                    break
    return found


# 自检样本：(说明, 行, 期望解析出的配对数)。锁住各体例的解析行为，
# 尤其是「体例变了必须大声失败、不得静默放行」——期望 0 即表示该行须落入嗅探器。
_SELF_TEST: tuple[tuple[str, str, int], ...] = (
    ("事件表行", "| 3000 | `DOCUMENT_ERROR` - 文档操作错误 |", 1),
    ("事件表行·加粗名称", "| 3000 | **`DOCUMENT_ERROR`** - x |", 1),
    ("规范层映射表·单码", "| 定位失败 | `3010` | `ELEMENT_NOT_FOUND` | kind | x |", 1),
    ("规范层映射表·并列多码", "| 参数 | `4001`/`4002` | `MISSING_PARAM` / `INVALID_PARAM` | — |", 2),
    ("同格写法", "| 缺 script | `4001 MISSING_PARAM` | — |", 1),
    ("同格·码名分置同格", "| 条件 | `3004` `OPERATION_FAILED` | x |", 1),
    ("名称+括号编号", "返回错误码 `SELECTION_EMPTY` (3002)。", 1),
    ("锚点", "见 [错误处理](error-handling.md#element_not_found-3010)。", 1),
    # 以下必须解析不出配对（→ 嗅探器接管、大声失败），不得静默放行
    ("体例脱节·名称改链接", "| 3000 | [`DOCUMENT_ERROR`](x.md) - x |", 0),
    ("体例脱节·码数≠名数", "| 参数 | `4001`/`4002` | `MISSING_PARAM` / `INVALID_PARAM` / `PARAM_OUT_OF_RANGE` | — |", 0),
)


def _reachable_from(line: str, registry_file: bool) -> set[str]:
    """该行**为可达性**贡献的错误码编号。承重判据集中在此，故可被自检锁住。

    两处排除缺一，闸口即静默失效——孤儿码会自我满足，#23 一类缺陷重新放行：
    1. 注册表文件内的引用不算——否则 `### TIMEOUT (1002)` 这类**详解小节标题**
       会让码给自己背书（每个码都有详解小节 = 每个码都「可达」）。
    2. 只认**表格行**——可达性要断言的是「**有事件会发出它**」，而事件×码映射
       只存在于表格里。一条锚点链接（`error-handling.md#search_no_match-3012`）
       不构成任何事件会发出该码的证据，不得用来续命。
    """
    if registry_file or not line.lstrip().startswith("|"):
        return set()
    return {code for code, _ in _citations(line)}


def _orphans(registry: set[tuple[str, str]], reachable: set[str]) -> list[tuple[str, str]]:
    """注册了却无人引用的 3xxx/4xxx 码。`reachable` 应由 `_reachable_from` 累积。"""
    return sorted(
        (code, name)
        for code, name in registry
        if code[0] not in _REACHABILITY_EXEMPT_BANDS and code not in reachable
    )


# 可达性判据样本：(说明, 行, 是否注册表文件, 期望贡献的码)
_REACH_LINE_SELF_TEST: tuple[tuple[str, str, bool, set[str]], ...] = (
    ("事件表行 → 贡献可达性", "| 3012 | `SEARCH_NO_MATCH` - x |", False, {"3012"}),
    ("规范层映射表行 → 贡献", "| 定位失败 | `3010` | `ELEMENT_NOT_FOUND` | k | x |", False, {"3010"}),
    # 以下均**不得**贡献，否则孤儿码可自我满足
    ("注册表内的详解小节标题 → 不贡献", "### TIMEOUT (1002)", True, set()),
    ("注册表内的表格行 → 不贡献", "| 3012 | `SEARCH_NO_MATCH` - x |", True, set()),
    ("注册表外的锚点链接 → 不贡献（非事件×码证据）", "见 [x](error-handling.md#search_no_match-3012)。", False, set()),
    ("注册表外的散文提及 → 不贡献", "本码承 `4003 INVALID_PARAM_TYPE` 的先例。", False, set()),
)

# 孤儿判定样本：(说明, 注册表, 可达集, 期望孤儿数)
_REACH_SELF_TEST: tuple[tuple[str, set[tuple[str, str]], set[str], int], ...] = (
    ("3xxx 有引用 → 不是孤儿", {("3012", "SEARCH_NO_MATCH")}, {"3012"}, 0),
    ("3xxx 零引用 → 孤儿", {("3012", "SEARCH_NO_MATCH")}, set(), 1),
    ("4xxx 零引用 → 孤儿", {("4004", "PARAM_OUT_OF_RANGE")}, set(), 1),
    ("1xxx 零引用 → 豁免（通用失败，跨切面）", {("1000", "UNKNOWN")}, set(), 0),
    ("2xxx 零引用 → 豁免（连接生命周期）", {("2005", "CONNECTION_LOST")}, set(), 0),
)


def _run_self_test() -> None:
    """用合成样本断言各体例的解析行为与可达性判法，与主逻辑共用同一组正则/函数。"""
    known = {n for _, n in _load_registry()}
    failures = []
    for label, line, expect in _SELF_TEST:
        got = len(_citations(line))
        if got != expect:
            failures.append(f"  {label}：期望解析 {expect} 处，实际 {got} 处 —— {line}")
            continue
        # 期望 0 的样本，必须能被嗅探器接住（否则就是静默放行）
        if expect == 0:
            sniffed = _CODE_TOKEN.search(line) and any(n in line for n in known)
            if not sniffed:
                failures.append(f"  {label}：解析不出配对且嗅探器未触发 —— 会静默放行！{line}")
    for label, line, in_registry, expect in _REACH_LINE_SELF_TEST:
        got = _reachable_from(line, in_registry)
        if got != expect:
            failures.append(f"  可达性判据·{label}：期望贡献 {expect or '{}'}，实际 {got or '{}'} —— {line}")
    for label, reg, reach, expect in _REACH_SELF_TEST:
        got = len(_orphans(reg, reach))
        if got != expect:
            failures.append(f"  孤儿判定·{label}：期望 {expect} 个孤儿，实际 {got} 个")
    if failures:
        print("❌ 守护器自检失败（判法已偏离预期）：\n" + "\n".join(failures))
        raise SystemExit(1)
    reach_n = len(_REACH_LINE_SELF_TEST) + len(_REACH_SELF_TEST)
    print(f"✅ 守护器自检通过（{len(_SELF_TEST)} 个体例样本 + {reach_n} 个可达性样本）")


@task(help={"self_test": "仅跑守护器自身的体例自检，不扫描文档"})
def check_error_codes(c: Context, self_test: bool = False) -> None:
    """校验规范内引用的 (错误码, 名称) 配对精确命中 error-handling.md 权威注册表。

    守护 #17 / #20 一类的历史漂移：注册表重排编号后引用方未跟进，
    导致同一名称在不同文件/段落挂着不同编号。
    """
    _run_self_test()
    if self_test:
        return

    registry = _load_registry()
    by_code = {code: name for code, name in registry}
    by_name = {name: code for code, name in registry}
    known_names = set(by_name)

    violations: list[tuple[str, int, str]] = []
    unparsed: list[tuple[str, int, str]] = []
    reachable: set[str] = set()
    seen = 0

    for path in sorted(_SPEC_DIR.glob("*.md")):
        rel = str(path.relative_to(_PROJECT_ROOT))
        registry_file = path == _REGISTRY_FILE
        for lineno, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            # 注册表自身的定义行是权威，不是对它的引用
            if registry_file and _REGISTRY_ROW.match(line):
                continue
            cites = _citations(line)
            seen += len(cites)
            reachable |= _reachable_from(line, registry_file)
            for code, name in cites:
                if (code, name) in registry:
                    continue
                if name not in by_name:
                    hint = f"注册表无名称 `{name}`" + (
                        f"；编号 {code} 实为 `{by_code[code]}`"
                        if code in by_code
                        else f"；编号 {code} 亦未定义"
                    )
                elif code not in by_code:
                    hint = f"编号 {code} 未定义；`{name}` 实为 {by_name[name]}"
                else:
                    hint = f"编号 {code} 实为 `{by_code[code]}`；`{name}` 实为 {by_name[name]}"
                violations.append((rel, lineno, f"{code} `{name}`  →  {hint}"))
            # 嗅探：行内同时出现已知名称与码段数字（即像一处配对），却一处也没解析出来
            # —— 多半是体例变了而正则没跟，此时必须大声失败，不能以「零违规」放行
            if not cites and _CODE_TOKEN.search(line) and any(n in line for n in known_names):
                unparsed.append((rel, lineno, line.strip()[:110]))

    if unparsed:
        print(f"❌ {len(unparsed)} 行疑似错误码引用但未被任何体例正则识别（正则可能已与文档体例脱节）：\n")
        for rel, lineno, text in unparsed:
            print(f"  {rel}:{lineno}  {text}")
        raise SystemExit(1)

    if seen < _CITATION_FLOOR:
        raise RuntimeError(
            f"仅解析出 {seen} 处错误码引用，低于下限 {_CITATION_FLOOR}——"
            "正则可能已与文档体例脱节，拒绝以「零违规」放行"
        )

    if violations:
        print(f"❌ {len(violations)} 处错误码引用与权威注册表不符：\n")
        for rel, lineno, detail in violations:
            print(f"  {rel}:{lineno}  {detail}")
        raise SystemExit(1)

    orphans = _orphans(registry, reachable)
    if orphans:
        print(f"❌ {len(orphans)} 个孤儿错误码：已在注册表定义，但无任何事件/规范引用（调用方永远收不到）：\n")
        for code, name in orphans:
            print(f"  {code} `{name}`")
        print(
            "\n  每个 3xxx/4xxx 码都须至少被一处事件表或规范层映射表引用。"
            "\n  要么接上（补引用），要么退役（从注册表删除）——不得两头挂空。"
        )
        raise SystemExit(1)

    exempt = sum(1 for code, _ in registry if code[0] in _REACHABILITY_EXEMPT_BANDS)
    print(
        f"✅ {seen} 处错误码引用全部命中注册表（{len(registry)} 个已注册码）"
        f"；{len(registry) - exempt} 个 3xxx/4xxx 码均可达（{exempt} 个 1xxx/2xxx 跨切面码豁免）"
    )


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
docs_tasks.add_task(check_error_codes, name="check-error-codes")
docs_tasks.add_task(clean)
