from __future__ import annotations

import argparse
import asyncio
import contextlib
import csv
import hashlib
import html
import json
import os
import re
import shutil
import sqlite3
import sys
import traceback
from collections import deque
from io import BytesIO
from dataclasses import dataclass
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Any, AsyncIterator, Callable

import httpx

from wechat_decrypt_tool.modules.constants import (
    MEDIA_TYPE_IMAGE, MEDIA_TYPE_VIDEO, MEDIA_TYPE_LIVE_PHOTO,
    POST_TYPE_NORMAL, POST_TYPE_ARTICLE, POST_TYPE_LINK, POST_TYPE_COVER,
    POST_TYPE_FINDER, POST_TYPE_MUSIC,
    DEFAULT_PAGE_LIMIT, MAX_PAGE_LIMIT,
    DB_KEY_HEX_LENGTH,
    IMAGE_AES_KEY_LENGTH, IMAGE_XOR_KEY_MAX,
    PDF_FONT_NAME, PDF_FONT_FALLBACK, PDF_DPI,
    PDF_MARGIN_LEFT, PDF_MARGIN_RIGHT, PDF_MARGIN_TOP, PDF_MARGIN_BOTTOM,
)
from wechat_decrypt_tool.modules.wechat_emoji import emojify_wechat_shortcodes

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
CONFIG_DIR = PROJECT_ROOT / "config"
DEFAULT_CONFIG = CONFIG_DIR / "config.json"
RUNTIME_DIR = PROJECT_ROOT / "runtime"
OUTPUT_RUNTIME_DIR = RUNTIME_DIR / "output"

os.environ.setdefault("WECHAT_TOOL_DATA_DIR", str(RUNTIME_DIR))
os.environ.setdefault("WECHAT_TOOL_OUTPUT_DIR", str(OUTPUT_RUNTIME_DIR))
os.environ.setdefault("WECHAT_TOOL_BUILD_SESSION_LAST_MESSAGE", "0")


@dataclass(frozen=True)
class AccountInfo:
    account: str
    wxid_dir: Path
    db_storage_dir: Path


@dataclass
class ExportedPost:
    time_text: str
    display: str
    location: str
    body: str
    images: list[str]
    interactions: str = ""


@dataclass(frozen=True)
class ContactEntry:
    username: str
    remark: str
    nickname: str
    other_info: str = ""
    privacy_status: str = ""
    source_table: str = ""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="导出微信朋友圈为 Markdown、HTML 和 PDF")
    parser.add_argument("--key", default="", help=argparse.SUPPRESS)
    parser.add_argument("--start", default="", help="起始时间，例如 20260606；留空表示不限制")
    parser.add_argument("--end", default="", help="终止时间，例如 20260606；留空表示不限制")
    parser.add_argument("--only-self", default="", help="是否只导出自己的朋友圈: y/n，默认 y（y=只自己，n=自己+好友）")
    parser.add_argument("--export-scope", default="", help="导出范围: self(只自己)/friends(只好友)/all(自己+好友)；优先于 --only-self")
    parser.add_argument("--keep-interactions", default="", help="是否保留点赞评论: y/n，默认 n")
    parser.add_argument("--export-contacts", default="", help="是否导出好友列表: y/n，默认 n")
    parser.add_argument("--md-only", default="", help="是否只导出 Markdown 和图片（跳过 PDF）: y/n，默认 n")
    parser.add_argument("--export-excel", default="", help="是否导出 Excel 表格（moments.xlsx）: y/n，默认 y")
    parser.add_argument("--clean-data", default="", help="是否清洗数据（剔除无图/非商品，生成 moments_clean.xlsx）: y/n，默认 n")
    parser.add_argument("--skip-non-product", default="", help="下载前预判，非商品朋友圈跳过图片下载（省流量/OSS）: y/n，默认 y")
    parser.add_argument("--skip-qrcode", default="", help="跳过含二维码的商品图（先取缩略图检测，命中不下载原图）: y/n，默认 y")
    parser.add_argument("--image-inspect", default="", help="清洗第三阶段：对已下载图片做二维码+OCR手机号检测，命中图隔离、无图商品剔除: y/n，默认 y")
    parser.add_argument("--export-merchant", default="", help="是否导出商家导入表格（商家导入.xlsx）: y/n，默认 n")
    parser.add_argument("--export-type", default="", help="导出类型: friends(好友/商家表格)/moments(朋友圈商品)；交互模式可直接选择，CLI 不填默认 moments")
    parser.add_argument("--merchant-template", default="", help="商家导入模板路径，默认 config/商家导入模板.xlsx")
    parser.add_argument("--friend-filter-mode", default="", help="好友筛选方式: all(全部)/prefix(按前缀)/manual(手动)；仅在导出范围为 friends/all 时生效")
    parser.add_argument("--friend-prefix", default="", help="按好友备注名前缀筛选导出，例如 服装")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help=argparse.SUPPRESS)
    parser.add_argument("--output-root", default="", help=argparse.SUPPRESS)
    parser.add_argument("--no-download", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--quiet", action="store_true", help=argparse.SUPPRESS)
    return parser.parse_args()


def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        raise ValueError(f"配置文件读取失败: {path}: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError("配置文件必须是 JSON 对象")
    return data


def save_config(path: Path, config: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def ensure_imports() -> None:
    if (SRC_ROOT / "wechat_decrypt_tool").exists():
        sys.path.insert(0, str(SRC_ROOT))
        return
    raise FileNotFoundError("缺少 src/wechat_decrypt_tool，请确认项目文件完整")


class RestartFlow(Exception):
    """用户在任意输入处输入 r，要求放弃本次、从头重新开始。"""


def ask_input(value: str, label: str, default: str = "") -> str:
    if value:
        return value.strip()
    suffix = f" [{default}]" if default else ""
    answer = input(f"{label}{suffix}: ").strip()
    if answer.lower() in {"r", "reset"}:
        raise RestartFlow
    return answer or default


def print_input_rules() -> None:
    print(
        "\n输入规则：\n"
        "- 第一步选择导出类型：d=默认导出（昨天+库存联盟，图片+清洗表格）；"
        "f=好友/商家表格（contacts.csv + 商家导入.xlsx，不导出朋友圈）；"
        "m=朋友圈商品（moments.md/图片/PDF/Excel）。直接回车表示 d。\n"
        "- 朋友圈商品模式：先选时间范围与导出范围（s=只自己；f=只好友；a=全部，回车=s），再按需筛选好友。\n"
        "- 好友/商家模式：先选好友筛选方式，再决定是否导出好友列表与商家导入表格。\n"
        "- 时间格式示例：20260606 或 2026-06-06；直接回车表示不限制。\n"
        "- 图片下载前会询问是否清洗识别手机号：y（默认）=下载后检测二维码/手机号并隔离命中图片；n=不检测，仅保留下载时的二维码快检。\n"
        "- 好友筛选方式：a=全部好友；p=按备注关键词筛选（输入 服装 或 库存联盟 等，匹配备注中包含该词的好友）；m=手动逐个选择。直接回车默认 a。\n"
        "- 任意输入处输入 r 回车 = 放弃本次、从头重新开始（包括时间/筛选/导出类型）。\n"
        "- 导出中途出错不会退出程序，自动回到开头；结束时输入 q 可退出。\n"
        "- 商家导入表格含商家名称/手机号(备注末尾)/微信号/省市区；地区缺省时自动补“市辖区”，无地区默认浙江金华义乌。\n"
        "- 手动选择：第1位直接回车表示导出所有人；输入 /me 表示自己；输入 /enough 表示结束选择。\n"
        "- 批量粘贴：可一次粘贴多个好友的昵称或备注，用逗号（, 或，）、分号（; 或；）、顿号（、）、竖线（|）或换行分隔，自动识别并逐个匹配。\n",
        flush=True,
    )


def parse_datetime(value: str, *, end_of_day: bool) -> datetime | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    for fmt in ("%Y%m%d%H%M%S", "%Y%m%d%H%M", "%Y%m%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            parsed = datetime.strptime(raw, fmt)
            if fmt in {"%Y%m%d", "%Y-%m-%d"} and end_of_day:
                return datetime.combine(parsed.date(), time.max)
            return parsed
        except ValueError:
            continue
    raise ValueError(f"时间格式不正确: {raw}。示例: 20260606 或 2026-06-06")


def parse_yes_no(value: str) -> bool:
    raw = str(value or "").strip().lower()
    if not raw:
        return True
    if raw in {"y", "yes", "1", "true", "on"}:
        return True
    if raw in {"n", "no", "0", "false", "off"}:
        return False
    raise ValueError("只导出自己的朋友圈请输入 y 或 n")


def parse_yes_no_default_no(value: str) -> bool:
    raw = str(value or "").strip().lower()
    if not raw:
        return False
    if raw in {"y", "yes", "1", "true", "on"}:
        return True
    if raw in {"n", "no", "0", "false", "off"}:
        return False
    raise ValueError("请输入 y 或 n")


def parse_export_scope(value: str) -> str:
    """Parse export scope into one of: self, friends, all."""
    raw = str(value or "").strip().lower()
    if raw in {"s", "self", "self_only"}:
        return "self"
    if raw in {"f", "friends", "friends_only"}:
        return "friends"
    if raw in {"a", "all"}:
        return "all"
    raise ValueError("导出范围请输入 s（只自己）、f（只好友）或 a（全部）")


def ask_export_scope(cli_value: str = "") -> str:
    """Determine export scope from CLI value or interactive prompt.

    Returns one of: "self", "friends", "all". Defaults to "self" on empty
    input.
    """
    scope = str(cli_value or "").strip().lower()
    valid = {
        "s": "self", "self": "self",
        "f": "friends", "friends": "friends",
        "a": "all", "all": "all",
    }
    if scope in valid:
        return valid[scope]
    while True:
        raw = input(
            "请选择导出范围（s=只导出自己，f=只导出好友，"
            "a=全部(自己+好友)；回车=s）: "
        ).strip().lower()
        if not raw:
            return "self"
        if raw in valid:
            return valid[raw]
        print("输入无效，请输入 s、f 或 a。", flush=True)


def resolve_scope_from_args(args: argparse.Namespace, *, use_cli_args: bool) -> str:
    """Resolve export scope from CLI args.

    --export-scope takes precedence; otherwise --only-self is mapped
    (y -> self, n -> all). Returns "" when neither CLI arg is set, so the
    caller can prompt interactively.
    """
    if not use_cli_args:
        return ""
    if args.export_scope:
        return parse_export_scope(args.export_scope)
    if args.only_self:
        return "self" if parse_yes_no(args.only_self) else "all"
    return ""


def clean_account_name(path_name: str) -> str:
    name = str(path_name or "").strip()
    match = re.match(r"^(wxid_[^_]+)(?:_[0-9a-f]{4})?$", name, flags=re.IGNORECASE)
    return match.group(1) if match else name


def default_wechat_roots() -> list[Path]:
    roots: list[Path] = []
    configured = os.environ.get("WXMOMENTS_WECHAT_DATA_ROOT", "")
    if configured:
        for item in re.split(r"[;\n]", configured):
            if item.strip():
                roots.append(Path(item.strip()).expanduser())

    try:
        from wechat_decrypt_tool.modules.wechat_detection import auto_detect_wechat_data_dirs

        roots.extend(Path(item).expanduser() for item in auto_detect_wechat_data_dirs())
    except Exception as exc:
        print(f"[warning] {exc}", file=sys.stderr)

    for raw in (
        Path.home() / "Documents" / "xwechat_files",
        Path.home() / "Documents" / "WeChat Files",
        Path.home() / "Documents" / "Weixin Files",
    ):
        text = str(raw or "").strip()
        if text:
            roots.append(Path(text).expanduser())
    uniq: list[Path] = []
    seen: set[str] = set()
    for root in roots:
        key = str(root).lower()
        if key not in seen:
            seen.add(key)
            uniq.append(root)
    return uniq


_ACCOUNT_SCAN_SKIP_NAMES = {
    "$recycle.bin",
    "$winreagent",
    ".git",
    ".venv",
    "appdata",
    "config.msi",
    "documents and settings",
    "intel",
    "node_modules",
    "onedrivetemp",
    "perflogs",
    "program files",
    "program files (x86)",
    "programdata",
    "recovery",
    "runtime",
    "system volume information",
    "windows",
    "windows.old",
    "windows.old(1)",
}


def _iter_sns_db_paths(root: Path, *, max_depth: int = 6) -> list[Path]:
    root = root.expanduser()
    found: list[Path] = []
    if not root.exists():
        return found

    stack: list[tuple[Path, int]] = [(root, 0)]
    seen: set[str] = set()
    while stack:
        current, depth = stack.pop()
        try:
            key = os.path.normcase(str(current.resolve()))
        except OSError:
            key = os.path.normcase(str(current))
        if key in seen:
            continue
        seen.add(key)

        sns_direct = current / "db_storage" / "sns" / "sns.db"
        if sns_direct.exists():
            found.append(sns_direct)
        sns_flat = current / "db_storage" / "sns.db"
        if sns_flat.exists():
            found.append(sns_flat)

        if current.name.lower() == "db_storage":
            sns_child = current / "sns" / "sns.db"
            if sns_child.exists():
                found.append(sns_child)
            sns_child = current / "sns.db"
            if sns_child.exists():
                found.append(sns_child)

        if depth >= max_depth:
            continue

        try:
            with os.scandir(current) as entries:
                children = [
                    Path(entry.path)
                    for entry in entries
                    if entry.is_dir()
                    and entry.name.strip().lower() not in _ACCOUNT_SCAN_SKIP_NAMES
                ]
        except (PermissionError, OSError):
            continue
        stack.extend((child, depth + 1) for child in reversed(children))

    uniq: list[Path] = []
    seen_paths: set[str] = set()
    for item in found:
        try:
            key = os.path.normcase(str(item.resolve()))
        except OSError:
            key = os.path.normcase(str(item))
        if key not in seen_paths:
            seen_paths.add(key)
            uniq.append(item)
    return uniq


def _account_from_sns_db_path(sns_path: Path, account_hint: str = "") -> AccountInfo | None:
    if sns_path.name.lower() != "sns.db":
        return None
    db_dir = sns_path.parent.parent if sns_path.parent.name.lower() == "sns" else sns_path.parent
    if db_dir.name.lower() != "db_storage":
        return None
    wxid_dir = db_dir.parent
    account = clean_account_name(wxid_dir.name)
    if account_hint and account_hint.lower() not in {account.lower(), wxid_dir.name.lower()}:
        return None
    return AccountInfo(account=account, wxid_dir=wxid_dir, db_storage_dir=db_dir)


def iter_account_candidates(root: Path, account_hint: str = "") -> list[AccountInfo]:
    candidates: list[AccountInfo] = []
    if not root.exists():
        return candidates
    roots = [root]
    try:
        roots.extend([p for p in root.iterdir() if p.is_dir() and p.name.lower().startswith("wxid_")])
    except Exception as exc:
        print(f"[warning] {exc}", file=sys.stderr)
    for wxid_dir in roots:
        db_dir = wxid_dir / "db_storage"
        sns = db_dir / "sns" / "sns.db"
        if not sns.exists():
            sns = db_dir / "sns.db"
        if not sns.exists():
            continue
        account = clean_account_name(wxid_dir.name)
        if account_hint and account_hint.lower() not in {account.lower(), wxid_dir.name.lower()}:
            continue
        candidates.append(AccountInfo(account=account, wxid_dir=wxid_dir, db_storage_dir=db_dir))

    for sns_path in _iter_sns_db_paths(root):
        item = _account_from_sns_db_path(sns_path, account_hint)
        if item is not None:
            candidates.append(item)

    uniq: list[AccountInfo] = []
    seen: set[str] = set()
    for item in candidates:
        try:
            key = os.path.normcase(str(item.db_storage_dir.resolve()))
        except OSError:
            key = os.path.normcase(str(item.db_storage_dir))
        if key not in seen:
            seen.add(key)
            uniq.append(item)
    return uniq


def find_account(config: dict[str, Any]) -> AccountInfo:
    account_hint = str(config.get("account") or "").strip()
    configured_root = str(config.get("wechat_data_root") or "").strip()
    roots = [
        Path(item.strip()).expanduser()
        for item in re.split(r"[;\n]", configured_root)
        if item.strip()
    ] if configured_root else default_wechat_roots()
    candidates: list[tuple[int, AccountInfo]] = []
    for root in roots:
        for item in iter_account_candidates(root, account_hint):
            sns_path = item.db_storage_dir / "sns" / "sns.db"
            if not sns_path.exists():
                sns_path = item.db_storage_dir / "sns.db"
            try:
                size = int(sns_path.stat().st_size)
            except OSError:
                size = 0
            candidates.append((size, item))
    if not candidates:
        searched = "、".join(str(p) for p in roots)
        raise FileNotFoundError(
            "没有找到微信朋友圈数据库。\n"
            f"已查找: {searched}\n"
            "请确认这台电脑登录过微信并已产生朋友圈缓存；如果微信文件保存位置不是默认值，"
            "请在 config/config.json 的 wechat_data_root 中填写微信数据目录，例如 "
            r"C:\Users\你的用户名\Documents\WeChat Files 或 D:\微信文件。"
        )
    return max(candidates, key=lambda item: item[0])[1]


def _normalize_pasted_path(value: str) -> Path:
    raw = str(value or "").strip().strip('"').strip("'")
    if raw.startswith("& "):
        raw = raw[2:].strip().strip('"').strip("'")
    raw = os.path.expandvars(raw)
    path = Path(raw).expanduser()
    if path.is_file() and path.name.lower() == "sns.db":
        path = path.parent
    return path


def find_account_with_interactive_retry(
    config: dict[str, Any],
    config_path: Path,
    original_error: FileNotFoundError,
) -> AccountInfo:
    print(f"\n{original_error}", flush=True)
    print(
        "\n自动定位失败，但可以现场补救：\n"
        "1. 打开电脑版微信 → 设置 → 文件管理 → 打开文件夹。\n"
        "2. 把打开的文件夹路径复制/拖入到这里。\n"
        "3. 也可以直接粘贴包含 wxid_xxxxx 或 db_storage 的目录。\n"
        "直接回车则退出。\n",
        flush=True,
    )

    while True:
        raw = input("请输入微信数据目录路径: ").strip()
        if not raw:
            raise original_error
        candidate = _normalize_pasted_path(raw)
        if not candidate.exists():
            print(f"这个路径不存在：{candidate}", flush=True)
            continue

        trial_config = dict(config)
        trial_config["wechat_data_root"] = str(candidate)
        try:
            account_info = find_account(trial_config)
        except FileNotFoundError:
            print(
                "这个路径里仍然没有找到朋友圈数据库 sns.db。请尝试上一级目录、"
                "WeChat Files / xwechat_files 目录，或直接搜索 sns.db 后粘贴其所在目录。",
                flush=True,
            )
            continue

        config["wechat_data_root"] = str(candidate)
        save_config(config_path, config)
        print(
            f"已找到账号目录：{account_info.wxid_dir}\n"
            f"已保存 wechat_data_root 到：{config_path}",
            flush=True,
        )
        return account_info


def is_wechat_running() -> bool:
    try:
        import psutil
    except Exception:
        return False
    targets = {"weixin.exe", "wechat.exe"}
    for proc in psutil.process_iter(["name"]):
        try:
            if str(proc.info.get("name") or "").strip().lower() in targets:
                return True
        except Exception:
            continue
    return False


def load_saved_db_key(account_info: AccountInfo) -> str:
    key_store = OUTPUT_RUNTIME_DIR / "account_keys.json"
    if not key_store.exists():
        return ""
    try:
        data = json.loads(key_store.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"[warning] {exc}", file=sys.stderr)
        return ""
    if not isinstance(data, dict):
        return ""
    candidates = [account_info.account, account_info.wxid_dir.name]
    for candidate in candidates:
        item = data.get(candidate)
        if not isinstance(item, dict):
            continue
        key = str(item.get("db_key") or "").strip()
        if re.fullmatch(rf"[0-9a-fA-F]{{{DB_KEY_HEX_LENGTH}}}", key):
            return key
    return ""


def acquire_db_key(account_info: AccountInfo, config: dict[str, Any], args: argparse.Namespace) -> str:
    key = (args.key or str(config.get("db_key") or "") or os.environ.get("WXMOMENTS_KEY", "")).strip()
    if key:
        return key

    saved = load_saved_db_key(account_info)
    if saved:
        print("已使用本地保存的数据库密钥。")
        return saved

    if not is_wechat_running():
        input("请先登录电脑版微信，确认已进入微信主界面后按回车继续...")
    else:
        input("请确认电脑版微信已登录并停留在主界面，然后按回车自动获取密钥...")

    from wechat_decrypt_tool.modules.key_service import get_db_key_workflow

    try:
        result = get_db_key_workflow(db_storage_path=str(account_info.db_storage_dir))
        return str(result.get("db_key") or "").strip()
    except Exception as exc:
        print(f"自动获取密钥失败: {exc}")
        return input("请手动输入数据库密钥（64位十六进制）: ").strip()


async def save_image_keys(account: str, wxid_dir: Path, db_storage_dir: Path) -> None:
    from wechat_decrypt_tool.modules.key_service import get_image_key_integrated_workflow

    account_dir = OUTPUT_RUNTIME_DIR / "databases" / account
    account_dir.mkdir(parents=True, exist_ok=True)

    # Reuse cached image keys if present and structurally valid. Image keys
    # are per-WeChat-session and may change after a WeChat restart; if cached
    # keys fail to decrypt images the export just logs missing-image warnings,
    # so this trade-off keeps repeated exports fast.
    keys_path = account_dir / "_media_keys.json"
    if keys_path.exists():
        try:
            cached = json.loads(keys_path.read_text(encoding="utf-8"))
            if isinstance(cached, dict) and "xor" in cached and "aes" in cached:
                xor_key = int(cached["xor"])
                aes_key = str(cached["aes"] or "")
                if 0 <= xor_key <= IMAGE_XOR_KEY_MAX and len(aes_key) == IMAGE_AES_KEY_LENGTH:
                    print(f"  复用已缓存的图片密钥: {keys_path}")
                    return
        except Exception:
            pass

    data = await get_image_key_integrated_workflow(
        account,
        db_storage_path=str(db_storage_dir),
        wxid_dir=str(wxid_dir),
    )
    if data.get("verified") is not True:
        raise RuntimeError("无法取得已验证的图片密钥，请保持微信已登录后重试")
    xor_raw = data.get("xor_key")
    try:
        xor_key = int(str(xor_raw), 16)
    except (TypeError, ValueError):
        xor_key = int(xor_raw)
    aes_key = str(data.get("aes_key") or "")[:16]
    if not (0 <= xor_key <= IMAGE_XOR_KEY_MAX) or len(aes_key) != IMAGE_AES_KEY_LENGTH:
        raise RuntimeError("图片密钥无效")
    keys_path.write_text(
        json.dumps({"xor": xor_key, "aes": aes_key}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _source_dbs_newer_than_cached(source_dir: Path, cached_dir: Path) -> bool:
    """Return True if any source WeChat .db is newer than its cached copy.

    WeChat rewrites its local databases when remarks change or new moments are
    cached. Comparing modification times lets us skip re-decryption only when
    the decrypted cache is still fresh.
    """
    try:
        newest_source = 0.0
        for db_file in source_dir.rglob("*.db"):
            try:
                mtime = db_file.stat().st_mtime
                if mtime > newest_source:
                    newest_source = mtime
            except OSError:
                continue
        if newest_source <= 0:
            return False
        for cached_file in cached_dir.rglob("*.db"):
            try:
                if cached_file.stat().st_mtime < newest_source:
                    return True
            except OSError:
                continue
    except Exception:
        return True  # be safe and re-decrypt on any error
    return False


def decrypt_databases(account_info: AccountInfo, key: str) -> Path:
    from wechat_decrypt_tool.modules.wechat_decrypt import decrypt_wechat_databases

    # Reuse already-decrypted databases only if the cached copies are not older
    # than the source WeChat databases. WeChat updates its local *.db files
    # when remarks are edited or new moments are cached, so a stale cache would
    # otherwise hide those changes.
    existing_dir = OUTPUT_RUNTIME_DIR / "databases" / account_info.account
    if (existing_dir / "sns.db").exists() and not _source_dbs_newer_than_cached(
        account_info.db_storage_dir, existing_dir
    ):
        try:
            conn = sqlite3.connect(str(existing_dir / "sns.db"))
            conn.execute("SELECT 1 FROM SnsTimeLine LIMIT 1").fetchone()
            conn.close()
            print(f"  复用已解密的数据库: {existing_dir}")
            return existing_dir
        except Exception:
            # Cached db is corrupt or key changed; fall through to re-decrypt.
            pass

    result = decrypt_wechat_databases(db_storage_path=str(account_info.db_storage_dir), key=key)
    if str(result.get("status") or "").lower() not in {"success", "ok"}:
        raise RuntimeError(str(result.get("message") or "数据库解密失败"))
    account_results = result.get("account_results") if isinstance(result.get("account_results"), dict) else {}
    for account_name, detail in account_results.items():
        if account_info.account.lower() == str(account_name or "").lower() and isinstance(detail, dict):
            out = Path(str(detail.get("output_dir") or ""))
            if (out / "sns.db").exists():
                return out
    fallback = OUTPUT_RUNTIME_DIR / "databases" / account_info.account
    if (fallback / "sns.db").exists():
        return fallback
    for candidate in (OUTPUT_RUNTIME_DIR / "databases").glob("*/sns.db"):
        return candidate.parent
    raise FileNotFoundError("解密完成后没有找到 sns.db")


def image_size(payload: bytes, media_type: str) -> tuple[int, int]:
    from wechat_decrypt_tool.modules.sns_reader import _image_size_from_bytes

    try:
        return _image_size_from_bytes(payload, media_type)
    except Exception:
        return (0, 0)


def expected_image_size(media: dict[str, Any]) -> tuple[int, int]:
    size = media.get("size") if isinstance(media.get("size"), dict) else {}
    for w_raw, h_raw in (
        (size.get("width"), size.get("height")),
        (media.get("width"), media.get("height")),
        (media.get("w"), media.get("h")),
    ):
        try:
            w = int(float(w_raw or 0))
            h = int(float(h_raw or 0))
        except Exception:
            continue
        if w > 0 and h > 0:
            return w, h
    return (0, 0)


def meets_expected_size(payload: bytes, media_type: str, expected: tuple[int, int]) -> bool:
    exp_w, exp_h = expected
    if exp_w <= 0 or exp_h <= 0:
        return True
    width, height = image_size(payload, media_type)
    if width <= 0 or height <= 0:
        return True
    return width * 4 >= exp_w * 3 and height * 4 >= exp_h * 3


def timestamp_folder_name(created: datetime) -> str:
    return created.strftime("%Y%m%d_%H%M%S")


def unique_post_dir(figure: Path, created: datetime) -> Path:
    base = timestamp_folder_name(created)
    candidate = figure / base
    if not candidate.exists():
        return candidate
    index = 2
    while True:
        candidate = figure / f"{base}_{index:02d}"
        if not candidate.exists():
            return candidate
        index += 1


def normalize_hex32(value: Any) -> str:
    from wechat_decrypt_tool.modules.sns_reader import normalize_hex32 as _impl
    return _impl(value)


def generate_sns_cache_key(tid: str, media_id: str, media_type: int = MEDIA_TYPE_IMAGE) -> str:
    from wechat_decrypt_tool.modules.sns_reader import generate_sns_cache_key as _impl
    return _impl(tid, media_id, media_type)


def resolve_sns_cached_image_path_by_cache_key(wxid_dir: Path, cache_key: str, create_time: int = 0) -> Path | None:
    from wechat_decrypt_tool.modules.sns_reader import resolve_sns_cached_image_path_by_cache_key as _impl
    return _impl(wxid_dir, cache_key, create_time)


def pick_sns_media_str(*values: Any) -> str:
    for value in values:
        text = str(value or "").strip()
        if text:
            return text
    return ""


def sns_media_md5_value(media: dict[str, Any], raw_url: str) -> str:
    url_attrs = media.get("urlAttrs") if isinstance(media.get("urlAttrs"), dict) else {}
    thumb_attrs = media.get("thumbAttrs") if isinstance(media.get("thumbAttrs"), dict) else {}
    md5_raw = pick_sns_media_str(
        url_attrs.get("md5"),
        thumb_attrs.get("md5"),
        url_attrs.get("MD5"),
        thumb_attrs.get("MD5"),
    )
    if not md5_raw:
        match = re.search(r"[?&]md5=([0-9a-fA-F]{16,32})", str(raw_url or ""))
        if match:
            md5_raw = match.group(1)
    return normalize_hex32(md5_raw)


def sns_image_source(media: dict[str, Any], *, prefer_thumb: bool) -> tuple[str, str, str, bool]:
    url_attrs = media.get("urlAttrs") if isinstance(media.get("urlAttrs"), dict) else {}
    thumb_attrs = media.get("thumbAttrs") if isinstance(media.get("thumbAttrs"), dict) else {}
    original_url = pick_sns_media_str(
        media.get("url"),
        media.get("originUrl"),
        media.get("originalUrl"),
        media.get("origin_url"),
        media.get("original_url"),
    )
    thumb_url = pick_sns_media_str(media.get("thumb"), media.get("thumbUrl"), media.get("thumb_url"))
    use_thumb = bool(prefer_thumb and thumb_url) or not original_url
    if use_thumb:
        return (
            thumb_url or original_url,
            pick_sns_media_str(media.get("thumbKey"), media.get("thumb_key"), thumb_attrs.get("key"), media.get("key"), url_attrs.get("key")),
            pick_sns_media_str(
                media.get("thumbToken"),
                media.get("thumbUrlToken"),
                media.get("thumb_url_token"),
                thumb_attrs.get("token"),
                media.get("token"),
                url_attrs.get("token"),
            ),
            False,
        )
    return (
        original_url,
        pick_sns_media_str(media.get("key"), url_attrs.get("key"), media.get("thumbKey"), thumb_attrs.get("key")),
        pick_sns_media_str(
            media.get("token"),
            media.get("urlToken"),
            media.get("url_token"),
            url_attrs.get("token"),
            media.get("thumbToken"),
            thumb_attrs.get("token"),
        ),
        True,
    )


def resolve_sns_exact_cached_image_path(wxid_dir: Path, post: dict[str, Any], media: dict[str, Any], raw_url: str) -> Path | None:
    post_id = str(post.get("id") or post.get("tid") or "").strip()
    media_id = str(media.get("id") or "").strip()
    try:
        post_type = int(post.get("type") or POST_TYPE_NORMAL)
    except Exception:
        post_type = POST_TYPE_NORMAL
    try:
        media_type = int(media.get("type") or MEDIA_TYPE_IMAGE)
    except Exception:
        media_type = MEDIA_TYPE_IMAGE
    try:
        create_time = int(post.get("createTime") or 0)
    except Exception:
        create_time = 0

    if post_id and media_id and post_type == POST_TYPE_COVER:
        bkg_md5 = hashlib.md5(f"{post_id}_{media_id}_4".encode("utf-8", errors="ignore")).hexdigest()
        bkg_path = wxid_dir / "business" / "sns" / "bkg" / bkg_md5[:2] / bkg_md5
        if bkg_path.is_file():
            return bkg_path

    if post_id and media_id:
        for candidate_type in (media_type, MEDIA_TYPE_IMAGE, MEDIA_TYPE_LIVE_PHOTO, MEDIA_TYPE_VIDEO, 0):
            cache_key = generate_sns_cache_key(post_id, media_id, candidate_type)
            found = resolve_sns_cached_image_path_by_cache_key(wxid_dir, cache_key, create_time)
            if found is not None:
                return found

    md5_32 = sns_media_md5_value(media, raw_url)
    if md5_32:
        return resolve_sns_cached_image_path_by_cache_key(wxid_dir, md5_32, create_time)
    return None


def load_timeline(account_dir: Path, usernames: list[str] | None = None, *, source: str = "auto") -> list[dict[str, Any]]:
    from wechat_decrypt_tool.modules.sns_reader import list_sns_timeline

    posts: list[dict[str, Any]] = []
    offset = 0
    limit = DEFAULT_PAGE_LIMIT
    usernames_text = ",".join(usernames or []) or None
    while True:
        response = list_sns_timeline(
            account=account_dir.name,
            limit=limit,
            offset=offset,
            usernames=usernames_text,
            keyword=None,
            source=source,
        )
        page = response.get("items") or response.get("timeline") or []
        posts.extend([item for item in page if isinstance(item, dict)])
        if not response.get("hasMore") or not page:
            break
        offset += len(page)
    return posts


def post_created_datetime(post: dict[str, Any]) -> datetime | None:
    try:
        created_ts = int(post.get("createTime") or 0)
    except Exception:
        return None
    if created_ts <= 0:
        return None
    return datetime.fromtimestamp(created_ts)


def build_coverage_report(
    posts: list[dict[str, Any]],
    filtered: list[dict[str, Any]],
    start: datetime | None,
    end: datetime | None,
) -> dict[str, Any]:
    dated = sorted(dt for post in posts if (dt := post_created_datetime(post)) is not None)
    filtered_dated = sorted(dt for post in filtered if (dt := post_created_datetime(post)) is not None)
    gaps: list[dict[str, Any]] = []
    for prev, curr in zip(dated, dated[1:]):
        gap_days = (curr.date() - prev.date()).days - 1
        if gap_days >= 21:
            gaps.append(
                {
                    "from": prev.strftime("%Y-%m-%d %H:%M:%S"),
                    "to": curr.strftime("%Y-%m-%d %H:%M:%S"),
                    "gap_days": gap_days,
                }
            )
    gaps.sort(key=lambda item: int(item.get("gap_days") or 0), reverse=True)
    report = {
        "local_posts_with_time": len(dated),
        "exported_posts_with_time": len(filtered_dated),
        "local_earliest": dated[0].strftime("%Y-%m-%d %H:%M:%S") if dated else "",
        "local_latest": dated[-1].strftime("%Y-%m-%d %H:%M:%S") if dated else "",
        "export_earliest": filtered_dated[0].strftime("%Y-%m-%d %H:%M:%S") if filtered_dated else "",
        "export_latest": filtered_dated[-1].strftime("%Y-%m-%d %H:%M:%S") if filtered_dated else "",
        "requested_start": start.strftime("%Y-%m-%d %H:%M:%S") if start else "",
        "requested_end": end.strftime("%Y-%m-%d %H:%M:%S") if end else "",
        "large_gaps": gaps[:20],
        "large_gap_threshold_days": 21,
    }
    warnings: list[str] = []
    if start and dated and start < dated[0]:
        warnings.append(
            f"请求起始时间早于本地缓存最早记录：{start:%Y-%m-%d %H:%M:%S} < {dated[0]:%Y-%m-%d %H:%M:%S}"
        )
    if end and dated and end > dated[-1]:
        warnings.append(
            f"请求结束时间晚于本地缓存最新记录：{end:%Y-%m-%d %H:%M:%S} > {dated[-1]:%Y-%m-%d %H:%M:%S}"
        )
    if gaps:
        largest = gaps[0]
        warnings.append(
            f"本地缓存存在明显断档：{largest['from']} 至 {largest['to']} 之间约 {largest['gap_days']} 天没有记录"
        )
    report["warnings"] = warnings
    return report


def write_coverage_report(output: Path, report: dict[str, Any]) -> None:
    (output / "coverage.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    earliest = str(report.get("local_earliest") or "")
    latest = str(report.get("local_latest") or "")
    if earliest or latest:
        print(f"  本地缓存范围: {earliest or '未知'} ~ {latest or '未知'}", flush=True)
    for warning in report.get("warnings") or []:
        print(f"  [coverage] {warning}", flush=True)


def save_db_key(account_info: AccountInfo, account_dir: Path, key: str) -> None:
    try:
        from wechat_decrypt_tool.modules.key_store import upsert_account_keys_in_store

        aliases = [account_info.wxid_dir.name, account_dir.name]
        upsert_account_keys_in_store(
            account_dir.name,
            db_key=key,
            aliases=aliases,
            db_key_source_wxid_dir=str(account_info.wxid_dir),
            db_key_source_db_storage_path=str(account_info.db_storage_dir),
        )
    except Exception as exc:
        print(f"[warning] {exc}", file=sys.stderr)


def self_username_candidates(account_info: AccountInfo, config: dict[str, Any]) -> list[str]:
    values: list[str] = []
    raw_one = str(config.get("self_username") or "").strip()
    if raw_one:
        values.append(raw_one)
    values.extend([account_info.account, account_info.wxid_dir.name])
    out: list[str] = []
    seen: set[str] = set()
    for value in values:
        if value and value not in seen:
            seen.add(value)
            out.append(value)
    return out


def normalize_contact_lookup_text(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip().casefold()


def decode_contact_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, memoryview):
        value = value.tobytes()
    if isinstance(value, (bytes, bytearray)):
        raw = bytes(value)
        for encoding in ("utf-8", "utf-16-le", "utf-16"):
            try:
                text = raw.decode(encoding, errors="ignore")
                text = text.replace("\x00", "").replace("\xa0", " ").strip()
                if text:
                    return text
            except Exception:
                continue
        return ""
    return str(value).replace("\xa0", " ").strip()


def sqlite_table_columns(conn: sqlite3.Connection, table: str) -> list[str]:
    try:
        rows = conn.execute(f"PRAGMA table_info({table})").fetchall()
    except Exception:
        return []
    out: list[str] = []
    for row in rows:
        try:
            name = str(row[1] or "").strip()
        except Exception:
            name = ""
        if name:
            out.append(name)
    return out


def row_value(row: sqlite3.Row, *names: str) -> Any:
    lower_to_name = {str(key).lower(): str(key) for key in row.keys()}
    for name in names:
        actual = lower_to_name.get(str(name).lower())
        if actual is not None:
            return row[actual]
    return None


def int_row_value(row: sqlite3.Row, *names: str) -> int:
    value = row_value(row, *names)
    try:
        return int(value or 0)
    except Exception:
        return 0


def infer_contact_privacy_status(flag: int, local_type: int, delete_flag: int) -> str:
    statuses: list[str] = []
    if flag & 0x100:
        statuses.append("不让ta看我的朋友圈")
    if flag & 0x10000:
        statuses.append("不看ta的朋友圈")
    if flag & 0x8:
        statuses.append("黑名单")
    if delete_flag:
        statuses.append("已删除或非当前好友")
    if local_type and not (local_type & 0x1):
        statuses.append("可能不是通讯录好友")
    return "；".join(dict.fromkeys(statuses))


def contact_is_current_friend(row: sqlite3.Row, table: str, self_usernames: set[str]) -> bool:
    """Return whether a contact row represents a current address-book friend.

    WeChat's contact.db is a broad profile cache, not a pure friend list. In
    recent PC WeChat builds, many non-friend cached profiles appear in
    contact.local_type=3 with flag=4, which previously made exports balloon
    into thousands of unrelated-looking rows. Current address-book contacts
    are the rows in contact.local_type=1 that are not deleted and not official
    accounts.
    """
    if str(table or "").lower() != "contact":
        return False
    username = decode_contact_text(row_value(row, "username"))
    if username in self_usernames:
        return False
    if not should_export_contact(username):
        return False
    if int_row_value(row, "delete_flag"):
        return False
    if int_row_value(row, "verify_flag"):
        return False
    return int_row_value(row, "local_type") == 1


# ---------------------------------------------------------------------------
# 好友资料地区：微信 4.x 把资料地区以 protobuf 存在 contact.extra_buffer，
# field5=国家代码(CN)、field6=省份拼音(Guangdong)、field7=城市拼音(Guangzhou)。
# ---------------------------------------------------------------------------

# 省份/直辖市/自治区拼音 -> 中文简称
_REGION_PROVINCE_PINYIN: dict[str, str] = {
    "Beijing": "北京", "Shanghai": "上海", "Tianjin": "天津", "Chongqing": "重庆",
    "Zhejiang": "浙江", "Jiangsu": "江苏", "Guangdong": "广东", "Anhui": "安徽",
    "Henan": "河南", "Jiangxi": "江西", "Shandong": "山东", "Fujian": "福建",
    "Hebei": "河北", "Hunan": "湖南", "Sichuan": "四川", "Hubei": "湖北",
    "Shaanxi": "陕西", "Shanxi": "山西", "Heilongjiang": "黑龙江", "Jilin": "吉林",
    "Liaoning": "辽宁", "Guizhou": "贵州", "Yunnan": "云南", "Gansu": "甘肃",
    "Qinghai": "青海", "Hainan": "海南", "Guangxi": "广西", "Ningxia": "宁夏",
    "Xinjiang": "新疆", "Inner Mongolia": "内蒙古", "Tibet": "西藏",
    "Hong Kong": "香港", "Macao": "澳门", "Taiwan": "台湾",
}
_REGION_DIRECT_CITIES = {"北京", "上海", "天津", "重庆"}
_REGION_AUTONOMOUS_SUFFIX = {
    "广西": "广西壮族自治区", "宁夏": "宁夏回族自治区", "新疆": "新疆维吾尔自治区",
    "内蒙古": "内蒙古自治区", "西藏": "西藏自治区",
}

# 地级市拼音 -> 中文简称
_REGION_CITY_PINYIN: dict[str, str] = {
    "Anqing": "安庆", "Anshan": "鞍山", "Anshun": "安顺", "Anyang": "安阳",
    "Baoding": "保定", "Baoshan": "保山", "Baotou": "包头", "Binzhou": "滨州",
    "Bozhou": "亳州", "Cangzhou": "沧州", "Changchun": "长春", "Changde": "常德",
    "Changning": "常宁", "Changsha": "长沙", "Changzhou": "常州", "Chaozhou": "潮州",
    "Chengdu": "成都", "Chenzhou": "郴州", "Chifeng": "赤峰", "Chizhou": "池州",
    "Chuzhou": "滁州", "Dalian": "大连", "Daqing": "大庆", "Datong": "大同",
    "Dongguan": "东莞", "Ezhou": "鄂州", "Foshan": "佛山", "Fuyang": "阜阳",
    "Fuzhou": "福州", "Ganzhou": "赣州", "Guangzhou": "广州", "Guigang": "贵港",
    "Guilin": "桂林", "Guiyang": "贵阳", "Haikou": "海口", "Handan": "邯郸",
    "Hangzhou": "杭州", "Hanzhong": "汉中", "Harbin": "哈尔滨", "Hefei": "合肥",
    "Hengshui": "衡水", "Hengyang": "衡阳", "Heze": "菏泽", "Hezhou": "贺州",
    "Hotan": "和田", "Huaibei": "淮北", "Huaihua": "怀化", "Huainan": "淮南",
    "Huangshan": "黄山", "Huizhou": "惠州", "Hulunbeier": "呼伦贝尔", "Huzhou": "湖州",
    "Huai": "淮安", "Jian": "吉安", "Luan": "六安",
    "Jiaxing": "嘉兴", "Jieyang": "揭阳", "Jilin": "吉林", "Jinan": "济南",
    "Jingdezhen": "景德镇", "Jinhua": "金华", "Jining": "济宁", "Jiujiang": "九江",
    "Jixi": "鸡西", "Kaifeng": "开封", "Kunming": "昆明", "Lanzhou": "兰州",
    "Lhasa": "拉萨", "Lianyungang": "连云港", "Liaocheng": "聊城", "Linfen": "临汾",
    "Linyi": "临沂", "Lishui": "丽水", "Liuzhou": "柳州", "Longyan": "龙岩",
    "Luoyang": "洛阳", "Luzhou": "泸州", "Maoming": "茂名", "Mianyang": "绵阳",
    "Mudanjiang": "牡丹江", "Nanchang": "南昌", "Nanchong": "南充", "Nanjing": "南京",
    "Nanning": "南宁", "Nanping": "南平", "Nantong": "南通", "Nanyang": "南阳",
    "Ningbo": "宁波", "Pingliang": "平凉", "Putian": "莆田", "Puyang": "濮阳",
    "Qingdao": "青岛", "Quanzhou": "泉州", "Quzhou": "衢州", "Rizhao": "日照",
    "Sanming": "三明", "Shangqiu": "商丘", "Shangrao": "上饶", "Shantou": "汕头",
    "Shaoxing": "绍兴", "Shaoyang": "邵阳", "Shenyang": "沈阳", "Shenzhen": "深圳",
    "Shijiazhuang": "石家庄", "Suining": "遂宁", "Suizhou": "随州", "Suqian": "宿迁",
    "Suzhou": "苏州", "Taiyuan": "太原", "Tangshan": "唐山", "Tongling": "铜陵",
    "Urumqi": "乌鲁木齐", "Weifang": "潍坊", "Weihai": "威海", "Weinan": "渭南",
    "Wenzhou": "温州", "Wuhan": "武汉", "Wuhu": "芜湖", "Wuxi": "无锡",
    "Xiamen": "厦门", "Xiaogan": "孝感", "Xingtai": "邢台", "Xinxiang": "新乡",
    "Xinyang": "信阳", "Xinyu": "新余", "Xuancheng": "宣城", "Xuchang": "许昌",
    "Xuzhou": "徐州", "Yaan": "雅安", "Yancheng": "盐城", "Yibin": "宜宾",
    "Yingtan": "鹰潭", "Yongzhou": "永州", "Zaozhuang": "枣庄", "Zhangzhou": "漳州",
    "Zhanjiang": "湛江", "Zhengzhou": "郑州", "Zhongshan": "中山", "Zhoukou": "周口",
    "Zhuzhou": "株洲", "Ziyang": "资阳",
}

# 自治州/盟拼音 -> 全称
_REGION_AUTOPREF_PINYIN: dict[str, str] = {
    "Enshi": "恩施土家族苗族自治州", "Liangshan": "凉山彝族自治州",
    "Xiangxi": "湘西土家族苗族自治州", "Ili": "伊犁哈萨克自治州",
    "South Guizhou": "黔南布依族苗族自治州",
    "Southwest Guizhou": "黔西南布依族苗族自治州",
}
_REGION_AUTOPREF_CN = {
    "恩施": "恩施土家族苗族自治州", "凉山": "凉山彝族自治州",
    "湘西": "湘西土家族苗族自治州", "伊犁": "伊犁哈萨克自治州",
    "黔南": "黔南布依族苗族自治州", "黔西南": "黔西南布依族苗族自治州",
}

# 直辖市的市辖区（按省份拼音消歧，值为中文区名简称）
_REGION_DISTRICT_PINYIN: dict[str, dict[str, str]] = {
    "Beijing": {
        "Changping": "昌平", "Daxing": "大兴", "Fengtai": "丰台", "Haidian": "海淀",
        "Mentougou": "门头沟", "Chaoyang": "朝阳", "Tongzhou": "通州",
    },
    "Shanghai": {
        "Fengxian": "奉贤", "Hongkou": "虹口", "Huangpu": "黄浦", "Jinshan": "金山",
        "Pudong New District": "浦东新区", "Songjiang": "松江", "Xuhui": "徐汇",
        "Yangpu": "杨浦", "Baoshan": "宝山", "Changning": "长宁",
    },
    "Tianjin": {"Baodi": "宝坻", "Jinghai": "静海", "Nankai": "南开", "North Star": "北辰"},
    "Chongqing": {"Beibei": "北碚", "Dadukou": "大渡口", "Kaixian": "开州", "South Bank": "南岸"},
}


def _pb_read_varint(buf: bytes, pos: int) -> tuple[int, int]:
    result = 0
    shift = 0
    while pos < len(buf):
        byte = buf[pos]
        pos += 1
        result |= (byte & 0x7F) << shift
        if not byte & 0x80:
            return result, pos
        shift += 7
    return result, pos


def parse_extra_buffer_region(blob: Any) -> tuple[str, str, str]:
    """解析 contact.extra_buffer protobuf，返回 (国家, 省份拼音, 城市拼音)。"""
    if blob is None:
        return "", "", ""
    if isinstance(blob, memoryview):
        blob = blob.tobytes()
    if isinstance(blob, str):
        blob = blob.encode("utf-8", "ignore")
    if not isinstance(blob, (bytes, bytearray)):
        return "", "", ""
    raw = bytes(blob)
    found: dict[int, str] = {}
    pos, size = 0, len(raw)
    while pos < size:
        key, pos = _pb_read_varint(raw, pos)
        field, wire_type = key >> 3, key & 7
        if wire_type == 2:
            length, pos = _pb_read_varint(raw, pos)
            value = raw[pos:pos + length]
            pos += length
            if field in (5, 6, 7):
                text = value.decode("utf-8", "ignore").strip()
                if text and all(ch.isalpha() or ch == " " for ch in text):
                    found[field] = text
        elif wire_type == 0:
            _, pos = _pb_read_varint(raw, pos)
        elif wire_type == 1:
            pos += 8
        elif wire_type == 5:
            pos += 4
        else:
            break
    return found.get(5, ""), found.get(6, ""), found.get(7, "")


def format_contact_region(blob: Any) -> str:
    """把 extra_buffer 中的地区转为中文短格式，如 '广东 广州'、'北京 海淀'。

    无法识别（含境外、未知拼音）时返回空串。
    """
    country, prov_py, city_py = parse_extra_buffer_region(blob)
    if country and country != "CN":
        return ""
    province = _REGION_PROVINCE_PINYIN.get(prov_py)
    if not province:
        return ""
    # 直辖市：第二级是区（也可能直接是直辖市本身）
    if prov_py in _REGION_DISTRICT_PINYIN and city_py and city_py.lower() != prov_py.lower():
        district = _REGION_DISTRICT_PINYIN[prov_py].get(city_py)
        return f"{province} {district}" if district else province
    if city_py in _REGION_AUTOPREF_PINYIN:
        return f"{province} {city_py and _autopref_short(city_py)}"
    if city_py == "Taizhou":
        # 同音城市按省份消歧：浙江=台州，江苏=泰州
        city_cn = "台州" if province == "浙江" else "泰州"
        return f"{province} {city_cn}"
    if city_py == "Suzhou":
        # 同音城市按省份消歧：江苏=苏州，安徽=宿州
        city_cn = "宿州" if province == "安徽" else "苏州"
        return f"{province} {city_cn}"
    city = _REGION_CITY_PINYIN.get(city_py)
    if city:
        return f"{province} {city}"
    return province


def _autopref_short(city_py: str) -> str:
    full = _REGION_AUTOPREF_PINYIN.get(city_py, "")
    return full[:2] if full else ""


def contact_other_info(row: sqlite3.Row) -> str:
    fields = [
        ("微信号", ("alias",)),
        ("额外备注", ("description", "desc", "remark_desc", "remark_description")),
        ("手机号", ("phone", "phone_number", "mobile", "mobile_phone", "telephone")),
        ("标签", ("label_id_list", "label_ids", "labels", "tag", "tag_ids")),
        ("地区", ("region", "country", "province", "city")),
        ("个性签名", ("signature", "sign")),
    ]
    parts: list[str] = []
    for label, names in fields:
        values: list[str] = []
        if label == "地区":
            values = [decode_contact_text(row_value(row, name)) for name in names]
            values = [item for item in values if item]
            value = " ".join(dict.fromkeys(values))
            # 新版微信无独立地区列，回退解析 extra_buffer protobuf
            if not value:
                value = format_contact_region(row_value(row, "extra_buffer"))
        else:
            value = ""
            for name in names:
                value = decode_contact_text(row_value(row, name))
                if value:
                    break
        if value:
            parts.append(f"{label}: {value}")
    return "；".join(parts)


def should_export_contact(username: str) -> bool:
    lowered = str(username or "").strip().lower()
    if not lowered:
        return False
    if (
        lowered.endswith("@chatroom")
        or lowered.endswith("@openim")
        or lowered.endswith("@kefu.openim")
        or lowered.startswith("gh_")
    ):
        return False
    system_accounts = {
        "filehelper",
        "newsapp",
        "fmessage",
        "weibo",
        "qqmail",
        "tmessage",
        "qmessage",
        "medianote",
        "floatbottle",
        "lbsapp",
        "shakeapp",
        "feedsapp",
        "qqsync",
        "weixin",
        "weixinreminder",
        "officialaccounts",
        "brandsessionholder",
        "masssendapp",
    }
    return lowered not in system_accounts


def load_contact_entries(account_dir: Path) -> list[ContactEntry]:
    contact_db = account_dir / "contact.db"
    if not contact_db.exists():
        return []

    entries_by_username: dict[str, ContactEntry] = {}
    self_usernames = {clean_account_name(account_dir.name), account_dir.name}
    conn = sqlite3.connect(str(contact_db))
    conn.row_factory = sqlite3.Row
    try:
        for table in ("contact", "stranger"):
            columns = sqlite_table_columns(conn, table)
            if not columns:
                continue
            try:
                rows = conn.execute(f"SELECT * FROM {table}").fetchall()
            except Exception:
                rows = []
            for row in rows:
                username = decode_contact_text(row_value(row, "username"))
                if not contact_is_current_friend(row, table, self_usernames):
                    continue
                remark = decode_contact_text(row_value(row, "remark"))
                nickname = decode_contact_text(row_value(row, "nick_name", "nickname"))
                if not remark and not nickname:
                    continue
                flag = int_row_value(row, "flag")
                local_type = int_row_value(row, "local_type")
                delete_flag = int_row_value(row, "delete_flag")
                entries_by_username.setdefault(
                    username,
                    ContactEntry(
                        username=username,
                        remark=remark,
                        nickname=nickname,
                        other_info=contact_other_info(row),
                        privacy_status=infer_contact_privacy_status(flag, local_type, delete_flag),
                        source_table=table,
                    ),
                )
    finally:
        conn.close()

    return sorted(entries_by_username.values(), key=lambda item: (item.remark or item.nickname or item.username).casefold())


def write_contact_cache(account_dir: Path, contacts: list[ContactEntry]) -> Path:
    cache_path = account_dir / "_contacts_cache.json"
    payload = [
        {
            "wxid": item.username,
            "nickname": item.nickname,
            "remark": item.remark,
        }
        for item in contacts
    ]
    cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return cache_path


def export_contact_list(output: Path, contacts: list[ContactEntry]) -> dict[str, Any]:
    output.mkdir(parents=True, exist_ok=True)
    fieldnames = ["wxid", "昵称", "备注名"]
    has_other = any(c.other_info for c in contacts)
    has_privacy = any(c.privacy_status for c in contacts)
    if has_other:
        fieldnames.append("其他信息")
    if has_privacy:
        fieldnames.append("隐私状态")
    rows = [
        {
            "wxid": item.username,
            "昵称": item.nickname,
            "备注名": item.remark,
            **({"其他信息": item.other_info} if has_other else {}),
            **({"隐私状态": item.privacy_status} if has_privacy else {}),
        }
        for item in contacts
    ]
    json_path = output / "contacts.json"
    csv_path = output / "contacts.csv"
    json_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    with csv_path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    exported = len(rows)
    print(f"  共导出 {exported} 位好友", flush=True)
    return {"contacts_exported": len(rows), "contacts_json": str(json_path), "contacts_csv": str(csv_path)}


# 省份列表（含直辖市）
_MERCHANT_PROVINCES = [
    "北京", "上海", "天津", "重庆",
    "河北", "山西", "内蒙古", "辽宁", "吉林", "黑龙江",
    "江苏", "浙江", "安徽", "福建", "江西", "山东",
    "河南", "湖北", "湖南", "广东", "广西", "海南",
    "四川", "贵州", "云南", "西藏", "陕西", "甘肃",
    "青海", "宁夏", "新疆", "香港", "澳门", "台湾",
]
_MERCHANT_DIRECT_CITIES = {"北京", "上海", "天津", "重庆"}
_MERCHANT_DEFAULT_REGION = ("浙江省", "金华市", "义乌市")
# 省市区下拉框缺少下级时，默认选第一项“市辖区”
_MERCHANT_DEFAULT_DISTRICT = "市辖区"


def _extract_region_from_info(other_info: str) -> str:
    """从 other_info 中提取地区字符串"""
    match = re.search(r"地区[:：]\s*(.+?)(?:；|$)", str(other_info or ""))
    return match.group(1).strip() if match else ""


def _extract_merchant_phone(remark: str) -> str:
    """从备注末尾提取手机号。

    新格式：备注末尾为连续数字（≥7位），如 "新账号 xxxxx 13812345678"。
    与商品导出（UID 在末尾）不同，商家导出手机号在末尾。
    """
    text = str(remark or "").strip()
    if not text:
        return ""
    m = re.search(r"(\d{7,})\s*$", text)
    return m.group(1) if m else ""


def _parse_merchant_region(region: str) -> tuple[str, str, str]:
    """将地区字符串（如 '广东 广州'、'北京 海淀'、'湖北 恩施'）
    拆分为带行政区后缀的 (省, 市, 区)；无法解析时返回默认地址。

    仅在无法读取模板 cityData 时作为兜底；正常流程使用
    resolve_merchant_region() 以系统下拉框真实选项为准。
    """
    region = (region or "").strip()
    if not region:
        return _MERCHANT_DEFAULT_REGION
    tokens = region.split()
    prov = tokens[0]
    second = tokens[1] if len(tokens) > 1 else ""
    if prov not in _MERCHANT_PROVINCES:
        return _MERCHANT_DEFAULT_REGION
    # 直辖市：省=市=XX市，第二级为区
    if prov in _MERCHANT_DIRECT_CITIES:
        if not second:
            return f"{prov}市", f"{prov}市", _MERCHANT_DEFAULT_DISTRICT
        district = second if second.endswith("区") else f"{second}区"
        return f"{prov}市", f"{prov}市", district
    province_full = _REGION_AUTONOMOUS_SUFFIX.get(prov, f"{prov}省")
    if not second:
        # 只有省份，不足以定位城市，回退默认
        return _MERCHANT_DEFAULT_REGION
    # 自治州用全称，其余补“市”；区缺省默认下拉第一项“市辖区”
    city_full = _REGION_AUTOPREF_CN.get(second, f"{second}市")
    return province_full, city_full, _MERCHANT_DEFAULT_DISTRICT


class _RegionMap:
    """商家导入模板 cityData 的省市区联动数据。"""

    def __init__(
        self,
        provinces: list[str],
        province_cities: dict[str, list[str]],
        city_districts: dict[str, list[str]],
    ) -> None:
        self.provinces = provinces
        self.province_cities = province_cities
        self.city_districts = city_districts


def load_merchant_region_map(template_path: "str | Path") -> "_RegionMap | None":
    """从模板 cityData sheet 读取省/市/区可选项。

    布局：A 列行2起为省名顺序；B 列起每列一个省（行1省名、行2起市名），
    省列之后每列一个市（行1市名、行2起区县名）。
    """
    try:
        from openpyxl import load_workbook
    except ImportError:
        return None
    path = Path(template_path)
    if not path.exists():
        return None
    try:
        wb = load_workbook(str(path), data_only=True)
        if "cityData" not in wb.sheetnames:
            return None
        ws = wb["cityData"]
        max_r, max_c = ws.max_row or 0, ws.max_column or 0
        # 一次性读取全部单元格（cityData 很小，约 36 行 × 400 列）
        grid = [[ws.cell(r, c).value for c in range(1, max_c + 1)] for r in range(1, max_r + 1)]
        prov_order = [row[0] for row in grid[1:]]
        prov_order = [p for p in prov_order if p]

        province_cities: dict[str, list[str]] = {}
        city_districts: dict[str, list[str]] = {}
        idx = 0
        city_zone = False
        for ci in range(1, max_c):  # grid 列下标，0=A 标签列，从 B 开始
            head = grid[0][ci]
            if not head:
                continue
            values = [grid[ri][ci] for ri in range(1, max_r)]
            values = [v for v in values if v]
            if not city_zone and idx < len(prov_order) and head == prov_order[idx]:
                province_cities[head] = values
                idx += 1
                continue
            city_zone = True
            city_districts.setdefault(head, values)
        wb.close()
        if not province_cities or not city_districts:
            return None
        return _RegionMap(prov_order, province_cities, city_districts)
    except Exception:
        return None


def _split_region_tokens(region: str) -> "tuple[str, str, str] | None":
    """把 '广东 广州' / '北京 海淀' / '湖北 恩施' 拆为 (省短名, 市短名, 区提示)。

    直辖市第二级是区提示；普通省份第二级是市短名。无法识别返回 None。
    """
    tokens = str(region or "").split()
    if not tokens or tokens[0] not in _MERCHANT_PROVINCES:
        return None
    prov = tokens[0]
    second = tokens[1] if len(tokens) > 1 else ""
    if prov in _MERCHANT_DIRECT_CITIES:
        return prov, prov, second
    return prov, second, ""


def resolve_merchant_region(region: str, region_map: "_RegionMap | None") -> tuple[str, str, str]:
    """以模板 cityData 的真实可选项为准解析 (省, 市, 区)。

    - 省/市按短名前缀匹配系统全名，匹配不到则取该下拉框第一项；
    - 区：直辖市若能从资料定位到具体区则保留（须在可选列表中），
      否则一律取该市下拉框第一项（不同城市第一项可能是“市辖区/东城区/万州区”等）；
    - 无地区/境外/无法定位：默认 浙江省 金华市 义乌市。
    """
    if region_map is None:
        return _parse_merchant_region(region)
    tokens = _split_region_tokens(region)
    if tokens is None:
        return _MERCHANT_DEFAULT_REGION
    prov_short, city_short, district_hint = tokens

    # 省：短名 -> 系统全名（内蒙古->内蒙古自治区 等）
    prov_full = next(
        (p for p in region_map.provinces if p == prov_short or p.startswith(prov_short)),
        "",
    )
    if not prov_full:
        return _MERCHANT_DEFAULT_REGION

    cities = region_map.province_cities.get(prov_full) or [prov_full]
    # 市：短名前缀匹配（恩施 -> 恩施土家族苗族自治州）；失败取第一项
    city_full = ""
    if city_short:
        city_full = next(
            (c for c in cities if c == city_short or c.startswith(city_short)),
            "",
        )
    if not city_full:
        city_full = prov_full if prov_full in cities else cities[0]

    # 区：优先用资料中的区提示（需在可选列表内），否则取第一项
    districts = region_map.city_districts.get(city_full) or []
    district = ""
    if district_hint:
        hint = district_hint
        if not hint.endswith(("区", "县", "市", "旗")):
            hint = f"{hint}区"
        district = next(
            (d for d in districts if d == hint or d.startswith(district_hint)),
            "",
        )
    if not district and districts:
        district = districts[0]
    return prov_full, city_full, district



def _extract_wechat_id(other_info: str) -> str:
    """从 other_info 中提取微信号（可用于添加好友的 ID）"""
    match = re.search(r"微信号[:：]\s*([^；;]+)", str(other_info or ""))
    return match.group(1).strip() if match else ""


def _strip_phone_from_text(text: str) -> str:
    """去除文本中的手机号"""
    return _PHONE_RE.sub("", str(text or "")).strip()


def export_merchant_list(
    output: Path,
    contacts: list[ContactEntry],
    prefix: str,
    template_path: str = "",
) -> dict[str, Any]:
    """按备注关键词筛选好友，导出为商家导入模板格式的 Excel。

    列：商家名称(微信昵称)、手机号(备注中解析)、微信号(从other_info提取)、省、市、区。
    不排序，保持好友列表原始顺序。
    """
    import shutil

    # 定位模板文件（需提前加载 cityData 以保证省市区取系统可选值）
    project_root = Path(__file__).resolve().parent.parent
    tpl = Path(template_path).expanduser().resolve() if template_path else project_root / "config" / "商家导入模板.xlsx"
    region_map = load_merchant_region_map(tpl)
    if region_map is None:
        print("  [warning] 未读取到模板 cityData，地区将使用内置兜底规则", flush=True)

    filtered = filter_contacts_by_prefix(contacts, prefix)
    prefix_desc = f"备注关键词「{prefix}」" if prefix else "全部好友"
    print(f"  筛选到 {len(filtered)} 位好友（{prefix_desc}）", flush=True)

    # 解析备注字段，收集有效商家记录
    rows: list[dict] = []
    skipped = 0
    for contact in filtered:
        phone = _extract_merchant_phone(contact.remark)
        if not phone:
            skipped += 1
            continue
        province, city, district = resolve_merchant_region(
            _extract_region_from_info(contact.other_info), region_map
        )
        # 商家名称：从昵称中剥离手机号，若昵称为空则用备注
        raw_name = contact.nickname or contact.remark or contact.username
        merchant_name = _strip_phone_from_text(raw_name) or raw_name
        # 微信号：优先从 other_info 提取，没有则用 wxid
        wechat_id = _extract_wechat_id(contact.other_info) or contact.username
        rows.append({
            "商家名称": merchant_name,
            "手机号": phone,
            "微信号": wechat_id,
            "省": province,
            "市": city,
            "区": district,
        })
    if skipped:
        print(f"  跳过 {skipped} 位无手机号的好友", flush=True)

    try:
        from openpyxl import load_workbook, Workbook
        from openpyxl.styles import Alignment, Font
    except ImportError:
        print("  未安装 openpyxl，已跳过商家导入表格生成", flush=True)
        return {"merchant_exported": 0}

    if tpl.exists():
        out_path = output / "商家导入.xlsx"
        shutil.copy(tpl, out_path)
        wb = load_workbook(out_path)
        ws = wb["商家导入"] if "商家导入" in wb.sheetnames else wb.active
        # 清除模板中可能存在的示例数据行（保留表头）
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
    else:
        # 无模板则新建
        out_path = output / "商家导入.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "商家导入"
        ws.append(["商家名称", "手机号", "微信号", "省", "市", "区"])
        for cell in ws[1]:
            cell.font = Font(bold=True)

    for row in rows:
        ws.append([row["商家名称"], row["手机号"], row["微信号"], row["省"], row["市"], row["区"]])

    for r in ws.iter_rows(min_row=2):
        for cell in r:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    wb.save(str(out_path))
    print(f"  已生成商家导入表格：{out_path}（共 {len(rows)} 位商家）", flush=True)
    return {"merchant_exported": len(rows), "merchant_xlsx": str(out_path)}


def build_contact_lookup(contacts: list[ContactEntry]) -> dict[str, list[ContactEntry]]:
    lookup: dict[str, list[ContactEntry]] = {}
    for contact in contacts:
        for value in (contact.remark, contact.nickname):
            key = normalize_contact_lookup_text(value)
            if key:
                lookup.setdefault(key, []).append(contact)
    return lookup


def filter_contacts_by_prefix(contacts: list[ContactEntry], prefix: str) -> list[ContactEntry]:
    """Return contacts whose remark (备注名) contains the given prefix text.

    Only the remark name is matched; the WeChat nickname is ignored. Matching
    is case-insensitive and trims surrounding whitespace. This supports
    matching any part of the remark (e.g. "库存联盟" matches "服装 库存联盟 138...").
    Empty prefix returns all contacts.
    """
    key = normalize_contact_lookup_text(prefix)
    if not key:
        return list(contacts)
    matches: list[ContactEntry] = []
    for contact in contacts:
        remark_key = normalize_contact_lookup_text(contact.remark)
        if key in remark_key:
            matches.append(contact)
    return matches


def build_contact_display_names(contacts: list[ContactEntry]) -> dict[str, str]:
    out: dict[str, str] = {}
    for contact in contacts:
        display = contact.remark or contact.nickname or contact.username
        if contact.username and display:
            out[contact.username] = display
    return out


# Human-readable labels for WeChat moments post types.
POST_TYPE_LABELS: dict[int, str] = {
    1: "图文",
    3: "公众号文章",
    5: "分享链接",
    7: "朋友圈封面",
    28: "视频号",
    42: "音乐/链接卡片",
}

# Common mainland China mobile pattern; also tolerates spaces/dashes.
_PHONE_RE = re.compile(r"(?:\+?86[\s-]?)?(1[3-9]\d{9})")


def post_type_label(type_value: Any) -> str:
    try:
        return POST_TYPE_LABELS.get(int(type_value), f"其他({type_value})")
    except (TypeError, ValueError):
        return str(type_value or "")


def extract_phones(*texts: str) -> list[str]:
    """Extract unique phone numbers from the given text fragments."""
    found: list[str] = []
    for text in texts:
        if not text:
            continue
        for match in _PHONE_RE.finditer(str(text)):
            number = match.group(1)
            if number not in found:
                found.append(number)
    return found


def contact_phone(contact: ContactEntry | None) -> str:
    if contact is None or not contact.other_info:
        return ""
    phones = extract_phones(contact.other_info)
    return phones[0] if phones else ""


# A remark may be structured as "类别 平台名 手机号 UID", e.g. "服装 库存联盟 13800138000 wxid_abc".
_REMARK_SPLIT_RE = re.compile(r"[\s,，;；、|]+")


def parse_remark_fields(remark: str) -> tuple[str, str, str, str]:
    """Parse (category, platform_name, phone, uid) from a contact remark.

    Expected format: ``类别 平台名 手机号 UID`` (tokens separated by
    whitespace or common punctuation). The phone number is located by regex;
    the token immediately following the phone is treated as the UID. Missing
    fields return empty strings.
    """
    if not remark:
        return "", "", "", ""
    tokens = [t for t in _REMARK_SPLIT_RE.split(str(remark).strip()) if t]
    category = tokens[0] if len(tokens) >= 1 else ""
    platform = tokens[1] if len(tokens) >= 2 else ""
    phone = ""
    uid = ""
    for i, token in enumerate(tokens):
        if _PHONE_RE.fullmatch(token):
            phone = token
            if i + 1 < len(tokens):
                uid = tokens[i + 1]
            break
    # Fallback: if no standalone phone token, try regex over the whole remark.
    if not phone:
        match = _PHONE_RE.search(str(remark))
        if match:
            phone = match.group(1)
    return category, platform, phone, uid


def display_interaction_name(value: Any, contact_names: dict[str, str]) -> str:
    if isinstance(value, dict):
        username = str(value.get("username") or value.get("userName") or "").replace("\xa0", " ").strip()
        if username and username in contact_names:
            return contact_names[username]
        nickname = str(value.get("nickname") or value.get("nickName") or value.get("displayName") or "").replace("\xa0", " ").strip()
        return nickname or username
    raw = str(value or "").replace("\xa0", " ").strip()
    if not raw:
        return ""
    return contact_names.get(raw, raw)


def display_interaction_person(username: Any, nickname: Any, contact_names: dict[str, str]) -> str:
    user = str(username or "").replace("\xa0", " ").strip()
    if user and user in contact_names:
        return contact_names[user]
    nick = str(nickname or "").replace("\xa0", " ").strip()
    if nick:
        return nick
    return user


def build_self_display_lookup(
    account_info: AccountInfo,
    config: dict[str, Any],
    posts: list[dict[str, Any]],
    contact_names: dict[str, str],
) -> dict[str, str]:
    candidates = self_username_candidates(account_info, config)
    candidate_set = set(candidates)
    display = str(config.get("self_display_name") or config.get("self_nickname") or "").strip()
    if not display:
        for post in posts:
            username = str(post.get("username") or "").strip()
            if username not in candidate_set:
                continue
            display = str(post.get("displayName") or "").replace("\xa0", " ").strip()
            if display:
                break
    if not display:
        display = contact_names.get(account_info.account, "") or account_info.account
    return {candidate: display for candidate in candidates if candidate and display}


async def materialize_comment_stickers(
    account_info: AccountInfo,
    account_dir: Path,
    output: Path,
    post_dir: Path,
    post: dict[str, Any],
    comment: dict[str, Any],
    comment_index: int,
    stats: dict[str, int],
    *,
    allow_download: bool,
    client: httpx.AsyncClient | None = None,
) -> list[str]:
    rels: list[str] = []
    images = comment.get("images") if isinstance(comment.get("images"), list) else []
    for image_index, raw in enumerate(images, 1):
        media = raw if isinstance(raw, dict) else {}
        if not media:
            continue
        payload, mt, _source = await choose_image(
            account_info, account_dir, post, media, allow_download=allow_download, client=client
        )
        if not payload:
            stats["missing_stickers"] = stats.get("missing_stickers", 0) + 1
            continue
        post_dir.mkdir(parents=True, exist_ok=True)
        name = f"emoji_{comment_index:02d}_{image_index:02d}{mime_to_ext(mt)}"
        target = post_dir / name
        target.write_bytes(payload)
        rel = target.relative_to(output).as_posix()
        rels.append(rel)
        stats["stickers"] = stats.get("stickers", 0) + 1
    return rels


async def format_post_interactions(
    account_info: AccountInfo,
    account_dir: Path,
    output: Path,
    post_dir: Path,
    post: dict[str, Any],
    contact_names: dict[str, str],
    stats: dict[str, int],
    *,
    allow_download: bool,
    client: httpx.AsyncClient | None = None,
) -> str:
    lines: list[str] = []

    raw_likes = post.get("likes") if isinstance(post.get("likes"), list) else []
    likes = []
    for item in raw_likes:
        name = display_interaction_name(item, contact_names)
        if name and name not in likes:
            likes.append(name)
    if likes:
        lines.append(f"**❤️ 点赞**：{'、'.join(likes)}")

    raw_comments = post.get("comments") if isinstance(post.get("comments"), list) else []
    comment_lines: list[str] = []
    for comment_index, item in enumerate(raw_comments, 1):
        if not isinstance(item, dict):
            continue
        author = display_interaction_person(item.get("username"), item.get("nickname"), contact_names) or "未知好友"
        reply_to = display_interaction_person(item.get("refUsername"), item.get("refNickname"), contact_names)
        content = emojify_wechat_shortcodes(str(item.get("content") or "").replace("\xa0", " ").strip())
        sticker_rels = await materialize_comment_stickers(
            account_info,
            account_dir,
            output,
            post_dir,
            post,
            item,
            comment_index,
            stats,
            allow_download=allow_download,
            client=client,
        )
        sticker_md = " ".join(f"![表情包]({rel})" for rel in sticker_rels)
        if sticker_md:
            content = f"{content} {sticker_md}".strip()
        if not content:
            continue
        if reply_to:
            comment_lines.append(f"- {author} 回复 {reply_to}：{content}")
        else:
            comment_lines.append(f"- {author}：{content}")
    if comment_lines:
        if lines:
            lines.append("")
        lines.append("**💬 评论**")
        lines.extend(comment_lines)

    return "\n".join(lines).strip()


CONTACT_INPUT_SEPARATORS = re.compile(r"[\r\n,，;；|、]+")


def split_contact_inputs(raw: str) -> list[str]:
    """Split a pasted input into individual contact entries.

    Supports common separators: newlines, half/full-width commas, half/full-
    width semicolons, pipe, and the Chinese enumeration comma. Whitespace
    around each piece is trimmed, empty pieces are dropped, and duplicate
    pieces (case-insensitive) within the same paste are kept only once.
    Special commands (/me, /enough) are returned as-is without splitting.
    """
    text = str(raw or "").strip()
    if not text:
        return []
    if text in {"/me", "/enough"}:
        return [text]
    pieces: list[str] = []
    seen: set[str] = set()
    for part in CONTACT_INPUT_SEPARATORS.split(text):
        piece = part.strip()
        if not piece:
            continue
        key = normalize_contact_lookup_text(piece)
        if key in seen:
            continue
        seen.add(key)
        pieces.append(piece)
    return pieces


def resolve_contact_input(
    raw: str,
    lookup: dict[str, list[ContactEntry]],
    account_info: AccountInfo,
    config: dict[str, Any],
) -> list[str]:
    value = str(raw or "").strip()
    if value == "/me":
        return self_username_candidates(account_info, config)

    key = normalize_contact_lookup_text(value)
    matches = lookup.get(key) or []
    if not matches:
        raise ValueError(f"没有找到好友：{value}")

    usernames = sorted({item.username for item in matches if item.username})
    if len(usernames) > 1:
        names = []
        for item in matches[:8]:
            label = item.remark or item.nickname or item.username
            names.append(f"{label}({item.username})")
        raise ValueError(f"匹配到多个好友：{value}。请给目标好友设置更唯一的备注后重试。匹配项：{', '.join(names)}")
    return usernames


def ask_friend_filter_mode(cli_value: str = "") -> str:
    """Determine how to select friends when not exporting only self.

    Returns one of: "all", "prefix", "manual".
    Honors the provided CLI value when non-empty; otherwise asks
    interactively. An empty/unknown interactive input defaults to "all".
    """
    mode = str(cli_value or "").strip().lower()
    valid = {"all": "all", "a": "all", "prefix": "prefix", "p": "prefix",
             "manual": "manual", "m": "manual"}
    if mode in valid:
        return valid[mode]
    while True:
        raw = input(
            "请选择好友筛选方式（a=全部好友，p=按备注关键词筛选，"
            "m=手动逐个选择；回车=a）: "
        ).strip().lower()
        if not raw:
            return "all"
        if raw in valid:
            return valid[raw]
        print("输入无效，请输入 a、p 或 m。", flush=True)


def ask_export_type(cli_value: str = "") -> str:
    """先选导出类型：friends=好友/商家表格，moments=朋友圈商品，default=默认导出。

    默认导出 = 昨天 + 备注含“库存联盟”的朋友圈（图片+清洗后表格），零交互。
    返回 "friends" / "moments" / "default"。CLI 值非空时直接采用，否则交互询问。
    """
    raw = str(cli_value or "").strip().lower()
    friends_aliases = {"f", "friend", "friends", "contacts", "merchant"}
    moments_aliases = {"m", "moment", "moments", "sns", "product", "products"}
    default_aliases = {"d", "default", "默认"}
    if raw in friends_aliases:
        return "friends"
    if raw in moments_aliases:
        return "moments"
    if raw in default_aliases:
        return "default"
    if raw:
        # 未知 CLI 值不阻断交互，继续询问
        pass
    while True:
        try:
            answer = input(
                "请选择导出类型（d=默认导出：昨天+库存联盟+图片+清洗表格；"
                "f=好友/商家表格；m=朋友圈商品；回车=d）: "
            ).strip().lower()
        except EOFError:
            return "default"
        if not answer or answer in default_aliases:
            return "default"
        if answer in {"r", "reset"}:
            raise RestartFlow
        if answer in {"q", "quit", "exit"}:
            raise SystemExit(0)
        if answer in moments_aliases:
            return "moments"
        if answer in friends_aliases:
            return "friends"
        print("输入无效，请输入 d、f 或 m（r=重置，q=退出）。", flush=True)


def resolve_friends_by_prefix(
    contacts: list[ContactEntry],
    prefix: str,
) -> tuple[list[str], list[dict[str, str]]]:
    """Collect usernames of contacts whose remark (备注名) starts with prefix.

    Only the remark name is matched; the WeChat nickname is ignored. Returns
    (usernames, friend_inputs) compatible with the existing export flow.
    Prints the matched contacts so the user can confirm.
    """
    matches = filter_contacts_by_prefix(contacts, prefix)
    if not matches:
        raise ValueError(f"没有找到备注包含“{prefix}”的好友")
    usernames = [item.username for item in matches if item.username]
    friend_inputs = [{"input": f"prefix:{prefix}", "usernames": ",".join(usernames)}]
    print(f"  按前缀“{prefix}”匹配到 {len(usernames)} 位好友：", flush=True)
    for item in matches:
        label = item.remark or item.nickname or item.username
        print(f"    - {label}（{item.username}）", flush=True)
    return usernames, friend_inputs


def ask_friend_filters(
    account_info: AccountInfo,
    account_dir: Path,
    config: dict[str, Any],
    contacts: list[ContactEntry] | None = None,
) -> tuple[list[str] | None, list[dict[str, str]]]:
    contacts = contacts if contacts is not None else load_contact_entries(account_dir)
    write_contact_cache(account_dir, contacts)
    lookup = build_contact_lookup(contacts)

    selected_usernames: list[str] = []
    selected_contacts: list[dict[str, str]] = []
    index = 1
    while True:
        if index == 1:
            raw = input(
                "请输入第1位好友的昵称或备注（回车=所有人，可批量粘贴，"
                "用逗号（, 或，）、分号（; 或；）、顿号（、）、竖线（|）或换行分隔）: "
            ).strip()
            if not raw:
                return None, []
        else:
            raw = input(
                f"请输入第{index}位好友的昵称或备注（/enough=结束，可批量粘贴）: "
            ).strip()
            if not raw:
                continue
        if raw == "/enough":
            if not selected_usernames:
                raise ValueError("还没有选择任何好友")
            return selected_usernames, selected_contacts

        pieces = split_contact_inputs(raw)
        if not pieces:
            continue
        failures: list[str] = []
        added_in_batch = 0
        for piece in pieces:
            try:
                resolved = resolve_contact_input(piece, lookup, account_info, config)
            except ValueError as exc:
                failures.append(str(exc))
                continue
            for username in resolved:
                if username not in selected_usernames:
                    selected_usernames.append(username)
            selected_contacts.append({"input": piece, "usernames": ",".join(resolved)})
            added_in_batch += 1
        if failures:
            print("以下好友未能匹配，已跳过：", flush=True)
            for msg in failures:
                print(f"  - {msg}", flush=True)
        if added_in_batch:
            index += added_in_batch


async def fetch_remote_image(
    account_dir: Path,
    media: dict[str, Any],
    prefer_thumb: bool,
    *,
    client: httpx.AsyncClient | None = None,
) -> tuple[bytes, str, str]:
    from wechat_decrypt_tool.modules.sns_media import fix_sns_cdn_url, get_cached_sns_remote_image, try_fetch_and_decrypt_sns_image_remote

    raw_url, key, token, is_original = sns_image_source(media, prefer_thumb=prefer_thumb)
    if not raw_url:
        return b"", "", ""
    fixed = fix_sns_cdn_url(raw_url, token=token, is_video=False)
    expected = expected_image_size(media) if is_original else (0, 0)
    cached = get_cached_sns_remote_image(account_dir=account_dir, url=fixed, key=str(key or ""), token=str(token or ""))
    if cached is not None and meets_expected_size(bytes(cached.payload or b""), str(cached.media_type or ""), expected):
        return bytes(cached.payload or b""), str(cached.media_type or ""), "remote-cache-original" if is_original else "remote-cache-thumb"
    fetched = await try_fetch_and_decrypt_sns_image_remote(
        account_dir=account_dir,
        url=fixed,
        key=str(key or ""),
        token=str(token or ""),
        use_cache=True,
        client=client,
    )
    if fetched is not None and meets_expected_size(bytes(fetched.payload or b""), str(fetched.media_type or ""), expected):
        return bytes(fetched.payload or b""), str(fetched.media_type or ""), "remote-original" if is_original else "remote-thumb"
    return b"", "", ""


def read_local_image(account_info: AccountInfo, account_dir: Path, post: dict[str, Any], media: dict[str, Any]) -> tuple[bytes, str, str]:
    from wechat_decrypt_tool.modules.media_decrypt import _detect_image_media_type, _read_and_maybe_decrypt_media

    local = resolve_sns_exact_cached_image_path(account_info.wxid_dir, post, media, str(media.get("url") or media.get("thumb") or ""))
    if local is None:
        post_id = str(post.get("id") or post.get("tid") or "").strip()
        media_id = str(media.get("id") or "").strip()
        for candidate_type in (MEDIA_TYPE_IMAGE, MEDIA_TYPE_LIVE_PHOTO, MEDIA_TYPE_VIDEO, 0):
            key = generate_sns_cache_key(post_id, media_id, candidate_type)
            found = resolve_sns_cached_image_path_by_cache_key(account_info.wxid_dir, key)
            if found:
                local = found
                break
    if local is None:
        return b"", "", ""
    try:
        payload, media_type = _read_and_maybe_decrypt_media(Path(local), account_dir)
    except Exception as exc:
        print(f"[warning] {exc}", file=sys.stderr)
        return b"", "", ""
    mt = str(media_type or "").split(";", 1)[0].strip() or _detect_image_media_type(payload[:32])
    if not payload or not mt.startswith("image/"):
        return b"", "", ""
    return payload, mt, "local-cache"


async def choose_image(
    account_info: AccountInfo,
    account_dir: Path,
    post: dict[str, Any],
    media: dict[str, Any],
    *,
    allow_download: bool,
    client: httpx.AsyncClient | None = None,
) -> tuple[bytes, str, str]:
    if allow_download:
        payload, mt, source = await fetch_remote_image(account_dir, media, prefer_thumb=False, client=client)
        if payload:
            return payload, mt, source
    local_payload, local_mt, local_source = read_local_image(account_info, account_dir, post, media)
    expected = expected_image_size(media)
    if local_payload and meets_expected_size(local_payload, local_mt, expected):
        return local_payload, local_mt, local_source
    if allow_download:
        payload, mt, source = await fetch_remote_image(account_dir, media, prefer_thumb=True, client=client)
        if payload:
            return payload, mt, source
    if local_payload:
        return local_payload, local_mt, "local-cache-small"
    return b"", "", ""


def mime_to_ext(media_type: str) -> str:
    from wechat_decrypt_tool.modules.sns_media import mime_to_ext as _impl
    return _impl(media_type)


# ---------------------------------------------------------------------------
# 二维码检测：先取缩略图（本地缓存优先），用 OpenCV 检测，命中则跳过原图下载
# 判定标准：能解码出内容才算二维码，降低商品图误判率
# ---------------------------------------------------------------------------
_QR_STATE: dict[str, Any] = {"detector": None, "disabled": False, "warned": False}


def detect_qrcode(payload: bytes) -> bool:
    """检测图片字节流是否含可解码的二维码。未安装 opencv 或解码失败时返回 False。"""
    if not payload or _QR_STATE["disabled"]:
        return False
    try:
        import cv2
        import numpy as np
    except ImportError:
        _QR_STATE["disabled"] = True
        if not _QR_STATE["warned"]:
            print(
                "  [warning] 未安装 opencv-python-headless，二维码检测已禁用（pip install opencv-python-headless）",
                file=sys.stderr,
                flush=True,
            )
            _QR_STATE["warned"] = True
        return False
    try:
        arr = np.frombuffer(payload, np.uint8)
        img = cv2.imdecode(arr, cv2.IMREAD_GRAYSCALE)
        if img is None:
            return False
        # 大图检测慢且定位点易过密：先缩到 800px 以内再检测，兼顾速度与准确率
        h, w = img.shape[:2]
        if max(h, w) > 800:
            scale = 800.0 / max(h, w)
            img = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
        detector = _QR_STATE.get("detector")
        if detector is None:
            detector = cv2.QRCodeDetector()
            _QR_STATE["detector"] = detector
        data, points, _straight = detector.detectAndDecode(img)
        if data:
            return True
        # 点状/艺术二维码（微信名片码常见）数据区解码失败，但定位点可检出。
        # 定位点四点须通过几何校验：近正方形（透视容差 25%）+ 面积占比 2%~60%，
        # 否则条纹/针织纹理会大量误检（实测 5.5% 误杀，校验后 600 张商品图 0 误杀）
        if points is not None:
            p = points[0]
            d2 = [
                (p[i][0] - p[(i + 1) % 4][0]) ** 2 + (p[i][1] - p[(i + 1) % 4][1]) ** 2
                for i in range(4)
            ]
            if max(d2) <= 1.5625 * (min(d2) + 1e-6):
                h, w = img.shape[:2]
                frac = cv2.contourArea(p.astype(np.float32)) / (h * w)
                if 0.02 <= frac <= 0.6:
                    return True
        ok, decoded, _pts, _st = detector.detectAndDecodeMulti(img)
        return bool(ok and any(decoded))
    except Exception:
        return False


async def fetch_qrcode_probe_image(
    account_info: AccountInfo,
    account_dir: Path,
    post: dict[str, Any],
    media: dict[str, Any],
    *,
    allow_download: bool,
    client: httpx.AsyncClient | None = None,
) -> bytes:
    """取用于二维码检测的小图：本地缓存优先，没有再远程取缩略图（约 200px，流量极小）。"""
    payload, _mt, _source = read_local_image(account_info, account_dir, post, media)
    if payload:
        return payload
    if allow_download:
        payload, _mt, _source = await fetch_remote_image(account_dir, media, prefer_thumb=True, client=client)
        return payload
    return b""


def should_skip_post_download(post: dict[str, Any], contact_remark: str) -> tuple[bool, str]:
    """复用清洗规则预判朋友圈是否为商品；非商品则跳过整条图片下载，返回 (是否跳过, 原因)。"""
    from tools.clean_moments import is_product_record

    image_medias: list[dict[str, Any]] = []
    for media_raw in post.get("media") or []:
        media = media_raw if isinstance(media_raw, dict) else {}
        try:
            media_type = int(media.get("type") or 0)
        except Exception:
            media_type = 0
        if media_type == MEDIA_TYPE_IMAGE:
            image_medias.append(media)
    probe = {
        "content": str(post.get("contentDesc") or ""),
        "remark": contact_remark,
        "images": image_medias,
        "image_count": len(image_medias),
    }
    is_product, reasons, _category = is_product_record(probe)
    if is_product:
        return False, ""
    return True, ",".join(reasons)


_DOWNLOAD_CONCURRENCY = 6  # 同时在途的图片请求数；Wi-Fi/CDN 单 IP 再高容易触发限流
_DOWNLOAD_WINDOW = 12  # 有序预取窗口：最多提前下载 12 张，占用约几十 MB


async def export_markdown(
    account_info: AccountInfo,
    account_dir: Path,
    output: Path,
    start: datetime | None,
    end: datetime | None,
    usernames: list[str] | None,
    config: dict[str, Any],
    keep_interactions: bool,
    contact_names: dict[str, str],
    contact_lookup: dict[str, ContactEntry] | None = None,
    *,
    allow_download: bool,
    skip_non_product: bool = False,
    skip_qrcode: bool = False,
    write_markdown: bool = True,
    client: httpx.AsyncClient | None = None,
    on_image_saved: Callable[[Path], None] | None = None,
) -> tuple[dict[str, Any], list[ExportedPost]]:
    figure = output / "figure"
    figure.mkdir(parents=True, exist_ok=True)
    # 每批导出重置二维码检测器状态：同进程内补装 opencv 后下一批即可生效
    _QR_STATE["disabled"] = False
    posts = load_timeline(account_dir, usernames, source="auto")
    filtered: list[dict[str, Any]] = []
    for post in posts:
        created_ts = int(post.get("createTime") or 0)
        if created_ts <= 0:
            continue
        created = datetime.fromtimestamp(created_ts)
        if start and created < start:
            continue
        if end and created > end:
            continue
        filtered.append(post)
    filtered.sort(key=lambda item: int(item.get("createTime") or 0), reverse=True)
    interaction_contact_names = dict(contact_names)
    interaction_contact_names.update(build_self_display_lookup(account_info, config, filtered, contact_names))
    coverage = build_coverage_report(posts, filtered, start, end)
    write_coverage_report(output, coverage)

    lines = ["# 微信朋友圈备份", ""]
    stats = {
        "posts": len(filtered),
        "images": 0,
        "original_images": 0,
        "fallback_images": 0,
        "missing_images": 0,
        "stickers": 0,
        "missing_stickers": 0,
        "skipped_non_product_posts": 0,
        "qrcode_images": 0,
        "local_earliest": coverage.get("local_earliest", ""),
        "local_latest": coverage.get("local_latest", ""),
        "coverage_gap_count": len(coverage.get("large_gaps") or []),
    }
    exported_posts: list[ExportedPost] = []
    moments_records: list[dict[str, Any]] = []
    total = len(filtered)

    # 先按最终顺序收集待下载图片，再开启有限窗口的有序预取：消费端严格按帖子/图片顺序
    # 落盘，下载最多提前 12 张、6 路并发，输出顺序与统计和串行版完全一致。
    download_tasks: list[tuple[dict[str, Any], dict[str, Any]]] = []
    post_skip: dict[int, tuple[bool, str]] = {}
    for p_idx, post in enumerate(filtered):
        username = str(post.get("username") or "").strip()
        contact = (contact_lookup or {}).get(username)
        contact_remark = contact.remark if contact else ""
        skip_images, skip_reason = (False, "")
        if skip_non_product:
            skip_images, skip_reason = should_skip_post_download(post, contact_remark)
        post_skip[p_idx] = (skip_images, skip_reason)
        if skip_images:
            continue
        for media_raw in post.get("media") or []:
            media = media_raw if isinstance(media_raw, dict) else {}
            try:
                media_type = int(media.get("type") or 0)
            except Exception:
                media_type = 0
            if media_type == MEDIA_TYPE_IMAGE:
                download_tasks.append((post, media))

    download_sem = asyncio.Semaphore(_DOWNLOAD_CONCURRENCY)

    async def _fetch_task(post: dict[str, Any], media: dict[str, Any]) -> tuple[bytes, str, str, bool]:
        async with download_sem:
            if skip_qrcode:
                probe_payload = await fetch_qrcode_probe_image(
                    account_info, account_dir, post, media, allow_download=allow_download, client=client
                )
                if probe_payload and detect_qrcode(probe_payload):
                    return b"", "", "", True
            payload, mt, source = await choose_image(
                account_info, account_dir, post, media, allow_download=allow_download, client=client
            )
            return payload, mt, source, False

    async def _ordered_downloads() -> AsyncIterator[tuple[bytes, str, str, bool]]:
        task_iter = iter(download_tasks)
        pending: deque[asyncio.Task] = deque()
        try:
            while True:
                while len(pending) < _DOWNLOAD_WINDOW:
                    item = next(task_iter, None)
                    if item is None:
                        break
                    pending.append(asyncio.create_task(_fetch_task(item[0], item[1])))
                if not pending:
                    break
                yield await pending.popleft()
        finally:
            for task in pending:
                task.cancel()

    download_results = _ordered_downloads()

    for idx, post in enumerate(filtered, 1):
        created = datetime.fromtimestamp(int(post.get("createTime") or 0))
        time_text = created.strftime("%Y-%m-%d %H:%M:%S")
        username = str(post.get("username") or "").strip()
        fallback_display = str(post.get("displayName") or username or account_info.account).strip()
        display = interaction_contact_names.get(username, fallback_display)
        lines.extend([f"## {format_post_heading(time_text, display)}", ""])
        location = str(post.get("location") or "").strip()
        content = emojify_wechat_shortcodes(str(post.get("contentDesc") or "").strip())
        title = emojify_wechat_shortcodes(str(post.get("title") or "").strip())
        content_url = str(post.get("contentUrl") or "").strip()
        body = content or "（无文字）"
        if title:
            body = f"{body}\n\n🔗 **链接标题**：{title}" if body else f"🔗 **链接标题**：{title}"
        if content_url:
            body = f"{body}\n\n🔗 **链接**：{content_url}" if body else f"🔗 **链接**：{content_url}"
        lines.extend([body, ""])

        contact = (contact_lookup or {}).get(username)
        contact_remark = contact.remark if contact else ""
        skip_images, skip_reason = post_skip.get(idx - 1, (False, ""))
        if skip_images:
            stats["skipped_non_product_posts"] += 1

        post_dir = unique_post_dir(figure, created)
        image_index = 0
        post_qrcode_count = 0
        post_images: list[str] = []
        if not skip_images:
            for media_raw in post.get("media") or []:
                media = media_raw if isinstance(media_raw, dict) else {}
                try:
                    media_type = int(media.get("type") or 0)
                except Exception as exc:
                    print(f"[warning] {exc}", file=sys.stderr)
                    media_type = 0
                if media_type != MEDIA_TYPE_IMAGE:
                    continue
                try:
                    payload, mt, source, skipped_qr = await anext(download_results)
                except StopAsyncIteration:
                    stats["missing_images"] += 1
                    continue
                if skipped_qr:
                    post_qrcode_count += 1
                    stats["qrcode_images"] += 1
                    continue
                if not payload:
                    stats["missing_images"] += 1
                    continue
                image_index += 1
                post_dir.mkdir(parents=True, exist_ok=True)
                name = f"{image_index:02d}{mime_to_ext(mt)}"
                target = post_dir / name
                target.write_bytes(payload)
                if on_image_saved is not None:
                    on_image_saved(target)
                rel = target.relative_to(output).as_posix()
                post_images.append(rel)
                lines.extend([f"![{time_text} image {image_index}]({rel})", ""])
                stats["images"] += 1
                if source.startswith("remote") and "thumb" not in source:
                    stats["original_images"] += 1
                else:
                    stats["fallback_images"] += 1
        interactions = ""
        if keep_interactions and not skip_images:
            interactions = await format_post_interactions(
                account_info,
                account_dir,
                output,
                post_dir,
                post,
                interaction_contact_names,
                stats,
                allow_download=allow_download,
                client=client,
            )
            if interactions:
                lines.extend([interactions, ""])
        lines.extend(["---", ""])
        exported_posts.append(
            ExportedPost(
                time_text=time_text,
                display=display,
                location=location,
                body=body,
                images=post_images,
                interactions=interactions,
            )
        )
        remark_category, platform_name, remark_phone, remark_uid = parse_remark_fields(contact_remark)
        text_for_phone = str(post.get("contentDesc") or "")
        # Phone priority: remark > contact other_info > post body.
        phone = (
            remark_phone
            or contact_phone(contact)
            or (extract_phones(text_for_phone)[0] if extract_phones(text_for_phone) else "")
        )
        # UID priority: remark (after phone) > post username.
        record_uid = remark_uid or username
        moments_records.append(
            {
                "id": str(post.get("id") or ""),
                "tid": post.get("tid"),
                "uid": record_uid,
                "display_name": display,
                "remark": contact_remark,
                "remark_category": remark_category,
                "platform_name": platform_name,
                "nickname": contact.nickname if contact else str(post.get("displayName") or ""),
                "create_time": int(post.get("createTime") or 0),
                "time_text": time_text,
                "type": post.get("type"),
                "category": post_type_label(post.get("type")),
                "content": str(post.get("contentDesc") or ""),
                "title": str(post.get("title") or ""),
                "content_url": str(post.get("contentUrl") or ""),
                "location": location,
                "source_name": str(post.get("sourceName") or ""),
                "images": post_images,
                "image_count": len(post_images),
                "images_skipped": skip_reason,
                "qrcode_images": post_qrcode_count,
                "likes_count": len(post.get("likes") or []),
                "comments_count": len(post.get("comments") or []),
                "phone": phone,
            }
        )
        suffix = ""
        if skip_images:
            suffix = f" [非商品-跳过图片: {skip_reason}]"
        elif post_qrcode_count:
            suffix = f" [跳过二维码图 x{post_qrcode_count}]"
        print(f"  ({idx}/{total}) {time_text} {display}{suffix}", flush=True)
    await download_results.aclose()
    if write_markdown:
        (output / "moments.md").write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    (output / "moments.json").write_text(
        json.dumps(moments_records, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return stats, exported_posts


def image_grid_class(count: int) -> str:
    if count <= 1:
        return "grid-one"
    if count in (2, 3):
        return f"grid-{count}"
    if count == 4:
        return "grid-four"
    return "grid-three"


def markdown_text_to_html(text: str, output: Path | None = None, *, file_uris: bool = False) -> str:
    def inline_markup(raw: str) -> str:
        def image_repl(match: re.Match[str]) -> str:
            alt = match.group(1)
            rel = match.group(2)
            src = Path(output / rel).resolve().as_uri() if output is not None and file_uris else rel
            return f'<img class="inline-sticker" src="{html.escape(src, quote=True)}" alt="{html.escape(alt, quote=True)}">'

        escaped = html.escape(raw)
        escaped = re.sub(r"!\[([^\]]*)\]\(([^)]+)\)", image_repl, escaped)
        escaped = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", escaped)
        return escaped.replace("\n", "<br>")

    normalized = emojify_wechat_shortcodes(str(text or ""))
    paragraphs = [inline_markup(p) for p in re.split(r"\n{2,}", normalized) if p.strip()]
    return "".join(f"<p>{p}</p>" for p in paragraphs) or "<p>（无文字）</p>"


def interactions_to_html(text: str, output: Path | None = None, *, file_uris: bool = False) -> str:
    if not str(text or "").strip():
        return ""
    return f'<div class="interactions">{markdown_text_to_html(text, output, file_uris=file_uris)}</div>'


def format_post_heading(time_text: str, display: str) -> str:
    display = str(display or "").strip()
    if display and display != "未知":
        return f"{time_text} · {display}"
    return time_text


def post_heading_text(post: ExportedPost) -> str:
    return format_post_heading(post.time_text, post.display)


def optimized_pdf_image_uri(output: Path, rel: str, max_side: int = 640, quality: int = 60) -> str:
    source = output / rel
    if not source.exists():
        return source.resolve().as_uri()
    safe_rel = Path(rel)
    target = output / "_pdf_assets" / safe_rel.with_suffix(".jpg")
    try:
        from PIL import Image, ImageOps

        target.parent.mkdir(parents=True, exist_ok=True)
        with Image.open(source) as image:
            image = ImageOps.exif_transpose(image)
            image.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)
            if image.mode in {"RGBA", "LA"}:
                background = Image.new("RGB", image.size, "white")
                alpha = image.getchannel("A")
                background.paste(image.convert("RGBA"), mask=alpha)
                image = background
            elif image.mode != "RGB":
                image = image.convert("RGB")
            image.save(target, format="JPEG", quality=quality, optimize=True, progressive=True)
        return target.resolve().as_uri()
    except Exception as exc:
        print(f"[warning] PDF 图片压缩失败，使用原图: {source} ({exc})", file=sys.stderr)
        return source.resolve().as_uri()


def write_pdf_html(
    output: Path,
    posts: list[ExportedPost],
    filename: str = "moments.html",
    optimize_images: bool = False,
) -> Path:
    css = """
@page { size: A4; margin: 10mm 11mm 11mm; }
* { box-sizing: border-box; }
body { margin: 0; color: #1f2328; background: #fff; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI Emoji", "Segoe UI", "Microsoft YaHei", Arial, sans-serif; font-size: 13.5px; line-height: 1.48; }
.page { max-width: 760px; margin: 0 auto; }
.moment { padding: 0 0 10px; margin: 0 0 10px; border-bottom: 1px solid #e5e7eb; }
.moment:last-child { border-bottom: 0; }
.time { font-size: 18px; font-weight: 700; margin: 0 0 7px; letter-spacing: .01em; }
.body { font-size: 14px; }
.body p { margin: 0 0 6px; white-space: normal; overflow-wrap: anywhere; }
.interactions { margin-top: 7px; padding: 7px 9px; background: #f6f8fa; border-left: 3px solid #8ace9b; border-radius: 8px; color: #3f4b4f; font-size: 12px; }
.interactions p { margin: 0 0 3px; }
.inline-sticker { display: inline-block; width: 34px; height: 34px; object-fit: contain; vertical-align: middle; margin: 0 2px; border-radius: 6px; }
.images { display: grid; gap: 5px; margin-top: 7px; width: 100%; max-width: 410px; }
.grid-one { display: block; max-width: 300px; }
.grid-2 { grid-template-columns: repeat(2, 1fr); }
.grid-3, .grid-three { grid-template-columns: repeat(3, 1fr); }
.grid-four { grid-template-columns: repeat(2, 1fr); max-width: 272px; }
.image-cell { width: 100%; aspect-ratio: 1 / 1; overflow: hidden; background: #f3f4f6; border-radius: 6px; }
.grid-one .image-cell { aspect-ratio: auto; }
.image-cell img { display: block; width: 100%; height: 100%; object-fit: cover; }
.grid-one .image-cell img { width: auto; max-width: 100%; height: auto; max-height: 330px; object-fit: contain; }
"""
    parts = [
        "<!doctype html>",
        "<html>",
        "<head>",
        '<meta charset="utf-8">',
        "<title>微信朋友圈备份</title>",
        f"<style>{css}</style>",
        "</head>",
        "<body>",
        '<main class="page">',
    ]
    for post in posts:
        parts.append('<article class="moment">')
        parts.append(f'<h2 class="time">{html.escape(post_heading_text(post))}</h2>')
        parts.append(f'<div class="body">{markdown_text_to_html(post.body, output, file_uris=optimize_images)}</div>')
        if post.images:
            parts.append(f'<div class="images {image_grid_class(len(post.images))}">')
            for rel in post.images:
                src = optimized_pdf_image_uri(output, rel) if optimize_images else rel
                parts.append(f'<div class="image-cell"><img src="{html.escape(src, quote=True)}"></div>')
            parts.append("</div>")
        if post.interactions:
            parts.append(interactions_to_html(post.interactions, output, file_uris=optimize_images))
        parts.append("</article>")
    parts.extend(["</main>", "</body>", "</html>", ""])
    html_path = output / filename
    html_path.write_text("\n".join(parts), encoding="utf-8")
    return html_path


def pdf_image_rows(count: int) -> list[int]:
    if count <= 0:
        return []
    if count <= 3:
        return [count]
    if count == 4:
        return [2, 2]
    if count == 5:
        return [3, 2]
    if count == 6:
        return [3, 3]
    rows: list[int] = []
    remaining = count
    while remaining > 0:
        row_count = min(3, remaining)
        rows.append(row_count)
        remaining -= row_count
    return rows


def pdf_image_columns(count: int) -> int:
    if count <= 1:
        return 1
    if count in (2, 3):
        return count
    if count == 4:
        return 2
    return 3


def render_pdf_direct(output: Path, posts: list[ExportedPost], pdf_path: Path) -> None:
    from PIL import Image, ImageOps
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    font_name = PDF_FONT_NAME
    try:
        pdfmetrics.registerFont(UnicodeCIDFont(font_name))
    except Exception:
        font_name = PDF_FONT_FALLBACK
        print("警告: PDF 中文字体 STSong-Light 不可用，中文可能显示为方块。"
              "安装 Adobe Acrobat 或配置中文字体可解决此问题。")

    image_dpi = PDF_DPI

    page_width, page_height = A4
    margin_left = PDF_MARGIN_LEFT
    margin_right = PDF_MARGIN_RIGHT
    margin_top = PDF_MARGIN_TOP
    margin_bottom = PDF_MARGIN_BOTTOM
    content_width = page_width - margin_left - margin_right
    y = page_height - margin_top
    page_no = 1
    canv = canvas.Canvas(str(pdf_path), pagesize=A4)

    def draw_footer() -> None:
        canv.setFont(PDF_FONT_FALLBACK, 8)
        canv.setFillColor(colors.HexColor("#777777"))
        canv.drawCentredString(page_width / 2, 22, str(page_no))
        canv.setFillColor(colors.black)

    def new_page() -> None:
        nonlocal y, page_no
        draw_footer()
        canv.showPage()
        page_no += 1
        y = page_height - margin_top

    def ensure_space(height: float) -> None:
        if y - height < margin_bottom:
            new_page()

    def text_width(text: str, size: float) -> float:
        return pdfmetrics.stringWidth(text, font_name, size)

    def wrap_line(text: str, size: float, max_width: float) -> list[str]:
        raw = str(text or "")
        if not raw:
            return [""]
        lines: list[str] = []
        current = ""
        for char in raw:
            candidate = current + char
            if current and text_width(candidate, size) > max_width:
                lines.append(current)
                current = char
            else:
                current = candidate
        if current:
            lines.append(current)
        return lines or [""]

    def draw_wrapped(text: str, size: float, leading: float, color: str = "#1f2328") -> None:
        nonlocal y
        canv.setFont(font_name, size)
        canv.setFillColor(colors.HexColor(color))
        paragraphs = re.split(r"\n{2,}", str(text or "").strip() or "（无文字）")
        for paragraph in paragraphs:
            for source_line in paragraph.splitlines() or [""]:
                for line in wrap_line(source_line, size, content_width):
                    ensure_space(leading + 2)
                    canv.drawString(margin_left, y, line)
                    y -= leading
            y -= 3
        canv.setFillColor(colors.black)

    def load_pdf_image(path: Path, max_width_pt: float, max_height_pt: float) -> tuple[ImageReader, float, float] | None:
        try:
            with Image.open(path) as image:
                image = ImageOps.exif_transpose(image)
                width_px, height_px = image.size
                if width_px <= 0 or height_px <= 0:
                    return None
                scale = min(max_width_pt / width_px, max_height_pt / height_px)
                draw_width = max(1.0, width_px * scale)
                draw_height = max(1.0, height_px * scale)
                target_width_px = max(1, int(draw_width / 72 * image_dpi))
                target_height_px = max(1, int(draw_height / 72 * image_dpi))
                resize_scale = min(1.0, target_width_px / width_px, target_height_px / height_px)
                if resize_scale < 1.0:
                    image = image.resize(
                        (max(1, int(width_px * resize_scale)), max(1, int(height_px * resize_scale))),
                        Image.Resampling.LANCZOS,
                    )
                if image.mode in {"RGBA", "LA"}:
                    background = Image.new("RGB", image.size, "white")
                    alpha = image.getchannel("A") if image.mode == "RGBA" else image.getchannel("A")
                    background.paste(image.convert("RGBA"), mask=alpha)
                    image = background
                elif image.mode != "RGB":
                    image = image.convert("RGB")
                buffer = BytesIO()
                image.save(buffer, format="JPEG", quality=92, optimize=True)
                buffer.seek(0)
                return ImageReader(buffer), draw_width, draw_height
        except Exception as exc:
            print(f"[warning] {exc}", file=sys.stderr)
            return None

    def draw_single_image(path: Path) -> None:
        nonlocal y
        max_width = min(content_width, 320)
        max_height = 360
        item = load_pdf_image(path, max_width, max_height)
        if item is None:
            return
        reader, width, height = item
        ensure_space(height + 12)
        canv.drawImage(reader, margin_left, y - height, width=width, height=height, preserveAspectRatio=True, mask="auto")
        y -= height + 10

    def draw_image_grid(paths: list[Path]) -> None:
        nonlocal y
        gap = 5
        grid_width = min(content_width, 365)
        columns = pdf_image_columns(len(paths))
        cell = (grid_width - gap * (columns - 1)) / columns
        index = 0
        for row_count in pdf_image_rows(len(paths)):
            ensure_space(cell + 9)
            x = margin_left
            for path in paths[index:index + row_count]:
                item = load_pdf_image(path, cell, cell)
                if item is not None:
                    reader, width, height = item
                    img_x = x + (cell - width) / 2
                    img_y = y - cell + (cell - height) / 2
                    canv.drawImage(reader, img_x, img_y, width=width, height=height, preserveAspectRatio=True, mask="auto")
                canv.setStrokeColor(colors.HexColor("#eeeeee"))
                canv.rect(x, y - cell, cell, cell, stroke=1, fill=0)
                x += cell + gap
            y -= cell + gap
            index += row_count
        y -= 5

    for post in posts:
        ensure_space(55)
        canv.setFont(font_name, 14)
        canv.setFillColor(colors.HexColor("#1f2328"))
        canv.drawString(margin_left, y, post_heading_text(post))
        y -= 17
        draw_wrapped(post.body, 11, 14.5)
        image_paths = [output / rel for rel in post.images]
        if len(image_paths) == 1:
            draw_single_image(image_paths[0])
        elif image_paths:
            draw_image_grid(image_paths)
        if post.interactions:
            draw_wrapped(post.interactions, 9.5, 12, "#3f4b4f")
        ensure_space(8)
        canv.setStrokeColor(colors.HexColor("#e5e7eb"))
        canv.line(margin_left, y, page_width - margin_right, y)
        y -= 10

    draw_footer()
    canv.save()
    if not pdf_path.read_bytes().startswith(b"%PDF"):
        raise RuntimeError("直接生成的 PDF 数据无效")


def find_chromium_executable() -> str | None:
    env_value = os.environ.get("WX_MOMENTS_CHROME") or os.environ.get("CHROME_PATH")
    candidates: list[Path] = []
    if env_value:
        candidates.append(Path(env_value).expanduser())
    for command in ("chrome", "chrome.exe", "msedge", "msedge.exe", "chromium", "chromium-browser"):
        found = shutil.which(command)
        if found:
            candidates.append(Path(found))
    for root in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(root)
        if not base:
            continue
        candidates.extend(
            [
                Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe",
                Path(base) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
            ]
        )
    seen: set[Path] = set()
    for candidate in candidates:
        if candidate in seen:
            continue
        seen.add(candidate)
        if candidate.exists() and candidate.is_file():
            return str(candidate)
    return None


def render_pdf_chromium(html_path: Path, pdf_path: Path) -> None:
    import subprocess
    import tempfile

    browser = find_chromium_executable()
    if not browser:
        raise RuntimeError("未找到 Chrome/Edge/Chromium")
    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="chrome-pdf-", dir=str(RUNTIME_DIR)) as user_data_dir:
        command = [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--allow-file-access-from-files",
            "--no-pdf-header-footer",
            f"--user-data-dir={user_data_dir}",
            f"--print-to-pdf={pdf_path.resolve()}",
            html_path.resolve().as_uri(),
        ]
        timeout = int(os.environ.get("WX_MOMENTS_CHROME_TIMEOUT", "600"))
        proc = subprocess.run(command, text=True, capture_output=True, timeout=timeout)
        if proc.returncode != 0:
            message = (proc.stderr or proc.stdout or "").strip()
            raise RuntimeError("浏览器 PDF 生成失败: " + " / ".join(message.splitlines()[-3:]))
    if not pdf_path.exists() or not pdf_path.read_bytes().startswith(b"%PDF"):
        raise RuntimeError("浏览器生成的 PDF 数据无效")


def render_pdf(output: Path, posts: list[ExportedPost], pdf_path: Path) -> None:
    if os.environ.get("WX_MOMENTS_PDF_ENGINE", "").lower() != "reportlab":
        html_path = output / "_moments_pdf.html"
        try:
            write_pdf_html(output, posts, filename=html_path.name, optimize_images=True)
            render_pdf_chromium(html_path, pdf_path)
            with contextlib.suppress(OSError):
                html_path.unlink()
            with contextlib.suppress(OSError):
                shutil.rmtree(output / "_pdf_assets")
            return
        except Exception as exc:
            print(f"[warning] 浏览器 PDF 生成失败，回退到直接渲染: {exc}", file=sys.stderr)
            with contextlib.suppress(OSError):
                html_path.unlink()
            with contextlib.suppress(OSError):
                shutil.rmtree(output / "_pdf_assets")
    render_pdf_direct(output, posts, pdf_path)


_step_counter = 0


def progress(label: str) -> None:
    global _step_counter
    _step_counter += 1
    print(f"第{_step_counter}步：{label}", flush=True)



async def main() -> int:
    args = parse_args()
    try:
        if args.quiet:
            RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
            (RUNTIME_DIR / "last_run.log").write_text("", encoding="utf-8")
            log_path = RUNTIME_DIR / "last_run.log"
            with log_path.open("a", encoding="utf-8") as log, contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
                return await _main_impl(args)
        else:
            return await _main_impl(args)
    except (FileNotFoundError, RuntimeError, ValueError) as exc:
        print(f"\n运行失败：{exc}", file=sys.stderr, flush=True)
        if os.environ.get("WXMOMENTS_DEBUG_TRACEBACK", "").strip().lower() in {"1", "true", "yes", "on"}:
            raise
        return 1


async def _main_impl(args: argparse.Namespace) -> int:
    config = load_config(Path(args.config))
    ensure_imports()
    print("欢迎使用 wxMoments 微信朋友圈导出备份工具！\n", flush=True)
    print_input_rules()

    # ---- One-time setup: reused across every export in this session ----
    progress("定位微信并获取密钥")
    try:
        account_info = find_account(config)
    except FileNotFoundError as exc:
        if args.quiet:
            raise
        account_info = find_account_with_interactive_retry(config, Path(args.config), exc)
    key = acquire_db_key(account_info, config, args)
    if not re.fullmatch(rf"[0-9a-fA-F]{{{DB_KEY_HEX_LENGTH}}}", key):
        raise ValueError("数据库密钥必须是 64 位十六进制字符串")

    progress("解密数据库")
    account_dir = decrypt_databases(account_info, key)
    save_db_key(account_info, account_dir, key)

    # Contacts are loaded once: friend filtering and interaction name
    # resolution both rely on them, and re-reading the decrypted contact.db
    # is cheap enough to always do up front.
    contacts: list[ContactEntry] = load_contact_entries(account_dir)
    write_contact_cache(account_dir, contacts)
    contact_names: dict[str, str] = build_contact_display_names(contacts)
    contact_lookup: dict[str, ContactEntry] = {c.username: c for c in contacts if c.username}

    progress("准备图片")
    await save_image_keys(account_info.account, account_info.wxid_dir, account_info.db_storage_dir)

    output_root = Path(args.output_root or str(config.get("output_root") or "") or PROJECT_ROOT / "output").expanduser()

    # ---- Per-export loop: change time range / friends / format without re-decrypting ----
    use_cli_args = True

    async def run_one_export() -> bool:
        """执行一次完整导出流程。返回 True=再来一批，False=退出。

        任何输入处输入 r 会抛 RestartFlow；由外层循环捕获后从头再来，
        避免输错参数或中途报错导致整个程序（缓慢的解密初始化）重启。
        """
        nonlocal use_cli_args, account_dir, contacts, contact_names, contact_lookup
        # 每批开始检查源库是否有更新（启动后新浏览的朋友圈会写入微信源库），
        # 有更新则重新解密并刷新联系人；无更新则跳过（仅一次 mtime 扫描，秒过）。
        progress("检查数据库更新")
        if _source_dbs_newer_than_cached(account_info.db_storage_dir, account_dir):
            progress("检测到新数据，重新解密数据库")
            account_dir = decrypt_databases(account_info, key)
            contacts = load_contact_entries(account_dir)
            write_contact_cache(account_dir, contacts)
            contact_names = build_contact_display_names(contacts)
            contact_lookup = {c.username: c for c in contacts if c.username}
            print("  数据库已更新，本次导出将包含启动后新浏览的朋友圈", flush=True)
        # 1) 先选导出类型：默认导出 / 好友商家表格 / 朋友圈商品
        export_type = ask_export_type(args.export_type if use_cli_args else "")

        friend_inputs: list[dict[str, str]] = []
        friend_filter_mode = ""
        prefix = ""
        start_raw = end_raw = ""
        start = end = None
        scope = "friends"
        keep_interactions = False
        md_only = False
        export_excel = False
        clean_data = False
        export_contacts_enabled = False
        export_merchant = False
        skip_moments = False
        # 图片下载优化：默认开启，仅在朋友圈商品模式生效；CLI 参数可关闭
        skip_non_product = True
        skip_qrcode = True
        # 清洗第三阶段：图片内容检测（二维码+OCR手机号），默认开启
        image_inspect = True
        # 本次实际参与好友/商家导出的联系人（按筛选条件收窄）
        friend_contacts: list[ContactEntry] = list(contacts)

        if export_type == "friends":
            # ---- 好友/商家模式：先筛选好友，只出好友/商家表格，不碰朋友圈 ----
            friend_filter_mode = args.friend_filter_mode if args.friend_filter_mode else "prefix"
            if friend_filter_mode == "prefix":
                prefix = ask_input(args.friend_prefix if use_cli_args else "", "请输入备注关键词（例如 服装 或 库存联盟，回车=全部好友）", "")
            # 简化：好友/商家模式默认同时导出 contacts.csv 和商家导入表，不再交互询问
            export_contacts_enabled = parse_yes_no(args.export_contacts) if args.export_contacts else True
            export_merchant = parse_yes_no(args.export_merchant) if args.export_merchant else True
            skip_moments = True
            use_cli_args = False

            self_usernames = self_username_candidates(account_info, config)
            all_friend_usernames = [c.username for c in contacts]
            progress("准备好友列表")
            if friend_filter_mode == "prefix" and prefix.strip():
                friend_contacts = filter_contacts_by_prefix(contacts, prefix)
                picked_usernames = [c.username for c in friend_contacts if c.username]
                friend_inputs = [{"input": f"prefix:{prefix}", "usernames": ",".join(picked_usernames)}]
                print(f"  按关键词“{prefix}”匹配到 {len(friend_contacts)} 位好友", flush=True)
            elif friend_filter_mode == "manual":
                picked_usernames, friend_inputs = ask_friend_filters(account_info, account_dir, config, contacts)
                if picked_usernames is None:
                    friend_contacts = list(contacts)
                else:
                    chosen = set(picked_usernames)
                    friend_contacts = [c for c in contacts if c.username in chosen]
            else:
                friend_contacts = list(contacts)
                friend_inputs = [{"input": "friends:all", "usernames": ",".join(all_friend_usernames)}]
            usernames: list[str] | None = [c.username for c in friend_contacts if c.username]
        else:
            # ---- 朋友圈商品模式 ----
            if export_type == "default":
                # 默认导出：昨天 + 备注含“库存联盟” + 图片 + 清洗后表格，零交互
                yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
                start_raw = end_raw = yesterday
                scope = "friends"
                friend_filter_mode = "prefix"
                prefix = "库存联盟"
                keep_interactions = False
                md_only = True  # 默认导出只要图片+清洗后表格，不生成 PDF
                export_excel = True
                clean_data = True
                use_cli_args = False
                if args.image_inspect:
                    image_inspect = parse_yes_no(args.image_inspect)
                print(f"默认导出：{yesterday} 备注含“库存联盟”的朋友圈（图片+清洗后表格）", flush=True)
            else:
                start_raw = ask_input(args.start if use_cli_args else "", "起始时间（格式: 20260606 或 2026-06-06，回车=不限制）", "")
                end_raw = ask_input(args.end if use_cli_args else "", "终止时间（格式: 20260606 或 2026-06-06，回车=不限制）", "")
                scope_cli = resolve_scope_from_args(args, use_cli_args=use_cli_args)
                scope = scope_cli or "friends"  # 默认只导出好友朋友圈，CLI 可覆盖为 self/all

                if scope != "self":
                    # 简化：默认按前缀筛选，CLI 可覆盖为 all/manual
                    friend_filter_mode = args.friend_filter_mode if args.friend_filter_mode else "prefix"
                    if friend_filter_mode == "prefix":
                        prefix = ask_input(args.friend_prefix if use_cli_args else "", "请输入备注关键词（例如 服装 或 库存联盟，回车=全部好友）", "")

                # 简化：固定选项不再交互询问，CLI 参数可覆盖
                keep_interactions = parse_yes_no_default_no(args.keep_interactions)
                md_only = parse_yes_no_default_no(args.md_only)
                export_excel = parse_yes_no(args.export_excel) if args.export_excel else True
                clean_data = parse_yes_no_default_no(args.clean_data) if args.clean_data else True
                if args.skip_non_product:
                    skip_non_product = parse_yes_no(args.skip_non_product)
                if args.skip_qrcode:
                    skip_qrcode = parse_yes_no(args.skip_qrcode)
                if args.image_inspect:
                    image_inspect = parse_yes_no(args.image_inspect)
                use_cli_args = False

            start = parse_datetime(start_raw, end_of_day=False)
            end = parse_datetime(end_raw, end_of_day=True)

            self_usernames = self_username_candidates(account_info, config)
            all_friend_usernames = [c.username for c in contacts]

            if scope == "self":
                usernames = self_usernames
            elif scope == "friends":
                progress("准备好友列表")
                if friend_filter_mode == "all" or (friend_filter_mode == "prefix" and not prefix.strip()):
                    usernames = list(all_friend_usernames)
                    friend_inputs = [{"input": "friends:all", "usernames": ",".join(all_friend_usernames)}]
                elif friend_filter_mode == "prefix":
                    usernames, friend_inputs = resolve_friends_by_prefix(contacts, prefix)
                else:
                    usernames, friend_inputs = ask_friend_filters(account_info, account_dir, config, contacts)
                    # ask_friend_filters returns None when the user picks "所有人".
                    if usernames is None:
                        usernames = list(all_friend_usernames)
                        friend_inputs = [{"input": "friends:all", "usernames": ",".join(all_friend_usernames)}]
            else:  # all
                progress("准备好友列表")
                if friend_filter_mode == "all" or (friend_filter_mode == "prefix" and not prefix.strip()):
                    usernames = None  # no filter = self + all friends
                    friend_inputs = []
                elif friend_filter_mode == "prefix":
                    friend_usernames, friend_inputs = resolve_friends_by_prefix(contacts, prefix)
                    usernames = list(dict.fromkeys(self_usernames + friend_usernames))
                else:
                    friend_usernames, friend_inputs = ask_friend_filters(account_info, account_dir, config, contacts)
                    if friend_usernames is None:
                        usernames = None
                    else:
                        usernames = list(dict.fromkeys(self_usernames + friend_usernames))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output = output_root / timestamp
        if output.exists():
            shutil.rmtree(output)
        output.mkdir(parents=True, exist_ok=True)

        contact_export_meta: dict[str, Any] = {}
        if export_contacts_enabled:
            progress("导出好友列表")
            contact_export_meta = export_contact_list(output, friend_contacts)

        merchant_export_meta: dict[str, Any] = {}
        if export_merchant:
            progress("导出商家导入表格")
            # friend_contacts 已按筛选条件收窄；prefix 仅用于二次包含匹配（全选手动时为空=不再过滤）
            merchant_export_meta = export_merchant_list(output, friend_contacts, prefix, args.merchant_template)

        params = {
            "export_type": export_type,
            "start_time": start_raw,
            "end_time": end_raw,
            "scope": scope,
            "only_self": scope == "self",
            "keep_interactions": keep_interactions,
            "export_contacts": export_contacts_enabled,
            "md_only": md_only,
            "export_excel": export_excel,
            "clean_data": clean_data,
            "export_merchant": export_merchant,
            "skip_moments": skip_moments,
            "skip_non_product": skip_non_product,
            "skip_qrcode": skip_qrcode,
            "image_inspect": image_inspect,
            "friend_filter_mode": friend_filter_mode,
            "friend_prefix": prefix if friend_filter_mode == "prefix" else "",
            "friend_inputs": friend_inputs,
        }
        (output / "params.json").write_text(json.dumps(params, ensure_ascii=False, indent=2), encoding="utf-8")

        if skip_moments:
            progress("跳过朋友圈导出（仅好友/商家表格模式）")
            stats = {"posts_exported": 0, "images_downloaded": 0}
            exported_posts = []
        else:
            # 图片下载前询问：是否做图片内容清洗（识别手机号/二维码）
            if not args.image_inspect:
                image_inspect = parse_yes_no(
                    ask_input(
                        "",
                        "图片下载后是否清洗识别手机号（含手机号/二维码的图片将隔离到 quarantine，无图商品剔除，y/n，回车=y）",
                        "y",
                    )
                )
            if not image_inspect:
                print("  已关闭图片内容清洗：仅保留下载时的二维码快检", flush=True)
            progress("导出朋友圈")
            if scope == "self" and skip_non_product:
                print(
                    "  [提示] 当前导出自己朋友圈，非商品预判已开启：生活类动态图片将被跳过下载"
                    "（如需完整备份可加参数 --skip-non-product n）",
                    flush=True,
                )
            image_inspector = None
            # 仅在已开启非商品预判时做后台 OCR：此时下载的图片基本都会进入清洗，
            # 不会为最终被剔除的非商品图白跑 OCR；关闭预判时仍走清洗阶段检测。
            if clean_data and image_inspect and skip_non_product:
                try:
                    from tools.image_inspect import BackgroundInspector

                    image_inspector = BackgroundInspector(RUNTIME_DIR / "image_inspect_cache.json", output)
                except Exception as exc:
                    print(
                        f"  [warning] 后台图片检测初始化失败，改为清洗阶段检测: {exc}",
                        file=sys.stderr,
                        flush=True,
                    )
            async with httpx.AsyncClient(
                timeout=15.0,
                follow_redirects=False,
                limits=httpx.Limits(
                    max_connections=_DOWNLOAD_CONCURRENCY + 2,
                    max_keepalive_connections=_DOWNLOAD_CONCURRENCY,
                ),
            ) as download_client:
                try:
                    stats, exported_posts = await export_markdown(
                        account_info,
                        account_dir,
                        output,
                        start,
                        end,
                        usernames,
                        config,
                        keep_interactions,
                        contact_names,
                        contact_lookup,
                        allow_download=not args.no_download,
                        skip_non_product=skip_non_product,
                        skip_qrcode=skip_qrcode,
                        write_markdown=not clean_data,
                        client=download_client,
                        on_image_saved=image_inspector.submit if image_inspector is not None else None,
                    )
                finally:
                    if image_inspector is not None:
                        image_inspector.close()
            if md_only or clean_data:
                progress("跳过 PDF 生成（清洗模式仅保留图片+表格）")
            else:
                progress("生成 PDF")
                html_path = write_pdf_html(output, exported_posts)
                pdf_path = output / "moments.pdf"
                try:
                    render_pdf(output, exported_posts, pdf_path)
                except Exception as exc:
                    raise RuntimeError(f"PDF 生成失败: {exc}。HTML 已生成: {html_path}") from exc

        if export_excel and not skip_moments and not clean_data:
            progress("生成 Excel")
            try:
                from tools.moments_to_excel import load_records, write_xlsx
                records = load_records(output / "moments.json")
                write_xlsx(records, output / "moments.xlsx")
            except RuntimeError:
                print("  未安装 openpyxl，已跳过 Excel 生成（可手动运行：pip install openpyxl）", flush=True)
            except Exception as exc:
                print(f"  [warning] Excel 生成失败: {exc}", file=sys.stderr, flush=True)

        if clean_data and not skip_moments:
            progress("清洗数据")
            try:
                from tools.clean_moments import (
                    load_records,
                    clean_records,
                    write_xlsx as clean_write_xlsx,
                    apply_image_inspection,
                )
                records = load_records(output / "moments.json")
                kept, removed = clean_records(records)
                # 第三阶段：图片内容检测（二维码+OCR手机号），命中图隔离到 quarantine/，
                # 无图商品整条剔除；表格在其后生成，图片路径与磁盘文件始终一致
                if image_inspect:
                    progress("图片内容检测（二维码/手机号）")
                    kept, removed, img_stats = apply_image_inspection(
                        kept, removed, output, cache_path=RUNTIME_DIR / "image_inspect_cache.json"
                    )
                    if img_stats.get("inspected"):
                        print(
                            f"  图片检测: {img_stats['inspected']} 张，"
                            f"二维码 {img_stats['qrcode']} 张、含手机号 {img_stats['phone']} 张已隔离至 quarantine/，"
                            f"图片全部违规剔除 {img_stats['records_dropped']} 条",
                            flush=True,
                        )
                clean_write_xlsx(kept, output / "moments_clean.xlsx")
                print(f"  清洗结果: 原始 {len(records)} 条 → 保留 {len(kept)} 条，剔除 {len(removed)} 条", flush=True)
                # 只保留清洗后表格 + 其引用的图片：删除被剔除商品的已下载图片
                kept_images = {
                    str(img).replace("\\", "/")
                    for rec in kept
                    for img in (rec.get("images") or [])
                }
                figure_dir = output / "figure"
                removed_files = 0
                removed_bytes = 0
                if figure_dir.is_dir():
                    for f in sorted(figure_dir.rglob("*")):
                        if not f.is_file():
                            continue
                        if f.relative_to(output).as_posix() not in kept_images:
                            removed_bytes += f.stat().st_size
                            f.unlink()
                            removed_files += 1
                    # 清理删除后留下的空目录（自底向上）
                    for d in sorted(
                        (p for p in figure_dir.rglob("*") if p.is_dir()),
                        key=lambda p: len(p.parts), reverse=True,
                    ):
                        if not any(d.iterdir()):
                            d.rmdir()
                # moments.json 是清洗输入，用完即删（md/xlsx/pdf 清洗模式下已不生成，防御性清理）
                for name in (
                    "moments.md", "moments.json", "moments.xlsx",
                    "moments_clean.json", "moments_removed.json",
                    "moments.pdf", "moments.html", "_moments_pdf.html",
                ):
                    (output / name).unlink(missing_ok=True)
                print(
                    f"  已清理中间产物，删除冗余图片 {removed_files} 张"
                    f"（释放 {removed_bytes / 1024 / 1024:.1f} MB），"
                    f"仅保留清洗后表格与商品图",
                    flush=True,
                )
            except Exception as exc:
                print(f"  [warning] 数据清洗失败: {exc}", file=sys.stderr, flush=True)

        result = {"output": str(output), **params, **contact_export_meta, **merchant_export_meta, **stats}
        print(f"\n导出完成：{output}", flush=True)
        if not args.quiet:
            print(json.dumps(result, ensure_ascii=False, indent=2))

        # In quiet/automation mode only export once.
        if args.quiet:
            return False

        while True:
            again = input(
                "\n本次结束。y=再来一批（可改时间/筛选/类型，无需重新解密）；"
                "r=从头重置；q=退出（回车=q）: "
            ).strip().lower()
            if again in {"r", "reset"}:
                raise RestartFlow
            if not again or again in {"q", "quit", "exit", "n", "no"}:
                return False
            if again in {"y", "yes", "1", "true", "on"}:
                print("\n" + "=" * 60, flush=True)
                return True
            print("输入无效，请输入 y、r 或 q。", flush=True)

    # quiet 自动化模式：只跑一次，错误交给 main() 的外层处理
    if args.quiet:
        await run_one_export()
        return 0

    # 交互模式：出错/重置都不退出进程（解密和密钥准备很慢，避免重复初始化）
    while True:
        try:
            keep_going = await run_one_export()
        except RestartFlow:
            print("\n" + "=" * 60 + "\n已重置，回到第一步重新开始。\n" + "=" * 60, flush=True)
            continue
        except SystemExit:
            raise
        except KeyboardInterrupt:
            print("\n检测到 Ctrl+C，退出程序。", flush=True)
            break
        except (FileNotFoundError, RuntimeError, ValueError) as exc:
            print(f"\n本次导出失败（程序不会退出）：{exc}", file=sys.stderr, flush=True)
            if os.environ.get("WXMOMENTS_DEBUG_TRACEBACK", "").strip().lower() in {"1", "true", "yes", "on"}:
                traceback.print_exc()
            try:
                ans = input("按回车回到开头重试，输入 q 退出: ").strip().lower()
            except EOFError:
                break
            if ans in {"q", "quit", "exit"}:
                break
            print("\n" + "=" * 60, flush=True)
            continue
        except Exception as exc:  # noqa: BLE001 - 交互模式兜底，保证不闪退
            print(f"\n发生未预期的错误（程序不会退出）：{exc}", file=sys.stderr, flush=True)
            traceback.print_exc()
            try:
                ans = input("按回车回到开头重试，输入 q 退出: ").strip().lower()
            except EOFError:
                break
            if ans in {"q", "quit", "exit"}:
                break
            print("\n" + "=" * 60, flush=True)
            continue
        if not keep_going:
            break

    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
