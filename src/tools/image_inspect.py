#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""图片内容检测：二维码 + OCR 手机号。

供清洗流程第三阶段使用：对保留商品的已下载图片做内容检测，
命中（含二维码或手机号）的图片由调用方从记录中移除并隔离。

- 二维码判定与 wxmoments.detect_qrcode 同一标准：
  能解码出内容；或定位点四点近正方形（透视容差 25%）且面积占比 2%~60%（艺术码/点状码）
- 手机号判定：OCR 文本（去空格/横线后）匹配 1[3-9]\\d{9}
- OCR 前把长边缩到 _OCR_MAX_SIDE（默认 1280）：RapidOCR 检测默认只放大小图、
  不缩小大图，商品大图会按原分辨率推理；缩放后单图耗时约降一半，手机号仍可识别
- 2 路线程并行、共享同一份 OCR 引擎：多进程会让每进程的 ONNX Runtime 各自吃满线程，
  实测反而比串行更慢，故不使用多进程
- 结果按 (相对路径,大小) 缓存，重跑同一批不重复检测；内容完全相同的图片只检测一次
- 依赖缺失时优雅降级：无 opencv 跳过二维码检测，无 rapidocr 跳过手机号检测
"""

from __future__ import annotations

import hashlib
import json
import re
import sys
import threading
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path

# 前后不能再跟数字：避免 EAN-13 条码（69开头13位）等长数字串的 11 位子串误命中
PHONE_RE = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
# 北美电话格式（1-888-516-2631 / 1.733.654.2887）：去分隔符后正好撞手机号规则。
# 外贸原单包装常印此类厂家客服电话，属商品包装固有信息而非商家导流，先剔除再匹配手机号。
_NA_PHONE_RE = re.compile(r"1[\-\.\s]\d{3}[\-\.\s]\d{3}[\-\.\s]\d{4}")

# 模块级全局：检测引擎只初始化一次，整批复用
_CV2 = None
_NP = None
_OCR = None
_CLS_DISABLED = False  # 是否已关闭角度分类器（手机号只需正排文字，关闭省推理）
_INIT_DONE = False

# OCR 输入长边上限：RapidOCR 检测默认 limit_type=min，只会放大小图、不会缩小大图，
# 商品图动辄 2000px 以上，缩到 1280 后单图耗时约为原来的 1/2，手机号仍可识别。
_OCR_MAX_SIDE = 1280
# 共享单引擎的并行线程数：实测 2 路约 1.5x，再加线程收益递减且会与下载抢 CPU。
_OCR_THREADS = 2
# cv2.QRCodeDetector 非线程安全，按线程复用实例。
_QR_LOCAL = threading.local()


def _init_worker() -> None:
    """加载 cv2 / rapidocr，缺失时对应检测自动禁用。可在主进程或子进程调用。"""
    global _CV2, _NP, _OCR, _CLS_DISABLED, _INIT_DONE
    if _INIT_DONE:
        return
    _INIT_DONE = True
    try:
        import cv2
        import numpy as np

        _CV2, _NP = cv2, np
    except ImportError:
        _CV2 = _NP = None
    try:
        from rapidocr_onnxruntime import RapidOCR

        _OCR = RapidOCR()
    except Exception:
        _OCR = None
        return
    # 手机号检测只需正排文字：关闭角度分类器(cls)，省去每个文本框的分类推理。
    # 版本间该开关为实例属性（use_angle_cls），无此属性则保持默认。
    try:
        _OCR.use_angle_cls = False
        _CLS_DISABLED = True
    except Exception:
        _CLS_DISABLED = False


def capabilities() -> tuple[bool, bool]:
    """返回 (二维码检测可用, OCR手机号检测可用)，不初始化引擎（仅探测依赖）。"""
    qr_ok = ocr_ok = False
    try:
        import cv2  # noqa: F401
        import numpy  # noqa: F401

        qr_ok = True
    except ImportError:
        pass
    try:
        import rapidocr_onnxruntime  # noqa: F401

        ocr_ok = True
    except ImportError:
        pass
    return qr_ok, ocr_ok


def _is_qrcode(gray) -> bool:
    """与 wxmoments.detect_qrcode 同一判定标准（输入灰度图 ndarray）。"""
    cv2 = _CV2
    np = _NP
    h, w = gray.shape[:2]
    if max(h, w) > 800:
        scale = 800.0 / max(h, w)
        gray = cv2.resize(gray, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
    detector = getattr(_QR_LOCAL, "detector", None)
    if detector is None:
        detector = cv2.QRCodeDetector()
        _QR_LOCAL.detector = detector
    data, points, _straight = detector.detectAndDecode(gray)
    if data:
        return True
    if points is not None:
        p = points[0]
        d2 = [
            (p[i][0] - p[(i + 1) % 4][0]) ** 2 + (p[i][1] - p[(i + 1) % 4][1]) ** 2
            for i in range(4)
        ]
        if max(d2) <= 1.5625 * (min(d2) + 1e-6):
            h2, w2 = gray.shape[:2]
            frac = cv2.contourArea(p.astype(np.float32)) / (h2 * w2)
            if 0.02 <= frac <= 0.6:
                return True
    ok, decoded, _pts, _st = detector.detectAndDecodeMulti(gray)
    return bool(ok and any(decoded))


def _find_phones(texts: list[str]) -> list[str]:
    """从 OCR 文本块中找手机号：逐块 + 相邻两块拼接（应对同处换行排版的号码）。

    不做全文拼接：无关块的数字串会误拼出 11 位号码（实测外贸包装图误报）。
    """
    hits: list[str] = []
    blocks = [str(t or "") for t in texts if str(t or "").strip()]
    candidates = list(blocks)
    for i in range(len(blocks) - 1):
        candidates.append(blocks[i] + blocks[i + 1])
    for text in candidates:
        cleaned = _NA_PHONE_RE.sub("", text)
        compact = re.sub(r"[\s\-]", "", cleaned)
        hits.extend(PHONE_RE.findall(compact))
    return sorted(set(hits))


def _inspect_one(path_str: str) -> tuple[str, bool, list[str], str]:
    """检测单张图片，返回 (路径, 是否二维码, 手机号列表, OCR文本样例)。

    OpenCV 分配失败（OutOfMemoryError）时跳过该图返回空结果，不让整批中断。
    """
    _init_worker()
    is_qr = False
    phones: list[str] = []
    sample = ""
    try:
        data = Path(path_str).read_bytes()
    except OSError:
        return path_str, False, [], ""
    img = None
    if _CV2 is not None:
        try:
            arr = _NP.frombuffer(data, _NP.uint8)
            img = _CV2.imdecode(arr, _CV2.IMREAD_COLOR)
            if img is not None:
                try:
                    is_qr = _is_qrcode(_CV2.cvtColor(img, _CV2.COLOR_BGR2GRAY))
                except Exception:
                    is_qr = False
        except Exception:
            # OpenCV OutOfMemoryError / 解码失败，跳过该图
            return path_str, False, [], ""
    if _OCR is not None and img is not None and not is_qr:
        # 二维码图一定会被隔离，无需再做 OCR，省一次完整识别开销
        try:
            h, w = img.shape[:2]
            if max(h, w) > _OCR_MAX_SIDE:
                scale = _OCR_MAX_SIDE / max(h, w)
                img = _CV2.resize(
                    img,
                    (max(1, int(round(w * scale))), max(1, int(round(h * scale)))),
                    interpolation=_CV2.INTER_AREA,
                )
            result, _elapse = _OCR(img)
        except Exception:
            result = None
        if result:
            texts = [str(r[1]) for r in result if len(r) >= 2 and r[1]]
            phones = _find_phones(texts)
            if phones:
                sample = " ".join(texts)[:80]
    return path_str, is_qr, phones, sample


def _load_cache(cache_path: Path | None) -> dict:
    if cache_path and cache_path.exists():
        try:
            data = json.loads(cache_path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except Exception:
            pass
    return {}


def _save_cache(cache_path: Path | None, cache: dict) -> None:
    if not cache_path:
        return
    try:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def _cache_key(path: Path, root: Path | None = None) -> str:
    """缓存键：相对导出目录的路径 + 文件大小。

    图片目录按朋友圈发布时间命名，重导同一批时相对路径不变，
    因此缓存可跨批次命中；只用大小不读内容，判定开销几乎为零。
    """
    try:
        size = path.stat().st_size
    except OSError:
        return f"{path.resolve().as_posix()}|missing"
    rel = ""
    if root is not None:
        try:
            rel = path.resolve().relative_to(root).as_posix()
        except ValueError:
            rel = ""
    return f"{rel or path.resolve().as_posix()}|{size}"


def _content_digest(path: Path) -> str:
    """分块计算 md5：识别内容完全相同的图片（同一张图被多位商家转发）。"""
    try:
        digest = hashlib.md5()
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
        return digest.hexdigest()
    except OSError:
        return ""


class BackgroundInspector:
    """导出阶段后台并行 OCR，结果写入与 inspect_images 相同的缓存。

    每落盘一张商品图就 submit 一次，OCR 与后续下载重叠；导出结束 close() 等待
    剩余任务并落盘。之后清洗阶段调用 inspect_images 时直接命中缓存，不再重复识别。
    """

    def __init__(self, cache_path: Path | None, root: Path | None, workers: int = _OCR_THREADS) -> None:
        self._cache_path = cache_path
        self._root = root.resolve() if root is not None else None
        self._executor = ThreadPoolExecutor(max_workers=max(1, workers), thread_name_prefix="img-prefetch")
        self._lock = threading.Lock()
        self._cache = _load_cache(cache_path)
        self._submitted: set[str] = set()
        self._futures: list[Future] = []
        _init_worker()  # 主线程先加载引擎，避免多个工作线程同时初始化

    def submit(self, path: Path) -> None:
        p = Path(path)
        key = _cache_key(p, self._root)
        with self._lock:
            if key in self._cache or key in self._submitted:
                return
            self._submitted.add(key)
        future = self._executor.submit(_inspect_one, str(p))
        future.add_done_callback(lambda fut, k=key: self._on_done(k, fut))
        self._futures.append(future)

    def _on_done(self, key: str, future: Future) -> None:
        try:
            result = future.result()
        except Exception:
            return
        with self._lock:
            self._cache[key] = [result[1], result[2], result[3]]

    def close(self) -> None:
        self._executor.shutdown(wait=True)
        with self._lock:
            _save_cache(self._cache_path, self._cache)
        self._futures.clear()


def inspect_images(
    paths: list[Path],
    cache_path: Path | None = None,
    root: Path | None = None,
) -> dict[str, dict]:
    """批量检测图片，返回 {路径字符串: {"qrcode": bool, "phones": [...], "sample": str}}。

    root 用于生成跨批次稳定的缓存键（相对路径+文件大小），缺省退回绝对路径。
    同进程 _OCR_THREADS 路并行、共享一份 OCR 引擎；多进程会各自加载模型并让
    每进程的 ONNX Runtime 各自吃满线程，实测比串行更慢，故不采用。
    """
    uniq: list[Path] = []
    seen: set[str] = set()
    for p in paths:
        key = str(p)
        if key not in seen:
            seen.add(key)
            uniq.append(p)
    if not uniq:
        return {}

    qr_ok, ocr_ok = capabilities()
    if not qr_ok and not ocr_ok:
        print("  [warning] 未安装 opencv/rapidocr，图片内容检测整体跳过", file=sys.stderr, flush=True)
        return {}
    if not ocr_ok:
        print("  [warning] 未安装 rapidocr-onnxruntime，手机号检测禁用（仅检测二维码）", file=sys.stderr, flush=True)

    root = root.resolve() if root is not None else None
    cache = _load_cache(cache_path)
    todo: list[tuple[Path, str]] = []
    results: dict[str, dict] = {}
    for p in uniq:
        ck = _cache_key(p, root)
        hit = cache.get(ck)
        if hit is not None:
            results[str(p)] = {"qrcode": bool(hit[0]), "phones": list(hit[1]), "sample": hit[2] if len(hit) > 2 else ""}
        else:
            todo.append((p, ck))

    if not todo:
        return results

    # 内容去重：只有大小相同的图才计算 md5（哈希开销可忽略），相同内容的图只检测一次
    by_size: dict[int, list[tuple[Path, str]]] = {}
    for p, ck in todo:
        try:
            size = p.stat().st_size
        except OSError:
            size = -1
        by_size.setdefault(size, []).append((p, ck))
    detect: list[tuple[Path, str]] = []
    clones: dict[str, list[tuple[Path, str]]] = {}
    for size, items in by_size.items():
        if size < 0 or len(items) == 1:
            detect.extend(items)
            continue
        digests: dict[str, tuple[Path, str]] = {}
        for p, ck in items:
            digest = _content_digest(p)
            if not digest:
                detect.append((p, ck))
            elif digest in digests:
                clones.setdefault(str(digests[digest][0]), []).append((p, ck))
            else:
                digests[digest] = (p, ck)
                detect.append((p, ck))

    total = len(detect)
    duplicated = len(todo) - total
    print(
        f"  图片内容检测: {total} 张待检测（缓存命中 {len(results)} 张"
        + (f"，内容重复 {duplicated} 张" if duplicated else "")
        + f"），{_OCR_THREADS} 路并行",
        flush=True,
    )

    def _absorb(path_str: str, is_qr: bool, phones: list[str], sample: str, ck: str) -> None:
        results[path_str] = {"qrcode": is_qr, "phones": phones, "sample": sample}
        cache[ck] = [is_qr, phones, sample]
        # 同内容的其他路径复用本次结果，不再重复检测
        for clone_path, clone_ck in clones.get(path_str, ()):
            cache[clone_ck] = [is_qr, phones, sample]
            results[str(clone_path)] = {
                "qrcode": is_qr,
                "phones": list(phones),
                "sample": sample,
            }

    _init_worker()

    def _run(item: tuple[Path, str]) -> tuple[str, tuple]:
        p, ck = item
        return ck, _inspect_one(str(p))

    with ThreadPoolExecutor(max_workers=_OCR_THREADS, thread_name_prefix="img-inspect") as executor:
        for i, (ck, r) in enumerate(executor.map(_run, detect), 1):
            _absorb(r[0], r[1], r[2], r[3], ck)
            if i % 50 == 0 or i == total:
                print(f"    检测进度 {i}/{total}", flush=True)
    _save_cache(cache_path, cache)
    return results
