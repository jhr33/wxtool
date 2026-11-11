#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""朋友圈数据清洗脚本。

从 moments.json 中筛选出「商品类」朋友圈，剔除：
1. 没有图片的记录
2. 内容不像商品信息的记录（无数量/尺码/价格/颜色等商品特征）

用法:
    python src/tools/clean_moments.py <导出目录>
    python src/tools/clean_moments.py output/20260918_172020

输出:
    - moments_clean.json  (清洗后的 JSON)
    - moments_clean.xlsx  (清洗后的 Excel，与 moments.xlsx 列结构一致)
"""

import json
import re
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# 商品分类正则（按内容文字匹配）
# 优先级：内容正则 > 备注前缀 > 剔除
# ---------------------------------------------------------------------------
CATEGORY_PATTERNS: dict[str, re.Pattern] = {
    "服装": re.compile(
        r"(?:(?<!洗)(?<!干)衣|服|衫|裤|裙|袖|恤|羽绒|大衣|风衣|外套|夹克|卫衣|毛衣|毛衫|棉衣|棉服|牛仔|打底|衬|连衣裙|半身裙|针织|开衫|马甲|吊带|背心|套头|梭织|冲锋衣|阔腿裤|哈伦裤|小香风|蕾丝|呢子|羊毛|羊绒|皮草|貂|鹅|羽绒服)"
    ),
    "鞋子": re.compile(
        r"(?:鞋|靴|拖鞋|凉鞋|单鞋|板鞋|运动鞋|高跟鞋|老爹鞋|皮鞋|帆布鞋|休闲鞋|童鞋|男鞋|女鞋|鞋面|鞋底|鞋跟|鞋垫|马丁靴|雪地靴|豆豆鞋|切西尼|切尔西)"
    ),
    "家纺": re.compile(
        r"(?:床|被|枕|毯|窗帘|席|床单|被套|枕套|床垫|四件套|被芯|床笠|遮光|沙发套|床裙|地毯|毛毯|乳胶枕|鹅绒被|蚕丝被|棉花被|夏凉被|冬被|空调被|床品|布艺|沙发布|桌布|台布|餐垫|地垫|门垫|脚垫|枕芯|抱枕|靠垫|坐垫|椅套)"
    ),
    "百货": re.compile(
        r"(?:杯|壶|瓶|锅|碗|筷|勺|刀|黑板|白板|菜板|砧板|搓衣板|烫衣板|熨衣板|毛巾|纸|洗洁精|洗衣液|垃圾袋|保鲜膜|收纳|拖把|抹布|保温杯|牙刷|梳子|镜子|台灯|伞|包包|帽|袜|围巾|手套|口罩|耳罩|围裙|雨伞|袋子|置物架|架|夹|密封罐|保鲜盒|便当盒|饭盒|杯|牙缸|皂|框|洗发水|沐浴露|护肤品|面霜|乳液|口红|粉底|眉笔|眼影|碟|卸妆|防晒霜|面膜|膏|发圈|发夹|皮筋|耳环|项链|手链|戒指|胸针|挂件|吊坠|手镯|手环|手表|充电宝|数据线|手机壳|钢化膜|耳机|蓝牙|插座|电池|香薰|花瓶|装饰|摆件|收纳箱|收纳袋|密封袋|食品袋|垃圾筒|垃圾桶|洗衣机|冰箱|冰柜|空调|电视|电风扇|风扇|电饭煲|电饭锅|电水壶|微波炉|电磁炉|电陶炉|洗碗机|烘干机|干衣机|吹风机|电吹风|吸尘器|扫地机|扫地机器人|空气炸锅|烤箱|净水器|热水器|加湿器|除湿机|取暖器|电暖器|消毒柜|榨汁机|豆浆机|破壁机|料理机|绞肉机|搅拌机|电熨斗|挂烫机|燃气灶|油烟机|抽油烟机|咖啡机|蒸蛋器|电饼铛)"
    ),
}

# 备注前缀到分类的映射
REMARK_PREFIX_MAP = {
    "服装": "服装",
    "百货": "百货",
    "鞋子": "鞋子",
    "家纺": "家纺",
}

# 分类到 ID 的映射
CATEGORY_ID_MAP = {
    "服装": "23",
    "百货": "21",
    "鞋子": "26",
    "家纺": "27",
}

# ---------------------------------------------------------------------------
# 品牌词过滤：正文出现品牌名（耐克/阿迪达斯等）整条剔除
# 中文词直接匹配；英文词/缩写不区分大小写并要求英文边界
# 注意：有歧义的常用词（北面/冠军/椰子/飞跃/回力等）不收录，避免误删
# ---------------------------------------------------------------------------
_BRANDS_CN = [
    # 运动品牌（2026-11 放宽：库存尾货常见的斐乐/安踏/李宁/新百伦/纽巴伦/匡威/
    # 斯凯奇/安德玛/迪桑特/特步/鸿星尔克/乔丹/万斯/卡帕/美津浓/茵宝/鬼冢虎/
    # 亚瑟士/添柏岚已移出过滤，含品牌名的商品记录正常保留）
    "耐克", "阿迪达斯", "阿迪", "彪马", "始祖鸟", "蒙口", "加拿大鹅",
    # 奢侈/轻奢品牌
    "路易威登", "古驰", "香奈儿", "爱马仕", "普拉达", "巴宝莉", "博柏利",
    "蔻驰", "迪奥", "巴黎世家", "芬迪", "罗意威", "赛琳", "纪梵希", "阿玛尼", "范思哲",
    # 快时尚/户外/其他知名品牌（波司登/拉夫劳伦/李维斯/哥伦比亚/狼爪/凯乐石已移出）
    "优衣库",
]
_BRANDS_EN = [
    # 运动（fila/anta/li-ning/new balance/converse/vans/skechers/under armour/
    # descente/xtep/jordan/kappa/mizuno/umbro/onitsuka/asics/timberland 已移出）
    "nike", "adidas", "puma", "yeezy", "the north face", "arc'teryx", "moncler",
    # 奢侈/轻奢
    "louis vuitton", "gucci", "chanel", "hermes", "prada", "burberry", "coach",
    "dior", "balenciaga", "fendi", "loewe", "celine", "givenchy", "armani",
    "versace",
    # 快时尚/户外/其他（muji/ralph lauren/tommy hilfiger/levis/levi's/columbia/
    # jack wolfskin/lacoste 已移出）
    "uniqlo", "zara", "h&m", "canada goose", "calvin klein",
]
# 缩写（AJ/LV/CK/MK/YSL/MLB/TNF）：要求英文边界，避免粘连误匹配
_BRANDS_CODE = ["AJ", "LV", "CK", "MK", "YSL", "MLB", "TNF"]

BRAND_RE = re.compile(
    "|".join(_BRANDS_CN)
    + r"|(?<![A-Za-z])(?:"
    + "|".join(re.escape(w) for w in (_BRANDS_EN + _BRANDS_CODE))
    + r")(?![A-Za-z])",
    re.IGNORECASE,
)


def find_brand(content: str) -> str:
    """检查正文是否含品牌名，返回命中的品牌词，未命中返回空串"""
    match = BRAND_RE.search(str(content or ""))
    return match.group(0) if match else ""


# ---------------------------------------------------------------------------
# 复合词消歧：由「分类关键词 + 其他分类关键词」组成的复合词，整词归属如下分类。
# 分类匹配前先把复合词替换为占位符，避免内部单字被其他分类正则误命中
# （如"衣架"=衣(服装)+架(百货)，整词应归百货；"口袋"=口+袋(百货)，服装语境应归服装）
# ---------------------------------------------------------------------------
COMPOUND_WORDS: dict[str, tuple[str, ...]] = {
    # 注意：长词必须排在含其子串的短词之前（如 衣帽架 > 帽架、拖鞋架 > 鞋架）
    "百货": (
        "升降晾衣架", "晾衣架", "晾衣绳", "晾衣夹", "晾衣杆",
        "衣帽架", "衣帽钩", "衣撑", "衣刷", "衣钩",
        "烫衣板", "熨衣板", "搓衣板", "穿衣镜", "内衣皂",
        "围裙", "裤夹", "裤架", "裙架", "衣柜", "帽架",
        "拖鞋架", "鞋架", "鞋盒", "鞋油", "鞋刷",
        "床头柜", "床头灯", "床架",
    ),
    "服装": ("口袋",),
}


def _mask_compound_words(text: str) -> tuple[str, list[str]]:
    """把内容中的已知复合词替换为占位符，返回 (替换后文本, 命中的分类列表)。

    占位符不含任何分类关键词，防止复合词内部的字被单字正则跨分类误命中。
    """
    cats: list[str] = []
    for category, words in COMPOUND_WORDS.items():
        for word in words:
            if word in text:
                text = text.replace(word, "〇")
                if category not in cats:
                    cats.append(category)
    return text, cats


def classify_by_content_all(content: str) -> list[str]:
    """返回内容命中的所有分类（按固定优先级顺序排列），用于歧义判断。

    复合词按整词归属其真实分类，其余文本正常匹配四大分类正则。
    """
    text = str(content or "")
    if not text.strip():
        return []
    masked, compound_cats = _mask_compound_words(text)
    hits = [c for c in ("服装", "鞋子", "家纺", "百货") if CATEGORY_PATTERNS[c].search(masked)]
    for c in compound_cats:
        if c not in hits:
            hits.append(c)
    return [c for c in ("服装", "鞋子", "家纺", "百货") if c in hits]


def classify_by_content(content: str) -> str:
    """根据朋友圈文字内容匹配商品分类，返回分类名或空字符串"""
    matches = classify_by_content_all(content)
    return matches[0] if matches else ""


def classify_by_remark(remark: str) -> str:
    """根据备注中的分类词匹配商品分类，返回分类名或空字符串。

    备注格式如「新账号 百货 库存联盟 13800000000」，分类词可能在任意位置，
    故用包含匹配而非前缀匹配。
    """
    remark = str(remark or "").strip()
    for prefix, category in REMARK_PREFIX_MAP.items():
        if prefix in remark:
            return category
    return ""


def classify_product(record: dict) -> str:
    """综合分类：内容唯一命中直接用；内容命中多个分类（关键词歧义）时
    商家备注分类优先，无备注再按固定优先级取第一个；内容无命中时按备注兜底。"""
    content = str(record.get("content") or "")
    remark = str(record.get("remark") or "")

    matches = classify_by_content_all(content)
    if len(matches) == 1:
        return matches[0]

    remark_category = classify_by_remark(remark)
    if len(matches) > 1:
        # 关键词跨分类重复（如“毛衣+保温杯”），以商家备注分类为准
        return remark_category or matches[0]

    return remark_category


# ---------------------------------------------------------------------------
# 商品特征正则
# ---------------------------------------------------------------------------

# 1a) 数量 + 单位（强信号）
#    匹配：1500件、450件、5600套、4万多条、一包500条、2000pcs、一千条、大货530件
QUANTITY_RE = re.compile(
    r"\d+\s*(?:多)?\s*(?:万|千|百)?\s*"
    r"(?:件|条|套|pcs|包|对|双|盒|箱|组|款|扎|卷|匹|令|提|打|码|吨)"
    r"|[一二两三四五六七八九十百千万]+"
    r"(?:件|条|套|包|对|双|盒|箱|组|款|扎|卷|匹|令|提|吨)",
    re.IGNORECASE,
)

# 1b) 大额纯数字（强信号）—— 不要求量词，只要有 >=100 的整数或中文大额数即可
#    匹配：1500、2000、5600、4万、几千、一万、一千
#    不匹配：9.5（价格）、83（尺寸）、6.8（价格）等小额数字
LARGE_NUMBER_RE = re.compile(
    r"\d{3,}"  # 3位以上整数（>=100），小数点会自然断开不会误匹配
    r"|\d+\s*(?:万|千|百)"  # 4万、5千、3百
    r"|[一二三四五六七八九十]+\s*(?:万|千|百)"  # 一万、几千、几百
)

# 2) 尺码 / 规格尺寸（强信号）
#    匹配：尺码4-6-8-10、MLXL码、均码、SMLXL码、尺寸83*34、码数
SIZE_RE = re.compile(
    r"(?:尺码|码数|size|尺寸)\s*[\d\-*×xXSMLR]+"
    r"|(?:均码|均一码|统一码)"
    r"|[SMLXL]{2,5}\s*码"
    r"|(?:尺码|码数)齐全|码全"
    r"|[SMLX]{1,2}\s*[-~至]\s*\d{0,2}\s*[SMLX]{1,2}L?"
    r"|\d+\s*[-*×xX]\s*\d+(?:\s*[-*×xX]\s*\d+)*",
    re.IGNORECASE,
)

# 3) 价格（强信号）
#    匹配：9.5元、6毛8分、全清6.8元、价格便宜5.5一套、5.5一套、6毛8分
PRICE_RE = re.compile(
    r"\d+(?:\.\d+)?\s*(?:元|块|毛|分)(?:\s*(?:一(?:套|包|件|条))?)?"
    r"|(?:全清|清货|出货|特价|秒杀)\s*\d+(?:\.\d+)?"
    r"|价格\s*(?:便宜|美丽|实惠|低)?\s*\d*(?:\.\d+)?"
    r"|\d+(?:\.\d+)?\s*(?:元|块)?\s*/?\s*(?:一(?:套|包|件|条|双|对))",
    re.IGNORECASE,
)

# 4) 颜色（弱信号）
#    匹配：二个色、多色、单色、双拼色、色号
COLOR_RE = re.compile(r"(?:[一二三四五六七八九十两二三四五6-9]?(?:个)?色|多色|单色|双?拼色|色号|几?个?色)")

# 5) 营销/状态词（弱信号）
#    匹配：新款、爆款、热卖、上新、大货、全清、清仓、清货、出货、特价、秒杀、回购、现货、到货、库存、批发、统清、散批
#    注意：这类词只描述状态，不指向具体商品（如正文仅“到货”“现货”），
#          不能单独作为商品判定依据，必须配合数量/价格/尺码/物品名称等实体信号
KEYWORD_RE = re.compile(
    r"(?:新款|爆款|热卖|上新|大货|全清|清仓|清货|出货|特价|秒杀|回购|现货|到货|库存|批发|统清|散批|秒清|拿货|配货|补货|抢完|截单)"
)

# 6) 克重 / 重量规格（弱信号）
WEIGHT_RE = re.compile(r"\d+\s*(?:克|g|公斤|斤|kg|磅|lb)", re.IGNORECASE)

# 7) 批发 / 贸易关键词（弱信号）
TRADE_RE = re.compile(r"(?:批发|贸易|厂家|工厂|货源|出货|清仓|尾货|尾单|外贸|原单|库存|现货)")

# 8) 款式 / 型号 / 货号（强信号）
#    匹配：款式多样、型号X200、货号8836、款号、版型、材质、面料
MODEL_RE = re.compile(r"(?:款式|型号|货号|款号|编号|款型|版型|材质|面料)")

# 9) 物品名称（弱信号）—— 手套、框、杯这类具体商品名词，不能单独作为商品依据
#    基础词库复用四大分类正则（自动保持同步），另补充分类词表外的常见物品名
GOODS_NAME_RE = re.compile(
    "|".join(p.pattern for p in CATEGORY_PATTERNS.values())
    + r"|(?:相框|镜框|眼镜框|收纳框|塑料框|筐|眼镜|玩具|文具|饰品|箱包|皮具|童车|健身器材"
    + r"|化妆刷|笔芯|中性笔|球拍|帐篷|丝巾|串灯|彩灯|修剪器)"
)

# 收货/收购类帖子过滤：正文含收购、收货、求购等词的整条剔除
# 只保留正向出售/批发的商品，排除逆向收购需求
# 排除 #收购库存尾货联系方式 这类引流标签（该标签出现在出售帖中，非收购帖）
# 完整词：收购/收货/收到货/求购/回收；带前缀的“收”：大量收/高价收/长期收/上门收/现金收/急收/包收/低价收/全网收/厂家收
COLLECT_RE = re.compile(
    r"(?:收购|收货|收到货|求购|回收|包收|急收|全网收|低价收|厂家收|现金收|高价收|长期收|上门收|大量收|已出|已清)"
)

def is_collect_post(content: str) -> bool:
    """检查正文是否为收购/收货/求购类帖子"""
    text = str(content or "")
    if "#收购库存尾货联系方式" in text:
        return False
    return bool(COLLECT_RE.search(text))


# 借贷/金融类广告过滤：正文含贷款、抵押、急用钱等词的整条剔除
# 此类广告常夹带品类词伪装成商品帖（如“服装 抵押”“急用钱...服装 抵押 小马”）
FINANCE_RE = re.compile(
    r"(?:贷款|借款|放款|借钱|急用钱|车贷|房贷|信贷|黑户|上征信|抵押"
    r"|资金周转|短期周转|周转资金)"
)


def is_finance_ad(content: str) -> bool:
    """检查正文是否为借贷/金融类广告"""
    return bool(FINANCE_RE.search(str(content or "")))


def is_product_record(record: dict) -> tuple[bool, list[str], str]:
    """判断一条朋友圈是否为商品数据并分类。

    两步走：
    1. 先判断是否商品 —— 需出现商品特征信号：
       大额数字/数量、款式型号尺码、价格，或手套/框/杯等物品名称；
    2. 是商品后再分类 —— 内容正则优先，无法判断时按商家备注中的分类兜底；
       两者都无法判断则剔除。

    返回 (是否商品, 命中的信号列表, 分类名称)。
    """
    content = str(record.get("content") or "")
    images = record.get("images") or []
    image_count = int(record.get("image_count") or len(images) if isinstance(images, list) else 0)

    # 硬条件：必须要有图片
    if image_count <= 0 or not images:
        return False, ["无图片"], ""

    # 硬条件：内容不能为空
    if not content.strip():
        return False, ["内容为空"], ""

    # 硬条件：正文出现品牌名（耐克/阿迪达斯等）整条剔除
    brand = find_brand(content)
    if brand:
        return False, [f"品牌:{brand}"], ""

    # 硬条件：收购/收货类帖子剔除
    if is_collect_post(content):
        return False, ["收购/收货"], ""

    # 硬条件：借贷/金融类广告剔除（常夹带品类词，如“服装 抵押”）
    if is_finance_ad(content):
        return False, ["借贷广告"], ""

    # 第一步：判断是否商品（需命中至少一个商品特征信号）
    signals: list[str] = []
    if QUANTITY_RE.search(content):
        signals.append("数量")
    if LARGE_NUMBER_RE.search(content):
        signals.append("大额数字")
    if SIZE_RE.search(content):
        signals.append("尺码")
    if MODEL_RE.search(content):
        signals.append("款式型号")
    if PRICE_RE.search(content):
        signals.append("价格")
    if GOODS_NAME_RE.search(content):
        signals.append("物品名称")
    if KEYWORD_RE.search(content):
        signals.append("商品词")
    if not signals:
        return False, ["无商品特征"], ""

    # 物品名称是弱信号，单独命中不能证明是商品，需搭配其他商品信号
    if signals == ["物品名称"]:
        return False, ["仅物品名称无其他商品特征"], ""

    # 仅有营销/状态词（如正文只有“到货”“现货”）不构成商品信息，剔除
    if signals == ["商品词"]:
        return False, ["仅营销词无商品信息"], ""

    # 第二步：分类（内容正则 > 商家备注前缀）
    category = classify_product(record)
    if not category:
        return False, ["无法分类"], ""

    # 补充统计用弱信号（不影响判定结果）
    if COLOR_RE.search(content):
        signals.append("颜色")
    if WEIGHT_RE.search(content):
        signals.append("克重")
    if TRADE_RE.search(content):
        signals.append("贸易词")

    return True, signals, category


def load_records(json_path: Path) -> list[dict]:
    raw = json_path.read_text(encoding="utf-8-sig")
    records = json.loads(raw)
    if not isinstance(records, list):
        records = [records] if records else []
    return records


# ---------------------------------------------------------------------------
# 第三阶段：图片内容清洗（二维码 + OCR 手机号）
# 仅对文本判定+分类后保留的记录执行：命中图片从记录移除并隔离到 quarantine/，
# 图片被清空的记录整条剔除。表格在此之后才生成，天然避免死链与无图商品。
# ---------------------------------------------------------------------------


def apply_image_inspection(
    kept: list[dict],
    removed: list[dict],
    output_dir: Path,
    cache_path: Path | None = None,
) -> tuple[list[dict], list[dict], dict]:
    """对保留记录的已下载图片做内容检测，返回 (kept, removed, stats)。"""
    try:
        from tools.image_inspect import inspect_images
    except ImportError:
        sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
        from tools.image_inspect import inspect_images

    stats = {"inspected": 0, "qrcode": 0, "phone": 0, "records_dropped": 0}
    all_paths: list[Path] = []
    for rec in kept:
        for img in rec.get("images") or []:
            p = output_dir / str(img)
            if p.is_file():
                all_paths.append(p)
    if not all_paths:
        return kept, removed, stats

    results = inspect_images(all_paths, cache_path=cache_path, root=output_dir)
    stats["inspected"] = len(results)

    # 汇总命中图片 → {相对路径: 原因}
    bad: dict[str, str] = {}
    for p_str, info in results.items():
        reasons: list[str] = []
        if info.get("qrcode"):
            reasons.append("二维码")
            stats["qrcode"] += 1
        phones = info.get("phones") or []
        if phones:
            reasons.append("手机号:" + "/".join(phones))
            stats["phone"] += 1
        if reasons:
            rel = Path(p_str).relative_to(output_dir).as_posix()
            bad[rel] = ",".join(reasons)
    if not bad:
        return kept, removed, stats

    # 命中图片移入 quarantine/（保留目录结构，可人工复核；figure/ 内即干净）
    quarantine = output_dir / "quarantine"
    violations: list[dict] = []
    new_kept: list[dict] = []
    for rec in kept:
        images = [str(i).replace("\\", "/") for i in (rec.get("images") or [])]
        survivors = [img for img in images if img not in bad]
        for img in images:
            if img in survivors:
                continue
            src = output_dir / img
            dst = quarantine / img
            try:
                dst.parent.mkdir(parents=True, exist_ok=True)
                src.replace(dst)
            except OSError:
                pass
            violations.append(
                {
                    "image": img,
                    "reason": bad[img],
                    "uid": str(rec.get("uid") or ""),
                    "time_text": str(rec.get("time_text") or ""),
                    "content": str(rec.get("content") or "")[:40],
                }
            )
        if survivors:
            rec["images"] = survivors
            rec["image_count"] = len(survivors)
            new_kept.append(rec)
        else:
            rec["_remove_reason"] = "图片全部违规"
            removed.append(rec)
            stats["records_dropped"] += 1
    if violations:
        quarantine.mkdir(parents=True, exist_ok=True)
        (quarantine / "violations.json").write_text(
            json.dumps(violations, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    return new_kept, removed, stats


def clean_records(records: list[dict]) -> tuple[list[dict], list[dict]]:
    kept: list[dict] = []
    removed: list[dict] = []
    seen: set[str] = set()  # (uid + content) 去重
    for record in records:
        is_product, reasons, category = is_product_record(record)
        if not is_product:
            record_copy = dict(record)
            record_copy["_remove_reason"] = ",".join(reasons)
            removed.append(record_copy)
            continue

        # 去重：同商家(uid) + 相同文字内容只保留一条
        uid = str(record.get("uid") or record.get("username") or "")
        content = str(record.get("content") or "").strip()
        dedup_key = f"{uid}\x00{content}"
        if dedup_key in seen:
            record_copy = dict(record)
            record_copy["_remove_reason"] = "重复商品"
            removed.append(record_copy)
            continue
        seen.add(dedup_key)

        record_copy = dict(record)
        record_copy["_category"] = category
        record_copy["_category_id"] = CATEGORY_ID_MAP.get(category, "")
        kept.append(record_copy)
    return kept, removed


def write_xlsx(records: list[dict], out_path: Path) -> None:
    # moments_to_excel.py 位于 src/tools/ 目录下，与本脚本同级
    moments_to_excel_path = Path(__file__).parent / "moments_to_excel.py"
    try:
        from tools.moments_to_excel import COLUMNS, row_values, chunk_records, chunked_path, MAX_ROWS_PER_FILE
    except ImportError:
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            "moments_to_excel", moments_to_excel_path
        )
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        COLUMNS, row_values = mod.COLUMNS, mod.row_values
        chunk_records, chunked_path, MAX_ROWS_PER_FILE = mod.chunk_records, mod.chunked_path, mod.MAX_ROWS_PER_FILE

    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font
    except ImportError:
        # 回退 CSV
        import csv
        csv_path = out_path.with_suffix(".csv")
        with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([h for h, _ in COLUMNS])
            for i, record in enumerate(records, 1):
                writer.writerow(row_values(record, i))
        print(f"  (未安装 openpyxl，已生成 CSV: {csv_path.name})")
        return

    chunks = chunk_records(records, MAX_ROWS_PER_FILE)
    for index, chunk in enumerate(chunks):
        file_path = chunked_path(out_path, index)
        wb = Workbook()
        ws = wb.active
        ws.title = "朋友圈"
        headers = [h for h, _ in COLUMNS]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for i, record in enumerate(chunk, 1):
            ws.append(row_values(record, i))

        widths = {
            "时间": 20, "UID": 28, "手机号": 16, "朋友圈正文": 60, "图片路径": 50,
        }
        from openpyxl.utils import get_column_letter
        for col_idx, (header, _) in enumerate(COLUMNS, 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = widths.get(header, 16)
        ws.freeze_panes = "A2"
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        wb.save(str(file_path))

    if len(chunks) > 1:
        print(f"  共 {len(records)} 条，拆分为 {len(chunks)} 个 Excel（每个上限 {MAX_ROWS_PER_FILE} 条）")


def main() -> int:
    if len(sys.argv) < 2:
        print("用法: python src/tools/clean_moments.py <导出目录>")
        return 1

    target_dir = Path(sys.argv[1])
    json_path = target_dir / "moments.json"
    if not json_path.exists():
        print(f"未找到 {json_path}")
        return 1

    records = load_records(json_path)
    print(f"原始记录: {len(records)} 条")

    kept, removed = clean_records(records)
    print(f"保留商品: {len(kept)} 条")
    print(f"剔除记录: {len(removed)} 条")

    # 输出剔除原因统计
    if removed:
        from collections import Counter
        reason_counter: Counter[str] = Counter()
        for r in removed:
            for reason in r.get("_remove_reason", "").split(","):
                reason_counter[reason] += 1
        print("\n剔除原因统计:")
        for reason, count in reason_counter.most_common():
            print(f"  {reason}: {count} 条")

    # 写入清洗后的 JSON
    clean_json = target_dir / "moments_clean.json"
    clean_json.write_text(
        json.dumps(kept, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\n已生成: {clean_json}")

    # 写入清洗后的 Excel
    clean_xlsx = target_dir / "moments_clean.xlsx"
    try:
        write_xlsx(kept, clean_xlsx)
        print(f"已生成: {clean_xlsx}")
    except RuntimeError as exc:
        print(f"  [warning] Excel 生成失败: {exc}")

    # 导出被剔除的记录（便于人工核查）
    if removed:
        removed_json = target_dir / "moments_removed.json"
        removed_json.write_text(
            json.dumps(removed, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"剔除记录(供核查): {removed_json}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
