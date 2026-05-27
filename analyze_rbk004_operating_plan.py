from __future__ import annotations

import html
import math
import re
import sys
import warnings
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


LISTING_NAME = "RBK004"
PRODUCT_AGE_NOTE = "产品适用年龄为 3-7 岁；所有搜索词都需要按年龄段、对象、功能、场景和实际转化拆分判断，不能按单一词根一刀切。"
PRODUCT_AGE_MIN = 3
PRODUCT_AGE_MAX = 7


def safe_div(numerator: Any, denominator: Any, digits: int = 9) -> float | None:
    """除法统一入口：避免 0 除，并保留足够精度供后续复核。"""
    try:
        denominator = float(denominator)
        numerator = float(numerator)
    except (TypeError, ValueError):
        return None
    if denominator == 0 or math.isnan(denominator):
        return None
    return round(numerator / denominator, digits)


def describe_rank(best: Any, average: Any, worst: Any, rank_type: str = "自然") -> str:
    """排名数字越小越好，所以用最佳/平均/最差表达。"""
    best_text = "-" if pd.isna(best) or best == 0 else int(best)
    avg_text = "-" if pd.isna(average) or average == 0 else f"{float(average):.1f}"
    worst_text = "-" if pd.isna(worst) or worst == 0 else int(worst)
    return f"最佳{rank_type}排名第 {best_text}，平均{rank_type}排名第 {avg_text}，最差{rank_type}排名第 {worst_text}"


def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def pct(value: Any, digits: int = 2) -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"{float(value) * 100:.{digits}f}%"


def money(value: Any, digits: int = 2) -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"${float(value):,.{digits}f}"


def num_text(value: Any, digits: int = 0) -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"{float(value):,.{digits}f}"


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def is_source_excel(name: str) -> bool:
    return name.lower().endswith(".xlsx") and not name.startswith("~$")


def has_any(text: str, words: list[str]) -> bool:
    return any(word in text for word in words)


def extract_age_segment(term: str) -> str:
    t = term.lower().strip()
    has_infant_context = bool(re.search(r"\bbaby\b|\binfant\b|\btoddler\b|months?", t))
    month_patterns = [r"\b\d+\s*-\s*\d+\s*months?\b", r"\b\d+\s*months?\b", r"\binfant\b"]
    if any(re.search(pattern, t) for pattern in month_patterns):
        return "0-2"

    range_match = re.search(r"\b(\d{1,2})\s*[-–]\s*(\d{1,2})\b", t)
    if range_match:
        low = int(range_match.group(1))
        high = int(range_match.group(2))
        if high <= 2 or (has_infant_context and high > 12):
            return "0-2"
        return f"{low}-{high}"

    age_match = re.search(r"\bage\s*(\d{1,2})\b", t)
    if age_match:
        age = int(age_match.group(1))
        if age <= 2:
            return "0-2"
        if age <= 4:
            return "2-4"
        if age <= 7:
            return "3-7"
        return "8+"

    if re.search(r"\bbaby\b", t):
        return "baby泛"
    if re.search(r"\btoddler\b", t):
        return "toddler泛"
    if re.search(r"\b(kid|kids|child|children|childrens)\b", t):
        return "kids泛"
    return "未标明"


def age_segment_overlaps_product(segment: str) -> bool:
    match = re.fullmatch(r"(\d{1,2})-(\d{1,2})", clean_text(segment))
    if not match:
        return False
    low = int(match.group(1))
    high = int(match.group(2))
    return low <= PRODUCT_AGE_MAX and high >= PRODUCT_AGE_MIN


def extract_search_features(term: str) -> dict[str, str]:
    t = term.lower().strip()

    if "beach essentials" in t:
        root = "beach essentials"
    elif "toddler" in t and re.search(r"sunglass|sun glass|shades", t):
        root = "toddler sunglasses"
    elif "baby" in t and re.search(r"sunglass|sun glass|shades", t):
        root = "baby sunglasses"
    elif re.search(r"\bboy|boys\b", t) and re.search(r"sunglass|sun glass|shades", t):
        root = "boys sunglasses"
    elif re.search(r"\bgirl|girls\b", t) and re.search(r"sunglass|sun glass|shades", t):
        root = "girls sunglasses"
    elif "sunglasses for kids" in t:
        root = "sunglasses for kids"
    elif "polarized" in t and re.search(r"\b(kid|kids|child|children|childrens)\b", t) and re.search(r"sunglass|sun glass|shades", t):
        root = "kids polarized sunglasses"
    elif re.search(r"\b(kid|kids|child|children|childrens)\b", t) and re.search(r"sunglass|sun glass|shades", t):
        root = "kids sunglasses"
    elif re.search(r"sunglass|sun glass|shades", t):
        root = "sunglasses泛词"
    else:
        root = "其他"

    if re.search(r"\bboy|boys\b", t):
        gender = "男孩"
    elif re.search(r"\bgirl|girls\b", t):
        gender = "女孩"
    else:
        gender = "通用"

    function_tags = []
    if "polarized" in t:
        function_tags.append("polarized")
    if "strap" in t:
        function_tags.append("strap")
    if has_any(t, ["uv", "uv400"]):
        function_tags.append("uv")
    if has_any(t, ["sport", "baseball", "cycling", "running"]):
        function_tags.append("sports")
    if has_any(t, ["bulk", "pack", "party favors"]):
        function_tags.append("bulk")
    function = "+".join(function_tags) if function_tags else "基础"

    scene_tags = []
    for scene in ["beach", "travel", "vacation", "camp", "cruise", "summer"]:
        if scene in t:
            scene_tags.append(scene)
    scene = "+".join(scene_tags) if scene_tags else "无"

    return {
        "root": root,
        "age_segment": extract_age_segment(t),
        "gender": gender,
        "function": function,
        "scene": scene,
    }


def classify_relevance(term: str) -> str:
    t = term.lower().strip()
    features = extract_search_features(t)
    has_sun = bool(re.search(r"sunglass|sun glass|shades", t))
    has_core_audience = bool(
        re.search(r"\b(kid|kids|child|children|childrens|toddler|girl|girls|boy|boys)\b", t)
    )
    has_age = bool(re.search(r"\bage\s*\d|\d+\s*-\s*\d+", t))
    has_polarized = "polarized" in t
    has_scene = has_any(t, ["beach", "travel", "vacation", "camp", "cruise", "summer"])
    has_baby = bool(re.search(r"\bbaby\b|\binfant\b|0-6|6-12|12-24", t))
    has_teen = bool(re.search(r"\bteen\b|\byouth\b", t))
    has_brand = has_any(t, ["rivbos", "rayban", "oakley", "goodr", "knockaround", "disney"])

    if has_brand:
        return "品牌词"
    if has_teen:
        return "低相关词"
    if features["age_segment"] == "0-2":
        return "低相关词"
    if age_segment_overlaps_product(features["age_segment"]):
        return "年龄词"
    if features["root"] == "baby sunglasses":
        return "年龄泛词"
    if has_baby:
        return "年龄泛词"
    if has_sun and has_polarized and has_core_audience:
        return "核心词"
    if has_sun and has_core_audience and has_age:
        return "年龄词"
    if has_sun and has_any(t, ["girl", "girls", "boy", "boys"]):
        return "性别词"
    if has_sun and has_core_audience:
        return "核心词" if len(t.split()) <= 3 else "长尾词"
    if has_scene:
        return "场景词"
    if has_sun:
        return "泛流量词"
    return "低相关词"


def explain_search_term_cause(
    term: str,
    relevance: str,
    performance: str,
    action: str,
    clicks: float,
    orders: float,
    spend: float,
    sales: float,
    acos: Any,
    cvr: Any,
    aba_rank: Any,
) -> str:
    """把指标现象翻译成业务原因，避免只堆数字。"""
    reasons: list[str] = []
    t = term.lower()
    features = extract_search_features(term)
    if relevance == "低相关词":
        if features["age_segment"] == "0-2" or "infant" in t:
            reasons.append("搜索意图偏 0-2 岁或婴幼儿，和 Listing 的 3-7 岁儿童定位存在年龄段错配")
        elif "teen" in t or "youth" in t:
            reasons.append("搜索意图偏青少年，和 Listing 的 3-7 岁儿童定位不完全一致")
        else:
            reasons.append("词义和儿童太阳镜主购买意图不够贴合，流量进入后承接难度高")
    elif relevance == "年龄泛词":
        reasons.append("词根带年龄或对象倾向但年龄段不明确，需要用实际转化和 ABA 热度判断，不能按单一词根整体否定")
    elif relevance == "场景词":
        reasons.append("场景词表达的是出行/海滩需求，不一定已经明确要买儿童太阳镜，转化链路更长")
    elif relevance == "泛流量词":
        reasons.append("泛太阳镜词用户人群过宽，Listing 需要同时匹配年龄、性别和儿童场景才容易成交")
    else:
        reasons.append("词义和产品定位匹配，后续主要看流量成本与转化效率是否支撑放量")

    if clicks >= 10 and orders <= 1:
        reasons.append("点击样本已经出现但订单少，说明当前词的转化承接弱")
    elif clicks >= 10 and cvr is not None and not pd.isna(cvr) and cvr < 0.15:
        reasons.append("点击后转化率偏低，问题更可能在搜索意图匹配或 Listing 承接")
    elif orders >= 5 and cvr is not None and not pd.isna(cvr) and cvr >= 0.35:
        reasons.append("点击后转化效率较好，说明用户意图和页面承接基本成立")

    if acos is not None and not pd.isna(acos) and acos >= 0.8:
        reasons.append("ACOS 过高不是单一花费问题，而是 CPC、低转化和客单价共同导致回收不足")
    elif acos is not None and not pd.isna(acos) and acos >= 0.45:
        reasons.append("产出能证明词有一定相关性，但利润空间不足，需要先降价而不是扩量")
    elif acos is not None and not pd.isna(acos) and acos <= 0.30 and orders >= 5:
        reasons.append("在有订单规模的同时 ACOS 可控，说明这个词具备保量或转精确价值")

    if aba_rank is not None and not pd.isna(aba_rank) and aba_rank > 0:
        if aba_rank <= 100000:
            reasons.append("ABA 搜索热度较高，若相关性成立就值得纳入重点词池")
        elif aba_rank <= 300000:
            reasons.append("ABA 仍在 30 万内，可作为补充长尾词测试")

    return "；".join(reasons[:5])


def classify_search_term(row: dict[str, Any] | pd.Series) -> dict[str, Any]:
    term = clean_text(row.get("term", ""))
    relevance = classify_relevance(term)
    clicks = float(row.get("clicks", 0) or 0)
    orders = float(row.get("orders", 0) or 0)
    spend = float(row.get("spend", 0) or 0)
    sales = float(row.get("sales", 0) or 0)
    acos = row.get("acos", None)
    cvr = row.get("cvr", None)
    aba_rank = row.get("aba_rank", None)

    evidence: list[str] = []
    if relevance in {"核心词", "年龄词", "性别词", "长尾词"}:
        evidence.append(f"相关性={relevance}")
    elif relevance in {"低相关词", "场景词", "泛流量词", "品牌词", "年龄泛词"}:
        evidence.append(f"相关性偏弱={relevance}")
    if spend >= 20:
        evidence.append(f"花费较高={spend:.2f}")
    if clicks >= 10:
        evidence.append(f"点击样本={clicks:.0f}")
    if orders >= 5:
        evidence.append(f"订单={orders:.0f}")
    if sales >= 50:
        evidence.append(f"销售额={sales:.2f}")
    if acos is not None and not pd.isna(acos):
        if acos <= 0.25:
            evidence.append(f"ACOS低={pct(acos)}")
        elif acos >= 0.45:
            evidence.append(f"ACOS高={pct(acos)}")
    if cvr is not None and not pd.isna(cvr):
        if cvr >= 0.35:
            evidence.append(f"CVR高={pct(cvr)}")
        elif clicks >= 10 and cvr < 0.15:
            evidence.append(f"CVR低={pct(cvr)}")
    if aba_rank is not None and not pd.isna(aba_rank) and aba_rank > 0:
        if aba_rank <= 100000:
            evidence.append(f"ABA高价值={int(aba_rank)}")
        elif aba_rank <= 300000:
            evidence.append(f"ABA可参考={int(aba_rank)}")

    if orders == 0 and spend >= 5 and clicks >= 3:
        performance = "无订单浪费"
        action = "否定" if relevance in {"低相关词", "场景词", "泛流量词"} else "观察"
    elif acos is not None and not pd.isna(acos) and acos >= 0.8 and spend >= 5:
        performance = "高花费低产出"
        action = "暂停" if relevance in {"低相关词", "场景词", "泛流量词"} else "降价"
    elif acos is not None and not pd.isna(acos) and acos >= 0.45 and spend >= 5:
        performance = "高花费低产出"
        action = "降价"
    elif orders >= 5 and acos is not None and not pd.isna(acos) and acos <= 0.30:
        performance = "可放量"
        action = "转精确" if relevance in {"核心词", "年龄词", "性别词", "长尾词"} else "低预算测试"
    elif orders >= 5:
        performance = "高转化"
        action = "保持"
    elif clicks < 10:
        performance = "不足样本"
        action = "观察"
    else:
        performance = "控价观察"
        action = "降价" if relevance in {"低相关词", "场景词", "泛流量词"} else "观察"

    cause = explain_search_term_cause(term, relevance, performance, action, clicks, orders, spend, sales, acos, cvr, aba_rank)

    return {
        "相关性分类": relevance,
        "表现分类": performance,
        "动作建议": action,
        "归因解释": cause,
        "判断依据": "；".join(evidence[:8]),
        "证据数量": len(evidence),
    }


@dataclass
class Paths:
    root: Path
    src: Path
    listing_dir: Path
    output_dir: Path
    listing_file: Path
    activity_file: Path
    placement_file: Path
    target_file: Path
    search_file: Path
    rank_file: Path
    aba_file: Path
    product_file: Path | None


def find_one(paths: list[Path], marker: str) -> Path:
    matches = [path for path in paths if marker in path.name]
    if not matches:
        raise FileNotFoundError(f"未找到包含 {marker} 的文件")
    return matches[0]


def resolve_paths(root: Path) -> Paths:
    src = root / "src_data"
    listing_dir = src / LISTING_NAME
    if not listing_dir.exists():
        candidates = [p for p in src.iterdir() if p.is_dir()]
        if not candidates:
            raise FileNotFoundError("src_data 下未找到 Listing 目录")
        listing_dir = candidates[0]

    files = [path for path in listing_dir.glob("*.xlsx") if is_source_excel(path.name)]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = root / "outputs" / f"{LISTING_NAME}_operating_plan_{timestamp}"
    product_files = list(listing_dir.glob("*.md"))
    aba_files = [path for path in src.glob("*.xlsx") if is_source_excel(path.name)]

    return Paths(
        root=root,
        src=src,
        listing_dir=listing_dir,
        output_dir=output_dir,
        listing_file=find_one(files, "000048271"),
        activity_file=find_one(files, "广告活动"),
        placement_file=find_one(files, "广告位"),
        target_file=find_one(files, "投放报表"),
        search_file=find_one(files, "用户搜索词"),
        rank_file=find_one(files, "000048270"),
        aba_file=find_one(aba_files, "ABA"),
        product_file=product_files[0] if product_files else None,
    )


def read_product_intro(path: Path | None) -> str:
    if not path:
        return ""
    return path.read_text(encoding="utf-8", errors="ignore")


def read_listing(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path)
    df = pd.DataFrame(
        {
            "date": pd.to_datetime(raw["日期"]),
            "units": to_num(raw["销量"]),
            "orders": to_num(raw["订单量"]),
            "sales": to_num(raw["销售额"]),
            "actual_sales": to_num(raw["实际销售额"]),
            "sessions": to_num(raw["Sessions-Total"]),
            "pv": to_num(raw["PV-Total"]),
            "rating": to_num(raw["星级/评分"]),
            "rating_count": to_num(raw["评星数"]),
            "review_count": to_num(raw["评论数"]),
            "ad_impr": to_num(raw["广告曝光量"]),
            "ad_clicks": to_num(raw["广告点击量"]),
            "ad_orders": to_num(raw["广告订单量"]),
            "ad_spend": to_num(raw["广告花费"]).abs(),
            "ad_sales": to_num(raw["广告销售额"]),
            "natural_orders": to_num(raw["自然订单量"]),
            "gross_profit": to_num(raw["销售毛利"]),
            "gross_margin_rate_source": to_num(raw["销售毛利率"]),
            "net_profit": to_num(raw["销售净毛利"]),
            "net_margin_rate_source": to_num(raw["销售净毛利率"]),
            "promo_discount": to_num(raw["促销折扣"]).abs(),
            "coupon_spend": to_num(raw["Coupon花费"]).abs(),
            "deal_spend": to_num(raw["Deal花费"]).abs(),
            "avg_sale_price_source": to_num(raw["平均销售价格"]),
        }
    )
    return df.sort_values("date")


def exclude_incomplete_last_day(daily: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    if len(daily) < 4:
        return daily, "数据天数不足，未排除日期"
    last = daily.iloc[-1]
    previous = daily.iloc[:-1].tail(7)
    median_orders = previous["orders"].median()
    if last["orders"] < max(5, median_orders * 0.2):
        note = f"排除 {last['date'].date()}：订单 {last['orders']:.0f} 明显低于前 7 日中位数 {median_orders:.0f}，判断为未完整日。"
        return daily.iloc[:-1].copy(), note
    return daily, "未发现明显未完整日。"


def period_metrics(df: pd.DataFrame, label: str) -> dict[str, Any]:
    sales = df["sales"].sum()
    orders = df["orders"].sum()
    units = df["units"].sum()
    ad_spend = df["ad_spend"].sum()
    ad_sales = df["ad_sales"].sum()
    ad_orders = df["ad_orders"].sum()
    ad_clicks = df["ad_clicks"].sum()
    ad_impr = df["ad_impr"].sum()
    sessions = df["sessions"].sum()
    gross_profit = df["gross_profit"].sum()
    net_profit = df["net_profit"].sum()
    promo_total = df["promo_discount"].sum() + df["coupon_spend"].sum() + df["deal_spend"].sum()

    return {
        "周期": label,
        "开始日期": df["date"].min().date().isoformat(),
        "结束日期": df["date"].max().date().isoformat(),
        "天数": df["date"].nunique(),
        "销量": units,
        "订单量": orders,
        "销售额": sales,
        "客单价": safe_div(sales, orders),
        "件单价": safe_div(sales, units),
        "Sessions": sessions,
        "订单CVR": safe_div(orders, sessions),
        "广告曝光": ad_impr,
        "广告点击": ad_clicks,
        "广告CTR": safe_div(ad_clicks, ad_impr),
        "广告订单": ad_orders,
        "广告销售额": ad_sales,
        "广告花费": ad_spend,
        "广告CVR": safe_div(ad_orders, ad_clicks),
        "广告ACOS": safe_div(ad_spend, ad_sales),
        "广告ROAS": safe_div(ad_sales, ad_spend),
        "广告客单价": safe_div(ad_sales, ad_orders),
        "自然订单": df["natural_orders"].sum(),
        "自然订单占比": safe_div(df["natural_orders"].sum(), orders),
        "广告订单占比": safe_div(ad_orders, orders),
        "销售毛利": gross_profit,
        "销售毛利率": safe_div(gross_profit, sales),
        "销售净毛利": net_profit,
        "销售净毛利率": safe_div(net_profit, sales),
        "促销折扣+Coupon+Deal": promo_total,
        "促销费用率": safe_div(promo_total, sales),
    }


def compare_periods(prev: dict[str, Any], recent: dict[str, Any]) -> pd.DataFrame:
    rows = []
    metrics = [
        "销量",
        "订单量",
        "销售额",
        "客单价",
        "件单价",
        "Sessions",
        "订单CVR",
        "广告花费",
        "广告销售额",
        "广告ACOS",
        "广告ROAS",
        "广告客单价",
        "自然订单",
        "自然订单占比",
        "销售毛利率",
        "销售净毛利率",
        "促销折扣+Coupon+Deal",
        "促销费用率",
    ]
    for metric in metrics:
        p = prev.get(metric)
        r = recent.get(metric)
        rows.append({"指标": metric, "前期": p, "近期": r, "变化率": safe_div((r or 0) - (p or 0), p)})
    return pd.DataFrame(rows)


def read_ad_table(path: Path, kind: str) -> pd.DataFrame:
    raw = pd.read_excel(path)
    if kind == "activity":
        cols = {"date": "日期", "campaign": "广告活动"}
    elif kind == "placement":
        cols = {"date": "日期", "campaign": "广告活动", "placement": "广告位"}
    elif kind == "target":
        cols = {"date": "日期", "campaign": "所属广告活动", "target": "投放名称", "match": "匹配类型", "bid": "竞价", "aba_rank": "ABA排名(按日)"}
    elif kind == "search":
        cols = {
            "date": "日期",
            "campaign": "所属广告活动",
            "term": "用户搜索词",
            "target": "投放名称",
            "match": "匹配类型",
            "bid": "竞价",
            "aba_rank": "ABA排名(按日)",
        }
    else:
        raise ValueError(kind)

    data: dict[str, Any] = {k: raw[v] for k, v in cols.items()}
    data.update(
        {
            "impr": to_num(raw["曝光量"]),
            "clicks": to_num(raw["点击量"]),
            "spend": to_num(raw["花费"]),
            "sales": to_num(raw["广告总销售额"]),
            "orders": to_num(raw["广告总订单量"]),
        }
    )
    df = pd.DataFrame(data)
    df["date"] = pd.to_datetime(df["date"])
    return df


def aggregate_ad(df: pd.DataFrame, keys: list[str], start: pd.Timestamp) -> pd.DataFrame:
    numeric = {"impr": "sum", "clicks": "sum", "spend": "sum", "sales": "sum", "orders": "sum"}
    if "bid" in df.columns:
        numeric["bid"] = "mean"
    if "aba_rank" in df.columns:
        numeric["aba_rank"] = lambda s: to_num(s).replace(0, pd.NA).min()
    grouped = df[df["date"] >= start].groupby(keys, dropna=False).agg(numeric).reset_index()
    grouped["CTR"] = grouped.apply(lambda r: safe_div(r["clicks"], r["impr"]), axis=1)
    grouped["CPC"] = grouped.apply(lambda r: safe_div(r["spend"], r["clicks"]), axis=1)
    grouped["CVR"] = grouped.apply(lambda r: safe_div(r["orders"], r["clicks"]), axis=1)
    grouped["ACOS"] = grouped.apply(lambda r: safe_div(r["spend"], r["sales"]), axis=1)
    grouped["ROAS"] = grouped.apply(lambda r: safe_div(r["sales"], r["spend"]), axis=1)
    grouped["CPA"] = grouped.apply(lambda r: safe_div(r["spend"], r["orders"]), axis=1)
    return grouped


def add_activity_decision(df: pd.DataFrame, prev: pd.DataFrame | None = None) -> pd.DataFrame:
    if prev is not None and not prev.empty:
        prev_small = prev[["campaign", "spend", "sales", "orders", "ACOS"]].rename(
            columns={"spend": "前期花费", "sales": "前期销售额", "orders": "前期订单", "ACOS": "前期ACOS"}
        )
        df = df.merge(prev_small, on="campaign", how="left")
        df["花费变化"] = df.apply(lambda r: safe_div(r["spend"] - (r["前期花费"] or 0), r["前期花费"]), axis=1)
        df["销售额变化"] = df.apply(lambda r: safe_div(r["sales"] - (r["前期销售额"] or 0), r["前期销售额"]), axis=1)
    else:
        df["花费变化"] = None
        df["销售额变化"] = None

    decisions = []
    causes = []
    evidence = []
    for _, r in df.iterrows():
        ev = [
            f"花费 {money(r['spend'])}",
            f"订单 {r['orders']:.0f}",
            f"ACOS {pct(r['ACOS'])}",
            f"CVR {pct(r['CVR'])}",
            f"ROAS {num_text(r['ROAS'], 2)}",
        ]
        if not pd.isna(r.get("花费变化")):
            ev.append(f"花费变化 {pct(r['花费变化'])}")
        if not pd.isna(r.get("销售额变化")):
            ev.append(f"销售额变化 {pct(r['销售额变化'])}")

        if r["orders"] >= 100 and r["ACOS"] <= 0.32 and r["CVR"] >= 0.30:
            decision = "继续保量"
            cause = "订单规模、转化效率和 ACOS 同时成立，说明活动不仅能拿量，也能在当前利润约束下承接转化"
        elif r["orders"] >= 100 and r["ACOS"] <= 0.38:
            decision = "控价保量"
            cause = "活动有订单规模，但 ACOS 已接近利润约束，需要保住流量同时控制竞价"
        elif r["ACOS"] >= 0.40 and r["spend"] >= 200:
            decision = "降预算/降竞价"
            cause = "花费规模已经不小，但销售回收不足，问题通常来自 CPC 偏高、词意图过宽或转化承接下降"
        elif r["orders"] <= 5 and r["spend"] >= 50:
            decision = "暂停或大幅降预算"
            cause = "预算已经消耗但订单不足，说明当前活动还没有证明可规模化，继续放量会扩大亏损"
        else:
            decision = "观察"
            cause = "数据规模或效率信号还不够一致，先保持观察，避免只凭单一指标调整"
        decisions.append(decision)
        causes.append(cause)
        evidence.append("；".join(ev))
    df["动作建议"] = decisions
    df["归因解释"] = causes
    df["判断依据"] = evidence
    return df


def add_placement_decision(df: pd.DataFrame) -> pd.DataFrame:
    decisions, causes, evidence = [], [], []
    for _, r in df.iterrows():
        ev = [
            f"花费 {money(r['spend'])}",
            f"订单 {r['orders']:.0f}",
            f"CTR {pct(r['CTR'])}",
            f"CVR {pct(r['CVR'])}",
            f"ACOS {pct(r['ACOS'])}",
        ]
        placement = clean_text(r.get("placement", ""))
        if placement in {"商品页面", "亚马逊站外"} and (r["ACOS"] >= 0.40 or r["CVR"] < 0.22):
            decision = "下调广告位"
            cause = "该位置更多承接浏览或竞品页流量，购买意图弱于搜索结果页；当 CVR 和 ACOS 同时变差时，应先降位置权重"
        elif placement == "搜索结果顶部(首页)" and r["orders"] >= 100:
            decision = "保留顶部"
            cause = "顶部位置虽然 CPC 通常更高，但订单规模和转化效率证明它能承接核心需求"
        elif r["ACOS"] >= 0.38:
            decision = "小幅下调"
            cause = "该位置仍有订单，但利润回收偏弱，先小幅降价验证是否能改善 ACOS"
        else:
            decision = "保持"
            cause = "花费、订单和效率没有形成明确负面信号，暂不做剧烈调整"
        decisions.append(decision)
        causes.append(cause)
        evidence.append("；".join(ev))
    df["动作建议"] = decisions
    df["归因解释"] = causes
    df["判断依据"] = evidence
    return df


def add_target_decision(df: pd.DataFrame) -> pd.DataFrame:
    decisions, causes, evidence = [], [], []
    for _, r in df.iterrows():
        ev = [
            f"花费 {money(r['spend'])}",
            f"订单 {r['orders']:.0f}",
            f"CTR {pct(r['CTR'])}",
            f"CVR {pct(r['CVR'])}",
            f"ACOS {pct(r['ACOS'])}",
        ]
        target = clean_text(r.get("target", ""))
        relevance = classify_relevance(target)
        ev.append(f"相关性 {relevance}")
        if r["orders"] >= 20 and r["ACOS"] <= 0.30:
            decision = "保量/可小幅提价"
            cause = "投放词有订单规模且 ACOS 在可控区间，说明词意图、竞价和 Listing 承接匹配"
        elif r["orders"] >= 20 and r["ACOS"] <= 0.38:
            decision = "控价保量"
            cause = "词能出单但利润空间被压缩，当前优先保流量并控制 CPC"
        elif r["spend"] >= 20 and (r["orders"] == 0 or r["ACOS"] >= 0.45):
            decision = "降价/暂停"
            cause = "花费已形成样本但产出不足，主要问题是转化承接弱或点击成本高于可承受水平"
        elif relevance in {"低相关词", "场景词", "泛流量词"} and r["ACOS"] >= 0.38:
            decision = "降价"
            cause = "词意图偏宽，能带来点击但转化链路不够短，适合低价保留而不是重点放量"
        else:
            decision = "观察"
            cause = "相关性、规模和效率尚未形成一致结论，需要继续积累数据"
        decisions.append(decision)
        causes.append(cause)
        evidence.append("；".join(ev))
    df["相关性分类"] = df["target"].map(lambda x: classify_relevance(clean_text(x)))
    df["动作建议"] = decisions
    df["归因解释"] = causes
    df["判断依据"] = evidence
    return df


def read_rank_table(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path)
    df = pd.DataFrame(
        {
            "keyword": raw["关键词"],
            "ad_rank": to_num(raw["广告总排名"]),
            "ad_change": pd.to_numeric(raw["广告排名变化"], errors="coerce"),
            "natural_rank": to_num(raw["自然总排名"]),
            "natural_change": pd.to_numeric(raw["自然排名变化"], errors="coerce"),
            "other": raw["其他流量位置"].fillna(""),
            "activity": raw["活动"].fillna(""),
        }
    )

    def positive_min(s: pd.Series) -> float:
        s = s[s > 0]
        return float(s.min()) if len(s) else 0

    def positive_mean(s: pd.Series) -> float:
        s = s[s > 0]
        return float(s.mean()) if len(s) else 0

    def positive_max(s: pd.Series) -> float:
        s = s[s > 0]
        return float(s.max()) if len(s) else 0

    grouped = df.groupby("keyword", dropna=False).agg(
        监控次数=("keyword", "size"),
        最佳广告排名=("ad_rank", positive_min),
        平均广告排名=("ad_rank", positive_mean),
        最差广告排名=("ad_rank", positive_max),
        广告排名变化均值=("ad_change", "mean"),
        最佳自然排名=("natural_rank", positive_min),
        平均自然排名=("natural_rank", positive_mean),
        最差自然排名=("natural_rank", positive_max),
        自然排名变化均值=("natural_change", "mean"),
        其他流量位置=("other", lambda s: "；".join(sorted({clean_text(x) for x in s if clean_text(x)}))[:120]),
        活动=("activity", lambda s: "；".join(sorted({clean_text(x) for x in s if clean_text(x)}))[:120]),
    ).reset_index()
    grouped["相关性分类"] = grouped["keyword"].map(lambda x: classify_relevance(clean_text(x)))

    actions, causes, evidence = [], [], []
    for _, r in grouped.iterrows():
        rank_desc = describe_rank(r["最佳自然排名"], r["平均自然排名"], r["最差自然排名"], "自然")
        ad_desc = describe_rank(r["最佳广告排名"], r["平均广告排名"], r["最差广告排名"], "广告")
        relevance = r["相关性分类"]
        ev = [rank_desc, ad_desc, f"相关性 {relevance}"]
        best_natural = r["最佳自然排名"]
        avg_natural = r["平均自然排名"]
        best_ad = r["最佳广告排名"]
        if relevance == "低相关词":
            action = "不作为主线，观察自然排名"
            cause = "虽然排名可能不错，但词义和 3-7 岁儿童太阳镜定位不完全一致，广告不应为了排名信号盲目扩量"
        elif best_natural and best_natural <= 5 and best_ad:
            action = "保自然排名，广告控价防守"
            cause = "自然排名已经靠前，广告的主要任务是防守核心位置，不是继续用高价冲排名"
        elif best_natural and best_natural <= 10 and not best_ad:
            action = "低价补广告位"
            cause = "自然结果已有承接，低价广告可补充页面占位，但需要避免挤压自然利润"
        elif avg_natural and avg_natural > 10 and best_ad:
            action = "观察广告是否带动自然"
            cause = "广告有位置但自然排名仍偏后，需要观察广告订单是否能沉淀为自然权重"
        else:
            action = "观察"
            cause = "排名和广告占位信号不够明确，暂不把它作为预算调整主因"
        actions.append(action)
        causes.append(cause)
        evidence.append("；".join(ev))
    grouped["动作建议"] = actions
    grouped["归因解释"] = causes
    grouped["判断依据"] = evidence
    return grouped.sort_values(["最佳自然排名", "平均自然排名"], ascending=True)


def read_aba(path: Path, current_terms: set[str]) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=0, usecols=list(range(22)))
    df = pd.DataFrame(
        {
            "country": raw.iloc[:, 0],
            "keyword": raw.iloc[:, 1].astype("string"),
            "搜索量排名": pd.to_numeric(raw.iloc[:, 4], errors="coerce"),
            "排名变化方向": raw.iloc[:, 6],
            "变化名次": pd.to_numeric(raw.iloc[:, 7], errors="coerce"),
            "前三ASIN点击份额": pd.to_numeric(raw.iloc[:, 8], errors="coerce"),
            "前三ASIN转化份额": pd.to_numeric(raw.iloc[:, 9], errors="coerce"),
            "#1 ASIN": raw.iloc[:, 10],
            "#2 ASIN": raw.iloc[:, 14],
            "#3 ASIN": raw.iloc[:, 18],
        }
    )
    df = df[df["country"].eq("US") & df["keyword"].notna()].copy()
    kw = df["keyword"].str.lower()
    sun = kw.str.contains(r"sunglass|sun glass|shades", regex=True, na=False)
    audience = kw.str.contains(
        r"\b(?:kid|kids|child|children|childrens|toddler|baby|girl|girls|boy|boys|youth|teen|junior)\b|age\s*\d|\d+\s*-\s*\d+",
        regex=True,
        na=False,
    )
    scenario = kw.str.contains(
        r"beach essentials|kids beach essentials|beach essentials for kids|summer essentials|vacation essentials|travel essentials|camp essentials|cruise essentials",
        regex=True,
        na=False,
    )
    df = df[(sun & audience) | scenario].copy()
    df["相关性分类"] = df["keyword"].map(lambda x: classify_relevance(clean_text(x)))
    df["当前广告覆盖"] = df["keyword"].str.lower().map(lambda x: "已覆盖" if x in current_terms else "未精确覆盖")

    actions, causes, evidence = [], [], []
    for _, r in df.iterrows():
        rank = r["搜索量排名"]
        click_share = r["前三ASIN点击份额"]
        conv_share = r["前三ASIN转化份额"]
        relevance = r["相关性分类"]
        covered = r["当前广告覆盖"]
        ev = [
            f"ABA排名 {int(rank) if not pd.isna(rank) else '-'}",
            f"排名变化 {int(r['变化名次']) if not pd.isna(r['变化名次']) else '-'}",
            f"前三点击份额 {pct(click_share)}",
            f"前三转化份额 {pct(conv_share)}",
            f"相关性 {relevance}",
            f"广告覆盖 {covered}",
        ]
        if relevance == "低相关词":
            action = "暂不重点布局"
            cause = "市场热度不能抵消产品定位不匹配，相关性不足会让广告点击更难转化"
        elif rank <= 100000 and click_share <= 0.45 and covered == "未精确覆盖":
            action = "新增重点测试"
            cause = "搜索热度高、头部集中度不极端且当前未精确覆盖，说明存在可切入的增量空间"
        elif rank <= 100000 and click_share <= 0.45:
            action = "保量优化"
            cause = "市场需求明确且竞争没有极端垄断，已有覆盖时重点是提升承接和控制成本"
        elif rank <= 300000 and click_share <= 0.45:
            action = "低预算测试"
            cause = "仍在 ABA 30 万内但热度不如核心词，适合作为长尾补充而不是主预算"
        else:
            action = "观察/不抢头部"
            cause = "头部份额或相关性信号不利，强行抢位容易变成高 CPC 低回收"
        actions.append(action)
        causes.append(cause)
        evidence.append("；".join(ev))
    df["动作建议"] = actions
    df["归因解释"] = causes
    df["判断依据"] = evidence
    return df.sort_values("搜索量排名")


def add_search_feature_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    features = result["term"].map(lambda x: extract_search_features(clean_text(x)))
    result["词根聚类"] = features.map(lambda x: x["root"])
    result["年龄段"] = features.map(lambda x: x["age_segment"])
    result["性别"] = features.map(lambda x: x["gender"])
    result["功能属性"] = features.map(lambda x: x["function"])
    result["场景"] = features.map(lambda x: x["scene"])
    return result


def build_aba_feature_table(aba_df: pd.DataFrame) -> pd.DataFrame:
    if aba_df.empty or "keyword" not in aba_df.columns:
        return pd.DataFrame()
    result = aba_df.copy()
    features = result["keyword"].map(lambda x: extract_search_features(clean_text(x)))
    result["词根聚类"] = features.map(lambda x: x["root"])
    result["年龄段"] = features.map(lambda x: x["age_segment"])
    result["性别"] = features.map(lambda x: x["gender"])
    result["功能属性"] = features.map(lambda x: x["function"])
    result["场景"] = features.map(lambda x: x["scene"])
    for col in ["搜索量排名", "前三ASIN点击份额", "前三ASIN转化份额"]:
        if col not in result.columns:
            result[col] = pd.NA
    result["搜索量排名"] = pd.to_numeric(result["搜索量排名"], errors="coerce")
    return result


def aba_best_for_root(aba_features: pd.DataFrame, root: str) -> dict[str, Any]:
    if aba_features.empty:
        return {}
    candidates = aba_features[aba_features["词根聚类"].eq(root)].copy()
    if candidates.empty:
        return {}
    candidates = candidates.sort_values("搜索量排名", na_position="last")
    row = candidates.iloc[0]
    return {
        "词根ABA代表词": clean_text(row.get("keyword", "")),
        "词根ABA最佳排名": row.get("搜索量排名", pd.NA),
        "词根ABA排名变化": row.get("变化名次", pd.NA),
        "词根ABA头部点击份额": row.get("前三ASIN点击份额", pd.NA),
        "词根ABA头部转化份额": row.get("前三ASIN转化份额", pd.NA),
    }


def add_search_aba_columns(search_df: pd.DataFrame, aba_df: pd.DataFrame) -> pd.DataFrame:
    result = add_search_feature_columns(search_df)
    aba_features = build_aba_feature_table(aba_df)
    if aba_features.empty:
        result["精确ABA排名"] = pd.NA
        result["词根ABA代表词"] = ""
        result["词根ABA最佳排名"] = pd.NA
        result["词根ABA头部点击份额"] = pd.NA
        result["词根ABA头部转化份额"] = pd.NA
        return result

    exact_rank_source = aba_features.dropna(subset=["搜索量排名"]).sort_values("搜索量排名").drop_duplicates("keyword").copy()
    exact_rank_source["keyword_key"] = exact_rank_source["keyword"].str.lower()
    exact_rank = exact_rank_source.set_index("keyword_key")["搜索量排名"]
    result["精确ABA排名"] = result["term"].str.lower().map(exact_rank)
    root_rows = result["词根聚类"].map(lambda root: aba_best_for_root(aba_features, root))
    for col in ["词根ABA代表词", "词根ABA最佳排名", "词根ABA排名变化", "词根ABA头部点击份额", "词根ABA头部转化份额"]:
        result[col] = root_rows.map(lambda x, c=col: x.get(c, pd.NA) if x else pd.NA)
    if "aba_rank" not in result.columns:
        result["aba_rank"] = pd.NA
    result["aba_rank"] = result["aba_rank"].where(result["aba_rank"].notna() & (result["aba_rank"] > 0), result["精确ABA排名"])
    result["aba_rank"] = result["aba_rank"].where(result["aba_rank"].notna() & (result["aba_rank"] > 0), result["词根ABA最佳排名"])
    return result


def summarize_top_value(df: pd.DataFrame, value_col: str, weight_col: str = "spend", limit: int = 3) -> str:
    if df.empty or value_col not in df.columns:
        return ""
    top = (
        df.groupby(value_col, dropna=False)[weight_col]
        .sum()
        .sort_values(ascending=False)
        .head(limit)
        .index
        .map(clean_text)
    )
    return "、".join([x for x in top if x])


def campaign_placement_lookup(placement_campaign: pd.DataFrame) -> pd.DataFrame:
    if placement_campaign.empty or "campaign" not in placement_campaign.columns or "placement" not in placement_campaign.columns:
        return pd.DataFrame(columns=["campaign", "主要广告位", "广告位归因"])
    rows = []
    for campaign, group in placement_campaign.groupby("campaign", dropna=False):
        top = group.sort_values("spend", ascending=False).iloc[0]
        rows.append(
            {
                "campaign": campaign,
                "主要广告位": top["placement"],
                "广告位归因": f"该活动主要花费在{top['placement']}，花费{money(top['spend'])}，CVR {pct(top['CVR'])}，ACOS {pct(top['ACOS'])}",
            }
        )
    return pd.DataFrame(rows)


def build_search_clusters(search: pd.DataFrame, aba_df: pd.DataFrame, placement_campaign: pd.DataFrame) -> pd.DataFrame:
    if search.empty:
        return pd.DataFrame()
    detail = add_search_aba_columns(search, aba_df)
    placement_lookup = campaign_placement_lookup(placement_campaign)
    if not placement_lookup.empty:
        detail = detail.merge(placement_lookup, on="campaign", how="left")
    else:
        detail["主要广告位"] = ""
        detail["广告位归因"] = ""

    group_cols = ["词根聚类", "年龄段", "性别", "功能属性", "场景"]
    grouped = (
        detail.groupby(group_cols, dropna=False)
        .agg({"impr": "sum", "clicks": "sum", "spend": "sum", "sales": "sum", "orders": "sum", "term": pd.Series.nunique})
        .reset_index()
        .rename(columns={"term": "搜索词数量"})
    )
    grouped["CTR"] = grouped.apply(lambda r: safe_div(r["clicks"], r["impr"]), axis=1)
    grouped["CPC"] = grouped.apply(lambda r: safe_div(r["spend"], r["clicks"]), axis=1)
    grouped["CVR"] = grouped.apply(lambda r: safe_div(r["orders"], r["clicks"]), axis=1)
    grouped["ACOS"] = grouped.apply(lambda r: safe_div(r["spend"], r["sales"]), axis=1)
    grouped["ROAS"] = grouped.apply(lambda r: safe_div(r["sales"], r["spend"]), axis=1)
    grouped["CPA"] = grouped.apply(lambda r: safe_div(r["spend"], r["orders"]), axis=1)

    aba_features = build_aba_feature_table(aba_df)
    rows = []
    for _, r in grouped.iterrows():
        mask = pd.Series(True, index=detail.index)
        for col in group_cols:
            mask &= detail[col].eq(r[col])
        g = detail[mask].copy()
        top_terms = g.sort_values(["orders", "spend"], ascending=False)["term"].head(5).map(clean_text).tolist()
        aba = aba_best_for_root(aba_features, r["词根聚类"])
        placement = summarize_top_value(g, "主要广告位")
        campaign = summarize_top_value(g, "campaign")
        relevance = classify_relevance(f"{r['词根聚类']} {r['年龄段']}")
        rank = aba.get("词根ABA最佳排名", pd.NA)
        click_share = aba.get("词根ABA头部点击份额", pd.NA)
        acos = r["ACOS"]
        cvr = r["CVR"]
        orders = r["orders"]
        spend = r["spend"]

        evidence = [
            f"花费 {money(spend)}",
            f"订单 {orders:.0f}",
            f"CVR {pct(cvr)}",
            f"ACOS {pct(acos)}",
            f"词根ABA排名 {int(rank) if not pd.isna(rank) else '-'}",
            f"头部点击份额 {pct(click_share)}",
        ]

        if relevance == "低相关词" and spend >= 5:
            conclusion = "低相关浪费簇"
            action = "否定或暂停"
            cause = "该簇的年龄或人群意图与 Listing 不匹配，广告继续买量会优先放大无效点击"
        elif orders >= 5 and acos is not None and not pd.isna(acos) and acos <= 0.30:
            conclusion = "广告有效且可承接"
            action = "保量并拆精确"
            cause = "订单规模、转化效率和 ACOS 同时成立，说明词根、用户意图和页面承接一致"
        elif rank is not None and not pd.isna(rank) and rank <= 100000 and spend >= 5 and (orders == 0 or (acos is not None and not pd.isna(acos) and acos >= 0.45)):
            conclusion = "市场有机会但当前承接弱"
            action = "降价拆分测试"
            cause = "ABA 证明有需求，但广告端回收不足，问题更可能在匹配过宽、竞价过高或 Listing 承接不足"
        elif spend >= 10 and (orders == 0 or (acos is not None and not pd.isna(acos) and acos >= 0.45)):
            conclusion = "广告效率偏弱"
            action = "降价或收窄"
            cause = "该簇已经消耗预算但订单或回收不足，优先控制 CPC 和匹配范围"
        elif rank is not None and not pd.isna(rank) and rank <= 100000:
            conclusion = "ABA 高价值待验证"
            action = "小预算保留"
            cause = "市场热度较高，但广告端样本或效率还不足以证明可放量"
        else:
            conclusion = "观察簇"
            action = "继续观察"
            cause = "市场热度、广告效率和样本规模没有形成一致信号，暂不做激进调整"

        row = r.to_dict()
        row.update(
            {
                "相关性分类": relevance,
                "代表搜索词": "、".join(top_terms),
                "主要活动": campaign,
                "主要广告位": placement,
                "广告位归因": "；".join([x for x in g["广告位归因"].dropna().map(clean_text).unique().tolist() if x][:3]),
                "词根ABA代表词": aba.get("词根ABA代表词", ""),
                "词根ABA最佳排名": rank,
                "词根ABA排名变化": aba.get("词根ABA排名变化", pd.NA),
                "词根ABA头部点击份额": click_share,
                "词根ABA头部转化份额": aba.get("词根ABA头部转化份额", pd.NA),
                "ABA归因": "ABA有搜索热度且头部集中度不极端" if rank is not None and not pd.isna(rank) and rank <= 100000 and (pd.isna(click_share) or click_share <= 0.45) else "ABA热度弱、未匹配或头部集中偏强",
                "聚类结论": conclusion,
                "动作建议": action,
                "归因解释": cause,
                "判断依据": "；".join(evidence),
            }
        )
        rows.append(row)
    return pd.DataFrame(rows).sort_values(["动作建议", "spend"], ascending=[True, False])


def build_search_aba_cross(search_detail: pd.DataFrame, clusters: pd.DataFrame, aba_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    if not clusters.empty:
        for _, r in clusters.iterrows():
            rank = r.get("词根ABA最佳排名", pd.NA)
            acos = r.get("ACOS", pd.NA)
            cvr = r.get("CVR", pd.NA)
            orders = r.get("orders", 0)
            if rank is not None and not pd.isna(rank) and rank <= 100000 and orders >= 5 and (pd.isna(acos) or acos <= 0.35):
                relation = "广告表现好 + ABA高价值"
                action = "重点保位并拆精确"
                cause = "广告端已证明能转化，ABA 也证明市场需求存在，适合做核心词池"
            elif rank is not None and not pd.isna(rank) and rank <= 100000 and (orders == 0 or (not pd.isna(acos) and acos >= 0.45)):
                relation = "广告表现差 + ABA高价值"
                action = "降价、收窄匹配、复核页面承接"
                cause = "市场词没错，但当前广告买法或 Listing 承接没有把需求转成订单"
            elif pd.isna(rank) and r.get("spend", 0) >= 10:
                relation = "广告在打但ABA弱/未匹配"
                action = "降预算或暂停"
                cause = "广告端已经消耗，但 ABA 没有对应市场热度证据，优先把预算迁回高价值词"
            else:
                relation = "广告与ABA信号待验证"
                action = "小预算观察"
                cause = "广告效率和市场热度都没有形成强信号，需要继续积累样本"
            rows.append(
                {
                    "词根聚类": r.get("词根聚类", ""),
                    "年龄段": r.get("年龄段", ""),
                    "性别": r.get("性别", ""),
                    "功能属性": r.get("功能属性", ""),
                    "场景": r.get("场景", ""),
                    "广告覆盖状态": "已产生广告搜索词",
                    "spend": r.get("spend", 0),
                    "sales": r.get("sales", 0),
                    "orders": orders,
                    "CVR": cvr,
                    "ACOS": acos,
                    "词根ABA代表词": r.get("词根ABA代表词", ""),
                    "词根ABA最佳排名": rank,
                    "词根ABA头部点击份额": r.get("词根ABA头部点击份额", pd.NA),
                    "市场-广告关系": relation,
                    "动作建议": action,
                    "归因解释": cause,
                    "判断依据": r.get("判断依据", ""),
                }
            )

    aba_features = build_aba_feature_table(aba_df)
    if not aba_features.empty:
        covered_roots = set(clusters["词根聚类"].dropna().map(clean_text)) if not clusters.empty else set()
        missing = aba_features[
            (~aba_features["词根聚类"].isin(covered_roots))
            & (aba_features["搜索量排名"].notna())
            & (aba_features["搜索量排名"] <= 100000)
            & (aba_features["相关性分类"].ne("低相关词") if "相关性分类" in aba_features.columns else True)
        ].sort_values("搜索量排名").head(30)
        for _, r in missing.iterrows():
            rows.append(
                {
                    "词根聚类": r.get("词根聚类", ""),
                    "年龄段": r.get("年龄段", ""),
                    "性别": r.get("性别", ""),
                    "功能属性": r.get("功能属性", ""),
                    "场景": r.get("场景", ""),
                    "广告覆盖状态": "有ABA机会但当前搜索词未覆盖",
                    "spend": 0,
                    "sales": 0,
                    "orders": 0,
                    "CVR": pd.NA,
                    "ACOS": pd.NA,
                    "词根ABA代表词": r.get("keyword", ""),
                    "词根ABA最佳排名": r.get("搜索量排名", pd.NA),
                    "词根ABA头部点击份额": r.get("前三ASIN点击份额", pd.NA),
                    "市场-广告关系": "有潜力ABA词未重点覆盖",
                    "动作建议": "新建小预算精确测试",
                    "归因解释": "ABA 有需求但当前搜索词报告没有对应承接，说明广告词池可能漏掉市场词",
                    "判断依据": f"ABA排名 {int(r['搜索量排名']) if not pd.isna(r['搜索量排名']) else '-'}；前三点击份额 {pct(r.get('前三ASIN点击份额', pd.NA))}",
                }
            )

    return pd.DataFrame(rows)


def build_source_check(paths: Paths, tables: dict[str, pd.DataFrame], incomplete_note: str) -> pd.DataFrame:
    rows = []
    mapping = {
        "Listing经营数据": paths.listing_file,
        "广告活动数据": paths.activity_file,
        "广告位数据": paths.placement_file,
        "投放数据": paths.target_file,
        "用户搜索词数据": paths.search_file,
        "关键词排名监控": paths.rank_file,
        "ABA Top30万搜索词": paths.aba_file,
    }
    for name, path in mapping.items():
        df = tables.get(name)
        rows.append(
            {
                "数据": name,
                "文件": str(path),
                "行数": len(df) if df is not None else "",
                "说明": incomplete_note if name == "Listing经营数据" else "",
            }
        )
    return pd.DataFrame(rows)


def html_table(df: pd.DataFrame, columns: list[str], max_rows: int = 12) -> str:
    ratio_markers = ("ACOS", "CVR", "CTR", "ROAS", "率", "占比", "份额", "变化率")

    def is_ratio_column(col_name: str) -> bool:
        return any(marker in col_name for marker in ratio_markers)

    def is_ratio_metric(row: pd.Series) -> bool:
        metric = clean_text(row.get("指标", ""))
        return any(marker in metric for marker in ratio_markers)

    rows = []
    for _, r in df.head(max_rows).iterrows():
        cells = []
        for col in columns:
            value = r.get(col, "")
            if isinstance(value, float):
                if is_ratio_column(col) or (col in {"前期", "近期"} and is_ratio_metric(r)):
                    text = pct(value)
                elif "花费" in col or "销售额" in col or "客单价" in col:
                    text = money(value)
                else:
                    text = num_text(value, 2)
            else:
                text = clean_text(value)
            cells.append(f"<td>{html.escape(text)}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    head = "".join(f"<th>{html.escape(col)}</th>" for col in columns)
    return f"<table><thead><tr>{head}</tr></thead><tbody>{''.join(rows)}</tbody></table>"


def build_evidence_chain(
    listing_recent: dict[str, Any],
    listing_prev: dict[str, Any],
    period_compare: pd.DataFrame,
    activity: pd.DataFrame,
    placement: pd.DataFrame,
    search_actions: pd.DataFrame,
    rank_df: pd.DataFrame,
    aba_df: pd.DataFrame,
) -> pd.DataFrame:
    def change(metric: str) -> str:
        row = period_compare[period_compare["指标"].eq(metric)]
        if row.empty:
            return "-"
        return pct(row.iloc[0]["变化率"])

    rows = [
        {
            "结论": "Listing 当前是订单增长但利润承压",
            "依据1": f"订单量变化 {change('订单量')}",
            "依据2": f"销售额变化 {change('销售额')}",
            "依据3": f"净毛利率从 {pct(listing_prev['销售净毛利率'])} 到 {pct(listing_recent['销售净毛利率'])}",
            "归因解释": "订单增长快于销售额，说明增量订单的客单价或成交结构变弱；净毛利率同步下降，说明增长质量被促销、成本或广告效率稀释",
            "动作": "先控广告浪费，再保核心排名",
        },
        {
            "结论": "近期存在价格/促销让利迹象",
            "依据1": f"客单价变化 {change('客单价')}",
            "依据2": f"件单价变化 {change('件单价')}",
            "依据3": f"促销费用率变化 {change('促销费用率')}",
            "归因解释": "客单价和件单价同时下降时，即使促销费用率只小幅变化，也可能说明折扣、组合销售或低价变体占比在拉低利润",
            "动作": "复核 Coupon、Deal、价格和促销折扣",
        },
        {
            "结论": "商品页面广告位拖累效率",
            "依据1": f"商品页面 ACOS {pct(placement.loc[placement['placement'].eq('商品页面'), 'ACOS'].iloc[0]) if not placement.loc[placement['placement'].eq('商品页面')].empty else '-'}",
            "依据2": f"商品页面 CVR {pct(placement.loc[placement['placement'].eq('商品页面'), 'CVR'].iloc[0]) if not placement.loc[placement['placement'].eq('商品页面')].empty else '-'}",
            "依据3": f"商品页面花费 {money(placement.loc[placement['placement'].eq('商品页面'), 'spend'].iloc[0]) if not placement.loc[placement['placement'].eq('商品页面')].empty else '-'}",
            "归因解释": "商品页面流量更接近浏览和竞品比较场景，购买意图弱于搜索结果页；当 CTR、CVR 和 ACOS 同时弱时，说明位置本身不适合作为主预算入口",
            "动作": "下调商品页面广告位加价",
        },
        {
            "结论": "高价值词集中在高相关、可转化且有市场热度的意图簇",
            "依据1": "搜索词中多个高相关意图簇已有订单",
            "依据2": "关键词监控显示核心词自然排名靠前",
            "依据3": "ABA 中部分相关词热度高且集中度不极端",
            "归因解释": "这类意图簇同时满足市场需求、产品相关性和转化承接，广告投入更容易沉淀到自然排名和自然订单",
            "动作": "保高相关核心词，新增长尾精确词",
        },
        {
            "结论": "搜索词不能按单一词根整体否定或整体放量",
            "依据1": "同一词根下不同年龄段、功能或场景会对应不同购买意图",
            "依据2": "搜索词聚类表已按词根、年龄段、性别、功能、场景拆分",
            "依据3": "动作需要同时看转化、ACOS、广告位和 ABA 市场热度",
            "归因解释": "底层逻辑是先拆搜索意图，再看广告承接和市场热度；词根只是聚类入口，不是判断好坏的唯一依据",
            "动作": "按聚类结果分别否定、降价、保量或小预算验证",
        },
    ]
    return pd.DataFrame(rows)


def build_execution_plan(search_clusters: pd.DataFrame | None = None, placement: pd.DataFrame | None = None) -> pd.DataFrame:
    rows = [
        {"优先级": "P0", "时间": "当天", "动作": "下调低转化、高 ACOS 的广告位", "验证指标": "低效广告位花费占比下降，整体 ACOS 下降", "原因": "广告位层如果同时出现低 CTR、低 CVR、高 ACOS，会持续放大无效点击"},
        {"优先级": "P0", "时间": "当天", "动作": "按搜索词聚类表处理错配意图簇", "验证指标": "浪费花费减少，低相关簇 ACOS 下降", "原因": "同一词根下不同年龄段、对象、功能、场景可能对应不同购买意图，不能按词根整体判断"},
        {"优先级": "P1", "时间": "1-2 天", "动作": "保留有订单、低 ACOS、自然排名可承接的核心词预算", "验证指标": "核心自然排名保持前 10，广告 ACOS 不恶化", "原因": "这类词同时具备广告转化和自然承接，预算更容易沉淀为自然权重"},
        {"优先级": "P1", "时间": "1-2 天", "动作": "把高转化搜索词拆成精确投放", "验证指标": "新精确词 3 单以上且 ACOS < 30.00%", "原因": "搜索词已有转化时，拆精确能减少宽泛匹配带来的意图噪音"},
        {"优先级": "P2", "时间": "3-5 天", "动作": "把高热度但宽泛的场景词改为更贴近产品对象的长尾词小预算测试", "验证指标": "CTR、CVR、ACOS 达标后再加预算", "原因": "场景词有需求但购买意图不一定足够短，需要用更明确的长尾词承接"},
        {"优先级": "P2", "时间": "3-7 天", "动作": "复核 Coupon、Deal、促销折扣和售价", "验证指标": "客单价、件单价、净毛利率回升", "原因": "近期订单增长但净毛利率下降时，需要确认是否由让利或成交结构变化导致"},
    ]
    if search_clusters is not None and not search_clusters.empty:
        weak = search_clusters.sort_values(["spend", "ACOS"], ascending=[False, False]).iloc[0]
        rows.insert(
            1,
            {
                "优先级": "P0",
                "时间": "当天",
                "动作": f"优先处理 {clean_text(weak.get('词根聚类'))}/{clean_text(weak.get('年龄段'))}/{clean_text(weak.get('功能属性'))} 搜索意图簇",
                "验证指标": f"该簇花费下降，CVR 从 {pct(weak.get('CVR'))} 改善，ACOS 从 {pct(weak.get('ACOS'))} 下降",
                "原因": clean_text(weak.get("归因解释", "该簇消耗和产出不匹配，需要先按意图拆分后再决定否定、降价或保留")),
            },
        )
    return pd.DataFrame(rows)


def write_excel(
    output_path: Path,
    sheets: dict[str, pd.DataFrame],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
            ws = writer.book[safe_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            header_fill = PatternFill("solid", fgColor="1F4E78")
            header_font = Font(color="FFFFFF", bold=True)
            thin = Side(style="thin", color="D9E2F3")
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(bottom=thin)
            for column_cells in ws.columns:
                max_len = 8
                column_letter = get_column_letter(column_cells[0].column)
                for cell in column_cells[:200]:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, min(len(value), 60))
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                ws.column_dimensions[column_letter].width = min(max(max_len + 2, 10), 42)


def build_detail_html_report(
    output_path: Path,
    product_intro: str,
    all_metrics: dict[str, Any],
    prev_metrics: dict[str, Any],
    recent_metrics: dict[str, Any],
    period_compare: pd.DataFrame,
    activity: pd.DataFrame,
    placement: pd.DataFrame,
    target: pd.DataFrame,
    search_actions: pd.DataFrame,
    search_clusters: pd.DataFrame,
    search_aba_cross: pd.DataFrame,
    rank_df: pd.DataFrame,
    aba_df: pd.DataFrame,
    evidence_chain: pd.DataFrame,
    execution_plan: pd.DataFrame,
) -> None:
    top_actions = search_actions.sort_values(["动作排序", "spend"], ascending=[True, False]).head(18)
    top_clusters = search_clusters.sort_values("spend", ascending=False).head(16) if not search_clusters.empty else search_clusters
    top_cross = search_aba_cross.head(18) if not search_aba_cross.empty else search_aba_cross
    html_text = f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>{LISTING_NAME} 完整运营分析报告</title>
<style>
body {{ margin:0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", Arial, sans-serif; color:#1f2933; background:#f6f8fb; }}
.wrap {{ max-width:1180px; margin:0 auto; padding:28px 28px 56px; }}
.hero {{ background:#16324f; color:white; padding:28px 32px; border-radius:8px; }}
.hero h1 {{ margin:0 0 8px; font-size:30px; letter-spacing:0; }}
.hero p {{ margin:6px 0 0; color:#d9e8f5; line-height:1.7; }}
.grid {{ display:grid; grid-template-columns: repeat(4, minmax(0,1fr)); gap:14px; margin:18px 0; }}
.card {{ background:white; border:1px solid #e5e9f0; border-radius:8px; padding:16px; }}
.kpi .label {{ color:#66788a; font-size:13px; }}
.kpi .value {{ font-size:24px; font-weight:700; margin-top:6px; }}
.section {{ background:white; border:1px solid #e5e9f0; border-radius:8px; padding:22px; margin-top:18px; }}
h2 {{ margin:0 0 12px; font-size:21px; }}
h3 {{ margin:18px 0 8px; font-size:16px; }}
p, li {{ line-height:1.75; }}
.badge {{ display:inline-block; padding:3px 8px; border-radius:999px; background:#e8f2ff; color:#184e7a; font-size:12px; margin-right:6px; }}
table {{ width:100%; border-collapse:collapse; margin-top:10px; font-size:13px; }}
th {{ background:#eef3f8; color:#22384c; text-align:left; }}
th, td {{ border:1px solid #dbe3eb; padding:8px 9px; vertical-align:top; }}
.warn {{ background:#fff7ed; border-left:4px solid #f59e0b; padding:12px 14px; margin:12px 0; }}
.good {{ background:#eefbf3; border-left:4px solid #2f9e44; padding:12px 14px; margin:12px 0; }}
.bad {{ background:#fff1f2; border-left:4px solid #e11d48; padding:12px 14px; margin:12px 0; }}
@media (max-width:900px) {{ .grid {{ grid-template-columns: repeat(2, minmax(0,1fr)); }} }}
@media (max-width:560px) {{ .grid {{ grid-template-columns: 1fr; }} .wrap {{ padding:16px; }} }}
</style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <h1>{LISTING_NAME} 完整运营分析报告</h1>
    <p>核心判断：订单增长，但净毛利和广告效率承压。当前主线不是无脑扩量，而是控浪费、保核心排名、把预算迁移到更相关且有证据链支持的搜索词。</p>
    <p>排名口径：数字越小排名越好；报告统一使用最佳排名、平均排名、最差排名。</p>
  </div>

  <div class="grid">
    <div class="card kpi"><div class="label">近完整周期销售额</div><div class="value">{money(all_metrics['销售额'])}</div></div>
    <div class="card kpi"><div class="label">近期订单变化</div><div class="value">{pct(period_compare.loc[period_compare['指标'].eq('订单量'), '变化率'].iloc[0])}</div></div>
    <div class="card kpi"><div class="label">近期净毛利率</div><div class="value">{pct(recent_metrics['销售净毛利率'])}</div></div>
    <div class="card kpi"><div class="label">近期广告 ACOS</div><div class="value">{pct(recent_metrics['广告ACOS'])}</div></div>
  </div>

  <div class="section">
    <h2>1. Listing 经营状态</h2>
    <p>{PRODUCT_AGE_NOTE}</p>
    <div class="warn">近周期订单增长 {pct(period_compare.loc[period_compare['指标'].eq('订单量'), '变化率'].iloc[0])}，但销售额只增长 {pct(period_compare.loc[period_compare['指标'].eq('销售额'), '变化率'].iloc[0])}，净毛利率从 {pct(prev_metrics['销售净毛利率'])} 降至 {pct(recent_metrics['销售净毛利率'])}。这三个指标共同说明增长质量下降。</div>
    {html_table(period_compare, ['指标','前期','近期','变化率'], 18)}
  </div>

  <div class="section">
    <h2>2. 客单价与促销判断</h2>
    <p>促销判断不是只看 Coupon 或 Deal，而是同时看客单价、件单价、促销费用率、毛利率和净毛利率。</p>
    <div class="bad">近期客单价 {money(recent_metrics['客单价'])}，前期 {money(prev_metrics['客单价'])}；近期件单价 {money(recent_metrics['件单价'])}，前期 {money(prev_metrics['件单价'])}；近期净毛利率 {pct(recent_metrics['销售净毛利率'])}。如果客单价/件单价下降且促销费用率或利润率恶化，需要复核售价、Coupon、Deal 和折扣。</div>
  </div>

  <div class="section">
    <h2>3. 广告整体与活动层</h2>
    <p>活动判断至少联看花费、订单、ACOS、CVR、ROAS、前后期变化，不用单一 ACOS 下结论。</p>
    {html_table(activity.sort_values('spend', ascending=False), ['campaign','spend','sales','orders','CVR','ACOS','ROAS','动作建议','归因解释','判断依据'], 12)}
  </div>

  <div class="section">
    <h2>4. 广告位层</h2>
    {html_table(placement.sort_values('spend', ascending=False), ['placement','spend','sales','orders','CTR','CVR','ACOS','动作建议','归因解释','判断依据'], 8)}
  </div>

  <div class="section">
    <h2>5. 投放与搜索词动作</h2>
    <h3>投放层</h3>
    {html_table(target.sort_values('spend', ascending=False), ['campaign','target','match','spend','sales','orders','CVR','ACOS','相关性分类','动作建议','归因解释','判断依据'], 14)}
    <h3>搜索词动作</h3>
    {html_table(top_actions, ['term','campaign','target','match','spend','sales','orders','CVR','ACOS','相关性分类','表现分类','动作建议','归因解释','判断依据'], 18)}
  </div>

  <div class="section">
    <h2>6. 搜索词聚类与 ABA 交叉归因</h2>
    <p>这里不再把搜索词动作和用户搜索词归类分析做成同一张表。搜索词动作解决单词怎么调，聚类分析解决共同词根、年龄段、功能和场景到底哪里在赚钱、哪里在浪费，并联看 ABA 市场热度和广告位归因。</p>
    <h3>用户搜索词归类分析</h3>
    {html_table(top_clusters, ['词根聚类','年龄段','性别','功能属性','场景','spend','sales','orders','CVR','ACOS','词根ABA代表词','词根ABA最佳排名','主要广告位','聚类结论','动作建议','归因解释','判断依据'], 16)}
    <h3>搜索词 ABA 交叉分析</h3>
    {html_table(top_cross, ['词根聚类','年龄段','广告覆盖状态','spend','orders','CVR','ACOS','词根ABA代表词','词根ABA最佳排名','词根ABA头部点击份额','市场-广告关系','动作建议','归因解释'], 18)}
  </div>

  <div class="section">
    <h2>7. 关键词排名承接</h2>
    <p>这里用排名监控验证广告是否有自然承接。排名数字越小越好。</p>
    {html_table(rank_df, ['keyword','相关性分类','最佳自然排名','平均自然排名','最佳广告排名','平均广告排名','动作建议','归因解释','判断依据'], 16)}
  </div>

  <div class="section">
    <h2>8. ABA 市场机会</h2>
    <p>ABA 只筛与儿童太阳镜或明确儿童海滩场景相关的词。高热度但相关性弱或头部集中度过高的词，不作为主线扩量。</p>
    {html_table(aba_df, ['keyword','搜索量排名','排名变化方向','变化名次','前三ASIN点击份额','前三ASIN转化份额','相关性分类','当前广告覆盖','动作建议','归因解释','判断依据'], 18)}
  </div>

  <div class="section">
    <h2>9. 结论证据链</h2>
    {html_table(evidence_chain, ['结论','依据1','依据2','依据3','归因解释','动作'], 10)}
  </div>

  <div class="section">
    <h2>10. 未来 3-7 天执行计划</h2>
    {html_table(execution_plan, ['优先级','时间','动作','验证指标','原因'], 12)}
  </div>

  <div class="section">
    <h2>产品信息</h2>
    <pre>{html.escape(product_intro[:1800])}</pre>
  </div>
</div>
</body>
</html>"""
    output_path.write_text(html_text, encoding="utf-8")


def build_html_report(
    output_path: Path,
    product_intro: str,
    all_metrics: dict[str, Any],
    prev_metrics: dict[str, Any],
    recent_metrics: dict[str, Any],
    period_compare: pd.DataFrame,
    activity: pd.DataFrame,
    placement: pd.DataFrame,
    target: pd.DataFrame,
    search_actions: pd.DataFrame,
    search_clusters: pd.DataFrame,
    search_aba_cross: pd.DataFrame,
    rank_df: pd.DataFrame,
    aba_df: pd.DataFrame,
    evidence_chain: pd.DataFrame,
    execution_plan: pd.DataFrame,
) -> None:
    def metric_change(metric: str) -> Any:
        row = period_compare[period_compare["指标"].eq(metric)]
        return None if row.empty else row.iloc[0]["变化率"]

    def plain(value: Any, fallback: str = "-") -> str:
        text = clean_text(value)
        return text if text else fallback

    def number(value: Any) -> str:
        return "-" if value is None or pd.isna(value) else num_text(value, 0)

    def cards(rows: list[dict[str, Any]], cls: str = "") -> str:
        if not rows:
            rows = [{"title": "暂无明确动作", "meta": "数据没有形成强信号", "body": "继续观察，不做激进预算调整。", "action": "动作：保守测试"}]
        blocks = []
        for row in rows:
            blocks.append(
                f"""<div class="insight {cls}">
  <div class="insight-title">{html.escape(plain(row.get('title')))}</div>
  <div class="insight-meta">{html.escape(plain(row.get('meta')))}</div>
  <p>{html.escape(plain(row.get('body')))}</p>
  <div class="next-action">{html.escape(plain(row.get('action')))}</div>
</div>"""
            )
        return "".join(blocks)

    evidence_rows = [
        {
            "title": r.get("结论", ""),
            "meta": " / ".join([plain(r.get("依据1"), ""), plain(r.get("依据2"), ""), plain(r.get("依据3"), "")]).strip(" /"),
            "body": r.get("归因解释", ""),
            "action": f"动作：{plain(r.get('动作'))}",
        }
        for _, r in evidence_chain.head(5).iterrows()
    ]

    problem_rows: list[dict[str, Any]] = [
        {
            "title": "增长质量弱于订单增长",
            "meta": f"订单变化 {pct(metric_change('订单量'))}，销售额变化 {pct(metric_change('销售额'))}，净毛利率 {pct(recent_metrics['销售净毛利率'])}",
            "body": "订单和销售额没有同步放大，且利润率承压，说明当前不是单纯缺流量，而是增量订单的价格、促销或广告效率需要先校正。",
            "action": "动作：先控浪费，再放量。",
        },
        {
            "title": "客单价和利润结构需要复核",
            "meta": f"近期客单价 {money(recent_metrics['客单价'])}，件单价 {money(recent_metrics['件单价'])}",
            "body": "客单价、件单价和净毛利率要一起看。若订单增长伴随件单价或净利率走低，优先排查 Coupon、Deal、售价和低价变体占比。",
            "action": "动作：复核促销、价格和利润结构。",
        },
    ]
    if not placement.empty:
        weak = placement.sort_values(["ACOS", "spend"], ascending=[False, False]).iloc[0]
        problem_rows.append(
            {
                "title": f"{plain(weak.get('placement'))} 广告位拖累效率",
                "meta": f"花费 {money(weak.get('spend'))}，CVR {pct(weak.get('CVR'))}，ACOS {pct(weak.get('ACOS'))}",
                "body": weak.get("归因解释", ""),
                "action": f"动作：{plain(weak.get('动作建议'))}",
            }
        )
    if not search_clusters.empty:
        weak_cluster = search_clusters.sort_values(["spend", "ACOS"], ascending=[False, False]).iloc[0]
        problem_rows.append(
            {
                "title": f"{plain(weak_cluster.get('词根聚类'))} / {plain(weak_cluster.get('年龄段'))} 需要收口",
                "meta": f"花费 {money(weak_cluster.get('spend'))}，订单 {number(weak_cluster.get('orders'))}，ACOS {pct(weak_cluster.get('ACOS'))}",
                "body": weak_cluster.get("归因解释", ""),
                "action": f"动作：{plain(weak_cluster.get('动作建议'))}",
            }
        )

    opportunity_rows: list[dict[str, Any]] = []
    if not search_clusters.empty:
        good = search_clusters[
            (search_clusters["orders"] >= 5)
            & (search_clusters["ACOS"].notna())
            & (search_clusters["ACOS"] <= 0.35)
        ].sort_values(["orders", "spend"], ascending=False).head(3)
        for _, r in good.iterrows():
            opportunity_rows.append(
                {
                    "title": f"{plain(r.get('词根聚类'))} 是优先承接方向",
                    "meta": f"订单 {number(r.get('orders'))}，CVR {pct(r.get('CVR'))}，ACOS {pct(r.get('ACOS'))}，ABA排名 {number(r.get('词根ABA最佳排名'))}",
                    "body": r.get("归因解释", ""),
                    "action": f"动作：{plain(r.get('动作建议'))}",
                }
            )
    if not aba_df.empty and len(opportunity_rows) < 3:
        aba_opps = aba_df[
            (aba_df["搜索量排名"].notna())
            & (aba_df["搜索量排名"] <= 100000)
            & (aba_df["相关性分类"].ne("低相关词"))
        ].sort_values("搜索量排名").head(3 - len(opportunity_rows))
        for _, r in aba_opps.iterrows():
            opportunity_rows.append(
                {
                    "title": f"ABA 机会词：{plain(r.get('keyword'))}",
                    "meta": f"搜索量排名 {number(r.get('搜索量排名'))}，前三点击份额 {pct(r.get('前三ASIN点击份额'))}",
                    "body": r.get("归因解释", ""),
                    "action": f"动作：{plain(r.get('动作建议'))}",
                }
            )

    ad_rows: list[dict[str, Any]] = []
    for df, label, name_col in [(placement, "广告位", "placement"), (activity, "活动", "campaign")]:
        if not df.empty:
            for _, r in df.sort_values("spend", ascending=False).head(2).iterrows():
                ad_rows.append(
                    {
                        "title": f"{label}：{plain(r.get(name_col))}",
                        "meta": f"花费 {money(r.get('spend'))}，订单 {number(r.get('orders'))}，ACOS {pct(r.get('ACOS'))}",
                        "body": r.get("归因解释", ""),
                        "action": f"动作：{plain(r.get('动作建议'))}",
                    }
                )
    if not search_actions.empty:
        for _, r in search_actions.sort_values(["动作排序", "spend"], ascending=[True, False]).head(3).iterrows():
            ad_rows.append(
                {
                    "title": f"搜索词：{plain(r.get('term'))}",
                    "meta": f"花费 {money(r.get('spend'))}，订单 {number(r.get('orders'))}，ACOS {pct(r.get('ACOS'))}",
                    "body": r.get("归因解释", ""),
                    "action": f"动作：{plain(r.get('动作建议'))}",
                }
            )

    keyword_rows: list[dict[str, Any]] = []
    if not search_aba_cross.empty:
        for _, r in search_aba_cross.head(4).iterrows():
            keyword_rows.append(
                {
                    "title": plain(r.get("词根聚类")),
                    "meta": f"{plain(r.get('市场-广告关系'))}；ABA排名 {number(r.get('词根ABA最佳排名'))}",
                    "body": r.get("归因解释", ""),
                    "action": f"动作：{plain(r.get('动作建议'))}",
                }
            )
    if not rank_df.empty:
        for _, r in rank_df.sort_values(["最佳自然排名", "平均自然排名"], ascending=True).head(3).iterrows():
            keyword_rows.append(
                {
                    "title": f"排名词：{plain(r.get('keyword'))}",
                    "meta": f"最佳自然第 {number(r.get('最佳自然排名'))}，平均自然第 {number(r.get('平均自然排名'))}",
                    "body": r.get("归因解释", ""),
                    "action": f"动作：{plain(r.get('动作建议'))}",
                }
            )

    plan_rows = [
        {
            "title": f"{plain(r.get('优先级'))} / {plain(r.get('时间'))}",
            "meta": r.get("动作", ""),
            "body": r.get("原因", ""),
            "action": f"验证：{plain(r.get('验证指标'))}",
        }
        for _, r in execution_plan.head(7).iterrows()
    ]

    html_text = f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>{LISTING_NAME} 运营洞察报告</title>
<style>
body {{ margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",Arial,sans-serif; color:#17212b; background:#f5f7fa; }}
.wrap {{ max-width:1160px; margin:0 auto; padding:28px 28px 54px; }}
.hero {{ background:#14324a; color:#fff; border-radius:8px; padding:30px 34px; }}
.hero h1 {{ margin:0 0 10px; font-size:30px; letter-spacing:0; }}
.hero p {{ margin:7px 0 0; color:#d8e7f3; line-height:1.75; max-width:960px; }}
.kpis {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:16px 0 4px; }}
.kpi {{ background:#fff; border:1px solid #e0e6ed; border-radius:8px; padding:15px; }}
.kpi .label {{ color:#617487; font-size:13px; }}
.kpi .value {{ font-size:23px; font-weight:700; margin-top:6px; }}
.section {{ background:#fff; border:1px solid #e0e6ed; border-radius:8px; padding:22px; margin-top:16px; }}
.section h2 {{ margin:0 0 12px; font-size:21px; }}
.section-lead {{ color:#526579; line-height:1.75; margin:0 0 14px; }}
.insight-grid {{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }}
.insight {{ border:1px solid #dfe7ef; border-left:4px solid #2f6f9f; border-radius:8px; padding:14px 15px; background:#fbfdff; }}
.insight-title {{ font-size:16px; font-weight:700; margin-bottom:6px; }}
.insight-meta {{ color:#5f7286; font-size:13px; line-height:1.55; margin-bottom:8px; }}
.insight p {{ margin:0; line-height:1.7; }}
.next-action {{ margin-top:10px; color:#173f5f; font-weight:700; line-height:1.6; }}
.problem {{ border-left-color:#d9480f; background:#fffaf6; }}
.opportunity {{ border-left-color:#2f9e44; background:#f6fff8; }}
.note {{ background:#eef5fb; border:1px solid #d7e6f2; border-radius:8px; padding:13px 15px; color:#27465f; line-height:1.7; }}
@media (max-width:900px) {{ .kpis,.insight-grid {{ grid-template-columns:repeat(2,minmax(0,1fr)); }} }}
@media (max-width:620px) {{ .wrap {{ padding:16px; }} .kpis,.insight-grid {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <h1>{LISTING_NAME} 运营洞察报告</h1>
    <p>执行摘要：当前 Listing 有订单增长，但增长质量被净毛利、广告位效率和部分搜索词承接拖累。下一步不建议直接扩预算，而是先把浪费广告位和错配年龄词收口，再把预算迁移到有转化、ABA 有热度、自然排名能承接的词根。</p>
    <p>明细追溯请看 Excel：HTML 只保留结论、洞察和动作；活动、广告位、投放、搜索词、ABA 的逐行过程已放在分析过程清单中。</p>
  </div>

  <div class="kpis">
    <div class="kpi"><div class="label">近完整周期销售额</div><div class="value">{money(all_metrics['销售额'])}</div></div>
    <div class="kpi"><div class="label">近期订单变化</div><div class="value">{pct(metric_change('订单量'))}</div></div>
    <div class="kpi"><div class="label">近期净毛利率</div><div class="value">{pct(recent_metrics['销售净毛利率'])}</div></div>
    <div class="kpi"><div class="label">近期广告 ACOS</div><div class="value">{pct(recent_metrics['广告ACOS'])}</div></div>
  </div>

  <div class="section">
    <h2>1. 执行摘要</h2>
    <p class="section-lead">{PRODUCT_AGE_NOTE} 排名口径：数字越小排名越好；报告统一使用最佳排名、平均排名、最差排名。</p>
    <div class="note">结论：当前重点不是继续买更多流量，而是提高流量质量。订单增长、销售额增长、净毛利率和广告 ACOS 需要一起判断；只看销量会高估经营状态，只看 ACOS 又会漏掉自然排名和市场词机会。</div>
  </div>

  <div class="section">
    <h2>2. 关键洞察</h2>
    <p class="section-lead">以下结论来自 Listing、广告、搜索词、排名和 ABA 的交叉判断，不使用单一指标定性。</p>
    <div class="insight-grid">{cards(evidence_rows)}</div>
  </div>

  <div class="section">
    <h2>3. 当前核心问题</h2>
    <p class="section-lead">先处理会持续吞预算、压利润或误导扩量判断的问题。</p>
    <div class="insight-grid">{cards(problem_rows, 'problem')}</div>
  </div>

  <div class="section">
    <h2>4. 当前增长机会</h2>
    <p class="section-lead">只把同时具备广告承接、市场热度或排名承接的方向列为机会。</p>
    <div class="insight-grid">{cards(opportunity_rows, 'opportunity')}</div>
  </div>

  <div class="section">
    <h2>5. 广告优化动作</h2>
    <p class="section-lead">这里列方向和原因，不展开逐活动、逐广告位或逐词明细。</p>
    <div class="insight-grid">{cards(ad_rows)}</div>
  </div>

  <div class="section">
    <h2>6. 搜索词聚类与 ABA 交叉归因</h2>
    <p class="section-lead">搜索词动作解决单个词怎么调；聚类与 ABA 交叉解决共同词根、年龄段、功能和场景是否值得继续买。</p>
    <div class="insight-grid">{cards(keyword_rows)}</div>
  </div>

  <div class="section">
    <h2>7. 未来 3-7 天执行重点</h2>
    <p class="section-lead">按优先级执行，先控损，再迁移预算，再验证新机会。</p>
    <div class="insight-grid">{cards(plan_rows)}</div>
  </div>

  <div class="section">
    <h2>8. 明细追溯</h2>
    <p class="section-lead">明细追溯请看 Excel。里面保留了 Listing 经营诊断、广告活动、广告位、投放、搜索词动作、用户搜索词明细归因、用户搜索词归类分析、搜索词 ABA 交叉分析、关键词排名承接和 ABA 机会词。</p>
  </div>
</div>
</body>
</html>"""
    output_path.write_text(html_text, encoding="utf-8")


def main() -> int:
    root = Path.cwd()
    paths = resolve_paths(root)
    paths.output_dir.mkdir(parents=True, exist_ok=True)

    product_intro = read_product_intro(paths.product_file)
    listing_daily = read_listing(paths.listing_file)
    listing_full, incomplete_note = exclude_incomplete_last_day(listing_daily)
    window_days = max(1, min(7, len(listing_full) // 2))
    recent_df = listing_full.tail(window_days)
    prev_df = listing_full.iloc[-window_days * 2 : -window_days]
    all_metrics = period_metrics(listing_full, "完整周期")
    prev_metrics = period_metrics(prev_df, "前期")
    recent_metrics = period_metrics(recent_df, "近期")
    period_compare = compare_periods(prev_metrics, recent_metrics)
    listing_diag = pd.DataFrame([all_metrics, prev_metrics, recent_metrics])

    recent_start = pd.Timestamp(recent_df["date"].min())
    prev_start = pd.Timestamp(prev_df["date"].min())
    prev_end = pd.Timestamp(prev_df["date"].max())

    activity_raw = read_ad_table(paths.activity_file, "activity")
    placement_raw = read_ad_table(paths.placement_file, "placement")
    target_raw = read_ad_table(paths.target_file, "target")
    search_raw = read_ad_table(paths.search_file, "search")

    activity_recent = aggregate_ad(activity_raw, ["campaign"], recent_start).sort_values("spend", ascending=False)
    activity_prev = aggregate_ad(activity_raw[(activity_raw["date"] >= prev_start) & (activity_raw["date"] <= prev_end)], ["campaign"], prev_start)
    activity_recent = add_activity_decision(activity_recent, activity_prev)

    placement_recent = aggregate_ad(placement_raw, ["placement"], recent_start).sort_values("spend", ascending=False)
    placement_recent = add_placement_decision(placement_recent)

    placement_campaign = aggregate_ad(placement_raw, ["campaign", "placement"], recent_start)
    placement_campaign = add_placement_decision(placement_campaign).sort_values("spend", ascending=False)

    target_recent = aggregate_ad(target_raw, ["campaign", "target", "match"], recent_start)
    target_recent = add_target_decision(target_recent).sort_values("spend", ascending=False)

    rank_df = read_rank_table(paths.rank_file)
    current_terms = {clean_text(x).lower() for x in pd.concat([target_raw.get("target", pd.Series(dtype=str)), search_raw.get("term", pd.Series(dtype=str))]).dropna()}
    aba_df = read_aba(paths.aba_file, current_terms)

    search_recent = aggregate_ad(search_raw, ["term", "target", "match", "campaign"], recent_start)
    search_recent = add_search_aba_columns(search_recent, aba_df)
    search_classes = search_recent.apply(lambda r: pd.Series(classify_search_term(r)), axis=1)
    search_recent = pd.concat([search_recent, search_classes], axis=1)
    action_order = {"否定": 1, "暂停": 2, "降价": 3, "转精确": 4, "低预算测试": 5, "保持": 6, "观察": 7}
    search_recent["动作排序"] = search_recent["动作建议"].map(action_order).fillna(99)
    search_actions = search_recent[
        (search_recent["spend"] >= 5)
        | (search_recent["orders"] >= 3)
        | (search_recent["动作建议"].isin(["转精确", "否定", "暂停", "降价"]))
    ].sort_values(["动作排序", "spend", "orders"], ascending=[True, False, False])
    search_clusters = build_search_clusters(search_recent, aba_df, placement_campaign)
    search_aba_cross = build_search_aba_cross(search_recent, search_clusters, aba_df)

    source_check = build_source_check(
        paths,
        {
            "Listing经营数据": listing_daily,
            "广告活动数据": activity_raw,
            "广告位数据": placement_raw,
            "投放数据": target_raw,
            "用户搜索词数据": search_raw,
            "关键词排名监控": rank_df,
            "ABA Top30万搜索词": aba_df,
        },
        incomplete_note,
    )

    method_notes = pd.DataFrame(
        [
            {"口径": "近期窗口", "说明": f"{recent_metrics['开始日期']} 至 {recent_metrics['结束日期']}，共 {recent_metrics['天数']} 天"},
            {"口径": "前期窗口", "说明": f"{prev_metrics['开始日期']} 至 {prev_metrics['结束日期']}，共 {prev_metrics['天数']} 天"},
            {"口径": "排名", "说明": "数字越小排名越好，报告使用最佳/平均/最差排名"},
            {"口径": "广告花费", "说明": "销售表现表广告花费为负数记账，脚本按绝对值计算 ACOS/ROAS"},
            {"口径": "动作判断", "说明": "每条动作至少尝试联看规模、效率、漏斗、利润、排名、ABA、相关性，不用单一指标定性"},
            {"口径": "产品定位", "说明": PRODUCT_AGE_NOTE},
        ]
    )

    promo_judgement = pd.DataFrame(
        [
            {"指标": "客单价", "前期": prev_metrics["客单价"], "近期": recent_metrics["客单价"], "变化率": period_compare.loc[period_compare["指标"].eq("客单价"), "变化率"].iloc[0]},
            {"指标": "件单价", "前期": prev_metrics["件单价"], "近期": recent_metrics["件单价"], "变化率": period_compare.loc[period_compare["指标"].eq("件单价"), "变化率"].iloc[0]},
            {"指标": "促销费用率", "前期": prev_metrics["促销费用率"], "近期": recent_metrics["促销费用率"], "变化率": period_compare.loc[period_compare["指标"].eq("促销费用率"), "变化率"].iloc[0]},
            {"指标": "销售毛利率", "前期": prev_metrics["销售毛利率"], "近期": recent_metrics["销售毛利率"], "变化率": period_compare.loc[period_compare["指标"].eq("销售毛利率"), "变化率"].iloc[0]},
            {"指标": "销售净毛利率", "前期": prev_metrics["销售净毛利率"], "近期": recent_metrics["销售净毛利率"], "变化率": period_compare.loc[period_compare["指标"].eq("销售净毛利率"), "变化率"].iloc[0]},
        ]
    )
    promo_judgement["判断"] = promo_judgement.apply(
        lambda r: "需复核促销/售价" if r["指标"] in {"客单价", "件单价", "销售净毛利率"} and (r["变化率"] or 0) < -0.05 else "观察",
        axis=1,
    )
    promo_judgement["归因解释"] = promo_judgement.apply(
        lambda r: (
            "该指标近期明显下降，说明增长可能来自更低成交价格、促销让利或低价变体占比提升，需要回到价格和促销设置核对"
            if r["判断"] == "需复核促销/售价"
            else "该指标没有单独构成促销异常，需要和客单价、件单价、利润率一起判断"
        ),
        axis=1,
    )

    evidence_chain = build_evidence_chain(
        recent_metrics,
        prev_metrics,
        period_compare,
        activity_recent,
        placement_recent,
        search_actions,
        rank_df,
        aba_df,
    )
    execution_plan = build_execution_plan(search_clusters, placement_recent)

    excel_path = paths.output_dir / f"{LISTING_NAME}分析过程清单.xlsx"
    html_path = paths.output_dir / f"{LISTING_NAME}完整运营分析报告.html"

    sheets = {
        "数据源校验": source_check,
        "口径说明": method_notes,
        "Listing经营诊断": listing_diag,
        "客单价促销判断": promo_judgement,
        "广告整体效率": pd.DataFrame([recent_metrics]),
        "活动层判断": activity_recent,
        "广告位判断": placement_recent,
        "广告位活动拆解": placement_campaign,
        "投放层判断": target_recent,
        "搜索词动作": search_actions,
        "用户搜索词归类分析": search_clusters,
        "用户搜索词明细归因": search_recent.sort_values(["动作排序", "spend"], ascending=[True, False]),
        "搜索词ABA交叉分析": search_aba_cross,
        "关键词排名承接": rank_df,
        "ABA机会词": aba_df.head(250),
        "结论证据链": evidence_chain,
        "3-7天执行计划": execution_plan,
    }
    write_excel(excel_path, sheets)
    build_html_report(
        html_path,
        product_intro,
        all_metrics,
        prev_metrics,
        recent_metrics,
        period_compare,
        activity_recent,
        placement_recent,
        target_recent,
        search_actions,
        search_clusters,
        search_aba_cross,
        rank_df,
        aba_df,
        evidence_chain,
        execution_plan,
    )

    print(f"HTML_REPORT={html_path}")
    print(f"EXCEL_PROCESS={excel_path}")
    print(f"OUTPUT_DIR={paths.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
