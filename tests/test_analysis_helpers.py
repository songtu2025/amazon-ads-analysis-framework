import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from analyze_rbk004_operating_plan import (
    build_search_clusters,
    build_html_report,
    classify_relevance,
    classify_search_term,
    describe_rank,
    extract_search_features,
    html_table,
    is_source_excel,
    pct,
    safe_div,
)


def zh(*codes):
    return "".join(chr(code) for code in codes)


COL_RELEVANCE = zh(0x76F8, 0x5173, 0x6027, 0x5206, 0x7C7B)
COL_PERFORMANCE = zh(0x8868, 0x73B0, 0x5206, 0x7C7B)
COL_ACTION = zh(0x52A8, 0x4F5C, 0x5EFA, 0x8BAE)
COL_EVIDENCE_COUNT = zh(0x8BC1, 0x636E, 0x6570, 0x91CF)
COL_CAUSE = zh(0x5F52, 0x56E0, 0x89E3, 0x91CA)


class AnalysisHelperTests(unittest.TestCase):
    def test_describe_rank_uses_best_average_worst_wording(self):
        text = describe_rank(2, 6.24, 18)
        self.assertIn(zh(0x6700, 0x4F73, 0x81EA, 0x7136, 0x6392, 0x540D, 0x7B2C) + " 2", text)
        self.assertIn(zh(0x5E73, 0x5747, 0x81EA, 0x7136, 0x6392, 0x540D, 0x7B2C) + " 6.2", text)
        self.assertIn(zh(0x6700, 0x5DEE, 0x81EA, 0x7136, 0x6392, 0x540D, 0x7B2C) + " 18", text)
        self.assertNotIn(zh(0x6700, 0x4F4E, 0x7B2C), text)

    def test_safe_div_calculates_average_order_value(self):
        self.assertEqual(safe_div(52481.12, 4663), 11.254797341)

    def test_baby_age_segment_is_not_blanket_low_relevance(self):
        self.assertEqual(extract_search_features("baby sunglasses 2-4")["age_segment"], "2-4")
        self.assertNotEqual(classify_relevance("baby sunglasses 2-4"), zh(0x4F4E, 0x76F8, 0x5173, 0x8BCD))
        self.assertEqual(classify_relevance("baby sunglasses 0-2"), zh(0x4F4E, 0x76F8, 0x5173, 0x8BCD))

    def test_baby_month_like_range_is_low_relevance(self):
        self.assertEqual(extract_search_features("baby sunglasses 18-24")["age_segment"], "0-2")
        self.assertEqual(classify_relevance("baby sunglasses 18-24"), zh(0x4F4E, 0x76F8, 0x5173, 0x8BCD))

    def test_classify_search_term_uses_relevance_and_multi_metric_action(self):
        row = {
            "term": "baby sunglasses 0-2",
            "clicks": 14,
            "orders": 1,
            "spend": 29.91,
            "sales": 8.98,
            "acos": 3.331,
            "cvr": 1 / 14,
            "aba_rank": 7566,
        }
        result = classify_search_term(row)
        self.assertEqual(result[COL_RELEVANCE], zh(0x4F4E, 0x76F8, 0x5173, 0x8BCD))
        self.assertEqual(result[COL_PERFORMANCE], zh(0x9AD8, 0x82B1, 0x8D39, 0x4F4E, 0x4EA7, 0x51FA))
        self.assertIn(result[COL_ACTION], {zh(0x6682, 0x505C), zh(0x964D, 0x4EF7)})
        self.assertGreaterEqual(result[COL_EVIDENCE_COUNT], 3)
        self.assertIn(COL_CAUSE, result)
        self.assertIn(zh(0x5E74, 0x9F84), result[COL_CAUSE])
        self.assertIn(zh(0x8F6C, 0x5316), result[COL_CAUSE])

    def test_pct_uses_two_decimal_places_by_default(self):
        self.assertEqual(pct(0.123456), "12.35%")

    def test_html_table_formats_ratio_metric_values_as_percentages(self):
        df = pd.DataFrame(
            [
                {
                    zh(0x6307, 0x6807): zh(0x5E7F, 0x544A) + "ACOS",
                    zh(0x524D, 0x671F): 0.2809281176,
                    zh(0x8FD1, 0x671F): 0.3557109212,
                    zh(0x53D8, 0x5316, 0x7387): 0.2661990700,
                }
            ]
        )
        rendered = html_table(df, [zh(0x6307, 0x6807), zh(0x524D, 0x671F), zh(0x8FD1, 0x671F), zh(0x53D8, 0x5316, 0x7387)])
        self.assertIn("28.09%", rendered)
        self.assertIn("35.57%", rendered)
        self.assertIn("26.62%", rendered)

    def test_is_source_excel_ignores_excel_lock_files(self):
        self.assertFalse(is_source_excel("~$ad-report.xlsx"))
        self.assertTrue(is_source_excel("ad-report.xlsx"))

    def test_search_cluster_table_contains_aba_and_feature_columns(self):
        search = pd.DataFrame(
            [
                {"term": "baby sunglasses 2-4", "campaign": "c1", "target": "baby sunglasses", "match": "broad", "impr": 100, "clicks": 10, "spend": 10.0, "sales": 50.0, "orders": 5},
                {"term": "baby sunglasses 0-2", "campaign": "c1", "target": "baby sunglasses", "match": "broad", "impr": 100, "clicks": 10, "spend": 20.0, "sales": 10.0, "orders": 1},
            ]
        )
        aba = pd.DataFrame(
            [
                {"keyword": "baby sunglasses", zh(0x641C, 0x7D22, 0x91CF, 0x6392, 0x540D): 7566, zh(0x524D, 0x4E09) + "ASIN" + zh(0x70B9, 0x51FB, 0x4EFD, 0x989D): 0.366},
            ]
        )
        clusters = build_search_clusters(search, aba, pd.DataFrame())
        self.assertIn(zh(0x8BCD, 0x6839, 0x805A, 0x7C7B), clusters.columns)
        self.assertIn(zh(0x5E74, 0x9F84, 0x6BB5), clusters.columns)
        self.assertIn(zh(0x8BCD, 0x6839) + "ABA" + zh(0x6700, 0x4F73, 0x6392, 0x540D), clusters.columns)
        self.assertEqual(len(clusters), 2)

    def test_html_report_is_insight_led_not_detail_dump(self):
        with TemporaryDirectory() as tmp:
            output = Path(tmp) / "report.html"
            period_compare = pd.DataFrame(
                [
                    {"指标": "订单量", "前期": 10, "近期": 13, "变化率": 0.3},
                    {"指标": "销售额", "前期": 100, "近期": 110, "变化率": 0.1},
                ]
            )
            metrics = {"销售额": 110.0, "销售净毛利率": 0.18, "广告ACOS": 0.36, "客单价": 10.0, "件单价": 9.0}
            activity = pd.DataFrame([{"campaign": "c1", "spend": 100.0, "sales": 300.0, "orders": 20, "CVR": 0.25, "ACOS": 0.3333, "ROAS": 3.0, "动作建议": "控价保量", "归因解释": "订单有规模但利润承压", "判断依据": "ACOS 33.33%"}])
            placement = pd.DataFrame([{"placement": "商品页面", "spend": 80.0, "sales": 120.0, "orders": 8, "CTR": 0.02, "CVR": 0.1, "ACOS": 0.6667, "动作建议": "下调广告位", "归因解释": "商品页面承接弱", "判断依据": "ACOS 66.67%"}])
            target = pd.DataFrame([{"campaign": "c1", "target": "kids sunglasses", "match": "exact", "spend": 50.0, "sales": 200.0, "orders": 10, "CVR": 0.4, "ACOS": 0.25, "相关性分类": "核心词", "动作建议": "保量", "归因解释": "核心词转化好", "判断依据": "CVR 40.00%"}])
            search_actions = pd.DataFrame([{"term": "baby sunglasses 0-2", "campaign": "c1", "target": "baby sunglasses", "match": "broad", "spend": 20.0, "sales": 10.0, "orders": 1, "CVR": 0.05, "ACOS": 2.0, "相关性分类": "低相关词", "表现分类": "高花费低产出", "动作建议": "暂停", "归因解释": "年龄错配且转化弱", "判断依据": "ACOS 200.00%", "动作排序": 1}])
            search_clusters = pd.DataFrame([{"词根聚类": "baby sunglasses", "年龄段": "0-2", "性别": "通用", "功能属性": "基础", "场景": "无", "spend": 20.0, "sales": 10.0, "orders": 1, "CVR": 0.05, "ACOS": 2.0, "词根ABA代表词": "baby sunglasses", "词根ABA最佳排名": 7566, "主要广告位": "商品页面", "聚类结论": "低相关浪费簇", "动作建议": "否定或暂停", "归因解释": "年龄错配", "判断依据": "ABA排名 7566"}])
            search_aba_cross = pd.DataFrame([{"词根聚类": "baby sunglasses", "年龄段": "0-2", "广告覆盖状态": "已产生广告搜索词", "spend": 20.0, "orders": 1, "CVR": 0.05, "ACOS": 2.0, "词根ABA代表词": "baby sunglasses", "词根ABA最佳排名": 7566, "词根ABA头部点击份额": 0.36, "市场-广告关系": "广告表现差 + ABA高价值", "动作建议": "降价、收窄匹配", "归因解释": "市场词没错但承接弱"}])
            rank_df = pd.DataFrame([{"keyword": "kids sunglasses", "相关性分类": "核心词", "最佳自然排名": 2, "平均自然排名": 6, "最佳广告排名": 1, "平均广告排名": 4, "动作建议": "保位", "归因解释": "自然靠前", "判断依据": "最佳第2"}])
            aba_df = pd.DataFrame([{"keyword": "kids sunglasses", "搜索量排名": 5000, "排名变化方向": "上升", "变化名次": 100, "前三ASIN点击份额": 0.3, "前三ASIN转化份额": 0.2, "相关性分类": "核心词", "当前广告覆盖": "已覆盖", "动作建议": "保量优化", "归因解释": "高热度且不极端", "判断依据": "ABA排名 5000"}])
            evidence_chain = pd.DataFrame([{"结论": "增长质量下降", "依据1": "订单增长", "依据2": "毛利下降", "依据3": "广告ACOS上升", "归因解释": "增量质量弱", "动作": "控浪费"}])
            execution_plan = pd.DataFrame([{"优先级": "P0", "时间": "当天", "动作": "下调商品页面", "验证指标": "ACOS下降", "原因": "广告位承接弱"}])

            build_html_report(output, "", metrics, metrics, metrics, period_compare, activity, placement, target, search_actions, search_clusters, search_aba_cross, rank_df, aba_df, evidence_chain, execution_plan)
            html = output.read_text(encoding="utf-8")

        self.assertIn("执行摘要", html)
        self.assertIn("关键洞察", html)
        self.assertIn("明细追溯请看 Excel", html)
        self.assertNotIn("<h3>投放层</h3>", html)
        self.assertLessEqual(html.count("<table"), 2)


if __name__ == "__main__":
    unittest.main()
