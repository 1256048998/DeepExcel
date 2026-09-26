"""提问卡：clarify_intent 的参数整理与发给宿主的消息。"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_tools import MAX_CLARIFY_OPTIONS, MAX_CLARIFY_QUESTIONS, normalize_questions
from ipc import clarify_message


def test_single_question_legacy_form():
    qs = normalize_questions({"question": " 求和还是计数？ ", "options": ["SUM", "COUNTA", ""]})
    assert qs == [{"question": "求和还是计数？", "header": "", "multi_select": False,
                   "options": [{"label": "SUM", "description": ""}, {"label": "COUNTA", "description": ""}]}]


def test_multiple_questions_with_descriptions_and_multi_select():
    qs = normalize_questions({"questions": [
        {"question": "按什么汇总？", "header": "汇总口径", "options": [
            {"label": "按月", "description": "每月一行"}, {"label": "按部门"}]},
        {"question": "要哪些指标？", "header": "指标", "multi_select": True, "options": ["销售额", "毛利"]},
        {"question": ""},             # 空问题丢掉
        "不是对象",                    # 非法条目丢掉
    ]})
    assert [q["header"] for q in qs] == ["汇总口径", "指标"]
    assert qs[0]["options"][0] == {"label": "按月", "description": "每月一行"}
    assert qs[1]["multi_select"] is True


def test_limits():
    many = {"questions": [{"question": f"问题{i}", "options": [str(j) for j in range(10)]} for i in range(9)]}
    qs = normalize_questions(many)
    assert len(qs) == MAX_CLARIFY_QUESTIONS
    assert all(len(q["options"]) == MAX_CLARIFY_OPTIONS for q in qs)
    assert normalize_questions({}) == []


def test_clarify_message_keeps_legacy_fields_for_hosts():
    one = clarify_message(normalize_questions({"question": "Q？", "options": ["A", "B"]}))
    assert one["question"] == "Q？" and one["options"] == ["A", "B"]
    two = clarify_message(normalize_questions({"questions": [
        {"question": "第一？", "options": ["A"]}, {"question": "第二？", "options": ["B"]}]}))
    # 多题时旧字段只用于记历史：问题逐行列出，选项留空（各题的选项在 questions 里）
    assert two["question"] == "1. 第一？\n2. 第二？"
    assert two["options"] == []
    assert len(two["questions"]) == 2
