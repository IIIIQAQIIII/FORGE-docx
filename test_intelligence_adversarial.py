"""Mission 02-A.5 — Classifier Adversarial Validation.

多类信号冲突、主题词干扰、计划/总结/通知/记录语义交叉下的对抗测试。
"""

from intelligence.classifier import classify_content


def _c(desc):
    return classify_content(desc)


def test_1_training_work_plan_is_official_plan():
    r = _c("师德师风培训工作实施方案，包括工作目标、重点任务、实施步骤、工作措施和保障措施。")
    assert r["intent"] == "official_plan"


def test_2_training_plan_with_time_place_is_official_plan():
    r = _c("教师培训工作方案，培训时间为9月，地点为会议室，主要包括工作目标、重点任务、实施步骤和保障措施。")
    assert r["intent"] == "official_plan"
    assert r["intent"] != "training_notice"


def test_3_training_notice():
    r = _c("关于开展师德师风专题培训的通知。培训时间9月10日，地点会议室，培训人为王老师，请全体教师准时参加并按要求做好学习记录。")
    assert r["intent"] == "training_notice"


def test_4_training_record():
    r = _c("师德师风专题培训活动记录，包括培训时间、培训地点、培训主题、主讲人、参培人员、培训内容和培训心得。")
    assert r["intent"] == "training_record"


def test_5_course_activity_plan_competes():
    r = _c("幼儿园课程活动实施方案")
    assert r["intent"] in ("activity_plan", "official_plan")
    assert r["confidence"] < 0.85
    assert r["status"] in ("recommended", "ambiguous")
    alts = {a["intent"] for a in r["alternatives"]}
    assert {"activity_plan", "official_plan"} & alts


def test_6_activity_plan_confirmed():
    r = _c("秋季亲子运动会活动方案，包括活动主题、活动目标、时间、地点、参加人员、活动准备、活动流程和人员分工。")
    assert r["intent"] == "activity_plan"
    assert r["status"] == "confirmed"


def test_7_work_summary_beats_activity_summary():
    r = _c("本学期活动工作总结，主要梳理本学期重点工作、取得成效、特色亮点、存在问题及下一步工作。")
    assert r["intent"] == "official_summary"
    scores = {a["intent"]: a["confidence"] for a in r["alternatives"]}
    scores[r["intent"]] = r["confidence"]
    assert scores.get("official_summary", 0) > scores.get("activity_summary", 0)


def test_8_activity_summary():
    r = _c("读书月活动总结，梳理活动开展情况、活动成效、活动亮点和活动反思。")
    assert r["intent"] == "activity_summary"


def test_9_work_report_not_activity_summary():
    r = _c("本学期保教工作汇报，重点汇报各类活动开展情况、课程建设情况、存在问题及下一阶段工作安排。")
    assert r["intent"] == "official_report"
    assert r["intent"] != "activity_summary"


def test_10_generic_ambiguous():
    r = _c("帮我把这个材料弄正规一点。")
    assert r["intent"] == "generic"
    assert r["status"] == "ambiguous"
