"""Mission 02-A Content Classifier 测试。"""

from intelligence.classifier import classify_content


def _classify(desc):
    return classify_content(desc)


def test_activity_plan_strong():
    result = _classify("六一儿童节活动方案，包含活动目标、活动时间、活动地点、活动流程和人员分工")
    assert result["intent"] == "activity_plan"
    assert result["status"] == "confirmed"
    assert result["confidence"] >= 0.85


def test_official_plan_work():
    result = _classify("制定本学期教研工作方案，明确工作目标、重点任务、工作措施和保障机制")
    assert result["intent"] == "official_plan"
    assert result["status"] == "confirmed"


def test_official_plan_curriculum():
    result = _classify("课程建设实施方案，包括重点任务、实施步骤和保障措施")
    assert result["intent"] == "official_plan"


def test_official_summary():
    result = _classify("本学期保教工作总结，梳理主要工作、特色亮点、存在问题和下一步工作")
    assert result["intent"] == "official_summary"


def test_training_record():
    result = _classify("师德师风培训活动记录，包括培训时间、地点、主讲人、参培人员、培训内容和培训心得")
    assert result["intent"] == "training_record"
    assert result["status"] == "confirmed"


def test_weekly_report():
    result = _classify("第12周行政周报：本周工作、存在问题、下周计划")
    assert result["intent"] == "weekly_report"
    assert result["status"] == "confirmed"


def test_academic_paper():
    result = _classify("幼儿园科学教育研究论文：摘要、关键词、研究方法、研究结果")
    assert result["intent"] == "academic_paper"
    assert result["status"] == "confirmed"


def test_ambiguous_generic():
    result = _classify("帮我整理一下这个文件")
    assert result["intent"] == "generic"
    assert result["status"] == "ambiguous"


def test_activity_plan_bare_lower_confidence():
    strong = _classify("六一儿童节活动方案，包含活动目标、活动时间、活动地点、活动流程和人员分工")
    bare = _classify("活动方案")
    assert bare["intent"] == "activity_plan"
    assert bare["confidence"] < strong["confidence"]
