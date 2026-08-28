"""Mission 02-B Format Resolver 测试。"""

from intelligence.classifier import classify_content
from intelligence.resolver import resolve_format


def test_1_activity_plan_confirmed_recommends_activity_profile():
    cls = classify_content("六一儿童节活动方案，包含活动目标、活动时间、活动地点、活动流程和人员分工")
    result = resolve_format(classification=cls)
    assert result["profile_id"] == "activity_plan_standard"
    assert result["decision_basis"] == "content_recommendation"


def test_2_explicit_profile_overrides_content():
    cls = classify_content("六一儿童节活动方案，包含活动目标、活动时间、活动地点、活动流程和人员分工")
    result = resolve_format(classification=cls, explicit_profile_id="official_standard")
    assert result["profile_id"] == "official_standard"
    assert result["decision_basis"] == "explicit_user_choice"


def test_3_explicit_profile_overrides_reference():
    result = resolve_format(
        classification={"intent": "activity_plan", "confidence": 0.9, "status": "confirmed"},
        explicit_profile_id="official_standard",
        reference_profile_id="academic_standard",
    )
    assert result["profile_id"] == "official_standard"
    assert result["decision_basis"] == "explicit_user_choice"


def test_4_reference_overrides_saved():
    result = resolve_format(
        classification={"intent": "activity_plan", "confidence": 0.9, "status": "confirmed"},
        reference_profile_id="academic_standard",
        saved_profile_id="official_standard",
    )
    assert result["profile_id"] == "academic_standard"
    assert result["decision_basis"] == "reference_profile"


def test_5_saved_overrides_recommendation():
    cls = classify_content("六一儿童节活动方案，包含活动目标、活动时间、活动地点、活动流程和人员分工")
    result = resolve_format(classification=cls, saved_profile_id="official_standard")
    assert result["profile_id"] == "official_standard"
    assert result["decision_basis"] == "saved_profile"


def test_6_generic_ambiguous_needs_guidance():
    cls = classify_content("帮我整理一下这个文件")
    result = resolve_format(classification=cls)
    assert result["status"] == "needs_guidance"
    assert result["decision_basis"] == "guided_required"


def test_7_close_competition_needs_guidance_with_candidates():
    cls = {
        "intent": "official_plan",
        "confidence": 0.5,
        "status": "ambiguous",
        "alternatives": [
            {"intent": "activity_plan", "confidence": 0.48},
            {"intent": "official_summary", "confidence": 0.3},
        ],
    }
    result = resolve_format(classification=cls)
    assert result["status"] == "needs_guidance"
    assert result["decision_basis"] == "guided_required"
    candidates = [c["profile_id"] for c in result["candidates"]]
    assert "activity_plan_standard" in candidates
    assert "official_standard" in candidates


def test_8_illegal_explicit_profile_not_found():
    result = resolve_format(explicit_profile_id="no_such_profile")
    assert result["status"] == "error"
    assert result["error"] == "PROFILE_NOT_FOUND"


def test_9_explicit_format_hint():
    result = resolve_format(classification={"intent": "official_plan", "status": "confirmed", "confidence": 0.9}, explicit_format_hint="传统公文")
    assert result["profile_id"] == "official_standard"
    assert result["decision_basis"] == "explicit_user_choice"


def test_10_hint_overrides_classified_activity():
    cls = classify_content("六一儿童节活动方案，包含活动目标、活动时间、活动地点、活动流程和人员分工")
    result = resolve_format(classification=cls, explicit_format_hint="正式公文")
    assert result["profile_id"] == "official_standard"
    assert result["decision_basis"] == "explicit_user_choice"
