"""Mission 02-C Format Intelligence MCP Integration 测试。"""

import asyncio
import json

import server


def call_sync(tool: str, args: dict) -> str:
    async def _call():
        output = await server.mcp.call_tool(tool, args)
        if isinstance(output, (list, tuple)):
            parts = output[0] if isinstance(output, tuple) else output
            if isinstance(parts, (list, tuple)) and parts:
                return parts[0].text if hasattr(parts[0], "text") else str(parts[0])
        return str(output)

    return asyncio.run(_call())


def resolve(description, **kwargs):
    args = {"description": description, **kwargs}
    return json.loads(call_sync("resolve_document_format", args))


def test_1_activity_plan():
    r = resolve("六一儿童节活动方案，包含活动目标、活动时间、活动地点、活动流程和人员分工")
    assert r["classification"]["intent"] == "activity_plan"
    assert r["resolution"]["profile_id"] == "activity_plan_standard"
    assert r["suggested_document_type"] == "活动方案"


def test_2_activity_plan_explicit_official():
    r = resolve("六一儿童节活动方案，包含活动目标、活动时间、活动地点、活动流程和人员分工", explicit_format_hint="正式公文")
    assert r["classification"]["intent"] == "activity_plan"
    assert r["resolution"]["profile_id"] == "official_standard"
    assert r["resolution"]["decision_basis"] == "explicit_user_choice"
    assert r["suggested_document_type"] == "传统公文"


def test_3_work_summary():
    r = resolve("本学期保教工作总结，梳理主要工作、特色亮点、存在问题和下一步工作")
    assert r["classification"]["intent"] == "official_summary"
    assert r["resolution"]["profile_id"] == "official_standard"


def test_4_activity_summary():
    r = resolve("读书月活动总结，梳理活动开展情况、活动成效、活动亮点和活动反思")
    assert r["classification"]["intent"] == "activity_summary"
    assert r["resolution"]["profile_id"] == "activity_summary_standard"
    assert r["suggested_document_type"] == "活动总结"


def test_5_activity_archive():
    r = resolve("活动影像资料，包括活动照片和影像")
    assert r["classification"]["intent"] == "activity_archive"
    assert r["resolution"]["profile_id"] == "activity_archive_standard"
    assert r["suggested_document_type"] == "活动影像"


def test_6_training_archive():
    r = resolve("培训活动影像，包括培训照片")
    assert r["classification"]["intent"] == "training_archive"
    assert r["resolution"]["profile_id"] == "training_archive_standard"
    assert r["suggested_document_type"] == "培训活动影像"


def test_7_training_record():
    r = resolve("师德师风培训活动记录，包括培训时间、地点、主讲人、参培人员、培训内容和培训心得")
    assert r["classification"]["intent"] == "training_record"
    assert r["resolution"]["profile_id"] == "training_record_standard"


def test_8_ambiguous_needs_guidance():
    r = resolve("帮我整理一下这个文件")
    assert r["resolution"]["status"] == "needs_guidance"
    assert r["resolution"]["decision_basis"] == "guided_required"


def test_9_unknown_explicit_profile():
    r = resolve("随便什么内容", explicit_profile_id="no_such_profile")
    assert r["resolution"]["status"] == "error"
    assert r["resolution"]["error"] == "PROFILE_NOT_FOUND"


def test_10_reference_over_saved():
    r = resolve(
        "六一儿童节活动方案",
        reference_profile_id="official_standard",
        saved_profile_id="activity_plan_standard",
    )
    assert r["resolution"]["profile_id"] == "official_standard"
    assert r["resolution"]["decision_basis"] == "reference_profile"


def test_11_explicit_over_reference():
    r = resolve(
        "六一儿童节活动方案",
        explicit_profile_id="official_standard",
        reference_profile_id="activity_plan_standard",
    )
    assert r["resolution"]["profile_id"] == "official_standard"
    assert r["resolution"]["decision_basis"] == "explicit_user_choice"


def test_12_legacy_recommend_unchanged():
    data = json.loads(call_sync("recommend_document_type", {"description": "写个评价计划"}))
    assert "recommended_document_type" in data
    assert data["recommended_document_type"] == "传统公文"
    assert "classification" not in data
    assert "resolution" not in data


def test_13_generate_regression():
    out = call_sync(
        "generate_by_type",
        {
            "document_type": "论文",
            "content": {
                "title": "回归测试论文",
                "author": "测试",
                "abstract": "回归测试摘要",
                "keywords": "回归；测试",
                "chapters": [{"heading": "第一章 引言", "sections": [{"heading": "1.1 背景", "subsections": [{"heading": "1.1.1 问题", "paragraphs": ["正文。"]}]}]}],
            },
            "output_name": "_intelligence_regression.docx",
        },
    )
    assert out.startswith("Created:")
    data = json.loads(call_sync("validate_docx", {"docx_path": "_intelligence_regression.docx"}))
    assert data["status"] in ("pass", "warning")
    import os

    os.remove(server.OUTPUTS_DIR / "_intelligence_regression.docx")
