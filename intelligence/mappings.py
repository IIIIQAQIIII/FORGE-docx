"""Mission 02-C — Centralized intelligence mappings.

集中维护：
- CONTENT_PROFILE_RECOMMENDATIONS      内容意图 → 推荐 FormatProfile
- PROFILE_DOCUMENT_TYPE_RECOMMENDATIONS FormatProfile → 建议 document_type

注意：这些只是 recommendation / adapter，不得写入 ContentIntent 模型。
"""

# 内容 → 推荐格式（仅推荐，不绑定）
CONTENT_PROFILE_RECOMMENDATIONS: dict[str, str] = {
    "official_plan": "official_standard",
    "official_summary": "official_standard",
    "official_report": "official_standard",
    "official_request": "official_standard",
    "activity_plan": "activity_plan_standard",
    "activity_summary": "activity_summary_standard",
    "activity_archive": "activity_archive_standard",
    "training_notice": "training_notice_standard",
    "training_record": "training_record_standard",
    "training_archive": "training_archive_standard",
    "weekly_report": "weekly_standard",
    "academic_paper": "academic_standard",
    "speech": "academic_standard",
    "long_form": "academic_standard",
    "generic": "",
}

# 推荐 profile → 建议 document_type（建议，不是强制模板绑定）
PROFILE_DOCUMENT_TYPE_RECOMMENDATIONS: dict[str, str | None] = {
    "official_standard": "传统公文",
    "activity_plan_standard": "活动方案",
    "activity_summary_standard": "活动总结",
    "activity_archive_standard": "活动影像",
    "academic_standard": "论文",
    "weekly_standard": "行政周报",
    "training_notice_standard": "培训通知",
    "training_record_standard": "培训活动记录",
    "training_archive_standard": "培训活动影像",
    "generic_document": None,
}
