"""Centralized configuration for the Open Format System."""

from __future__ import annotations

SCHEMA_VERSION = 1

# Reference evidence consensus: dominant value must reach this ratio to be
# adopted automatically. Below it the property becomes a conflict and the
# draft inherits the base profile for that property.
DOMINANCE_RATIO_THRESHOLD = 0.70

# Guided builder: maximum questions returned per round.
MAX_GUIDED_QUESTIONS_PER_ROUND = 5

# Slots a user profile may override. This mirrors FormatProfile slots.
ALLOWED_RULE_SLOTS = (
    "page",
    "title",
    "subtitle",
    "organization",
    "author",
    "body",
    "heading_1",
    "heading_2",
    "heading_3",
    "caption",
    "table",
    "image",
    "signature",
    "date",
    "page_number",
)

# Reference roles that map 1:1 to style slots.
REFERENCE_ROLE_SLOTS = {
    "title": "title",
    "subtitle": "subtitle",
    "organization": "organization",
    "author": "author",
    "heading_1": "heading_1",
    "heading_2": "heading_2",
    "heading_3": "heading_3",
    "body": "body",
    "caption": "caption",
    "signature": "signature",
    "date": "date",
}

# Guided question order (first round picks the first N).
GUIDED_FIELDS = (
    "page",
    "title",
    "body",
    "heading_1",
    "heading_2",
    "heading_3",
    "table",
    "image",
    "signature",
    "date",
    "page_number",
)

# Valid page-number alignment values.
ALIGNMENTS = ("left", "center", "right", "both", "center_right")
