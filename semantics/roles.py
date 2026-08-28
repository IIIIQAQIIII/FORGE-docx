"""Core semantic roles (Mission 03-C1)."""

from enum import Enum


class Role(str, Enum):
    TITLE = "title"
    SUBTITLE = "subtitle"
    HEADING_1 = "heading_1"
    HEADING_2 = "heading_2"
    HEADING_3 = "heading_3"
    BODY = "body"
    EMPTY = "empty"
    UNKNOWN = "unknown"
    ORGANIZATION = "organization"
    AUTHOR = "author"
    DATE = "date"
    SIGNATURE = "signature"
    CAPTION = "caption"
