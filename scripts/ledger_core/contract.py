"""共享输出契约及保守文本规范化。 @author denovochen"""
from __future__ import annotations

import hashlib
import json
import re
import unicodedata
from datetime import datetime

VERSION = "1.7.1"
FIELDS = [
    "序号", "项目名称", "项目编号", "标段名称", "标段编号", "公司名称", "中标与否", "投标排名",
    "文件类别", "依据文件路径", "来源页码", "提取方式", "证据文本", "置信度", "复核状态", "解析结果生成日期时间",
]
OUTPUTS = {"final.csv", "review_queue.csv", "ledger.json"}
ROLES = {
    "project_name", "project_code", "project_serial", "project_year", "project_owner", "agent",
    "lot_name", "lot_code", "bidder_name", "bidder_serial", "bidder_count", "bidder_price", "bidder_legal_person",
    "award_name", "award_status", "award_price", "award_legal_person", "rank", "notes",
}
BUSINESS_ROLES = {"project_name", "project_code", "lot_name", "lot_code", "bidder_name", "award_name"}


class LedgerError(ValueError):
    """输入、映射或产物校验失败；不得伪装为成功。"""


class MappingRevisionRequired(LedgerError):
    """结构映射需要局部修订；调用方可返回有界证据而不发布结果。"""

    def __init__(self, message: str, evidence: dict | None = None):
        super().__init__(message)
        self.evidence = evidence or {}


class RecoverableWorkbookError(LedgerError):
    """单个工作簿无法解析或真实范围超限，同批其他工作簿仍可继续交付。"""


def text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean(value: object) -> str:
    value = unicodedata.normalize("NFKC", text(value))
    value = re.sub(r"[\u200b-\u200d\ufeff]", "", value)
    return re.sub(r"\s+", " ", value).strip()


def name_key(value: object) -> str:
    return re.sub(r"\s+", "", clean(value)).casefold()


def stable_id(kind: str, *parts: object) -> str:
    encoded = json.dumps(parts, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return kind + "_" + hashlib.sha256(encoded).hexdigest()[:24]


def column_number(label: str) -> int:
    if not isinstance(label, str) or not re.fullmatch(r"[A-Z]{1,3}", label):
        raise LedgerError(f"列标无效: {label!r}")
    number = 0
    for char in label:
        number = number * 26 + ord(char) - 64
    return number


def column_label(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def coordinate(row: int, col: int) -> str:
    return f"{column_label(col)}{row}"
