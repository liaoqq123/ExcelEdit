"""设置和配置迁移共用的纯输入校验，不导入 Tk。"""

from decimal import Decimal, InvalidOperation
import re
import unicodedata

from excel_common import ExcelReadError


def parse_positive_int(value: object, field_label: str, maximum: int | None = None) -> int:
    """把用户输入解析为正整数，并给出带字段名的错误提示。"""

    normalized_value = normalize_number_text(value)
    try:
        decimal_value = Decimal(normalized_value)
    except (InvalidOperation, ValueError):
        decorated_match = re.fullmatch(r"(?:第)?([+-]?\d+)(?:行|列)?", normalized_value)
        if decorated_match is None:
            current_value = "空" if not normalized_value else normalized_value
            raise ExcelReadError(f"{field_label}必须是正整数，当前读取到：{current_value}")
        decimal_value = Decimal(decorated_match.group(1))

    if not decimal_value.is_finite() or decimal_value != decimal_value.to_integral_value():
        raise ExcelReadError(f"{field_label}必须是正整数，当前读取到：{normalized_value}")
    if decimal_value < 1:
        raise ExcelReadError(f"{field_label}必须大于等于 1")
    if maximum is not None and decimal_value > maximum:
        raise ExcelReadError(f"{field_label}不能大于 {maximum}")
    return int(decimal_value)


def normalize_number_text(value: object) -> str:
    """统一全角数字、逗号和不可见字符，降低用户输入格式要求。"""

    text = unicodedata.normalize("NFKC", str(value)).strip()
    text = text.replace(",", "").replace("，", "")
    return "".join(char for char in text if not char.isspace() and char not in {"\u200b", "\ufeff"})
