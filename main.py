#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Convert an insurance rate-table PDF to the fixed TXT/CSV-like format.

Output format:
    性别,男性,女性,...
    投保年龄|交费期间,一次性交纳,一次性交纳,...
    0,9391,9395,...

The parser detects payment periods and gender columns from the PDF text. When
the source contains surrender/paid-up amount sections such as "交清增额" or
"交清保额", that section and everything after it is ignored.
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import logging
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

try:
    from pypdf import PdfReader
    from pypdf.errors import PdfReadError, PdfStreamError
except ImportError as exc:  # pragma: no cover - environment guard
    raise SystemExit(
        "缺少依赖 pypdf，请先安装：pip install pypdf"
    ) from exc

# 导入 get_content_vl 模块中的 get_write_text 方法
try:
    from set_config import get_write_text
except ImportError:
    # 如果导入失败，使用本地定义的方法
    pass

STOP_MARKERS = ("交清增额", "交清保额","健康加费","交清增额费率表","交清增额保险费率表","交清保额费率表","交清保额的净保险费表")
PAYMENT_RE = re.compile(
    r"(?:趸\s*交|趸\s*缴|一次(?:性)?(?:交纳|缴纳|交清|交付|支付)|(?:[0-9０-９]+|[一二三四五六七八九十两]+)\s*年\s*(?:交|缴|期)?)"
)
GENDER_RE = re.compile(r"(男性|女性|男|女)")
AGE_RE = re.compile(r"^\s*([0-9０-９]{1,3})\s+(.+)$")
NUMBER_RE = re.compile(r"[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?")
GENDER_ORDER = ["男性", "女性"]
logging.getLogger("pypdf").setLevel(logging.ERROR)


@dataclass(frozen=True)
class Column:
    gender: str
    payment_period: str


@dataclass(frozen=True)
class RateTable:
    columns: list[Column]
    rows: list[list[str]]
    insurance_period: str | None = None
    sections: list["RateTable"] | None = None
    section_label: str = "保险期间"


def fullwidth_to_halfwidth(text: str) -> str:
    chars: list[str] = []
    for char in text:
        code = ord(char)
        if code == 0x3000:
            chars.append(" ")
        elif 0xFF10 <= code <= 0xFF19:
            chars.append(chr(code - 0xFF10 + ord("0")))
        else:
            chars.append(char)
    return "".join(chars)


def normalize_space(text: str) -> str:
    return re.sub(r"[ \t]+", " ", fullwidth_to_halfwidth(text)).strip()


def chinese_number_to_int(text: str) -> int | None:
    text = text.replace("两", "二")
    digits = {
        "零": 0,
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
    }
    if text.isdigit():
        return int(text)
    if text == "十":
        return 10
    if "十" in text:
        left, _, right = text.partition("十")
        tens = digits.get(left, 1) if left else 1
        ones = digits.get(right, 0) if right else 0
        return tens * 10 + ones
    if text in digits:
        return digits[text]
    return None


def normalize_payment_period(period: str) -> str:
    period = normalize_space(period).replace(" ", "")
    if "趸" in period or "一次" in period:
        return "一次性交纳"
    match = re.search(r"([0-9]+|[一二三四五六七八九十两]+)年", period)
    if match:
        number = chinese_number_to_int(match.group(1))
        if number is not None:
            return f"{number}年交"
    return period


def normalize_title_payment_period(text: str) -> str | None:
    if "一次交清" in text:
        return "一次交清"
    payments = extract_payments(text)
    if payments:
        return payments[0]
    return None


def normalize_gender(gender: str) -> str:
    if gender == "男":
        return "男性"
    if gender == "女":
        return "女性"
    return gender


def single_gender_line(line: str) -> str | None:
    normalized = normalize_space(line).replace(" ", "")
    if normalized in {"男", "男性"}:
        return "男性"
    if normalized in {"女", "女性"}:
        return "女性"
    return None


def title_gender(line: str) -> str | None:
    genders = unique_in_order(extract_genders(line))
    if len(genders) == 1 and ("保险期间" in line or "保险费" in line or "单位" in line):
        return genders[0]
    return None


def unique_in_order(items: Iterable[str]) -> list[str]:
    seen = set()
    result = []
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        result.append(item)
    return result


def payment_sort_key(period: str) -> tuple[int, int, str]:
    if period == "一次性交纳":
        return (0, 0, period)
    match = re.search(r"(\d+)年", period)
    if match:
        return (1, int(match.group(1)), period)
    return (2, 0, period)


def canonical_payment_order(periods: Iterable[str]) -> list[str]:
    return sorted(unique_in_order(periods), key=payment_sort_key)


def repair_pdf_bytes(data: bytes) -> bytes:
    if b"%%EOF" in data and b"startxref" in data:
        return data

    object_offsets = {
        int(match.group(1)): match.start()
        for match in re.finditer(rb"(?m)^(\d+)\s+0\s+obj", data)
    }
    if not object_offsets or 1 not in object_offsets:
        return data

    max_object = max(object_offsets)
    startxref = len(data) + 1
    xref_lines = [
        b"xref",
        f"0 {max_object + 1}".encode("ascii"),
        b"0000000000 65535 f ",
    ]
    for object_number in range(1, max_object + 1):
        offset = object_offsets.get(object_number, 0)
        generation = b"00000" if offset else b"65535"
        flag = b"n" if offset else b"f"
        xref_lines.append(f"{offset:010d}".encode("ascii") + b" " + generation + b" " + flag + b" ")

    trailer = (
        f"trailer\n<</Size {max_object + 1}/Root 1 0 R>>\n"
        f"startxref\n{startxref}\n%%EOF\n"
    ).encode("ascii")
    return data + b"\n" + b"\n".join(xref_lines) + b"\n" + trailer


def open_pdf_reader(pdf_path: Path) -> PdfReader:
    try:
        return PdfReader(str(pdf_path), strict=False)
    except (PdfReadError, PdfStreamError):
        repaired = repair_pdf_bytes(pdf_path.read_bytes())
        return PdfReader(io.BytesIO(repaired), strict=False)


def read_pdf_text(pdf_path: Path) -> str:
    reader = open_pdf_reader(pdf_path)
    page_texts = []
    for page in reader.pages:
        page_text = page.extract_text() or ""
        stop_positions = [
            position
            for marker in STOP_MARKERS
            if (position := page_text.find(marker)) != -1
        ]
        if stop_positions:
            page_texts.append(page_text[: min(stop_positions)])
            break
        page_texts.append(page_text)
    return "\n".join(page_texts)


def normalize_period_value(value: str) -> str:
    value = normalize_space(value)
    if value == "终身":
        return value
    age_match = re.match(r"至?\s*([0-9]+|[一二三四五六七八九十两]+)\s*周岁", value)
    if age_match:
        number = chinese_number_to_int(age_match.group(1))
        return f"至{number}周岁" if number is not None else value.replace(" ", "")
    number_match = re.match(r"([0-9]+|[一二三四五六七八九十两]+)年", value)
    if not number_match:
        return value
    number = chinese_number_to_int(number_match.group(1))
    return f"{number}年" if number is not None else value


def extract_insurance_period(text: str) -> str | None:
    match = re.search(
        r"保险期间\s*[:：]?\s*((?:至\s*)?[0-9０-９一二三四五六七八九十两]+\s*周岁|终身|[0-9０-９一二三四五六七八九十两]+年)",
        text,
    )
    if not match:
        return None
    return normalize_period_value(match.group(1))


def has_insurance_period(text: str) -> bool:
    return extract_insurance_period(text) is not None


def normalize_insured_count(value: str) -> str:
    value = normalize_space(value)
    match = re.search(r"([0-9０-９一二三四五六七八九十两]+)\s*人", value)
    if not match:
        return value
    raw_number = fullwidth_to_halfwidth(match.group(1))
    number = chinese_number_to_int(raw_number)
    if number is None:
        return f"被保人为{raw_number}人"
    chinese_digits = {
        1: "一",
        2: "两",
        3: "三",
        4: "四",
        5: "五",
        6: "六",
        7: "七",
        8: "八",
        9: "九",
        10: "十",
    }
    return f"被保人为{chinese_digits.get(number, str(number))}人"


def extract_insured_count(text: str) -> str | None:
    match = re.search(r"被保险?人为\s*([0-9０-９一二三四五六七八九十两]+人)", text)
    if not match:
        return None
    return normalize_insured_count(match.group(1))


def extract_insured_counts(text: str) -> list[str]:
    return [
        normalize_insured_count(match.group(1))
        for match in re.finditer(r"被保险?人为\s*([0-9０-９一二三四五六七八九十两]+人)", text)
    ]


def extract_payments(line: str) -> list[str]:
    return [normalize_payment_period(item) for item in PAYMENT_RE.findall(line)]


def extract_genders(line: str) -> list[str]:
    return [normalize_gender(item) for item in GENDER_RE.findall(line)]


def looks_like_data_row(line: str) -> bool:
    match = AGE_RE.match(line)
    if not match:
        return False
    age = int(fullwidth_to_halfwidth(match.group(1)))
    return 0 <= age <= 120 and len(extract_numbers(match.group(2))) >= 2


def find_first_data_row(lines: Sequence[str]) -> int:
    for index, line in enumerate(lines):
        if looks_like_data_row(line):
            return index
    raise ValueError("未找到年龄数据起始行。")


def infer_expected_values(header_lines: list[str], first_data_numbers: list[str]) -> int:
    payments = canonical_payment_order(
        payment
        for line in header_lines
        for payment in extract_payments(line)
    )
    genders = unique_in_order(
        gender
        for line in header_lines
        for gender in extract_genders(line)
    )

    if payments and genders:
        expected = len(payments) * len(genders)
        if expected <= len(first_data_numbers):
            return expected
    if genders and len(first_data_numbers) % len(genders) == 0:
        return len(first_data_numbers)
    return len(first_data_numbers)


def repeat_each(items: list[str], count: int) -> list[str]:
    repeated: list[str] = []
    for item in items:
        repeated.extend([item] * count)
    return repeated


def repeat_all(items: list[str], count: int) -> list[str]:
    return items * count


def infer_source_columns(header_lines: list[str], expected_values: int) -> list[Column]:
    payment_tokens: list[str] = []
    gender_tokens: list[str] = []
    gender_group_tokens: list[str] = []

    for line in header_lines:
        payments_seen_before = bool(payment_tokens)
        line_payments = extract_payments(line)
        if line_payments:
            payment_tokens.extend(line_payments)

        line_genders = extract_genders(line)
        # print('line_genders=======',line_genders)
        has_gender_sections(line_genders)
        if line_genders:
            gender_tokens.extend(line_genders)
            compact_line = line.replace(" ", "")
            if (
                (
                    "性别" in line
                    or "投保年龄" in line
                    or set(compact_line) <= {"男", "女", "性", "别"}
                    or (len(line_genders) >= 2 and not line_payments and not payments_seen_before)
                )
                and len(line_genders) < expected_values
            ):
                gender_group_tokens.extend(line_genders)

    payment_tokens = [item for item in payment_tokens if item]
    unique_payments = canonical_payment_order(payment_tokens)
    unique_genders = unique_in_order(gender_tokens)
    group_genders = unique_in_order(gender_group_tokens)

    if not unique_payments:
        raise ValueError("未识别到缴费期间。")

    # Some layouts put a single gender in the title, e.g. "（男性 保险期间终身 ...）".
    if len(unique_genders) == 1 and len(unique_payments) == expected_values:
        return [Column(unique_genders[0], period) for period in unique_payments]

    if not unique_genders:
        if len(payment_tokens) == expected_values and expected_values % 2 == 0:
            half = expected_values // 2
            if payment_tokens[:half] == payment_tokens[half:]:
                periods = canonical_payment_order(payment_tokens[:half])
                return [
                    Column(gender, payment)
                    for gender in GENDER_ORDER
                    for payment in periods
                ]
        if expected_values == len(unique_payments) * 2:
            unique_genders = GENDER_ORDER.copy()
        else:
            raise ValueError("未识别到性别信息，且无法从列数推断。")

    if len(payment_tokens) == expected_values and len(group_genders) in (1, 2):
        per_gender = expected_values // len(group_genders)
        if per_gender and payment_tokens[:per_gender] == payment_tokens[per_gender : per_gender * 2]:
            periods = canonical_payment_order(payment_tokens[:per_gender])
            return [
                Column(gender, payment)
                for gender in group_genders
                for payment in periods
            ]

    if len(unique_payments) * len(unique_genders) == expected_values:
        if group_genders and len(group_genders) == len(unique_genders):
            return [
                Column(gender, payment)
                for gender in group_genders
                for payment in unique_payments
            ]
        return [
            Column(gender, payment)
            for payment in unique_payments
            for gender in unique_genders
        ]

    if len(payment_tokens) == expected_values and len(unique_genders) == expected_values:
        return [
            Column(gender, payment)
            for gender, payment in zip(unique_genders, payment_tokens)
        ]

    if len(payment_tokens) == expected_values and len(unique_genders) == 1:
        return [Column(unique_genders[0], payment) for payment in payment_tokens]

    raise ValueError(
        "无法对应表头和数据列："
        f"缴费期间候选 {len(unique_payments)} 个，性别候选 {len(unique_genders)} 个，"
        f"数据列 {expected_values} 个。"
    )


def has_gender_sections(text: str) -> bool:
    print('text=======',text)
    return  True

def parse_single_gender_sections(lines: list[str]) -> RateTable | None:
    parsed_from_reordered = parse_title_trailing_gender_sections(lines)
    if parsed_from_reordered is not None:
        return parsed_from_reordered

    current_gender: str | None = None
    current_payments: list[str] = []
    by_gender: dict[str, dict[str, list[str]]] = {}
    payment_order: list[str] = []

    for line in lines:
        gender = single_gender_line(line) or title_gender(line)
        if gender:
            current_gender = gender
            current_payments = []
            by_gender.setdefault(gender, {})
            continue

        payments = canonical_payment_order(extract_payments(line))
        if current_gender and payments:
            current_payments = payments
            payment_order = unique_in_order([*payment_order, *payments])
            continue

        if not current_gender or not current_payments:
            continue

        if not AGE_RE.match(line):
            continue

        entries = split_age_entries(line, len(current_payments))
        if not entries:
            continue

        gender_rows = by_gender.setdefault(current_gender, {})
        for entry in entries:
            age = entry[0]
            values = align_missing_values(entry[1:], [Column(current_gender, p) for p in current_payments])
            gender_rows[age] = values

    genders = [gender for gender in GENDER_ORDER if gender in by_gender and by_gender[gender]]
    if not genders or not payment_order:
        return None

    ages = sorted(
        {age for gender in genders for age in by_gender[gender]},
        key=lambda value: int(value),
    )
    if not ages:
        return None

    columns = [
        Column(gender, payment)
        for payment in payment_order
        for gender in genders
    ]
    rows: list[list[str]] = []
    for age in ages:
        row = [age]
        for payment in payment_order:
            payment_index = payment_order.index(payment)
            for gender in genders:
                values = by_gender[gender].get(age, [])
                row.append(values[payment_index] if payment_index < len(values) else "")
        rows.append(row)

    return RateTable(columns=columns, rows=rows)


def parse_title_trailing_gender_sections(lines: list[str]) -> RateTable | None:
    sections: list[tuple[str, list[str]]] = []
    pending: list[str] = []

    for line in lines:
        pending.append(line)
        gender = title_gender(line)
        if gender:
            sections.append((gender, pending[:-1]))
            pending = []

    if len(sections) < 2:
        return None

    reordered: list[str] = []
    for gender, section_lines in sections:
        reordered.append(gender)
        reordered.extend(section_lines)

    return parse_single_gender_sections_without_reorder(reordered)


def parse_single_gender_sections_without_reorder(lines: list[str]) -> RateTable | None:
    current_gender: str | None = None
    current_payments: list[str] = []
    by_gender: dict[str, dict[str, list[str]]] = {}
    payment_order: list[str] = []

    for line in lines:
        gender = single_gender_line(line) or title_gender(line)
        if gender:
            current_gender = gender
            current_payments = []
            by_gender.setdefault(gender, {})
            continue

        payments = canonical_payment_order(extract_payments(line))
        if current_gender and payments:
            current_payments = payments
            payment_order = unique_in_order([*payment_order, *payments])
            continue

        if not current_gender or not current_payments or not AGE_RE.match(line):
            continue

        entries = split_age_entries(line, len(current_payments))
        if not entries:
            continue

        gender_rows = by_gender.setdefault(current_gender, {})
        for entry in entries:
            age = entry[0]
            values = align_missing_values(entry[1:], [Column(current_gender, p) for p in current_payments])
            gender_rows[age] = values

    return build_table_from_gender_rows(by_gender, payment_order)


def build_table_from_gender_rows(
    by_gender: dict[str, dict[str, list[str]]],
    payment_order: list[str],
) -> RateTable | None:
    genders = [gender for gender in GENDER_ORDER if gender in by_gender and by_gender[gender]]
    if not genders or not payment_order:
        return None

    ages = sorted(
        {age for gender in genders for age in by_gender[gender]},
        key=lambda value: int(value),
    )
    if not ages:
        return None

    columns = [
        Column(gender, payment)
        for payment in payment_order
        for gender in genders
    ]
    rows: list[list[str]] = []
    for age in ages:
        row = [age]
        for payment in payment_order:
            payment_index = payment_order.index(payment)
            for gender in genders:
                values = by_gender[gender].get(age, [])
                row.append(values[payment_index] if payment_index < len(values) else "")
        rows.append(row)

    return RateTable(columns=columns, rows=rows)


def is_header_candidate(line: str) -> bool:
    if looks_like_data_row(line):
        return False
    if line.startswith("注") or "保险费＝" in line or "保险费=" in line:
        return False
    payments = extract_payments(line)
    genders = extract_genders(line)
    return bool(
        "投保年龄" in line
        or "交费期间" in line
        or "缴费期间" in line
        or "性别" in line
        or single_gender_line(line)
        or len(genders) >= 2
        or (payments and ("保险费" not in line or "一次" in line or "趸" in line))
    )


def collect_header_candidates(lines: list[str]) -> list[str]:
    return [line for line in lines if is_header_candidate(line)]


def extract_numbers(text: str) -> list[str]:
    return [item.replace(",", "") for item in NUMBER_RE.findall(text)]


def is_integer_age_token(value: str) -> bool:
    try:
        numeric = float(value)
    except ValueError:
        return False
    return numeric.is_integer() and 0 <= int(numeric) <= 120


def split_age_entries(line: str, expected_values: int) -> list[list[str]]:
    numbers = extract_numbers(line)
    entries: list[list[str]] = []
    index = 0

    while index < len(numbers):
        age_token = numbers[index]
        if not is_integer_age_token(age_token):
            index += 1
            continue

        next_age_index: int | None = None
        candidate = index + expected_values + 1
        if candidate < len(numbers) and is_integer_age_token(numbers[candidate]):
            next_age_index = candidate

        end_index = next_age_index if next_age_index is not None else len(numbers)
        values = numbers[index + 1 : end_index]
        if values:
            entries.append([str(int(float(age_token))), *values])

        if next_age_index is None:
            break
        index = next_age_index

    return entries


def is_numeric_continuation(line: str) -> bool:
    if not extract_numbers(line):
        return False
    remaining = NUMBER_RE.sub("", line)
    return not remaining.strip(" ,，.\t")


def iter_data_rows(lines: list[str], start_index: int, expected_values: int) -> Iterable[list[str]]:
    pending_age: str | None = None
    pending_values: list[str] = []

    def flush_row(allow_incomplete: bool = False) -> list[str] | None:
        if pending_age is None:
            return None
        if len(pending_values) >= expected_values:
            return [pending_age, *pending_values[:expected_values]]
        if allow_incomplete and pending_values:
            return [pending_age, *pending_values]
        return None

    for line in lines[start_index + 1 :]:
        if not line:
            continue
        if "交费期间" in line or "投保年龄" in line or "性别" in line:
            continue

        match = AGE_RE.match(line)
        if match:
            completed = flush_row(allow_incomplete=True)
            if completed:
                yield completed

            entries = split_age_entries(line, expected_values)
            if len(entries) > 1:
                for entry in entries:
                    yield entry
                pending_age = None
                pending_values = []
                continue

            pending_age = entries[0][0] if entries else str(int(fullwidth_to_halfwidth(match.group(1))))
            pending_values = entries[0][1:] if entries else extract_numbers(match.group(2))
            completed = flush_row()
            if completed:
                yield completed
                pending_age = None
                pending_values = []
            continue

        if pending_age is not None:
            if not is_numeric_continuation(line):
                completed = flush_row(allow_incomplete=True)
                if completed:
                    yield completed
                pending_age = None
                pending_values = []
                continue

            pending_values.extend(extract_numbers(line))
            completed = flush_row()
            if completed:
                yield completed
                pending_age = None
                pending_values = []

    completed = flush_row(allow_incomplete=True)
    if completed:
        yield completed


def contiguous_column_groups(columns: list[Column]) -> list[list[int]]:
    groups: list[list[int]] = []
    for index, column in enumerate(columns):
        if not groups or columns[groups[-1][-1]].gender != column.gender:
            groups.append([index])
        else:
            groups[-1].append(index)
    return groups


def align_missing_values(values: list[str], columns: list[Column]) -> list[str]:
    expected_values = len(columns)
    if len(values) >= expected_values:
        return values[:expected_values]

    aligned = [""] * expected_values
    groups = contiguous_column_groups(columns)
    group_periods = [
        [columns[index].payment_period for index in group]
        for group in groups
    ]

    if (
        len(groups) > 1
        and len(values) % len(groups) == 0
        and all(periods == group_periods[0] for periods in group_periods)
    ):
        values_per_group = len(values) // len(groups)
        if values_per_group <= len(groups[0]):
            value_index = 0
            for group in groups:
                for column_index in group[:values_per_group]:
                    aligned[column_index] = values[value_index]
                    value_index += 1
            return aligned

    aligned[: len(values)] = values
    return aligned


def to_float(value: str) -> float | None:
    if value == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None


def continuity_align_values(values: list[str], reference_values: list[str]) -> list[str] | None:
    expected_values = len(reference_values)
    if len(values) >= expected_values:
        return values[:expected_values]

    numeric_values = [to_float(value) for value in values]
    numeric_reference = [to_float(value) for value in reference_values]
    if any(value is None for value in numeric_values):
        return None

    candidates = [index for index, value in enumerate(numeric_reference) if value is not None]
    if len(candidates) < len(values):
        return None

    states: dict[int, tuple[float, list[int]]] = {-1: (0.0, [])}
    for value in numeric_values:
        next_states: dict[int, tuple[float, list[int]]] = {}
        for previous_index, (cost, path) in states.items():
            for index in candidates:
                if index <= previous_index:
                    continue
                reference = numeric_reference[index]
                if reference is None:
                    continue
                step_cost = abs(value - reference) / max(abs(reference), 1.0)  # type: ignore[operator]
                new_cost = cost + step_cost
                if index not in next_states or new_cost < next_states[index][0]:
                    next_states[index] = (new_cost, [*path, index])
        states = next_states
        if not states:
            return None

    _, path = min(states.values(), key=lambda item: item[0])
    aligned = [""] * expected_values
    for value, index in zip(values, path):
        aligned[index] = value
    return aligned


def align_rows(raw_rows: list[list[str]], columns: list[Column]) -> list[list[str]]:
    rows: list[list[str]] = []
    reference_values: list[str] | None = None

    for row in raw_rows:
        age, values = row[0], row[1:]
        if len(values) < len(columns) and reference_values is not None:
            aligned_values = continuity_align_values(values, reference_values)
            if aligned_values is None:
                aligned_values = align_missing_values(values, columns)
        else:
            aligned_values = align_missing_values(values, columns)
        rows.append([age, *aligned_values])
        reference_values = aligned_values

    return rows


def first_unique_age_table(raw_rows: list[list[str]]) -> list[list[str]]:
    seen_ages: set[str] = set()
    table_rows: list[list[str]] = []
    for row in raw_rows:
        age = row[0]
        if age in seen_ages:
            break
        seen_ages.add(age)
        table_rows.append(row)
    return sorted(table_rows, key=lambda row: int(row[0]))


@dataclass(frozen=True)
class PositionedText:
    x: float
    y: float
    text: str


def page_positioned_text(page) -> list[PositionedText]:
    items: list[PositionedText] = []

    def visitor(text, cm, tm, font, size) -> None:
        for line in text.splitlines():
            value = normalize_space(line)
            if value:
                items.append(PositionedText(float(tm[4]), float(tm[5]), value))

    page.extract_text(visitor_text=visitor)
    return items


def group_positioned_rows(items: list[PositionedText], tolerance: float = 1.2) -> list[list[PositionedText]]:
    groups: list[list[PositionedText]] = []
    for item in sorted(items, key=lambda value: -value.y):
        for group in groups:
            if abs(group[0].y - item.y) <= tolerance:
                group.append(item)
                break
        else:
            groups.append([item])
    return [sorted(group, key=lambda value: value.x) for group in groups]


def first_number(text: str) -> str | None:
    numbers = extract_numbers(text)
    return numbers[0] if numbers else None


def is_age_value(value: str) -> bool:
    try:
        age = int(float(value))
    except ValueError:
        return False
    return 0 <= age <= 120 and float(value) == age


def row_entries_from_positioned_group(group: list[PositionedText]) -> list[tuple[int, list[str]]]:
    numeric_items: list[tuple[float, str]] = []
    for item in group:
        number = first_number(item.text)
        if number is not None:
            numeric_items.append((item.x, number))

    entries: list[tuple[int, list[str]]] = []
    index = 0
    while index < len(numeric_items):
        _, age_value = numeric_items[index]
        if not is_age_value(age_value):
            index += 1
            continue
        age = int(float(age_value))
        values: list[str] = []
        value_index = index + 1
        while value_index < len(numeric_items) and len(values) < 2:
            _, candidate = numeric_items[value_index]
            if not is_age_value(candidate):
                values.append(candidate)
            value_index += 1
        if len(values) == 2:
            entries.append((age, values))
            index = value_index
        else:
            index += 1
    return entries


def parse_insurance_period_pdf(pdf_path: Path) -> RateTable:
    reader = open_pdf_reader(pdf_path)
    sections: list[RateTable] = []

    for page in reader.pages:
        page_text = page.extract_text() or ""
        insurance_period = extract_insurance_period(page_text)
        if insurance_period is None:
            continue

        payment_period = normalize_title_payment_period(page_text)
        if payment_period is None:
            raise ValueError("检测到保险期间，但未识别到交费期间。")

        age_rows: dict[int, list[str]] = {}
        for group in group_positioned_rows(page_positioned_text(page)):
            for age, values in row_entries_from_positioned_group(group):
                if len(values) == 2:
                    age_rows[age] = values

        if not age_rows:
            raise ValueError(f"保险期间 {insurance_period} 未识别到年龄数据。")

        rows = [[str(age), *age_rows[age]] for age in sorted(age_rows)]
        sections.append(
            RateTable(
                columns=[
                    Column("男性", payment_period),
                    Column("女性", payment_period),
                ],
                rows=rows,
                insurance_period=insurance_period,
            )
        )

    if not sections:
        raise ValueError("未识别到含保险期间的数据表。")

    return RateTable(columns=[], rows=[], sections=sections)


def parse_multi_insurance_period_pdf(pdf_path: Path) -> RateTable | None:
    reader = open_pdf_reader(pdf_path)
    sections_by_period: list[tuple[str, list[str]]] = []

    for page in reader.pages:
        page_text = page.extract_text() or ""
        insurance_period = extract_insurance_period(page_text)
        if insurance_period is not None:
            sections_by_period.append((insurance_period, [page_text]))
        elif sections_by_period:
            sections_by_period[-1][1].append(page_text)

    unique_periods = unique_in_order(period for period, _ in sections_by_period)
    if len(unique_periods) < 2:
        return None

    sections: list[RateTable] = []
    for insurance_period, page_texts in sections_by_period:
        section = parse_rate_table("\n".join(page_texts), allow_section_parsers=False)
        sections.append(
            RateTable(
                columns=section.columns,
                rows=section.rows,
                insurance_period=insurance_period,
            )
        )

    return RateTable(columns=[], rows=[], sections=sections, section_label="保险期间")


def parse_insured_count_pdf(pdf_path: Path) -> RateTable | None:
    horizontal = parse_horizontal_insured_count_pdf(pdf_path)
    if horizontal is not None:
        return horizontal

    reader = open_pdf_reader(pdf_path)
    sections_by_value: list[tuple[str, list[str]]] = []
    current_value: str | None = None

    for page in reader.pages:
        page_text = page.extract_text() or ""
        insured_count = extract_insured_count(page_text)
        if insured_count:
            current_value = insured_count
        if current_value:
            if sections_by_value and sections_by_value[-1][0] == current_value:
                sections_by_value[-1][1].append(page_text)
            else:
                sections_by_value.append((current_value, [page_text]))

    if not sections_by_value:
        return None

    sections: list[RateTable] = []
    for value, page_texts in sections_by_value:
        section = parse_rate_table("\n".join(page_texts), allow_section_parsers=False)
        sections.append(
            RateTable(
                columns=section.columns,
                rows=section.rows,
                insurance_period=value,
            )
        )

    return RateTable(columns=[], rows=[], sections=sections, section_label="投保计划")


def split_values_for_groups(values: list[str], group_count: int, group_width: int) -> list[list[str]]:
    groups = [[""] * group_width for _ in range(group_count)]
    if group_count <= 0 or group_width <= 0:
        return groups

    if len(values) >= group_count * group_width:
        for group_index in range(group_count):
            start = group_index * group_width
            groups[group_index] = values[start : start + group_width]
        return groups

    base_count, extra_count = divmod(len(values), group_count)
    value_index = 0
    for group_index in range(group_count):
        values_per_group = min(group_width, base_count + (1 if group_index < extra_count else 0))
        groups[group_index][:values_per_group] = values[value_index : value_index + values_per_group]
        value_index += values_per_group
    return groups


def parse_horizontal_insured_count_pdf(pdf_path: Path) -> RateTable | None:
    reader = open_pdf_reader(pdf_path)
    current_gender: str | None = None
    current_counts: list[str] = []
    current_payments: list[str] = []
    payment_order: list[str] = []
    by_count_gender: dict[str, dict[str, dict[str, list[str]]]] = {}

    for page in reader.pages:
        page_text = page.extract_text() or ""
        lines = [normalize_space(line) for line in page_text.splitlines()]
        lines = [line for line in lines if line]

        page_genders = [
            gender
            for line in lines[:8]
            if (gender := single_gender_line(line) or title_gender(line))
        ]
        if page_genders:
            current_gender = page_genders[0]

        for line in lines:
            counts = extract_insured_counts(line)
            if len(counts) >= 2:
                current_counts = counts
                for count in current_counts:
                    by_count_gender.setdefault(count, {})
                continue

            raw_payments = extract_payments(line)
            if current_counts and len(raw_payments) >= len(current_counts) and len(raw_payments) % len(current_counts) == 0:
                group_width = len(raw_payments) // len(current_counts)
                first_group = raw_payments[:group_width]
                if all(raw_payments[i * group_width : (i + 1) * group_width] == first_group for i in range(len(current_counts))):
                    current_payments = first_group
                else:
                    current_payments = canonical_payment_order(first_group)
                payment_order = unique_in_order([*payment_order, *current_payments])
                continue

            if not current_gender or not current_counts or not current_payments or not AGE_RE.match(line):
                continue

            entries = split_age_entries(line, len(current_counts) * len(current_payments))
            for entry in entries:
                age = entry[0]
                grouped_values = split_values_for_groups(entry[1:], len(current_counts), len(current_payments))
                for count, values in zip(current_counts, grouped_values):
                    by_count_gender.setdefault(count, {}).setdefault(current_gender, {})[age] = values

    valid_counts = [count for count in current_counts if count in by_count_gender and by_count_gender[count]]
    if len(valid_counts) < 2 or not payment_order:
        return None

    sections: list[RateTable] = []
    for count in valid_counts:
        section = build_table_from_gender_rows(by_count_gender[count], payment_order)
        if section is not None:
            sections.append(RateTable(columns=section.columns, rows=section.rows, insurance_period=count))

    if len(sections) < 2:
        return None
    return RateTable(columns=[], rows=[], sections=sections, section_label="投保计划")


def parse_rate_table(text: str, allow_section_parsers: bool = True) -> RateTable:
    lines = [normalize_space(line) for line in text.splitlines()]
    lines = [line for line in lines if line]

    if allow_section_parsers:
        single_gender_table = parse_single_gender_sections(lines)
        if single_gender_table is not None:
            return single_gender_table

    first_data_index = find_first_data_row(lines)
    first_data_numbers = extract_numbers(AGE_RE.match(lines[first_data_index]).group(2))  # type: ignore[union-attr]
    header_lines = collect_header_candidates(lines)
    expected_values = infer_expected_values(header_lines, first_data_numbers)
    columns = infer_source_columns(header_lines, expected_values)

    if len(columns) != expected_values:
        raise ValueError(f"表头列数 {len(columns)} 与数据列数 {expected_values} 不一致。")

    raw_rows = list(iter_data_rows(lines, first_data_index - 1, expected_values))
    rows = align_rows(first_unique_age_table(raw_rows), columns)
    if not rows:
        raise ValueError("未识别到年龄数据行。")

    return RateTable(columns=columns, rows=rows, insurance_period=extract_insurance_period(text))


def common_payment_periods(count: int) -> list[str]:
    common_orders = {
        1: ["一次性交纳"],
        2: ["一次性交纳", "3年交"],
        3: ["一次性交纳", "3年交", "5年交"],
        4: ["一次性交纳", "3年交", "5年交", "10年交"],
        5: ["一次性交纳", "3年交", "5年交", "10年交", "20年交"],
        6: ["一次性交纳", "3年交", "5年交", "10年交", "15年交", "20年交"],
    }
    return common_orders.get(count, [f"第{index + 1}列" for index in range(count)])


def grouped_centers(indexes: Sequence[int], max_gap: int = 2) -> list[int]:
    if not indexes:
        return []
    centers: list[int] = []
    start = previous = indexes[0]
    for index in indexes[1:]:
        if index <= previous + max_gap:
            previous = index
            continue
        centers.append((start + previous) // 2)
        start = previous = index
    centers.append((start + previous) // 2)
    return centers


def detect_image_table_grid(image) -> tuple[list[int], list[int]] | None:
    try:
        import numpy as np
    except ImportError:
        return None

    gray = image.convert("L")
    pixels = np.array(gray)
    dark = pixels < 80
    height, width = pixels.shape

    horizontal_region = dark[:, int(width * 0.12) : int(width * 0.88)]
    horizontal_scores = horizontal_region.mean(axis=1)
    horizontal_indexes = [int(index) for index, score in enumerate(horizontal_scores) if score > 0.5]
    horizontal_lines = grouped_centers(horizontal_indexes)
    if len(horizontal_lines) < 5:
        return None

    y1, y2 = horizontal_lines[0], horizontal_lines[-1]
    vertical_region = dark[y1:y2, :]
    vertical_scores = vertical_region.mean(axis=0)
    vertical_indexes = [int(index) for index, score in enumerate(vertical_scores) if score > 0.9]
    vertical_lines = grouped_centers(vertical_indexes)
    if len(vertical_lines) < 4:
        vertical_indexes = [int(index) for index, score in enumerate(vertical_scores) if score > 0.75]
        vertical_lines = grouped_centers(vertical_indexes)

    vertical_lines = [
        line
        for line in vertical_lines
        if vertical_scores[line] > 0.75
    ]
    if len(vertical_lines) < 4:
        return None
    return horizontal_lines, vertical_lines


def prepare_ocr_cell(image, box: tuple[int, int, int, int]):
    try:
        from PIL import ImageOps
    except ImportError:
        return None

    x1, y1, x2, y2 = box
    if x2 <= x1 or y2 <= y1:
        return None
    cell = image.convert("L").crop((x1 + 2, y1 + 1, x2 - 2, y2 - 1))
    cell = ImageOps.autocontrast(cell)
    return cell.resize((max(1, cell.width * 8), max(1, cell.height * 8)))


def prepare_ocr_cell_variants(image, box: tuple[int, int, int, int]) -> list:
    cell = prepare_ocr_cell(image, box)
    if cell is None:
        return []
    return [
        cell,
        cell.point(lambda pixel: 0 if pixel < 130 else 255),
        cell.point(lambda pixel: 0 if pixel < 150 else 255),
    ]


def windows_ocr_images(image_paths: list[Path]) -> list[str]:
    if not image_paths:
        return []

    powershell = "powershell.exe"
    script = r'''
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.Streams.IRandomAccessStreamWithContentType, Windows.Storage.Streams, ContentType = WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function AwaitOp($op, $type) {
    $asTask = $asTaskGeneric.MakeGenericMethod($type)
    $task = $asTask.Invoke($null, @($op))
    $task.Wait()
    $task.Result
}
$lang = New-Object Windows.Globalization.Language 'zh-Hans-CN'
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if ($null -eq $engine) { throw 'Windows OCR engine is unavailable.' }
$paths = Get-Content -LiteralPath $args[0] -Raw | ConvertFrom-Json
$items = @()
foreach ($path in $paths) {
    $file = AwaitOp ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile])
    $stream = AwaitOp ($file.OpenReadAsync()) ([Windows.Storage.Streams.IRandomAccessStreamWithContentType])
    $decoder = AwaitOp ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = AwaitOp ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $result = AwaitOp ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    $items += $result.Text
}
$items | ConvertTo-Json -Compress
'''

    with tempfile.TemporaryDirectory(prefix="rate_ocr_") as temp_name:
        temp_dir = Path(temp_name)
        script_path = temp_dir / "ocr.ps1"
        paths_path = temp_dir / "paths.json"
        script_path.write_text(script, encoding="utf-8")
        paths_path.write_text(
            json.dumps([str(path) for path in image_paths], ensure_ascii=False),
            encoding="utf-8",
        )
        result = subprocess.run(
            [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script_path), str(paths_path)],
            text=True,
            encoding="utf-8",
            capture_output=True,
            check=True,
        )
    data = json.loads(result.stdout or "[]")
    if isinstance(data, str):
        return [data]
    return [str(item) for item in data]


def clean_ocr_number(text: str) -> str:
    text = fullwidth_to_halfwidth(text)
    replacements = {
        "．": ".",
        "。": ".",
        "·": ".",
        "，": ".",
        ",": ".",
        "O": "0",
        "o": "0",
        "L": "1",
        "l": "1",
        "I": "1",
        "|": "1",
        "H": "1",
        "了": "1",
        "刀": "1",
        "《": "1",
        "引": "5",
        "巧": "15",
        "在": "6",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^0-9.\-]", "", text)
    if not text:
        return ""

    if "." in text:
        integer, decimal = text.split(".", 1)
        integer = re.sub(r"\D", "", integer)
        decimal = re.sub(r"\D", "", decimal)
        if not integer:
            return ""
        if not decimal and len(integer) > 6:
            return f"{integer[:-6]}.{integer[-6:]}"
        return f"{integer}.{decimal[:6]}" if decimal else integer

    digits = re.sub(r"\D", "", text)
    if not digits:
        return ""
    if len(digits) > 6:
        return f"{digits[:-6]}.{digits[-6:]}"
    return digits


def ocr_number_score(value: str) -> tuple[int, float]:
    if not value:
        return (0, 0.0)
    try:
        number = float(value)
    except ValueError:
        return (1, 0.0)
    score = 1
    if re.fullmatch(r"\d{3,5}\.\d{6}", value):
        score += 100
    elif re.fullmatch(r"\d+\.\d{6}", value):
        score += 60
    if number >= 100:
        score += 20
    else:
        score -= 20
    return (score, number)


def select_ocr_number(texts: list[str]) -> str:
    candidates = [clean_ocr_number(text) for text in texts]
    candidates = [candidate for candidate in candidates if candidate]
    if not candidates:
        return ""
    return max(candidates, key=ocr_number_score)


def parse_image_pdf_rate_table(pdf_path: Path) -> RateTable | None:
    try:
        from PIL import Image
    except ImportError:
        return None

    reader = open_pdf_reader(pdf_path)
    if not reader.pages:
        return None
    page = reader.pages[0]
    images = list(page.images)
    if not images:
        return None

    image = Image.open(io.BytesIO(images[0].data)).convert("RGB")
    grid = detect_image_table_grid(image)
    if grid is None:
        return None
    horizontal_lines, vertical_lines = grid
    if len(horizontal_lines) < 4 or len(vertical_lines) < 4:
        return None

    data_column_count = len(vertical_lines) - 2
    if data_column_count <= 0 or data_column_count % 2 != 0:
        return None
    per_gender_count = data_column_count // 2
    payment_periods = common_payment_periods(per_gender_count)
    columns = [
        Column(gender, payment)
        for gender in GENDER_ORDER
        for payment in payment_periods
    ]

    data_start_line = 2
    row_count = len(horizontal_lines) - data_start_line - 1
    if row_count <= 0:
        return None

    cell_meta: list[tuple[int, int]] = []
    cell_paths: list[Path] = []
    with tempfile.TemporaryDirectory(prefix="rate_cells_") as temp_name:
        temp_dir = Path(temp_name)
        for row_index in range(row_count):
            y1 = horizontal_lines[data_start_line + row_index]
            y2 = horizontal_lines[data_start_line + row_index + 1]
            for column_index in range(len(vertical_lines) - 1):
                x1 = vertical_lines[column_index]
                x2 = vertical_lines[column_index + 1]
                variants = prepare_ocr_cell_variants(image, (x1, y1, x2, y2))
                if not variants:
                    continue
                for variant_index, cell in enumerate(variants):
                    path = temp_dir / f"r{row_index:03d}_c{column_index:02d}_v{variant_index}.png"
                    cell.save(path)
                    cell_meta.append((row_index, column_index))
                    cell_paths.append(path)

        ocr_texts = windows_ocr_images(cell_paths)

    cell_texts: dict[tuple[int, int], list[str]] = {}
    for meta, text in zip(cell_meta, ocr_texts):
        cell_texts.setdefault(meta, []).append(text)
    cell_values: dict[tuple[int, int], str] = {
        meta: select_ocr_number(texts)
        for meta, texts in cell_texts.items()
    }

    rows: list[list[str]] = []
    previous_age: int | None = None
    for row_index in range(row_count):
        raw_age = cell_values.get((row_index, 0), "")
        if raw_age and raw_age.isdigit():
            age = int(raw_age)
        elif previous_age is None:
            age = row_index
        else:
            age = previous_age + 1
        previous_age = age
        values = [
            cell_values.get((row_index, column_index), "")
            for column_index in range(1, len(vertical_lines) - 1)
        ]
        if any(values):
            rows.append([str(age), *values])

    return RateTable(columns=columns, rows=rows) if rows else None


def output_columns(source_columns: list[Column]) -> list[Column]:
    return source_columns


def parse_ocr_script_rate_table(pdf_path: Path) -> RateTable | None:
    ocr_script = Path(__file__).with_name("ocr_pdf_rate_table.py")
    if not ocr_script.exists():
        print(f"OCR脚本不存在：{ocr_script}", file=sys.stderr)
        return None

    print(f"文本解析失败，尝试调用OCR脚本：{ocr_script}", file=sys.stderr)
    result = subprocess.run(
        [sys.executable, str(ocr_script), str(pdf_path), "--json"],
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        check=False,
    )
    if result.returncode != 0 or not result.stdout.strip():
        message = (result.stdout or "").strip()
        if message:
            print(f"OCR脚本调用失败：{message}", file=sys.stderr)
        else:
            print("OCR脚本调用失败：没有返回可解析结果。", file=sys.stderr)
        return None

    payload = json.loads(result.stdout)
    columns = [
        Column(item["gender"], item["payment_period"])
        for item in payload.get("columns", [])
    ]
    rows = payload.get("rows", [])
    if not columns or not rows:
        return None
    return RateTable(columns=columns, rows=rows)


def build_output_rows(table: RateTable) -> list[list[str]]:
    if table.sections:
        rows: list[list[str]] = []
        for section in table.sections:
            rows.append([table.section_label, *([section.insurance_period or ""] * len(section.columns))])
            rows.extend(build_output_rows(RateTable(columns=section.columns, rows=section.rows)))
        return rows

    columns = output_columns(table.columns)
    source_index = {
        (column.gender, column.payment_period): index
        for index, column in enumerate(table.columns)
    }

    output_rows = [
        ["性别", *[column.gender for column in columns]],
        ["投保年龄|交费期间", *[column.payment_period for column in columns]],
    ]
    if table.insurance_period:
        output_rows.insert(0, ["保险期间", *([table.insurance_period] * len(columns))])

    for row in table.rows:
        age, values = row[0], row[1:]
        output_rows.append(
            [
                age,
                *[
                    values[source_index[(column.gender, column.payment_period)]]
                    for column in columns
                ],
            ]
        )
    return output_rows


def write_txt(rows: list[list[str]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file, lineterminator="\n")
        writer.writerows(rows)


def convert(pdf_path: Path, output_path: Path) -> RateTable:
    text = read_pdf_text(pdf_path)
    # 给获取set_config.py 函数   用的
    get_write_text(text, pdf_filename=str(pdf_path))
    try:
        insurance_period_table = parse_multi_insurance_period_pdf(pdf_path) if has_insurance_period(text) else None
        insured_count_table = parse_insured_count_pdf(pdf_path) if extract_insured_count(text) else None
        table = insurance_period_table or insured_count_table or parse_rate_table(text)
    except Exception:
        image_table = parse_ocr_script_rate_table(pdf_path)
        if image_table is not None:
            table = image_table
        else:
            if not has_insurance_period(text):
                raise
            table = parse_insurance_period_pdf(pdf_path)
    write_txt(build_output_rows(table), output_path)
    return table


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="将保险费率表 PDF 转换为示例格式的 txt 文件。"
    )
    parser.add_argument("pdf", type=Path, nargs="?", help="输入 PDF 文件路径（省略则自动扫描 stare/ 目录）")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="输出 txt 文件路径，默认与 PDF 同名",
    )
    return parser.parse_args(argv)


def _process_one(pdf_path: Path) -> int:
    """处理单个 PDF：生成 txt + xlsx，返回 0 表示成功。"""
    output_path = pdf_path.with_suffix(".txt")

    if not pdf_path.exists():
        print(f"输入 PDF 不存在：{pdf_path}", file=sys.stderr)
        return 2

    try:
        table = convert(pdf_path, output_path)
    except Exception as exc:
        print(f"转换失败：{exc}", file=sys.stderr)
        return 1

    num_rows = (
        sum(len(section.rows) for section in table.sections)
        if table.sections else len(table.rows)
    )
    num_cols = (
        len(table.sections[0].columns)
        if table.sections else len(table.columns)
    )
    print(f"转换完成：{output_path} ({num_rows} 行，{num_cols} 个数据列)")

    # ── 自动生成 Excel ────────────────────────────────────────────────────────
    try:
        from set_excel import create_config_sheet
        excel_path = output_path.with_suffix(".xlsx")
        create_config_sheet(output_path=excel_path, txt_path=output_path)
    except Exception as exc:
        print(f"⚠️ Excel 生成失败：{exc}", file=sys.stderr)

    return 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)

    if args.pdf is not None:
        # ── 单文件模式 ────────────────────────────────────────────────────────
        output_path = args.output or args.pdf.with_suffix(".txt")
        pdf_path = args.pdf
        if not pdf_path.exists():
            print(f"输入 PDF 不存在：{pdf_path}", file=sys.stderr)
            return 2
        try:
            table = convert(pdf_path, output_path)
        except Exception as exc:
            print(f"转换失败：{exc}", file=sys.stderr)
            return 1
        num_rows = (
            sum(len(section.rows) for section in table.sections)
            if table.sections else len(table.rows)
        )
        num_cols = (
            len(table.sections[0].columns)
            if table.sections else len(table.columns)
        )
        print(f"转换完成：{output_path} ({num_rows} 行，{num_cols} 个数据列)")
        try:
            from set_excel import create_config_sheet
            excel_path = output_path.with_suffix(".xlsx")
            create_config_sheet(output_path=excel_path, txt_path=output_path)
        except Exception as exc:
            print(f"⚠️ Excel 生成失败：{exc}", file=sys.stderr)
        return 0

    # ── 批量模式：自动扫描 stare/ 目录 ────────────────────────────────────────
    stare_dir = Path(__file__).resolve().parent / "stare"
    if not stare_dir.is_dir():
        print(f"未找到 stare 目录：{stare_dir}", file=sys.stderr)
        return 2

    pdfs = sorted(stare_dir.glob("*.pdf"))
    if not pdfs:
        print(f"stare 目录下没有 PDF 文件：{stare_dir}", file=sys.stderr)
        return 2

    print(f"📂 扫描到 {len(pdfs)} 个 PDF，开始批量处理...")
    failed = 0
    for pdf_path in pdfs:
        print(f"\n{'─'*50}\n处理：{pdf_path.name}")
        result = _process_one(pdf_path)
        if result != 0:
            failed += 1

    print(f"\n{'─'*50}\n批量处理完成：成功 {len(pdfs) - failed}/{len(pdfs)} 个")
    return 1 if failed else 0



if __name__ == "__main__":
    raise SystemExit(main())
