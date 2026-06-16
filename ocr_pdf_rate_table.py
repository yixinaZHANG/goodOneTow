#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
OCR fallback for image-only insurance rate-table PDFs.

This script is intentionally independent from pdf_rate_table_to_txt.py. It
extracts the embedded page image, detects table lines, OCRs each cell with
Windows OCR, and returns a structured rate table.
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Sequence

try:
    from pypdf import PdfReader
except ImportError as exc:  # pragma: no cover - environment guard
    raise SystemExit("缺少依赖 pypdf，请先安装：pip install pypdf") from exc


GENDER_ORDER = ["男性", "女性"]


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
    except ImportError as exc:
        raise RuntimeError("缺少依赖 numpy，请先安装：pip install numpy") from exc

    gray = image.convert("L")
    pixels = np.array(gray)
    _, width = pixels.shape

    for dark_threshold in (80, 110, 140, 170, 200):
        dark = pixels < dark_threshold
        horizontal_region = dark[:, int(width * 0.08) : int(width * 0.92)]
        horizontal_scores = horizontal_region.mean(axis=1)

        for horizontal_threshold in (0.55, 0.40, 0.28, 0.18, 0.10):
            horizontal_indexes = [
                int(index)
                for index, score in enumerate(horizontal_scores)
                if score > horizontal_threshold
            ]
            horizontal_lines = grouped_centers(horizontal_indexes, max_gap=3)
            if len(horizontal_lines) < 5:
                continue
            if len(horizontal_lines) > 130:
                continue

            y1, y2 = horizontal_lines[0], horizontal_lines[-1]
            if y2 <= y1:
                continue

            vertical_region = dark[y1:y2, :]
            vertical_scores = vertical_region.mean(axis=0)
            for vertical_threshold in (0.85, 0.65, 0.45, 0.30, 0.18):
                vertical_indexes = [
                    int(index)
                    for index, score in enumerate(vertical_scores)
                    if score > vertical_threshold
                ]
                vertical_lines = grouped_centers(vertical_indexes, max_gap=3)
                vertical_lines = [
                    line
                    for line in vertical_lines
                    if vertical_scores[line] > vertical_threshold
                ]
                if len(vertical_lines) < 4:
                    continue
                if len(vertical_lines) > 40:
                    continue

                return horizontal_lines, vertical_lines

    return None


def extract_or_render_first_page_image(pdf_path: Path):
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("缺少依赖 Pillow，请先安装：pip install pillow") from exc

    reader = PdfReader(str(pdf_path), strict=False)
    if reader.pages:
        images = list(reader.pages[0].images)
        if images:
            return Image.open(io.BytesIO(images[0].data)).convert("RGB")

    try:
        import fitz  # PyMuPDF
    except ImportError:
        fitz = None

    if fitz is not None:
        document = fitz.open(str(pdf_path))
        page = document[0]
        pixmap = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        return Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")

    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        with tempfile.TemporaryDirectory(prefix="rate_render_") as temp_name:
            prefix = Path(temp_name) / "page"
            subprocess.run(
                [pdftoppm, "-f", "1", "-singlefile", "-r", "300", "-png", str(pdf_path), str(prefix)],
                capture_output=True,
                text=True,
                check=True,
            )
            rendered = prefix.with_suffix(".png")
            if rendered.exists():
                return Image.open(rendered).convert("RGB")

    raise RuntimeError("PDF没有可直接提取的页面图片，且未找到 PyMuPDF 或 pdftoppm，无法渲染页面。")


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
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script_path), str(paths_path)],
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            check=True,
        )
    data = json.loads(result.stdout or "[]")
    if isinstance(data, str):
        return [data]
    return [str(item) for item in data]


def tesseract_ocr_images(image_paths: list[Path]) -> list[str]:
    tesseract = shutil.which("tesseract")
    if not tesseract:
        raise RuntimeError("当前系统未找到 tesseract。macOS 可执行：brew install tesseract tesseract-lang")

    outputs: list[str] = []
    total = len(image_paths)
    print(f"OCR开始：共 {total} 个单元格，macOS/Linux 使用 tesseract 单版本识别。", file=sys.stderr)
    for index, path in enumerate(image_paths, 1):
        if index == 1 or index % 50 == 0 or index == total:
            print(f"OCR进度：{index}/{total}", file=sys.stderr)
        result = subprocess.run(
            [
                tesseract,
                str(path),
                "stdout",
                "-l",
                "eng",
                "--psm",
                "7",
                "-c",
                "tessedit_char_whitelist=0123456789.,",
            ],
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            check=False,
            timeout=20,
        )
        text = result.stdout.strip() if result.returncode == 0 else ""
        outputs.append(text)
    return outputs


def ocr_images(image_paths: list[Path]) -> list[str]:
    if platform.system().lower().startswith("win"):
        return windows_ocr_images(image_paths)
    return tesseract_ocr_images(image_paths)


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


def parse_image_pdf_rate_table(pdf_path: Path, debug_image_path: Path | None = None) -> dict | None:
    image = extract_or_render_first_page_image(pdf_path)
    if debug_image_path is not None:
        debug_image_path.parent.mkdir(parents=True, exist_ok=True)
        image.save(debug_image_path)

    grid = detect_image_table_grid(image)
    if grid is None:
        raise RuntimeError("已取得PDF页面图片，但未检测到清晰表格线。")
    horizontal_lines, vertical_lines = grid
    if len(horizontal_lines) < 4 or len(vertical_lines) < 4:
        return None

    data_column_count = len(vertical_lines) - 2
    if data_column_count <= 0 or data_column_count % 2 != 0:
        return None
    per_gender_count = data_column_count // 2
    payment_periods = common_payment_periods(per_gender_count)
    columns = [
        {"gender": gender, "payment_period": payment}
        for gender in GENDER_ORDER
        for payment in payment_periods
    ]

    data_start_line = 2
    row_count = len(horizontal_lines) - data_start_line - 1
    if row_count <= 0:
        return None

    cell_meta: list[tuple[int, int]] = []
    cell_paths: list[Path] = []
    use_multiple_cell_variants = platform.system().lower().startswith("win")
    with tempfile.TemporaryDirectory(prefix="rate_cells_") as temp_name:
        temp_dir = Path(temp_name)
        for row_index in range(row_count):
            y1 = horizontal_lines[data_start_line + row_index]
            y2 = horizontal_lines[data_start_line + row_index + 1]
            for column_index in range(len(vertical_lines) - 1):
                x1 = vertical_lines[column_index]
                x2 = vertical_lines[column_index + 1]
                if use_multiple_cell_variants:
                    variants = prepare_ocr_cell_variants(image, (x1, y1, x2, y2))
                else:
                    cell = prepare_ocr_cell(image, (x1, y1, x2, y2))
                    variants = [cell] if cell is not None else []
                if not variants:
                    continue
                for variant_index, cell in enumerate(variants):
                    path = temp_dir / f"r{row_index:03d}_c{column_index:02d}_v{variant_index}.png"
                    cell.save(path)
                    cell_meta.append((row_index, column_index))
                    cell_paths.append(path)

        ocr_texts = ocr_images(cell_paths)

    cell_texts: dict[tuple[int, int], list[str]] = {}
    for meta, text in zip(cell_meta, ocr_texts):
        cell_texts.setdefault(meta, []).append(text)
    cell_values = {
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

    if not rows:
        return None
    return {"columns": columns, "rows": rows}


def build_output_rows(table: dict) -> list[list[str]]:
    columns = table["columns"]
    output_rows = [
        ["性别", *[column["gender"] for column in columns]],
        ["投保年龄|交费期间", *[column["payment_period"] for column in columns]],
    ]
    output_rows.extend(table["rows"])
    return output_rows


def write_txt(rows: list[list[str]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file, lineterminator="\n")
        writer.writerows(rows)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="OCR image-only insurance rate-table PDFs.")
    parser.add_argument("pdf", type=Path, help="输入 PDF 文件路径")
    parser.add_argument("-o", "--output", type=Path, help="输出 txt 文件路径")
    parser.add_argument("--json", action="store_true", help="输出结构化 JSON 到 stdout")
    parser.add_argument("--debug-image", type=Path, help="保存OCR实际处理的页面图片，便于排查")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        table = parse_image_pdf_rate_table(args.pdf, args.debug_image)
    except Exception as exc:
        print(f"OCR处理失败：{exc}", file=sys.stderr)
        return 1
    if table is None:
        print("OCR 未识别到可转换的表格。", file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(table))
    if args.output:
        write_txt(build_output_rows(table), args.output)
    if not args.json and not args.output:
        write_txt(build_output_rows(table), args.pdf.with_suffix(".txt"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
