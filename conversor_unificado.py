#!/usr/bin/env python3
"""
Conversor Universal PDF/Texto/Imagem/Office -> Markdown v4.2
- Usa PyMuPDF para texto/links e pdfplumber para tabelas.
- Suporte a arquivos do MS Office via markitdown.
- Escrita em streaming (evita OOM em PDFs grandes).
- Heurísticas universais de limpeza (sem hardcoded strings).
- Regex de listas e romanos corrigida.
- CLI e Batch reincluídos (bug do arquivo não gerado resolvido).
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import logging
import re
import sys
import unicodedata
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Optional, Sequence

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError as exc:
        raise SystemExit("PyMuPDF não encontrado. Instale com: pip install pymupdf") from exc

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from charset_normalizer import from_bytes
except ImportError:
    from_bytes = None

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

LOGGER = logging.getLogger("conversor_unificado")

EXTENSOES_IMAGEM = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp"}
EXTENSOES_TEXTO = {".txt", ".log", ".md", ".csv", ".json", ".xml", ".html", ".htm"}

# Adicionado suporte aos formatos do Microsoft Office
EXTENSOES_OFFICE = {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".odt", ".ods", ".odp"}

RE_HIFEN_FINAL = re.compile(r"(?<=[A-Za-zÀ-ÖØ-öø-ÿ])-\s*$")
RE_INICIO_MINUSCULA = re.compile(r"^[a-zà-öø-ÿ]")
RE_RODAPE = re.compile(r"^\s*(\d{1,3})\s*$")
RE_CAPITULO = re.compile(r"^CAPÍTULO\s+\d+\b", re.IGNORECASE)
RE_ITEM = re.compile(r"^(\d+(?:\.\d+)+)\s+(.+)$")
RE_LISTA_LETRA = re.compile(r"^([a-z])\)\s+(.+)$", re.IGNORECASE)
RE_LISTA_ROMANA = re.compile(r"^(X{0,3}(IX|IV|V?I{0,3}))\.\s+(.+)$", re.IGNORECASE)

# Regex para detecção de text-decoration: line-through (case-insensitive, tolerante a espaços)
RE_LINE_THROUGH = re.compile(r"text-decoration\s*:\s*line-through", re.IGNORECASE)


def remove_striked_paragraphs(html_content: str) -> str:
    """
    Remove conteúdo revogado/tachado de HTML legislativo:
      - tags <s>, <strike>, <del> (tag + conteúdo)
      - <span style="text-decoration: line-through"> (case-insensitive)
      - qualquer elemento com class="revogado" — sobe até o <p> ancestral e remove o bloco inteiro

    Retorna o HTML limpo como string.
    """
    if not BeautifulSoup:
        raise RuntimeError("BeautifulSoup é obrigatório para limpeza de HTML. Instale com: pip install beautifulsoup4")

    soup = BeautifulSoup(html_content, "html.parser")

    # Tags tachadas explícitas
    for tag in soup.find_all(["s", "strike", "del"]):
        tag.decompose()

    # Spans com estilo line-through (case-insensitive, tolerante a espaços)
    for par in soup.find_all("span", style=RE_LINE_THROUGH):
        par.decompose()

    # Conteúdo revogado (class="revogado") — remove o <p> ancestral inteiro
    for tag in soup.find_all(class_="revogado"):
        parent = tag.find_parent("p")
        if parent:
            parent.decompose()
        else:
            tag.decompose()

    return str(soup)


@dataclass
class TextBlock:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str

    @property
    def bbox(self) -> tuple[float, float, float, float]:
        return self.x0, self.y0, self.x1, self.y1


@dataclass
class PdfLink:
    rect: tuple[float, float, float, float]
    uri: str
    label: str


@dataclass
class TableState:
    headers: list[str]
    column_count: int
    signature: str


@dataclass
class ConversionStats:
    source: str
    output: str
    format: str
    pages: int = 0
    image_only_pages: int = 0
    tables: int = 0
    links: int = 0
    headings: int = 0
    list_items: int = 0
    page_markers: int = 0
    unresolved_link_urls: int = 0
    warnings: list[str] | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "source": self.source, "output": self.output, "format": self.format,
            "pages": self.pages, "image_only_pages": self.image_only_pages,
            "tables": self.tables, "links": self.links, "headings": self.headings,
            "list_items": self.list_items, "page_markers": self.page_markers,
            "unresolved_link_urls": self.unresolved_link_urls, "warnings": self.warnings or [],
        }


def nfc(text: str) -> str:
    text = unicodedata.normalize("NFC", text.replace("\ufeff", ""))
    return text.replace("\u00a0", " ").replace("\u200b", "")


def collapse_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", nfc(text)).strip()


def dehyphenate_lines(lines: Sequence[str]) -> list[str]:
    result: list[str] = []
    for raw in lines:
        line = collapse_spaces(raw)
        if not line:
            continue
        if result and RE_HIFEN_FINAL.search(result[-1]) and RE_INICIO_MINUSCULA.match(line):
            result[-1] = result[-1][:-1] + line
        else:
            result.append(line)
    return result


def clean_prose(text: str) -> str:
    lines = dehyphenate_lines(nfc(text).splitlines())
    value = collapse_spaces(" ".join(lines))
    value = re.sub(r"\s+([,.;:!?])", r"\1", value)
    value = re.sub(r"\(\s+", "(", value)
    value = re.sub(r"\s+\)", ")", value)
    return value.strip()


def clean_cell(text: Optional[str]) -> str:
    if not text:
        return ""
    lines = dehyphenate_lines(nfc(text).splitlines())
    value = collapse_spaces(" ".join(lines))
    value = re.sub(r"\s+([,.;:!?])", r"\1", value)
    return value.replace("|", r"\|").strip()


def rect_tuple(rect: Any) -> tuple[float, float, float, float]:
    return float(rect[0]), float(rect[1]), float(rect[2]), float(rect[3])


def intersection_area(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    x0, y0 = max(a[0], b[0]), max(a[1], b[1])
    x1, y1 = min(a[2], b[2]), min(a[3], b[3])
    return max(0.0, x1 - x0) * max(0.0, y1 - y0)


def overlaps(a: tuple, b: tuple, minimum: float = 0.01) -> bool:
    return intersection_area(a, b) > minimum


def union_rect(rects: Iterable[tuple]) -> tuple[float, float, float, float]:
    values = list(rects)
    if not values: return 0.0, 0.0, 0.0, 0.0
    return (min(r[0] for r in values), min(r[1] for r in values),
            max(r[2] for r in values), max(r[3] for r in values))


def markdown_link(label: str, uri: str) -> str:
    safe_label = label.replace("[", r"\[").replace("]", r"\]")
    return f"[{safe_label}](<{uri}>)"


def words_for_link(page: Any, rect: tuple[float, float, float, float]) -> str:
    raw_text = page.get_textbox(rect)
    return collapse_spaces(raw_text) if raw_text else ""


def normalize_uri(uri: str) -> str:
    value = str(uri).strip()
    match = re.match(r"^chrome-extension://[^/]+/https:/+(.*)$", value, flags=re.IGNORECASE)
    if match:
        return "https://" + match.group(1)
    return value


def extract_links(page: Any) -> list[PdfLink]:
    links: list[PdfLink] = []
    for raw in page.get_links() or []:
        uri = raw.get("uri")
        if not uri: continue
        rect = rect_tuple(raw.get("from", (0, 0, 0, 0)))
        expanded_rect = (rect[0] - 1, rect[1] - 1, rect[2] + 1, rect[3] + 1)
        label = words_for_link(page, expanded_rect)
        if label:
            links.append(PdfLink(rect=rect, uri=normalize_uri(uri), label=label))
    links.sort(key=lambda link: (link.rect[1], link.rect[0]))
    return links


def normalize_link_label(label: str) -> str:
    label = collapse_spaces(label).strip()
    label = re.sub(r"^[([]+", "", label)
    label = re.sub(r"[.,;:!?\])]+$", "", label)
    return label.strip()


def apply_links(text: str, bbox: tuple[float, float, float, float], links: Sequence[PdfLink], emitted_urls: set[str]) -> \
tuple[str, int]:
    original = text
    candidates = [link for link in links if overlaps(link.rect, bbox, minimum=0.05)]
    candidates.sort(key=lambda link: (link.rect[1], link.rect[0]))

    matches: list[tuple[int, int, PdfLink, str]] = []
    cursor = 0

    for link in candidates:
        label = normalize_link_label(link.label)
        if not label: continue
        escaped_label = re.escape(label).replace(r"\ ", r"\s+")

        match = re.search(escaped_label, original[cursor:])
        if match:
            start, end = cursor + match.start(), cursor + match.end()
        else:
            match = re.search(escaped_label, original)
            if not match: continue
            start, end = match.start(), match.end()

        if any(start < other_end and end > other_start for other_start, other_end, _, _ in matches):
            continue
        matches.append((start, end, link, original[start:end]))
        emitted_urls.add(link.uri)
        cursor = end

    if not matches: return original, 0

    parts, last = [], 0
    for start, end, link, matched_text in sorted(matches, key=lambda item: item[0]):
        parts.append(original[last:start])
        parts.append(markdown_link(matched_text, link.uri))
        last = end
    parts.append(original[last:])
    return "".join(parts), len(matches)


def extract_footer_number(page: Any) -> Optional[int]:
    height = float(page.rect.height)
    for block in page.get_text("blocks", sort=True):
        x0, y0, x1, y1, text = block[:5]
        if y0 < height - 65: continue
        match = RE_RODAPE.match(clean_prose(text))
        if match: return int(match.group(1))
    return None


def is_table_header_row(row: Sequence[str]) -> bool:
    cells = [clean_cell(cell).casefold() for cell in row if clean_cell(cell)]
    if not cells: return False
    short_cells = sum(1 for c in cells if len(c.split()) <= 3 and not c.endswith(('.', ';')))
    return short_cells >= len(cells) * 0.8


def table_rows_with_boxes(table: Any) -> tuple[list[list[str]], list[list[Optional[tuple]]]]:
    raw_rows = table.extract() or []
    boxes = []
    for row in getattr(table, "rows", []):
        boxes.append([rect_tuple(cell) if cell else None for cell in row.cells])
    return raw_rows, boxes


def canonical_table_rows(table: Any) -> tuple[list[list[str]], list[list[Optional[tuple]]]]:
    raw_rows, raw_boxes = table_rows_with_boxes(table)
    if not raw_rows: return [], []

    header_boxes = raw_boxes[0] if raw_rows else []
    groups = [box for box in header_boxes if box is not None]
    if not groups: groups = [rect_tuple(table.bbox)]

    def group_for_box(box: Optional[tuple]) -> int:
        if box is None: return 0
        center = (box[0] + box[2]) / 2.0
        return min(range(len(groups)), key=lambda idx: abs(((groups[idx][0] + groups[idx][2]) / 2.0) - center))

    canonical, canonical_boxes = [], []
    for row_index, raw_row in enumerate(raw_rows):
        result = [""] * len(groups)
        result_boxes: list[list[tuple]] = [[] for _ in groups]
        row_boxes = raw_boxes[row_index] if row_index < len(raw_boxes) else []
        for col_index, raw_value in enumerate(raw_row):
            value = clean_cell(raw_value)
            if not value: continue
            box = row_boxes[col_index] if col_index < len(row_boxes) else None
            group_index = group_for_box(box)
            result[group_index] = f"{result[group_index]}<br>{value}" if result[group_index] else value
            if box: result_boxes[group_index].append(box)
        canonical.append(result)
        canonical_boxes.append([union_rect(rects) if rects else None for rects in result_boxes])
    return canonical, canonical_boxes


def merge_header_rows(rows: list[list[str]], count: int) -> list[str]:
    if not rows: return []
    width = max(len(row) for row in rows)
    headers = []
    for col in range(width):
        values = []
        for row in rows[:count]:
            if col < len(row) and row[col]:
                val = row[col].replace("<br>", " / ")
                if val not in values: values.append(val)
        headers.append(" / ".join(values) or f"Coluna {col + 1}")
    return headers


def table_signature(headers: Sequence[str]) -> str:
    return "|".join(collapse_spaces(h).lower() for h in headers)


def table_has_header(table: Any) -> bool:
    rows, _ = canonical_table_rows(table)
    return bool(rows) and is_table_header_row(rows[0])


def render_table(table: Any, page: Any, links: Sequence[PdfLink], previous_state: Optional[TableState],
                 emitted_urls: set[str]) -> tuple[str, TableState, int]:
    rows, boxes = canonical_table_rows(table)
    if not rows: return "", previous_state or TableState([], 0, ""), 0

    header_count = 1 if is_table_header_row(rows[0]) else 0

    candidate_headers = merge_header_rows(rows, header_count) if header_count else []
    continuation = bool(
        previous_state and candidate_headers and table_signature(candidate_headers) == previous_state.signature)

    if continuation:
        headers = previous_state.headers
        state = previous_state
    elif header_count:
        headers = candidate_headers
        state = TableState(headers, len(headers), table_signature(headers))
    elif previous_state and previous_state.column_count == len(rows[0]):
        headers = previous_state.headers
        state = previous_state
        continuation = True
    else:
        headers = [f"Coluna {i + 1}" for i in range(len(rows[0]))]
        state = TableState(headers, len(headers), table_signature(headers))

    def decorate_row(row_index: int, row: list[str]) -> list[str]:
        output = []
        for col, value in enumerate(row):
            value = value or ""
            bbox = boxes[row_index][col] if row_index < len(boxes) and col < len(boxes[row_index]) else None
            if bbox and any(overlaps(link.rect, bbox, minimum=0.05) for link in links):
                value, _ = apply_links(value.replace("<br>", " "), bbox, links, emitted_urls)
            output.append(value.replace("• ", "- ").replace("|", r"\|"))
        return output[:len(headers)] + [""] * (len(headers) - len(output))

    data_start = header_count if not continuation else 0
    output_rows = [decorate_row(i, rows[i]) for i in range(data_start, len(rows))]
    if not output_rows and header_count == 0:
        output_rows = [decorate_row(0, rows[0])]

    lines = [] if continuation else ["| " + " | ".join(headers) + " |",
                                     "| " + " | ".join("---" for _ in headers) + " |"]
    for row in output_rows:
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines), state, sum(
        1 for link in links if overlaps(link.rect, rect_tuple(table.bbox), minimum=0.05))


def should_skip_block(block: TextBlock, table_boxes: Sequence[tuple], page_height: float) -> bool:
    text = clean_prose(block.text)
    if not text: return True
    if block.y0 >= page_height - 65 and RE_RODAPE.match(text): return True
    return any(overlaps(block.bbox, table_box, minimum=1.0) for table_box in table_boxes)


def raw_text_blocks(page: Any, table_boxes: Sequence[tuple]) -> list[TextBlock]:
    blocks = []
    height = float(page.rect.height)
    for raw in page.get_text("blocks", sort=True):
        block = TextBlock(float(raw[0]), float(raw[1]), float(raw[2]), float(raw[3]), str(raw[4]))
        if not should_skip_block(block, table_boxes, height):
            blocks.append(block)
    return blocks


def should_merge_blocks(previous: TextBlock, current: TextBlock) -> bool:
    prev_text, curr_text = clean_prose(previous.text), clean_prose(current.text)
    if not prev_text or not curr_text: return False
    if curr_text.startswith(("•", "CAPÍTULO", "RESOLUÇÃO", "CONSIDERANDO", "RESOLVE", "Art.")): return False
    if current.y0 - previous.y1 > 13: return False
    if curr_text[0].islower() or curr_text.startswith(("(", "a)", "b)", "c)", "I.", "II.")): return True
    return prev_text.startswith("•") and current.x0 >= previous.x0


def merge_text_blocks(blocks: Sequence[TextBlock]) -> list[TextBlock]:
    result = []
    for block in blocks:
        if result and should_merge_blocks(result[-1], block):
            prev = result[-1]
            prev.text = prev.text.rstrip() + " " + block.text.lstrip()
            prev.x0, prev.y0 = min(prev.x0, block.x0), min(prev.y0, block.y0)
            prev.x1, prev.y1 = max(prev.x1, block.x1), max(prev.y1, block.y1)
        else:
            result.append(block)
    return result


def heading_for_text(text: str) -> Optional[tuple[int, str]]:
    clean = clean_prose(text)
    if not clean: return None
    upper = clean.upper()
    if RE_CAPITULO.match(upper): return 1, clean
    if upper in {"SUMÁRIO", "APRESENTAÇÃO", "REVISÃO TEXTUAL", "EDITORAÇÃO"}: return 1, clean
    if upper.startswith("RESOLUÇÃO CJF N."): return 1, clean
    match = RE_ITEM.match(clean)
    if match:
        number, rest = match.groups()
        if len(clean) <= 180 and not rest.endswith((".", ";", ":")):
            return min(6, number.count(".") + 1), clean
    return None


def render_text_block(block: TextBlock, links: Sequence[PdfLink], emitted_urls: set[str]) -> tuple[str, str]:
    text = clean_prose(block.text)
    text, _ = apply_links(text, block.bbox, links, emitted_urls)
    heading = heading_for_text(text)
    if heading:
        level, content = heading
        return "#" * level + " " + content, "heading"
    if text.startswith("•"):
        return "- " + text[1:].strip(), "list"
    letter = RE_LISTA_LETRA.match(text)
    if letter:
        return f"- **{letter.group(1)})** {letter.group(2)}", "list"
    roman = RE_LISTA_ROMANA.match(text)
    if roman:
        return f"- **{roman.group(1).upper()}.** {roman.group(3)}", "list"
    return text, "paragraph"


def render_page_image(page: Any, output_path: Path, dpi: int = 144) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    scale = dpi / 72.0
    page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False).save(str(output_path))


def relative_asset_path(asset: Path, output: Path) -> str:
    return asset.relative_to(output.parent).as_posix()


def convert_pdf(source: Path, output: Path, include_page_markers: bool = True) -> ConversionStats:
    if pdfplumber is None:
        raise RuntimeError("pdfplumber é obrigatório para reconstrução de tabelas em PDF")

    output.parent.mkdir(parents=True, exist_ok=True)
    assets_dir = output.parent / f"{output.stem}_assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    stats = ConversionStats(str(source), str(output), "pdf", warnings=[])

    table_state: Optional[TableState] = None
    previous_table_physical: Optional[int] = None
    previous_table_bottom: Optional[float] = None
    pdf_url_set: set[str] = set()
    emitted_url_set: set[str] = set()

    fitz_doc = fitz.open(str(source))

    with open(output, "w", encoding="utf-8", newline="\n") as out_file:
        with pdfplumber.open(str(source)) as plumber_doc:
            stats.pages = len(fitz_doc)
            for physical, (fitz_page, plumber_page) in enumerate(zip(fitz_doc, plumber_doc.pages), start=1):
                links = extract_links(fitz_page)
                stats.links += len(links)
                pdf_url_set.update(link.uri for link in links)

                printed_page = extract_footer_number(fitz_page)
                marker = None
                if include_page_markers:
                    marker = f"<!-- Página física {physical}"
                    if printed_page is not None: marker += f" | página impressa {printed_page}"
                    marker += " -->"
                    stats.page_markers += 1

                words = fitz_page.get_text("words", sort=True)
                tables = plumber_page.find_tables()
                table_boxes = [rect_tuple(t.bbox) for t in tables]

                if not tables:
                    table_state, previous_table_physical, previous_table_bottom = None, None, None

                is_continuation_page = bool(
                    table_state and tables and previous_table_physical == physical - 1
                    and previous_table_bottom and previous_table_bottom > float(fitz_page.rect.height) - 180
                    and float(tables[0].bbox[1]) < 150 and not (table_has_header(tables[0]) if tables else False)
                )

                if not words:
                    stats.image_only_pages += 1
                    asset = assets_dir / f"page-{physical:03d}.png"
                    render_page_image(fitz_page, asset)
                    if marker: out_file.write(marker + "\n")
                    out_file.write(f"![Página física {physical} do PDF]({relative_asset_path(asset, output)})\n")
                    out_file.write(f"> Conteúdo visual da página {physical}; camada textual ausente.\n\n")
                    table_state = None
                    continue

                if marker and not is_continuation_page:
                    out_file.write(marker + "\n")

                blocks = merge_text_blocks(raw_text_blocks(fitz_page, table_boxes))
                elements: list[tuple[float, int, str, str]] = []

                for block in blocks:
                    rendered, kind = render_text_block(block, links, emitted_url_set)
                    if rendered:
                        elements.append((block.y0, 0, rendered, kind))

                for table_index, table in enumerate(tables):
                    rendered_table, new_table_state, _ = render_table(table, fitz_page, links, table_state,
                                                                      emitted_url_set)
                    table_state = new_table_state
                    if rendered_table:
                        elements.append((float(table.bbox[1]), 1, rendered_table, "table"))
                        stats.tables += 1

                elements.sort(key=lambda item: (item[0], item[1]))
                marker_pending = is_continuation_page and bool(marker)

                for _, _, rendered, kind in elements:
                    if marker_pending and kind != "table" and marker:
                        out_file.write(marker + "\n\n")
                        marker_pending = False
                    out_file.write(rendered + "\n")
                    if kind != "table":
                        out_file.write("\n")
                    if kind == "heading":
                        stats.headings += 1
                    elif kind == "list":
                        stats.list_items += 1

                if tables:
                    previous_table_physical = physical
                    previous_table_bottom = max(float(t.bbox[3]) for t in tables)

        unresolved_urls = sorted(pdf_url_set - emitted_url_set)
        stats.unresolved_link_urls = len(unresolved_urls)

        if unresolved_urls:
            out_file.write("\n## URLs do PDF sem âncora textual\n\n> Destinos preservados como manifesto.\n\n")
            for uri in unresolved_urls:
                out_file.write(f"- [Acessar destino do PDF](<{uri}>)\n")

    fitz_doc.close()
    return stats


def read_text_safely(path: Path) -> str:
    data = path.read_bytes()
    if from_bytes:
        try:
            best = from_bytes(data).best()
            if best: return nfc(str(best))
        except Exception as exc:
            LOGGER.warning("Falha na detecção de charset %s: %s", path, exc)
    return nfc(data.decode("utf-8", errors="replace"))


def csv_to_markdown(text: str) -> str:
    rows = list(csv.reader(io.StringIO(text)))
    if not rows: return ""
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    escaped = [[c.replace("|", r"\|").replace("\n", "<br>") for c in row] for row in rows]
    lines = ["| " + " | ".join(escaped[0]) + " |", "| " + " | ".join("---" for _ in range(width)) + " |"]
    lines += ["| " + " | ".join(row) + " |" for row in escaped[1:]]
    return "\n".join(lines)


def organize_markdown(text: str) -> str:
    text = nfc(text).replace("\r\n", "\n").replace("\r", "\n").strip()
    result, paragraph, in_fence = [], [], False

    def flush():
        if paragraph:
            result.append(collapse_spaces(" ".join(paragraph)))
            paragraph.clear()

    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("```"):
            flush()
            result.append(stripped)
            in_fence = not in_fence
            continue
        if in_fence:
            result.append(line.rstrip())
            continue
        if not stripped:
            flush()
            if result and result[-1] != "": result.append("")
            continue
        if stripped.startswith(("#", ">", "- ", "* ", "+ ", "|", "---")) or re.match(r"^\d+\.\s", stripped):
            flush()
            result.append(stripped)
        else:
            paragraph.append(stripped)
    flush()
    while result and result[-1] == "": result.pop()
    return "\n".join(result) + "\n"


def convert_image(source: Path, output: Path) -> ConversionStats:
    try:
        import pytesseract
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("OCR requer pytesseract e Pillow") from exc

    image = Image.open(source)
    try:
        text = pytesseract.image_to_string(image, lang="por+eng")
    except Exception as exc:
        raise RuntimeError("Falha no Tesseract; confira os idiomas por+eng") from exc

    body = organize_markdown(text) if text.strip() else "_Nenhum texto identificado por OCR._\n"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text("# Texto extraído por OCR\n\n" + body, encoding="utf-8", newline="\n")
    return ConversionStats(str(source), str(output), "image", warnings=[])


def convert_office(source: Path, output: Path) -> ConversionStats:
    """Converte arquivos do MS Office (docx, xlsx, pptx) para Markdown usando a biblioteca markitdown."""
    try:
        from markitdown import MarkItDown
    except ImportError as exc:
        raise RuntimeError(
            "A conversão de arquivos Office requer a biblioteca 'markitdown'. Instale com: pip install markitdown") from exc

    output.parent.mkdir(parents=True, exist_ok=True)

    try:
        md = MarkItDown()
        result = md.convert(str(source))
        text_content = result.text_content if result and result.text_content else ""
    except Exception as exc:
        raise RuntimeError(f"Falha ao processar o arquivo Office {source.name} com markitdown: {exc}") from exc

    body = organize_markdown(text_content) if text_content.strip() else "_Nenhum texto extraído do documento Office._\n"
    output.write_text(body, encoding="utf-8", newline="\n")

    return ConversionStats(str(source), str(output), source.suffix.lower().lstrip(".") or "office", warnings=[])


def convert_generic(source: Path, output: Path) -> ConversionStats:
    ext = source.suffix.lower()
    text = read_text_safely(source)
    if ext == ".md":
        body = organize_markdown(text)
    elif ext == ".csv":
        body = csv_to_markdown(text) + "\n"
    elif ext == ".json":
        try:
            body = "```json\n" + json.dumps(json.loads(text), ensure_ascii=False, indent=2) + "\n```\n"
        except json.JSONDecodeError:
            body = "```text\n" + text.strip() + "\n```\n"
    elif ext == ".xml":
        body = "```xml\n" + text.strip() + "\n```\n"
    elif ext in {".html", ".htm"} and BeautifulSoup:
        # Unificação com Remover.py: SEMPRE remove conteúdo revogado/tachado
        text = remove_striked_paragraphs(text)
        soup = BeautifulSoup(text, "html.parser")
        for tag in soup(["script", "style", "noscript"]): tag.decompose()
        body = organize_markdown(soup.get_text("\n"))
    else:
        body = organize_markdown(text)

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(body, encoding="utf-8", newline="\n")
    return ConversionStats(str(source), str(output), ext.lstrip(".") or "text", warnings=[])


def convert_html_text(html_text: str) -> str:
    """
    Converte uma string HTML diretamente para Markdown.
    SEMPRE aplica a remoção de conteúdo revogado/tachado (lógica do Remover.py).
    Usado pela API quando o conteúdo vem de upload ou de URL.
    """
    if not BeautifulSoup:
        raise RuntimeError("BeautifulSoup é obrigatório para processar HTML. Instale com: pip install beautifulsoup4")

    # 1. Remove conteúdo revogado/tachado (unificação com Remover.py)
    html_limpo = remove_striked_paragraphs(html_text)

    # 2. Limpa scripts/styles/noscript
    soup = BeautifulSoup(html_limpo, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    # 3. Extrai texto e organiza em Markdown
    texto = soup.get_text("\n")
    return organize_markdown(texto)


def detect_format(filename: str, content_type: str = "", content: bytes = b"") -> str:
    """
    Detecta o formato do conteúdo com base em (em ordem de prioridade):
      1. Content-Type HTTP (se fornecido)
      2. Extensão do nome do arquivo
      3. Assinatura mágica do conteúdo (magic bytes)

    Retorna um dos: 'pdf', 'html', 'office', 'image', 'text'
    """
    ct = (content_type or "").lower().split(";")[0].strip()

    # 1. Content-Type tem prioridade para URLs (mais confiável que extensão)
    if ct == "application/pdf":
        return "pdf"
    if ct in ("text/html", "application/xhtml+xml"):
        return "html"
    if ct.startswith("image/"):
        return "image"
    if ct in (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "application/msword", "application/vnd.ms-excel", "application/vnd.ms-powerpoint",
        "application/vnd.oasis.opendocument.text",
        "application/vnd.oasis.opendocument.spreadsheet",
        "application/vnd.oasis.opendocument.presentation",
    ):
        return "office"
    if ct.startswith("text/"):
        return "text"

    # 2. Extensão do nome do arquivo
    ext = Path(filename or "").suffix.lower()
    if ext == ".pdf":
        return "pdf"
    if ext in (".html", ".htm", ".xhtml"):
        return "html"
    if ext in EXTENSOES_IMAGEM:
        return "image"
    if ext in EXTENSOES_OFFICE:
        return "office"
    if ext in EXTENSOES_TEXTO:
        return "text"

    # 3. Magic bytes (fallback para URLs sem extensão clara)
    if content:
        if content[:5] == b"%PDF-":
            return "pdf"
        # ZIP-based (docx/xlsx/pptx/odt) — assinatura PK\x03\x04
        if content[:4] == b"PK\x03\x04":
            return "office"
        # PNG / JPEG / GIF / BMP / TIFF / WebP
        if content[:8] == b"\x89PNG\r\n\x1a\n":
            return "image"
        if content[:3] == b"\xff\xd8\xff":
            return "image"
        if content[:6] in (b"GIF87a", b"GIF89a"):
            return "image"
        if content[:2] == b"BM":
            return "image"
        if content[:4] in (b"II*\x00", b"MM\x00*"):
            return "image"
        if content[:4] == b"RIFF" and content[8:12] == b"WEBP":
            return "image"
        # Heurística de HTML: presença de <html, <!doctype html, <body
        amostra = content[:2048].decode("utf-8", errors="ignore").lower()
        if "<!doctype html" in amostra or "<html" in amostra or "<body" in amostra:
            return "html"

    # Padrão: trata como texto
    return "text"


def convert_bytes(
    content: bytes,
    filename: str,
    content_type: str = "",
    include_page_markers: bool = True,
    output: Optional[Path] = None,
) -> tuple[str, ConversionStats]:
    """
    Converte conteúdo em bytes para Markdown, roteando por detecção automática
    de formato (Content-Type + extensão + magic bytes).

    Usado pela API FastAPI quando o conteúdo vem de upload direto ou de URL remota.
    Retorna (markdown_text, stats).
    """
    formato = detect_format(filename, content_type, content)

    # Salva o conteúdo em arquivo temporário para reaproveitar toda a lógica existente
    import tempfile
    ext_map = {
        "pdf": ".pdf", "html": ".html", "office": Path(filename or "").suffix.lower() or ".docx",
        "image": Path(filename or "").suffix.lower() or ".png", "text": Path(filename or "").suffix.lower() or ".txt",
    }
    ext = ext_map.get(formato, ".bin")

    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
        tmp.write(content)
        tmp_path = Path(tmp.name)

    try:
        if output is None:
            output = tmp_path.with_suffix(".md")

        if formato == "html":
            # Caminho otimizado: processa HTML diretamente da string (sem passagem por disco)
            # para garantir que a remoção de revogados seja aplicada com a codificação correta.
            try:
                text = read_text_safely(tmp_path)
            except Exception:
                text = content.decode("utf-8", errors="replace")
            markdown = convert_html_text(text)
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_text(markdown, encoding="utf-8", newline="\n")
            stats = ConversionStats(str(filename or "<url>"), str(output), "html", warnings=[])
            return markdown, stats

        # Para outros formatos, delega para a função convert() existente
        stats = convert(tmp_path, output=output, include_page_markers=include_page_markers)
        markdown = output.read_text(encoding="utf-8")
        # Ajusta o campo source no stats para refletir a origem (filename/URL), não o tmp
        stats.source = filename or "<bytes>"
        stats.output = str(output)
        return markdown, stats
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass


def convert(source: Path, output: Optional[Path] = None, include_page_markers: bool = True) -> ConversionStats:
    source = source.expanduser().resolve()
    if not source.exists() or not source.is_file():
        raise FileNotFoundError(f"Arquivo não encontrado: {source}")
    output = (output or source.with_suffix(".md")).expanduser().resolve()
    ext = source.suffix.lower()

    if ext == ".pdf": return convert_pdf(source, output, include_page_markers)
    if ext in EXTENSOES_IMAGEM: return convert_image(source, output)
    if ext in EXTENSOES_TEXTO: return convert_generic(source, output)
    if ext in EXTENSOES_OFFICE: return convert_office(source, output)  # Roteamento adicionado aqui

    # Fallback genérico caso a extensão não esteja mapeada
    try:
        from markitdown import MarkItDown
        output.write_text(organize_markdown(MarkItDown().convert(str(source)).text_content), encoding="utf-8",
                          newline="\n")
        return ConversionStats(str(source), str(output), ext.lstrip(".") or "binary", warnings=[])
    except ImportError as exc:
        raise RuntimeError(f"Formato {ext} desconhecido ou biblioteca markitdown ausente.") from exc


# ==========================================
# FUNÇÕES DE CLI E BATCH (REINCLUÍDAS)
# ==========================================
def listar_arquivos_elegiveis(diretorio: Path) -> list[Path]:
    script_path = Path(__file__).resolve()
    extensoes = {".pdf"} | EXTENSOES_IMAGEM | EXTENSOES_TEXTO | EXTENSOES_OFFICE  # Adicionado OFFICE
    ignorar_sufixos = ("_unificado.md", "_otimizado.md", "_organizado.md", ".report.json", ".qa.json", ".render.json")
    candidatos = []
    for path in diretorio.expanduser().resolve().iterdir():
        if not path.is_file() or path.name.startswith("."): continue
        if path.resolve() == script_path: continue
        if path.suffix.lower() not in extensoes: continue
        if path.name.endswith(ignorar_sufixos): continue
        candidatos.append(path)

    stems_nao_markdown = {path.stem.casefold() for path in candidatos if path.suffix.lower() != ".md"}
    candidatos = [path for path in candidatos if
                  not (path.suffix.lower() == ".md" and path.stem.casefold() in stems_nao_markdown)]
    return sorted(candidatos, key=lambda path: path.name.casefold())


def construir_saidas_lote(fontes: Sequence[Path], output_arg: Optional[Path]) -> dict[Path, Path]:
    if output_arg is None:
        diretorio_saida = None
    elif output_arg.exists() and output_arg.is_dir():
        diretorio_saida = output_arg
    elif output_arg.suffix.lower() == ".md":
        diretorio_saida = output_arg.parent
    else:
        diretorio_saida = output_arg

    if diretorio_saida is not None:
        diretorio_saida = diretorio_saida.expanduser().resolve()
    #    diretorio_saida.mkdir(parents=True, exist_ok=True) - (cria pasta)

    contagem_stem = Counter(fonte.stem for fonte in fontes)
    usadas: set[Path] = set()
    saidas: dict[Path, Path] = {}
    for fonte in fontes:
        base_dir = diretorio_saida or fonte.parent
        nome_base = fonte.stem
        if contagem_stem[nome_base] > 1:
            nome_base = f"{nome_base}_{fonte.suffix.lower().lstrip('.') or 'arquivo'}"
        if fonte.suffix.lower() == ".md":
            nome_base = f"{nome_base}_organizado"
        destino = base_dir / f"{nome_base}.md"
        contador = 2
        while destino in usadas:
            destino = base_dir / f"{nome_base}_{contador}.md"
            contador += 1
        usadas.add(destino)
        saidas[fonte] = destino
    return saidas


def caminho_relatorio_lote(output_arg: Optional[Path], fontes: Sequence[Path], report_arg: Optional[Path]) -> Path:
    if report_arg is not None:
        if report_arg.exists() and report_arg.is_dir():
            destino = report_arg / "conversao_lote.report.json"
        elif report_arg.suffix.lower() == ".json":
            destino = report_arg
        else:
            destino = report_arg / "conversao_lote.report.json"
    elif output_arg is not None and output_arg.suffix.lower() != ".md":
        destino = output_arg / "conversao_lote.report.json"
    else:
        destino = fontes[0].parent / "conversao_lote.report.json"
    # destino.parent.mkdir(parents=True, exist_ok=True) (cria pasta)
    return destino


def processar_lote(fontes: Sequence[Path], output_arg: Optional[Path], report_arg: Optional[Path],
                   include_page_markers: bool) -> tuple[dict[str, Any], Path]:
    saidas = construir_saidas_lote(fontes, output_arg)
    resultados: list[dict[str, Any]] = []
    for fonte in fontes:
        destino = saidas[fonte]
        try:
            LOGGER.info("Processando %s -> %s", fonte.name, destino.name)
            stats = convert(fonte, output=destino, include_page_markers=include_page_markers)
            resultado = {"status": "sucesso", **stats.as_dict()}
        except Exception as exc:
            LOGGER.error("Falha em %s: %s", fonte.name, exc)
            resultado = {"status": "erro", "source": str(fonte.resolve()), "output": str(destino.resolve()),
                         "format": fonte.suffix.lower().lstrip("."), "error": str(exc)}
        resultados.append(resultado)

    report_path = caminho_relatorio_lote(output_arg, fontes, report_arg)
    payload = {
        "modo": "lote", "diretorio": str(Path.cwd().resolve()), "total_arquivos": len(fontes),
        "sucessos": sum(item["status"] == "sucesso" for item in resultados),
        "falhas": sum(item["status"] == "erro" for item in resultados),
        "arquivos": resultados, "relatorio": str(report_path.resolve())
    }
    report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return payload, report_path


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Converte um arquivo ou todos os arquivos elegíveis do diretório atual para Markdown.")
    parser.add_argument("input", nargs="?", type=Path, default=None,
                        help="Arquivo de origem; se omitido, processa todos do diretório")
    parser.add_argument("--output", type=Path, default=None, help="Arquivo de saída ou diretório")
    parser.add_argument("--no-page-markers", action="store_true", help="Não emitir comentários de página no Markdown")
    parser.add_argument("--report", type=Path, default=None, help="Relatório JSON")
    parser.add_argument("--verbose", action="store_true", help="Ativa logging detalhado")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(levelname)s | %(message)s")
    try:
        if args.input is None:
            fontes = listar_arquivos_elegiveis(Path.cwd())
            if not fontes:
                raise FileNotFoundError("Nenhum arquivo elegível encontrado.")
            payload, report_path = processar_lote(fontes, output_arg=args.output, report_arg=args.report,
                                                  include_page_markers=not args.no_page_markers)
            print(json.dumps(payload, ensure_ascii=False, indent=2))
            return 0 if payload["falhas"] == 0 else 1

        stats = convert(args.input, output=args.output, include_page_markers=not args.no_page_markers)
        if args.report:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(json.dumps(stats.as_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(stats.as_dict(), ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        LOGGER.error("Falha na conversão: %s", exc)
        if args.verbose:
            LOGGER.exception("Detalhes da falha")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())