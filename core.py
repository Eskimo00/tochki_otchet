import re
from pathlib import Path
from datetime import timedelta

from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Font


# Регулярные выражения для поиска данных
MODEL_RE = re.compile(r"Модель:\s*(.+?)(?:\s+Номер ТС:.*)?$", re.IGNORECASE)
PERIOD_RE = re.compile(r"Период:\s*с\s*(\d{2}\.\d{2}\.\d{4})", re.IGNORECASE)
POINT_RE = re.compile(r"Точка\s*(\d+)", re.IGNORECASE)

# Стили оформления
FILL_HEADER = PatternFill("solid", fgColor="FFF2CC")  # желтый для заголовков
FILL_BODY = PatternFill("solid", fgColor="FFFBEA")    # светло-желтый фон таблицы
FILL_GREEN = PatternFill("solid", fgColor="C6EFCE")   # зелёный для точек с проходами
FILL_RED = PatternFill("solid", fgColor="FFC7CE")     # красный для "нет"
THIN_BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)


def parse_duration(value) -> timedelta:
    """Парсим строку вида 00:05:12 в timedelta."""
    if value is None:
        return timedelta(0)
    s = str(value).strip()
    parts = s.split(":")
    if len(parts) != 3:
        return timedelta(0)
    try:
        h, m, sec = (int(float(p)) for p in parts)
    except ValueError:
        return timedelta(0)
    return timedelta(hours=h, minutes=m, seconds=sec)


def format_duration(td: timedelta) -> str:
    total_seconds = int(td.total_seconds())
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def detect_columns(ws):
    """Определяем номера колонок 'Название КТ' и 'Продолжительность' по заголовкам."""
    name_col = 3
    dur_col = 10
    max_scan_rows = min(40, ws.max_row)
    max_scan_cols = min(30, ws.max_column)
    for r in range(1, max_scan_rows + 1):
        for c in range(1, max_scan_cols + 1):
            text = ws.cell(r, c).value
            if not text:
                continue
            txt = str(text).strip().lower()
            if "название" in txt:
                name_col = c
            if "продолж" in txt:
                dur_col = c
    return name_col, dur_col


def ensure_unique_path(path: Path) -> Path:
    """Если файл существует — добавляем (1), (2) и т.д."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    i = 1
    while True:
        candidate = parent / f"{stem} ({i}){suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def build_default_output(source_path: Path) -> Path:
    base_dir = source_path.parent if source_path else Path.cwd()
    return ensure_unique_path(base_dir / "отчет прохождения точек.xlsx")


def process_file(source: Path, dest: Path):
    wb = load_workbook(source, data_only=True)
    ws = wb.active

    name_col, dur_col = detect_columns(ws)

    period_date = ""
    results = []
    current_model = None
    current_counts = None
    current_sums = None

    def flush_model():
        nonlocal current_model, current_counts, current_sums
        if current_model is None:
            return
        results.append({
            "name": current_model,
            "counts": current_counts,
            "sums": current_sums,
        })
        current_model = None
        current_counts = None
        current_sums = None

    for r in range(1, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        texts = [str(v) if v is not None else "" for v in row_vals]

        # Период
        for txt in texts:
            m = PERIOD_RE.search(txt)
            if m:
                period_date = m.group(1)

        # Начало блока Модель
        for txt in texts:
            m = MODEL_RE.search(txt)
            if m:
                flush_model()
                current_model = m.group(1).strip()
                current_counts = {i: 0 for i in range(1, 9)}
                current_sums = {i: timedelta(0) for i in range(1, 9)}
                break

        if current_model:
            name_txt = ""
            if name_col - 1 < len(row_vals):
                name_txt = texts[name_col - 1]
            m_pt = POINT_RE.search(name_txt)
            if m_pt:
                idx = int(m_pt.group(1))
                if 1 <= idx <= 8:
                    current_counts[idx] += 1
                    dur_txt = texts[dur_col - 1] if dur_col - 1 < len(texts) else ""
                    current_sums[idx] += parse_duration(dur_txt)

        # Конец блока
        for txt in texts:
            if "ИТОГО по ТС" in txt:
                flush_model()
                break

    flush_model()

    # Сортировка по названию объекта (по алфавиту)
    results.sort(key=lambda x: x["name"].lower())

    # Записываем итоговый файл
    out_wb = Workbook()
    out_ws = out_wb.active
    headers = ["№ п/п", "Название объекта", "Дата"] + [f"Точка {i}" for i in range(1, 9)]
    out_ws.append(headers)

    header_font = Font(bold=True)
    for cell in out_ws[1]:
        cell.fill = FILL_HEADER
        cell.border = THIN_BORDER
        cell.font = header_font

    for row_idx, row in enumerate(results, start=2):
        line_no = row_idx - 1
        line = [line_no, row["name"], period_date]
        point_values = []
        for p in range(1, 9):
            count = row["counts"][p]
            if count == 0:
                point_values.append("нет")
            else:
                point_values.append(f"({count}) / {format_duration(row['sums'][p])}")
        line.extend(point_values)
        out_ws.append(line)

        for col_idx, cell in enumerate(out_ws[row_idx], start=1):
            # Границы и базовая заливка
            cell.border = THIN_BORDER
            if col_idx >= 4 or (col_idx < 4 and row_idx > 1):
                cell.fill = FILL_BODY
            # Заливка только для столбцов точек
            if col_idx >= 4:
                value = cell.value
                if value == "нет":
                    cell.fill = FILL_RED
                else:
                    cell.fill = FILL_GREEN

    # Автоподбор ширины колонок
    for col_cells in out_ws.columns:
        max_len = 0
        col_letter = col_cells[0].column_letter
        for cell in col_cells:
            val = cell.value
            if val is None:
                continue
            length = len(str(val))
            if length > max_len:
                max_len = length
        out_ws.column_dimensions[col_letter].width = min(max_len + 2, 40)

    dest = ensure_unique_path(dest)
    out_wb.save(dest)
    return dest
