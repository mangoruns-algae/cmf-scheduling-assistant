from __future__ import annotations

import html
import io
import re
from dataclasses import dataclass
from datetime import date, datetime

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles.colors import COLOR_INDEX


SCHEDULE_HEADERS = ("日期", "时间", "工作内容", "现场负责人", "人员安排")
SOURCE_HEADER_VARIANTS = {
    ("日期", "时间", "工作内容", "现场负责人", "人员安排"),
    ("日期", "时间", "生产内容", "现场负责人", "人员安排"),
}
NAME_EXCLUSIONS = {
    "无安排", "待安排", "负责人", "现场负责人", "人员安排", "工作内容", "全天", "白班", "夜班",
    "早班", "中班", "晚班", "休息", "培训", "确认", "复核", "其他", "生产", "设备",
    "冻干机", "外壁洗", "洗烘", "灌装", "包装", "提前出", "待确认",
}


@dataclass
class ScheduleWorkbook:
    workbook: object
    schedule_sheets: list[str]


def load_schedule_workbook(file_bytes: bytes) -> ScheduleWorkbook:
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheets = [sheet.title for sheet in workbook.worksheets if _find_header_row(sheet) is not None]
    if not sheets:
        raise ValueError("未找到包含“日期、时间、工作内容、现场负责人、人员安排”的排班工作表。")
    return ScheduleWorkbook(workbook=workbook, schedule_sheets=sheets)


def _find_header_row(sheet):
    for row in range(1, min(sheet.max_row, 12) + 1):
        values = [str(sheet.cell(row, col).value or "").strip() for col in range(1, min(sheet.max_column, 12) + 1)]
        for start in range(0, max(1, len(values) - len(SCHEDULE_HEADERS) + 1)):
            if tuple(values[start : start + len(SCHEDULE_HEADERS)]) in SOURCE_HEADER_VARIANTS:
                return row, start + 1
    return None


def sheet_layout(sheet):
    location = _find_header_row(sheet)
    if location is None:
        raise ValueError("所选工作表不是可识别的生产排班表。")
    header_row, first_col = location
    last_col = first_col + len(SCHEDULE_HEADERS) - 1
    last_row = max(
        row
        for row in range(header_row, sheet.max_row + 1)
        if any(sheet.cell(row, col).value not in (None, "") for col in range(first_col, last_col + 1))
    )
    return header_row, first_col, last_col, last_row


def _display_value(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return str(value)


def _merged_value(sheet, row, col):
    cell = sheet.cell(row, col)
    if not isinstance(cell, MergedCell):
        return _display_value(cell.value)
    for merged in sheet.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
            return _display_value(sheet.cell(merged.min_row, merged.min_col).value)
    return ""


def schedule_records(sheet):
    header_row, first_col, _, last_row = sheet_layout(sheet)
    records = []
    for row in range(header_row + 1, last_row + 1):
        values = [_merged_value(sheet, row, first_col + offset).strip() for offset in range(5)]
        if any(values):
            records.append({"row": row, **dict(zip(SCHEDULE_HEADERS, values))})
    return records


def available_people(sheet):
    candidates = set()
    for record in schedule_records(sheet):
        for field in ("现场负责人", "人员安排"):
            text = record[field]
            for token in re.split(r"[、，,；;\n/（）()：:\s]+", text):
                token = token.strip()
                if (
                    re.fullmatch(r"[\u4e00-\u9fff]{2,4}", token)
                    and token not in NAME_EXCLUSIONS
                    and not any(word in token for word in ("人员", "安排", "现场", "负责", "清洗", "灌装", "冻干", "设备", "提前"))
                ):
                    candidates.add(token)
    return sorted(candidates)


def matching_records(sheet, person):
    if not person:
        return []
    return [
        record
        for record in schedule_records(sheet)
        if person in record["现场负责人"] or person in record["人员安排"]
    ]


def generated_schedule_records(result):
    """Convert the scheduler result into the same five-column read-only view."""
    if result is None or result.empty:
        return []
    assigned = result[result["status"] == "已排班"].copy()
    records = []
    group_columns = ["date", "production_line", "shift", "batch_id", "task_type", "required_role"]
    for values, group in assigned.groupby(group_columns, dropna=False, sort=True):
        schedule_date, line, shift, batch, task, role = values
        names = "、".join(dict.fromkeys(str(name) for name in group["employee_name"] if str(name).strip()))
        role_text = "" if role is None else str(role)
        records.append({
            "row": len(records) + 2,
            "日期": schedule_date.strftime("%Y-%m-%d") if hasattr(schedule_date, "strftime") else str(schedule_date),
            "时间": str(shift),
            "工作内容": " / ".join(part for part in (str(line), str(batch), str(task)) if part and part != "nan"),
            "现场负责人": names if "负责人" in role_text else "",
            "人员安排": "" if "负责人" in role_text else names,
        })
    return records


def people_from_records(records):
    candidates = set()
    for record in records:
        for field in ("现场负责人", "人员安排"):
            for token in re.split(r"[、，,；;\n/（）()：:\s]+", record.get(field, "")):
                if re.fullmatch(r"[\u4e00-\u9fff]{2,4}", token or ""):
                    candidates.add(token)
    return sorted(candidates)


def render_records_html(records, selected_person=""):
    header = "".join(f'<td class="schedule-cell schedule-header">{label}</td>' for label in SCHEDULE_HEADERS)
    rows = [f"<tr>{header}</tr>"]
    for record in records:
        match_type = _record_match_type(record, selected_person)
        cells = []
        for index, label in enumerate(SCHEDULE_HEADERS):
            classes = "schedule-cell"
            if match_type == "participant":
                classes += " schedule-match"
            elif match_type == "leader" and index == 0:
                classes += " schedule-leader-match"
            cells.append(f'<td class="{classes}">{html.escape(str(record.get(label, ""))).replace(chr(10), "<br>")}</td>')
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return f"""
    <div class="schedule-wrap"><table class="schedule-table">
      <colgroup><col class="date-col"><col class="time-col"><col class="work-col"><col class="lead-col"><col class="people-col"></colgroup>
      {''.join(rows)}
    </table></div>
    """


def export_records_pdf(records, title, selected_person=""):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import KeepInFrame, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    buffer = io.BytesIO()
    document = SimpleDocTemplate(buffer, pagesize=landscape(A3), rightMargin=10 * mm, leftMargin=10 * mm, topMargin=10 * mm, bottomMargin=10 * mm, title=title)
    styles = getSampleStyleSheet()
    body = ParagraphStyle("GeneratedBody", parent=styles["BodyText"], fontName="STSong-Light", fontSize=7.5, leading=10, alignment=TA_LEFT)
    center = ParagraphStyle("GeneratedCenter", parent=body, alignment=TA_CENTER)
    heading = ParagraphStyle("GeneratedTitle", parent=styles["Title"], fontName="STSong-Light", fontSize=16, leading=20, textColor=colors.HexColor("#17365D"))
    story = [Paragraph(html.escape(title), heading), Spacer(1, 3 * mm)]
    if selected_person:
        story.extend([Paragraph(f"人员定位：{html.escape(selected_person)}（黄色为参与班次，蓝色日期为现场负责人）", body), Spacer(1, 2 * mm)])
    data = [[Paragraph(label, center) for label in SCHEDULE_HEADERS]]
    for record in records:
        data.append([Paragraph(html.escape(str(record.get(label, ""))), center if label != "工作内容" else body) for label in SCHEDULE_HEADERS])
    table = Table(data, colWidths=[31 * mm, 33 * mm, 133 * mm, 39 * mm, 150 * mm], repeatRows=1)
    commands = [
        ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"), ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#AEB8C6")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for index, record in enumerate(records, start=1):
        match_type = _record_match_type(record, selected_person)
        if match_type == "participant":
            commands.extend([("BACKGROUND", (0, index), (-1, index), colors.HexColor("#FFF2B2")), ("BOX", (0, index), (-1, index), 1, colors.HexColor("#E0A800"))])
        elif match_type == "leader":
            commands.extend([("BACKGROUND", (0, index), (0, index), colors.HexColor("#DCEBFF")), ("BOX", (0, index), (0, index), 1, colors.HexColor("#2F6FED"))])
    table.setStyle(TableStyle(commands))
    story.append(KeepInFrame(document.width, document.height - 28 * mm, [table], mode="shrink"))
    document.build(story)
    return buffer.getvalue()


def _record_match_type(record, person):
    if not person:
        return ""
    is_leader = person in record["现场负责人"]
    is_participant = person in record["人员安排"]
    if is_participant:
        return "participant"
    if is_leader:
        return "leader"
    return ""


def _color_rgb(color, fallback):
    if color is None:
        return fallback
    if color.type == "rgb" and color.rgb:
        return f"#{color.rgb[-6:]}"
    if color.type == "indexed" and color.indexed is not None and color.indexed < len(COLOR_INDEX):
        return f"#{COLOR_INDEX[color.indexed][-6:]}"
    return fallback


def render_schedule_html(sheet, selected_person=""):
    header_row, first_col, last_col, last_row = sheet_layout(sheet)
    records_by_row = {record["row"]: record for record in schedule_records(sheet)}
    # A selected-person view deliberately expands merged cells. This prevents a
    # date cell shared by several tasks from making unrelated tasks look selected.
    if selected_person:
        rows = []
        header = "".join(f'<td class="schedule-cell schedule-header">{label}</td>' for label in SCHEDULE_HEADERS)
        rows.append(f"<tr>{header}</tr>")
        for record in records_by_row.values():
            match_type = _record_match_type(record, selected_person)
            cells = []
            for index, label in enumerate(SCHEDULE_HEADERS):
                classes = "schedule-cell"
                if match_type == "participant":
                    classes += " schedule-match"
                elif match_type == "leader" and index == 0:
                    classes += " schedule-leader-match"
                cells.append(f'<td class="{classes}">{html.escape(record[label]).replace(chr(10), "<br>")}</td>')
            rows.append("<tr>" + "".join(cells) + "</tr>")
        return f"""
        <div class="schedule-wrap">
          <table class="schedule-table">
            <colgroup><col class="date-col"><col class="time-col"><col class="work-col"><col class="lead-col"><col class="people-col"></colgroup>
            {''.join(rows)}
          </table>
        </div>
        """

    merged_starts = {}
    merged_covered = set()
    for merged in sheet.merged_cells.ranges:
        if merged.max_row < header_row or merged.min_row > last_row or merged.max_col < first_col or merged.min_col > last_col:
            continue
        start = (max(merged.min_row, header_row), max(merged.min_col, first_col))
        end_row = min(merged.max_row, last_row)
        end_col = min(merged.max_col, last_col)
        merged_starts[start] = (end_row - start[0] + 1, end_col - start[1] + 1)
        for row in range(start[0], end_row + 1):
            for col in range(start[1], end_col + 1):
                if (row, col) != start:
                    merged_covered.add((row, col))

    rows = []
    for row in range(header_row, last_row + 1):
        cells = []
        for col in range(first_col, last_col + 1):
            if (row, col) in merged_covered:
                continue
            source = sheet.cell(row, col)
            value = _display_value(source.value)
            rowspan, colspan = merged_starts.get((row, col), (1, 1))
            fill = _color_rgb(source.fill.fgColor, "#ffffff") if source.fill.fill_type else "#ffffff"
            color = _color_rgb(source.font.color, "#263142") if source.font.color else "#263142"
            align = source.alignment.horizontal or ("center" if col != first_col + 2 else "left")
            weight = "700" if source.font.bold or row == header_row else "400"
            classes = "schedule-cell"
            if row == header_row:
                classes += " schedule-header"
            attrs = f' rowspan="{rowspan}" colspan="{colspan}"' if rowspan > 1 or colspan > 1 else ""
            style = f"background:{fill};color:{color};text-align:{align};font-weight:{weight};"
            cells.append(
                f'<td class="{classes}"{attrs} style="{style}">{html.escape(value).replace(chr(10), "<br>")}</td>'
            )
        rows.append("<tr>" + "".join(cells) + "</tr>")

    return f"""
    <div class="schedule-wrap">
      <table class="schedule-table">
        <colgroup><col class="date-col"><col class="time-col"><col class="work-col"><col class="lead-col"><col class="people-col"></colgroup>
        {''.join(rows)}
      </table>
    </div>
    """


def export_schedule_pdf(sheet, selected_person=""):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import KeepInFrame, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    buffer = io.BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A3),
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
        title=f"{sheet.title} 生产排班",
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ChineseTitle", parent=styles["Title"], fontName="STSong-Light", fontSize=16, leading=21, textColor=colors.HexColor("#17365D")
    )
    body_style = ParagraphStyle(
        "ChineseBody", parent=styles["BodyText"], fontName="STSong-Light", fontSize=7.5, leading=10, alignment=TA_LEFT
    )
    center_style = ParagraphStyle(
        "ChineseCenter", parent=body_style, alignment=TA_CENTER
    )
    story = [Paragraph(html.escape(str(sheet.cell(2, 2).value or f"{sheet.title} 生产排班")), title_style), Spacer(1, 4 * mm)]
    if selected_person:
        story.extend([
            Paragraph(f"人员定位：{html.escape(selected_person)}（黄色行为该人员相关班次）", body_style),
            Spacer(1, 3 * mm),
        ])

    header_row, first_col, last_col, last_row = sheet_layout(sheet)
    records_by_row = {record["row"]: record for record in schedule_records(sheet)}
    table_data = []
    for row in range(header_row, last_row + 1):
        rendered = []
        for col in range(first_col, last_col + 1):
            value = _merged_value(sheet, row, col)
            rendered.append(Paragraph(html.escape(value).replace("\n", "<br/>"), center_style if col != first_col + 2 else body_style))
        table_data.append(rendered)

    table = Table(table_data, colWidths=[31 * mm, 33 * mm, 133 * mm, 39 * mm, 150 * mm], repeatRows=1)
    commands = [
        ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("LEADING", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#AEB8C6")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#17365D")),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for source_row, record in records_by_row.items():
        table_row = source_row - header_row
        match_type = _record_match_type(record, selected_person)
        if match_type == "participant":
            commands.extend([
                ("BACKGROUND", (0, table_row), (-1, table_row), colors.HexColor("#FFF2B2")),
                ("BOX", (0, table_row), (-1, table_row), 1.0, colors.HexColor("#E0A800")),
            ])
        elif match_type == "leader":
            commands.extend([
                ("BACKGROUND", (0, table_row), (0, table_row), colors.HexColor("#DCEBFF")),
                ("BOX", (0, table_row), (0, table_row), 1.0, colors.HexColor("#2F6FED")),
            ])
    table.setStyle(TableStyle(commands))
    # KeepInFrame guarantees that even long schedules remain a single-page PDF.
    story.append(KeepInFrame(document.width, document.height - 32 * mm, [table], mode="shrink"))

    document.build(story)
    return buffer.getvalue()
