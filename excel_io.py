from io import BytesIO

import pandas as pd


REQUIRED_SHEETS = ["人员信息", "人员可用性", "用工需求", "排班规则"]


def _date_only_frame(frame):
    """Copy a frame and normalize supported date columns to calendar dates."""
    normalized = frame.copy()
    for column in ("date", "日期"):
        if column not in normalized.columns:
            continue
        parsed = pd.to_datetime(normalized[column], errors="coerce")
        normalized[column] = parsed.dt.date.where(parsed.notna(), normalized[column])
    return normalized


def read_input_excel(file_obj):
    xls = pd.ExcelFile(file_obj)
    missing = [sheet for sheet in REQUIRED_SHEETS if sheet not in xls.sheet_names]
    if missing:
        raise ValueError(f"缺少必需工作表：{', '.join(missing)}")
    return {
        sheet: _date_only_frame(pd.read_excel(xls, sheet_name=sheet))
        for sheet in REQUIRED_SHEETS
    }


def export_schedule_to_excel(result_df, original_data=None):
    output = BytesIO()
    with pd.ExcelWriter(
        output,
        engine="openpyxl",
        date_format="YYYY-MM-DD",
        datetime_format="YYYY-MM-DD",
    ) as writer:
        if original_data:
            for name, df in original_data.items():
                _date_only_frame(df).to_excel(writer, sheet_name=name, index=False)
        _date_only_frame(result_df).to_excel(writer, sheet_name="排班结果", index=False)
        for worksheet in writer.book.worksheets:
            headers = {cell.value: cell.column for cell in worksheet[1]}
            for header in ("date", "日期"):
                column = headers.get(header)
                if column is None:
                    continue
                for row in range(2, worksheet.max_row + 1):
                    worksheet.cell(row, column).number_format = "yyyy-mm-dd"
    output.seek(0)
    return output
