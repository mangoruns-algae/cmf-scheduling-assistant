from io import BytesIO

import pandas as pd


REQUIRED_SHEETS = ["人员信息", "人员可用性", "用工需求", "排班规则"]


def read_input_excel(file_obj):
    xls = pd.ExcelFile(file_obj)
    missing = [sheet for sheet in REQUIRED_SHEETS if sheet not in xls.sheet_names]
    if missing:
        raise ValueError(f"缺少必需工作表：{', '.join(missing)}")
    return {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in REQUIRED_SHEETS}


def export_schedule_to_excel(result_df, original_data=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if original_data:
            for name, df in original_data.items():
                df.to_excel(writer, sheet_name=name, index=False)
        result_df.to_excel(writer, sheet_name="排班结果", index=False)
    output.seek(0)
    return output
