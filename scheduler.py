from collections import defaultdict

import pandas as pd


def _split_tags(value):
    if pd.isna(value):
        return set()
    return {item.strip() for item in str(value).replace("，", ";").split(";") if item.strip()}


def _is_yes(value):
    return str(value).strip().upper() in {"Y", "YES", "TRUE", "1"}


def _availability_map(df):
    mapping = {}
    if df is None or df.empty:
        return mapping
    df = df.copy()
    df["date"] = pd.to_datetime(df["date"]).dt.date
    for _, row in df.iterrows():
        mapping[(row["date"], str(row["employee_id"]))] = {
            "status": str(row.get("availability_status", "available")).strip().lower(),
            "shift": str(row.get("available_shift", "all")).strip().lower(),
            "fixed_line": "" if pd.isna(row.get("fixed_line")) else str(row.get("fixed_line")).strip(),
        }
    return mapping


def generate_schedule(staff_df, availability_df, demand_df, rules_df=None):
    staff = staff_df.copy()
    demand = demand_df.copy()
    staff["employee_id"] = staff["employee_id"].astype(str)
    demand["date"] = pd.to_datetime(demand["date"]).dt.date
    availability = _availability_map(availability_df)

    rows = []
    used = set()
    assignment_count = defaultdict(int)
    sort_cols = ["date"] + (["priority"] if "priority" in demand.columns else []) + ["production_line", "shift"]
    demand = demand.sort_values(sort_cols, kind="stable")

    for _, requirement in demand.iterrows():
        date = requirement["date"]
        line = str(requirement["production_line"]).strip()
        shift = str(requirement["shift"]).strip().lower()
        role = "" if pd.isna(requirement.get("required_role")) else str(requirement.get("required_role")).strip()
        skill = "" if pd.isna(requirement.get("required_skill")) else str(requirement.get("required_skill")).strip()
        need = int(requirement.get("required_headcount", 1))
        allow_borrowing = _is_yes(requirement.get("allow_borrowing", "Y"))

        candidates = []
        for _, person in staff.iterrows():
            employee_id = str(person["employee_id"])
            if not _is_yes(person.get("enabled", "Y")) or (date, shift, employee_id) in used:
                continue

            available = availability.get((date, employee_id), {"status": "available", "shift": "all", "fixed_line": ""})
            if available["status"] not in {"available", ""}:
                continue
            if available["shift"] not in {"all", shift, ""}:
                continue
            if available["fixed_line"] and available["fixed_line"] != line:
                continue

            home = "" if pd.isna(person.get("home_line")) else str(person.get("home_line")).strip()
            person_role = "" if pd.isna(person.get("role")) else str(person.get("role")).strip()
            if role and person_role != role:
                continue
            if skill and skill not in _split_tags(person.get("skills")):
                continue
            if line not in _split_tags(person.get("qualified_lines")) and line != home:
                continue

            is_home = home == line
            if not is_home and (not allow_borrowing or not _is_yes(person.get("can_cross_line", "N"))):
                continue
            if shift == "night" and not _is_yes(person.get("can_night_shift", "N")):
                continue
            candidates.append((0 if is_home else 1, assignment_count[employee_id], employee_id, person))

        candidates.sort(key=lambda item: (item[0], item[1], item[2]))
        selected = candidates[:need]
        for _, _, employee_id, person in selected:
            used.add((date, shift, employee_id))
            assignment_count[employee_id] += 1
            rows.append({
                "date": date, "production_line": line, "shift": shift,
                "batch_id": requirement.get("batch_id", ""), "task_type": requirement.get("task_type", ""),
                "required_role": role, "required_skill": skill, "employee_id": employee_id,
                "employee_name": person.get("employee_name", ""),
                "assignment_type": "本产线" if str(person.get("home_line", "")).strip() == line else "跨产线借调",
                "status": "已排班", "conflict_message": "",
            })

        shortage = need - len(selected)
        for _ in range(shortage):
            rows.append({
                "date": date, "production_line": line, "shift": shift,
                "batch_id": requirement.get("batch_id", ""), "task_type": requirement.get("task_type", ""),
                "required_role": role, "required_skill": skill, "employee_id": "", "employee_name": "",
                "assignment_type": "", "status": "缺员",
                "conflict_message": f"当前需求缺员 {shortage} 人；请检查人员可用性、岗位/技能资格或借调条件。",
            })

    result = pd.DataFrame(rows)
    if not result.empty:
        result["date"] = pd.to_datetime(result["date"])
    return result
