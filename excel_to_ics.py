import pandas as pd
from datetime import datetime
import uuid

REMINDER_MINUTES = 10

# ----------------------------
# ICS 文本处理：转义 + 标准折行
# ----------------------------
def escape_ics_text(text: str) -> str:
    """
    转义 ICS 不允许直接出现的字符：
    - "\" 变成 "\\\\"
    - ";" → "\\;"
    - "," → "\\,"
    - 换行 → "\\n"
    """
    if text is None:
        return ""
    text = str(text)
    text = text.replace("\\", "\\\\")
    text = text.replace(";", "\\;")
    text = text.replace(",", "\\,")
    text = text.replace("\n", "\\n")
    return text


def fold_ics_line(line: str) -> str:
    """
    ICS 每行最多 75 字节，超过需要折行
    折行为：
        原行
        空格开头继续内容
    """
    encoded = line.encode('utf-8')
    if len(encoded) <= 75:
        return line

    folded = []
    current = ""

    for ch in line:
        if len((current + ch).encode('utf-8')) > 75:
            folded.append(current)
            current = " " + ch  # 折行必须以空格开头
        else:
            current += ch
    folded.append(current)
    return "\r\n".join(folded)


# ----------------------------
# 主功能：生成 ICS
# ----------------------------
def generate_ics_from_excel(excel_file):
    ics = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//JicoHe//ClassTable//CN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:广工课程表",
        "X-WR-TIMEZONE:Asia/Shanghai"
    ]

    try:
        xls = pd.ExcelFile(excel_file)
    except FileNotFoundError:
        print(f"错误: 找不到文件 {excel_file}")
        return

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)

        required_columns = ['日期', '时间', '课程名称']
        if not all(col in df.columns for col in required_columns):
            print(f"警告: 工作表 {sheet_name} 缺少必要的列，跳过。")
            continue

        for index, row in df.iterrows():
            try:
                date_str = str(row['日期']).split(' ')[0]
                time_range = str(row['时间']).strip()
                course_name = str(row['课程名称'])

                if '-' not in time_range:
                    continue

                teacher = escape_ics_text(str(row['教师']) if pd.notna(row['教师']) else "")
                class_name = escape_ics_text(str(row['班级']) if pd.notna(row['班级']) else "")
                location = escape_ics_text(str(row['教室']) if pd.notna(row['教室']) else "")
                summary = escape_ics_text(str(row['授课内容']) if '授课内容' in row and pd.notna(row['授课内容']) else "")

                start_time_str, end_time_str = time_range.split('-')
                start_dt = datetime.strptime(f"{date_str} {start_time_str}", '%Y-%m-%d %H:%M')
                end_dt = datetime.strptime(f"{date_str} {end_time_str}", '%Y-%m-%d %H:%M')

                uid = str(uuid.uuid4())
                dtstamp = datetime.now().strftime('%Y%m%dT%H%M%S')

                description = f"教师: {teacher}\\n班级: {class_name}"
                if summary:
                    description += f"\\n内容: {summary}"

                # 使用折行
                event_lines = [
                    fold_ics_line("BEGIN:VEVENT"),
                    fold_ics_line(f"UID:{uid}"),
                    fold_ics_line(f"DTSTAMP:{dtstamp}"),
                    fold_ics_line(f"DTSTART:{start_dt:%Y%m%dT%H%M%S}"),
                    fold_ics_line(f"DTEND:{end_dt:%Y%m%dT%H%M%S}"),
                    fold_ics_line(f"SUMMARY:{escape_ics_text(course_name)}"),
                    fold_ics_line(f"LOCATION:{location}"),
                    fold_ics_line(f"DESCRIPTION:{description}")
                ]

                if REMINDER_MINUTES > 0:
                    event_lines.extend([
                        "BEGIN:VALARM",
                        "ACTION:DISPLAY",
                        "DESCRIPTION:该上课了",
                        f"TRIGGER:-PT{REMINDER_MINUTES}M",
                        "END:VALARM"
                    ])

                event_lines.append("END:VEVENT")
                ics.extend(event_lines)

            except Exception as e:
                print(f"处理 {sheet_name} 第 {index+2} 行时出错: {e}")
                continue

    ics.append("END:VCALENDAR")

    output = "ClassTable.ics"
    with open(output, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(ics))

    print(f"✓ ICS 文件已生成：{output}（已符合 macOS / iOS 标准）")


if __name__ == "__main__":
    generate_ics_from_excel("ClassTable.xlsx")
