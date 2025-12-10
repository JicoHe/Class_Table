import pandas as pd
from datetime import datetime
import uuid

REMINDER_MINUTES = 10

def generate_ics_from_excel(excel_file):
    ics = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//GDUT//CN',
        'X-WR-CALNAME:广工课程表',
        'X-WR-TIMEZONE:Asia/Shanghai',
    ]
    
    try:
        xls = pd.ExcelFile(excel_file)
    except FileNotFoundError:
        print(f"错误: 找不到文件 {excel_file}")
        return

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # 确保必要的列存在
        required_columns = ['日期', '时间', '课程名称']
        if not all(col in df.columns for col in required_columns):
            print(f"警告: 工作表 {sheet_name} 缺少必要的列，跳过。")
            continue
            
        for index, row in df.iterrows():
            try:
                date_str = str(row['日期']).split(' ')[0] # 处理可能的时间戳格式
                time_range = str(row['时间'])
                course_name = str(row['课程名称'])
                location = str(row['教室']) if pd.notna(row['教室']) else ''
                teacher = str(row['教师']) if pd.notna(row['教师']) else ''
                class_name = str(row['班级']) if pd.notna(row['班级']) else ''
                summary = str(row['授课内容']) if '授课内容' in row and pd.notna(row['授课内容']) else ''
                
                if '-' not in time_range:
                    continue
                    
                start_time_str, end_time_str = time_range.split('-')
                
                start_dt_str = f"{date_str} {start_time_str}"
                end_dt_str = f"{date_str} {end_time_str}"
                
                start_dt = datetime.strptime(start_dt_str, '%Y-%m-%d %H:%M')
                end_dt = datetime.strptime(end_dt_str, '%Y-%m-%d %H:%M')
                
                # 生成唯一ID
                uid = str(uuid.uuid4())
                
                description = f"教师: {teacher}\\n班级: {class_name}"
                if summary:
                    description += f"\\n内容: {summary}"
                
                # 构建事件
                event_block = [
                    'BEGIN:VEVENT',
                    f'UID:{uid}',
                    f'DTSTART:{start_dt:%Y%m%dT%H%M%S}',
                    f'DTEND:{end_dt:%Y%m%dT%H%M%S}',
                    f'SUMMARY:{course_name}',
                    f'LOCATION:{location}',
                    f'DESCRIPTION:{description}',
                ]

                if REMINDER_MINUTES > 0:
                    event_block.extend([
                        'BEGIN:VALARM',
                        'ACTION:DISPLAY',
                        'DESCRIPTION:该上课了',
                        f'TRIGGER:-PT{REMINDER_MINUTES}M',
                        'END:VALARM'
                    ])
                
                event_block.append('END:VEVENT')
                ics.extend(event_block)
            except Exception as e:
                print(f"处理 {sheet_name} 第 {index+2} 行时出错: {e}")
                continue

    ics.append('END:VCALENDAR')
    
    output_file = 'ClassTable.ics'
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(ics))
    print(f'✓ ICS文件已生成: {output_file} (含 {REMINDER_MINUTES} 分钟课前提醒)')

if __name__ == '__main__':
    generate_ics_from_excel('ClassTable.xlsx')
