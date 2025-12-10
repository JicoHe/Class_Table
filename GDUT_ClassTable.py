import pandas as pd
from datetime import datetime
from collections import defaultdict
import os

# ========== 配置 ==========
PERIOD_TIME = {
    1: ('08:30', '09:15'), 2: ('09:20', '10:05'),
    3: ('10:25', '11:10'), 4: ('11:15', '12:00'),
    5: ('13:50', '14:35'), 6: ('14:40', '15:25'),
    7: ('15:30', '16:15'), 8: ('16:30', '17:15'),
    9: ('17:20', '18:05'), 10: ('18:30', '19:15'),
    11: ('19:20', '20:05'), 12: ('20:10', '20:55'),
}
# =========================

def parse_period_str(p_str):
    """解析 '0102', '101112' 这种格式的节次字符串"""
    if pd.isna(p_str) or p_str == '':
        return 1, 1
    
    p_str = str(p_str).strip()
    # 如果是单数字 (e.g. 1)
    if len(p_str) == 1:
        return int(p_str), int(p_str)
        
    # 每两位截取
    periods = []
    for i in range(0, len(p_str), 2):
        try:
            periods.append(int(p_str[i:i+2]))
        except ValueError:
            pass
            
    if not periods:
        return 1, 1
        
    return min(periods), max(periods)

def parse_csv_file(file_path):
    """解析 CSV 文件"""
    print(f"正在解析 CSV 文件: {file_path}")
    try:
        # 尝试读取 CSV，跳过可能的坏行
        df = pd.read_csv(file_path, encoding='utf-8', on_bad_lines='skip')
        
        # 清理列名（去除可能的空格）
        df.columns = [c.strip() for c in df.columns]
        
        courses = []
        for _, row in df.iterrows():
            try:
                date_str = str(row['排课日期']).strip()
                date = datetime.strptime(date_str, '%Y-%m-%d')
                
                period_str = str(row['节次']).strip()
                # 补零，例如 '102' -> '0102'
                if len(period_str) % 2 != 0:
                    period_str = '0' + period_str
                
                start_p, end_p = parse_period_str(period_str)
                
                courses.append({
                    'date': date,
                    'week': int(row['周次']),
                    'start_period': start_p,
                    'end_period': end_p,
                    'course_name': str(row['课程名称']).strip(),
                    'location': str(row['上课地点']).strip(),
                    'teacher': str(row['教师']).strip(),
                    'class_name': str(row['班级名称']).strip(),
                    'summary': str(row['授课内容简介']).strip() if pd.notna(row['授课内容简介']) else ''
                })
            except Exception as e:
                # print(f"跳过一行: {e}")
                continue
                
        return courses
    except Exception as e:
        print(f"读取 CSV 失败: {e}")
        return []

def save_excel_from_list(courses):
    weeks_data = defaultdict(list)
    
    for c in courses:
        week_num = c['week']
        start_p = c['start_period']
        end_p = c['end_period']
        
        st, _ = PERIOD_TIME.get(start_p, ('08:30', '09:15'))
        _, et = PERIOD_TIME.get(end_p, ('08:30', '09:15'))
        time_str = f"{st}-{et}"
        
        row = {
            '日期': c['date'].strftime('%Y-%m-%d'),
            '星期': f"星期{c['date'].isoweekday()}",
            '时间': time_str,
            '课程名称': c['course_name'],
            '教室': c['location'],
            '教师': c['teacher'],
            '班级': c['class_name'],
            '授课内容': c['summary'],
            '原始节次': f"{start_p}-{end_p}"
        }
        weeks_data[week_num].append(row)
        
    # 写入Excel
    output_file = 'ClassTable.xlsx'
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        sorted_weeks = sorted(weeks_data.keys())
        for week in sorted_weeks:
            df = pd.DataFrame(weeks_data[week])
            # 按日期和时间排序
            df = df.sort_values(by=['日期', '时间'])
            sheet_name = f'第{week}周'
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
    print(f"✓ Excel: {output_file}")

def main():
    print('=' * 50)
    print('广工课表解析 & Excel生成')
    print('=' * 50)
    
    csv_file = 'ClassTable.csv'
    if os.path.exists(csv_file):
        courses = parse_csv_file(csv_file)
        if courses:
            print(f"✓ 解析到 {len(courses)} 条课程记录")
            save_excel_from_list(courses)
            print('\n下一步: 修改 Excel 文件后，运行 python excel_to_ics.py 生成日历')
            return
    
    print(f"未找到 {csv_file}，请确保文件存在。")

if __name__ == '__main__':
    main()
