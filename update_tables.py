import sys
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 读取CSV数据
df = pd.read_csv(r'C:\Users\Administrator\Desktop\survey_data_模拟数据.csv')

print('=== 数据统计 ===')
print(f'总样本量: {len(df)}')

# 表1: 基本情况
# 性别
male = len(df[df['Q1_性别'] == '男'])
female = len(df[df['Q1_性别'] == '女'])

# 身份
sports_major = len(df[df['Q2_身份'] == '体育学院篮球专项'])
amateur = len(df[df['Q2_身份'] == '普通篮球爱好者（公体课/社团）'])

# 运动年限
year1 = len(df[df['Q3_运动年限'] == '1年以下'])
year2 = len(df[df['Q3_运动年限'] == '1-3年'])
year3 = len(df[df['Q3_运动年限'] == '3-5年'])
year4 = len(df[df['Q3_运动年限'] == '5年以上'])

# 每周频率
freq1 = len(df[df['Q4_每周频率'] == '1-2次'])
freq2 = len(df[df['Q4_每周频率'] == '3-4次'])
freq3 = len(df[df['Q4_每周频率'] == '5次及以上'])

print(f'性别: 男{male}人({male/68*100:.1f}%), 女{female}人({female/68*100:.1f}%)')

# 表2: 损伤发生情况
injured = len(df[df['Q5_是否受伤'] == '是'])
not_injured = len(df[df['Q5_是否受伤'] == '否'])
print(f'受伤情况: 有{injured}人({injured/68*100:.1f}%), 无{not_injured}人({not_injured/68*100:.1f}%)')

# 表3: 损伤类型（统计受伤人群）
injured_df = df[df['Q5_是否受伤'] == '是']
print(f'受伤人群: {len(injured_df)}人')

# 统计各种损伤类型
meniscus = len(injured_df[injured_df['Q7_损伤类型'].str.contains('半月板', na=False)])
patella = len(injured_df[injured_df['Q7_损伤类型'].str.contains('髌骨劳损', na=False)])
ligament = len(injured_df[injured_df['Q7_损伤类型'].str.contains('韧带', na=False)])
tendon = len(injured_df[injured_df['Q7_损伤类型'].str.contains('髌腱炎', na=False)])
synovitis = len(injured_df[injured_df['Q7_损伤类型'].str.contains('滑膜炎', na=False)])
other = len(injured_df[injured_df['Q7_损伤类型'].str.contains('其他', na=False)])

print(f'损伤类型: 半月板{meniscus}人, 髌骨劳损{patella}人, 韧带{ligament}人, 髌腱炎{tendon}人')

# 表4: 损伤原因
cause1 = len(injured_df[injured_df['Q9_损伤原因'].str.contains('准备活动', na=False)])
cause2 = len(injured_df[injured_df['Q9_损伤原因'].str.contains('疲劳', na=False)])
cause3 = len(injured_df[injured_df['Q9_损伤原因'].str.contains('技术动作', na=False)])
cause4 = len(injured_df[injured_df['Q9_损伤原因'].str.contains('场地', na=False)])
cause5 = len(injured_df[injured_df['Q9_损伤原因'].str.contains('其他', na=False)])

print(f'损伤原因: 准备活动{cause1}人, 疲劳{cause2}人, 技术动作{cause3}人, 场地{cause4}人')

# 更新Word文档
doc_path = r'C:\Users\Administrator\Desktop\毕业论文_刘小龙_修正版.docx'
doc = Document(doc_path)

# 找到并更新4个表格
table_index = 0
for table in doc.tables:
    # 获取表格前的标题
    table_para = table._element.getprevious()
    if table_para is not None:
        title_text = table_para.text if hasattr(table_para, 'text') else ''
        
        if '表1' in title_text or '基本情况' in title_text:
            # 更新表1
            data = [
                ['项目', '类别', '人数', '百分比(%)'],
                ['性别', '男', str(male), f'{male/68*100:.1f}'],
                ['', '女', str(female), f'{female/68*100:.1f}'],
                ['身份', '体育学院', str(sports_major), f'{sports_major/68*100:.1f}'],
                ['', '普通爱好者', str(amateur), f'{amateur/68*100:.1f}'],
                ['运动年限', '1年以下', str(year1), f'{year1/68*100:.1f}'],
                ['', '1-3年', str(year2), f'{year2/68*100:.1f}'],
                ['', '3-5年', str(year3 if year3 > 0 else 0), f'{year3/68*100:.1f}' if year3 > 0 else '0.0'],
                ['', '5年以上', str(year4), f'{year4/68*100:.1f}'],
                ['每周频率', '1-2次', str(freq1), f'{freq1/68*100:.1f}'],
                ['', '3-4次', str(freq2), f'{freq2/68*100:.1f}'],
                ['', '5次及以上', str(freq3), f'{freq3/68*100:.1f}']
            ]
            for i, row_data in enumerate(data):
                if i < len(table.rows):
                    for j, cell_text in enumerate(row_data):
                        if j < len(table.rows[i].cells):
                            table.rows[i].cells[j].text = cell_text
            print('表1已更新')
            
        elif '表2' in title_text or '发生情况' in title_text:
            # 更新表2
            data = [
                ['损伤情况', '人数', '百分比(%)'],
                ['有损伤史', str(injured), f'{injured/68*100:.1f}'],
                ['无损伤史', str(not_injured), f'{not_injured/68*100:.1f}']
            ]
            for i, row_data in enumerate(data):
                if i < len(table.rows):
                    for j, cell_text in enumerate(row_data):
                        if j < len(table.rows[i].cells):
                            table.rows[i].cells[j].text = cell_text
            print('表2已更新')
            
        elif '表3' in title_text or '类型分布' in title_text:
            # 更新表3
            injured_count = len(injured_df)
            data = [
                ['损伤类型', '人数', '百分比(%)'],
                ['半月板损伤', str(meniscus), f'{meniscus/injured_count*100:.1f}' if injured_count > 0 else '0.0'],
                ['髌骨劳损', str(patella), f'{patella/injured_count*100:.1f}' if injured_count > 0 else '0.0'],
                ['韧带损伤', str(ligament), f'{ligament/injured_count*100:.1f}' if injured_count > 0 else '0.0'],
                ['髌腱炎', str(tendon), f'{tendon/injured_count*100:.1f}' if injured_count > 0 else '0.0'],
                ['其他', str(other), f'{other/injured_count*100:.1f}' if injured_count > 0 else '0.0']
            ]
            for i, row_data in enumerate(data):
                if i < len(table.rows):
                    for j, cell_text in enumerate(row_data):
                        if j < len(table.rows[i].cells):
                            table.rows[i].cells[j].text = cell_text
            print('表3已更新')
            
        elif '表4' in title_text or '主要原因' in title_text:
            # 更新表4
            injured_count = len(injured_df)
            data = [
                ['损伤原因', '人数', '百分比(%)'],
                ['准备活动不充分', str(cause1), f'{cause1/injured_count*100:.1f}' if injured_count > 0 else '0.0'],
                ['运动疲劳', str(cause2), f'{cause2/injured_count*100:.1f}' if injured_count > 0 else '0.0'],
                ['技术动作不规范', str(cause3), f'{cause3/injured_count*100:.1f}' if injured_count > 0 else '0.0'],
                ['场地因素', str(cause4), f'{cause4/injured_count*100:.1f}' if injured_count > 0 else '0.0'],
                ['其他', str(cause5), f'{cause5/injured_count*100:.1f}' if injured_count > 0 else '0.0']
            ]
            for i, row_data in enumerate(data):
                if i < len(table.rows):
                    for j, cell_text in enumerate(row_data):
                        if j < len(table.rows[i].cells):
                            table.rows[i].cells[j].text = cell_text
            print('表4已更新')

# 保存
doc.save(doc_path)
print(f'\n文档已保存: {doc_path}')
