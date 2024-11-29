import openpyxl

# 定义文件名、路径和工作表名称
file_name = 'example.xlsx'
file_path = 'D:/'
sheet_name = '明细'
column_names = ['交易时间', '月份', '来源', '收/支', '支付状态', '类型', '交易对方', '商品', '金额', '收/支',
       '计入/不计入', '乘后金额']

# 创建一个工作簿和一个工作表
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = sheet_name

# 写入列名到第一行
sheet.append(column_names)

# 保存文件
workbook.save(file_path + file_name)

print(f'Excel 文件已保存到: {file_path + file_name}')