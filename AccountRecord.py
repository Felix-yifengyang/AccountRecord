"""
需要创建一个新的xlsx文件，并且Sheet1重命名为'明细'
"""
import pandas as pd
import tkinter.filedialog
import datetime
import openpyxl
import msvcrt

def read_data_wx(path):  # 获取微信数据
    d_wx = pd.read_csv(path, header=16, skipfooter=0, encoding='GB2312', engine='python')
    d_wx = d_wx.iloc[:, [0, 4, 7, 1, 2, 3, 5]] # 时间，收/支，状态，交易类型，交易方，商品，金额

    # 转换数据类型
    d_wx.iloc[:, 0] = d_wx.iloc[:, 0].astype('datetime64[ns]')  # 数据类型更改
    d_wx.iloc[:, 6] = d_wx.iloc[:, 6].map(lambda date_str: date_str[1:]) # 删除符号 # 微信比支付宝多一个符号
    d_wx.iloc[:, 6] = d_wx.iloc[:, 6].astype('float64')  # 数据类型更改
    d_wx = d_wx.drop(d_wx[d_wx['收/支'] == '/'].index)  # 删除'收/支'为'/'的行
    d_wx.rename(columns={'当前状态': '支付状态', '交易类型': '类型', '金额(元)': '金额'}, inplace=True)  # 修改列名称
    d_wx.insert(1, '来源', "微信", allow_duplicates=True)  # 添加微信来源标识
    print("成功读取 " + str(len(d_wx)) + " 条「微信」账单数据\n")
    return d_wx

def read_data_alipay(path):  # 获取支付宝数据
    d_alipay = pd.read_csv(path,  header=4, skipfooter=7, encoding='GB2312', engine='python')
    d_alipay = d_alipay.rename(columns={column_name: column_name.strip() for column_name in d_alipay.columns}) # 支付宝比微信多了很多space
    d_alipay = d_alipay.iloc[:, [2, 10, 11, 6, 7, 8, 9]] # 时间，收/支，状态，类型，交易方，商品，金额

    # 转换数据类型
    d_alipay.iloc[:, 0] = d_alipay.iloc[:, 0].astype('datetime64[ns]')  # 数据类型更改
    d_alipay.iloc[:, 6] = d_alipay.iloc[:, 6].astype('float64')  # 数据类型更改
    d_alipay = d_alipay.drop(d_alipay[d_alipay['收/支'] == ''].index)  # 删除'收/支'为空的行
    d_alipay.rename(columns={'交易创建时间': '交易时间', '交易状态': '支付状态', '商品名称': '商品', '金额（元）': '金额'},
                 inplace=True)  # 修改列名称
    d_alipay.insert(1, '来源', "支付宝", allow_duplicates=True)  # 添加来源标识
    print("成功读取 " + str(len(d_alipay)) + " 条「支付宝」账单数据\n")
    return d_alipay

def add_cols(data):  # 增加3列数据
    # 逻辑1：取值-1 or 1。-1表示支出，1表示收入。
    data.insert(8, '收/支', -1, allow_duplicates=True)  # 插入列，默认值为-1
    for index in range(len(data.iloc[:, 2])):  # 遍历第3列的值，判断为收入，则改'逻辑1'为1
        if data.iloc[index, 2] == '收入':
            data.iloc[index, 8] = 1

    # 逻辑2：取值0 or 1。1表示计入，0表示不计入。
    data.insert(9, '计入/不计入', 1, allow_duplicates=True)  # 插入列，默认值为1
    for index in range(len(data.iloc[:, 3])):  # 遍历第4列的值，判断为资金流动，则改'逻辑2'为0
        col3 = data.iloc[index, 3]
        if (col3 == '提现已到账') or (col3 == '已全额退款') or (col3 == '已退款') or (col3 == '退款成功') or (col3 == '还款成功') or (
                col3 == '交易关闭'):
            data.iloc[index, 9] = 0

    # 月份
    data.insert(1, '月份', 0, allow_duplicates=True)  # 插入列，默认值为0
    for index in range(len(data.iloc[:, 0])):
        time = data.iloc[index, 0]
        data.iloc[index, 1] = time.month  # 访问月份属性的值，赋给这月份列

    # 乘后金额
    data.insert(11, '乘后金额', 0, allow_duplicates=True)  # 插入列，默认值为0
    for index in range(len(data.iloc[:, 8])):
        money = data.iloc[index, 8] * data.iloc[index, 9] * data.iloc[index, 10]
        data.iloc[index, 11] = float(money)
    return data


if __name__ == '__main__':

    # 路径设置
    print('提示：请在弹窗中选择要导入的【微信】账单文件\n')
    path_wx = tkinter.filedialog.askopenfilename(title='选择要导入的微信账单：', filetypes=[('所有文件', '.*'), ('csv文件', '.csv')])
    if path_wx == '':  # 判断是否只导入了微信或支付宝账单中的一个
        cancel_wx = 1
    else:
        cancel_wx = 0

    print('提示：请在弹窗中选择要导入的【支付宝】账单文件\n')
    path_alipay = tkinter.filedialog.askopenfilename(title='选择要导入的支付宝账单：', filetypes=[('所有文件', '.*'), ('csv文件', '.csv')])
    if path_alipay == '':  # 判断是否只导入了微信或支付宝账单中的一个
        cancel_alipay = 1
    else:
        cancel_alipay = 0

    while cancel_alipay == 1 and cancel_wx == 1:
        print('\n您没有选择任何一个账单！     请按任意键退出程序')
        ord(msvcrt.getch())

    path_account = tkinter.filedialog.askopenfilename(title='选择要导出的目标账本表格：', filetypes=[('所有文件', '.*'), ('Excel表格', '.xlsx')])
    while path_account == '':  # 判断是否选择了账本
        print('\n年轻人，不选账本怎么记账？      请按任意键退出程序')
        ord(msvcrt.getch())

    path_write = path_account

    # 判断是否只导入了微信或支付宝账单中的一个
    if cancel_wx == 1:
        data_wx = pd.DataFrame()
    else:
        data_wx = read_data_wx(path_wx)  # 读数据
    if cancel_alipay == 1:
        data_alipay = pd.DataFrame()
    else:
        data_alipay = read_data_alipay(path_alipay)  # 读数据

    data_merge = pd.concat([data_wx, data_alipay], axis=0)  # 上下拼接合并表格
    data_merge = add_cols(data_merge)  # 新增 逻辑、月份、乘后金额 3列
    # print(data_merge.columns)
    print("已自动计算乘后金额和交易月份，已合并数据")
    merge_list = data_merge.values.tolist()  # 格式转换，DataFrame->List
    workbook = openpyxl.load_workbook(path_account)  # openpyxl读取账本文件
    sheet = workbook['明细']
    maxrow = sheet.max_row  # 获取最大行
    print('\n「明细」 sheet 页已有 ' + str(maxrow) + ' 行数据，将在末尾写入数据')
    for row in merge_list:
        sheet.append(row)  # openpyxl写文件

    # 在最后1行写上导入时间，作为分割线
    now = datetime.datetime.now()
    now = '👆导入时间：' + str(now.strftime('%Y-%m-%d %H:%M:%S'))
    break_lines = [now, '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-']
    sheet.append(break_lines)

    workbook.save(path_write)  # 保存
    print("\n成功将数据写入到 " + path_write)
    print("\n运行成功！write successfully!    按任意键退出")
    ord(msvcrt.getch())