import os
import pandas as pd
import re

from cyberbrain import trace
import time
import sys

# 获取当前根路径
base_path = os.getcwd()
# 错误信息
err_info = []
success_files_list = []


# 输出txt文档
def text_save(filename, text, open_type='a'):  #filename为写入CSV文件的路径，data为要写入数据列表
    folder_output_path = os.path.join(base_path, 'output_info')
    file_save_path = os.path.join(folder_output_path, filename)
    file = open(file_save_path, open_type)
    for item in text:
        file.write(str(item) + '\n')
    file.close()


# 获取文件目录
def get_files_list(path, foldername=''):
    folder_path = os.path.join(path, foldername)
    files_list = os.listdir(folder_path)
    return files_list


# 集合比较，返回差集/交集
def set_compare(list_a, list_b):
    set_a = set(list_a)
    set_b = set(list_b)
    diff_a = sorted(list(set_a - set_b))
    diff_b = sorted(list(set_b - set_a))
    unite_ab = sorted(list(set_a & set_b))
    return diff_a, diff_b, unite_ab


# 判断是否为数字
def is_number(str):
    try:
        # if str == 'NaN':
        #     return False
        float(str)
        return True
    except ValueError:
        return False


# 内容比较
def content_diff(cont_a, cont_b):
    if is_number(cont_a) or is_number(cont_b):
        return True
    if re.compile(u'[\u4e00-\u9fa5]').search(cont_a) or re.compile(u'[\u4e00-\u9fa5]').search(cont_b):
        return True
    return False


# 获取模板表：集团经营月报
template_folder_name = '模板表'
template_folder_path = os.path.join(base_path, template_folder_name)
template_financial_report_name = '集团经营月报模板.xlsx'
template_financial_report_path = os.path.join(template_folder_path, template_financial_report_name)
template_financial_report = pd.read_excel(template_financial_report_path, sheet_name=None, header=None, index=None)
financial_report_sheetname_list = list(template_financial_report)

# 新建合并表
merge_dataframe = pd.read_excel(template_financial_report_path, sheet_name=None, header=None, index=None)
save_folder_path = os.path.join(base_path, '结果表')
save_path = os.path.join(save_folder_path, '集团经营月报汇总.xlsx')
writer = pd.ExcelWriter(save_path)

# 获取待汇总表格
original_folder_name = '待汇总报表'
original_files_list = get_files_list(base_path, original_folder_name)
original_folder_path = os.path.join(base_path, original_folder_name)

err_diff_msg = []
for file_item in original_files_list:
    file_item_path = os.path.join(original_folder_path, file_item)
    item_workbook = pd.read_excel(file_item_path, sheet_name=None, header=None, index=None)
    item_workbook_sheetname_list = list(item_workbook)
    # 比较待汇总报表与集团经营月报模板：所含工作表差异
    diff_a, diff_b, unite_sheetname_list = set_compare(financial_report_sheetname_list, item_workbook_sheetname_list)
    if diff_a:
        err_info.append(f'<{template_financial_report_name}>与<{file_item}>存在差集：{diff_a}')
    if diff_b:
        err_info.append(f'<{file_item}>与<{template_financial_report_name}>存在差集：{diff_b}')

    print(file_item)
    # 加总工作表数值
    for sheetname_item in unite_sheetname_list:
        sheet_item = item_workbook[sheetname_item]
        merge_sheet = merge_dataframe[sheetname_item]
        sheet_item.fillna('', inplace=True)
        merge_sheet.fillna('', inplace=True)
        # try:
        # 遍历工作表列
        for column_index, row in sheet_item.iteritems():
            # 遍历行
            for row_index, item in enumerate(row):
                # 所有者权益变动表调整
                if sheetname_item == '所有者权益变动表' and len(sheet_item[0]) > 32:
                    if row_index == 31:
                        merge_sheet.loc[row_index, column_index] += sheet_item.loc[row_index + 3, column_index]
                        break
                # 行次列不进行加操作
                if item == '行次':
                    break
                # 若为数值，则加总
                if is_number(item):
                    if merge_sheet.loc[row_index, column_index] == '':
                        merge_sheet.loc[row_index, column_index] = 0.00
                    merge_sheet.loc[row_index, column_index] += item
                    continue
                # 若非数值，比较是否一致
                temp_value = merge_sheet.loc[row_index, column_index]
                if (item == '--' and not (temp_value)) or (not (temp_value) and not (item)):
                    continue
                if str(temp_value).strip() != str(item).strip():
                    err_diff_msg.append(
                        f'{file_item}<{sheetname_item}>[{row_index+1},{column_index+1}]单元格不一致:{temp_value},{item};')
        merge_dataframe[sheetname_item].to_excel(writer, sheet_name=sheetname_item, header=None, index=None)
        # except:
        #     print(sheetname_item, row_index + 1, column_index)
        #     sys.exit()

    success_files_list.append(file_item)
    # break

writer.save()
writer.close()
print(f'<{success_files_list}>已取数！')
text_save('content_diff_info.txt', err_diff_msg, 'w+')
text_save('err_info.txt', err_info, 'w+')
