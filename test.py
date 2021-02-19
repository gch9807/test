# import openpyxl
# from openpyxl import Workbook
#
# workbook = Workbook()
# # workbook.save(r'C:\Users\admin-geng\Desktop\test.xlsx')
# wb = openpyxl.load_workbook(r'C:\Users\admin-geng\Desktop\e-Approval Requirements v4 - RPA Testing.xlsx')
# ws = wb['PO']
# # cell = ws.cell(row=2, column=1)
# # cell2 = ws.cell(row=2, column=2)
# # # print(cell.font.strike)
# # # print(cell2.font.strike)
# # # print(type(ws))
# a = 0
# tittle_data = ['row']
# openpyxl_data = []
# for inx1, i in enumerate(ws['J':'V']):
#     row_list = []
#     for inx, j in enumerate(i):
#         flag = 0
#         if inx == 0:
#             tittle_data.append(j.value)
#         row_list.append(inx+1)
#         if j.font.strike:
#             flag = 1
#             print('j.value>>>>',j.value)
#             row_list.append(j.value)
#             print([j.value])
#         row_list.append('')
#     print(f'row_list>>>>>>>{row_list}')
#     if flag:
#         openpyxl_data.append(row_list)
# ww = workbook.active
# rows = len(openpyxl_data)
# lines = len(tittle_data)
# print(tittle_data)
# print(openpyxl_data)
# # for i in range(rows):
# #     for j in range(lines):
# #         ww.cell(row=i + 1, column=j + 1).value = openpyxl_data[i][j]
# # workbook.save(r'C:\Users\admin-geng\Desktop\test.xlsx')
#
# # import openpyxl
# #
# # openpyxl_data = [
# #     ('我', '们', '在', '这', '寻', '找'),
# #     ('我', '们', '在', '这', '失', '去'),
# #     ('p', 'y', 't', 'h', 'o', 'n')
# # ]
# # output_file_name = 'openpyxl_file.xlsx'
# #
# #
# # def save_excel(target_list, output_file_name):
# #     """
# #     将数据写入xlsx文件
# #     """
# #     if not output_file_name.endswith('.xlsx'):
# #         output_file_name += '.xlsx'
# #
# #     # 创建一个workbook对象，而且会在workbook中至少创建一个表worksheet
# #     wb = openpyxl.Workbook()
# #     # 获取当前活跃的worksheet,默认就是第一个worksheet
# #     ws = wb.active
# #     title_data = ('a', 'b', 'c', 'd', 'e', 'f')
# #     target_list.insert(0, title_data)
# #     rows = len(target_list)
# #     lines = len(target_list[0])
# #     for i in range(rows):
# #         for j in range(lines):
# #             ws.cell(row=i + 1, column=j + 1).value = target_list[i][j]
# #
# #     # 保存表格
# #     wb.save(filename=output_file_name)
# #
# #
# # save_excel(openpyxl_data, output_file_name)
#
#
# # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# # 使用openpyxl打开表格，进行单元格
# wb = openpyxl.load_workbook('F:\公司文件\e-Approval Requirements v4 - RPA Testing.xlsx')
# ws = wb['PO']
# # 取读进来列表的行数，当做接下来循环的行数开始与结束
# index_l = len(L1)
# # print(index_l)
#
# # print(df_c.columns.values)
# # 获取接下来需要取取的列所在列数与列名
# output_excel_columns = []
# # 这个是写入文件的表头列表
# excel_columns = ['行号']
# # 需要存储的数据列表
# excel_date = []
# # 获取checker列索引，并制作表头
# for k, v in enumerate(df_c.columns.values):
#     if 'Checker' in v:
#         output_excel_columns.append((k + 1, v))
#         cell = ws.cell(1, column=k + 1)
#         excel_columns.append(v)
#
# for i in range(2, index_l + 2):
#     d1 = [i, ]
#     # 加入判断，是1的话，这行数据不需要写入
#     judge = 1
#     for col in output_excel_columns:
#         # 根据行数与列数进行单元格取值，判断
#         cell = ws.cell(i, col[0])
#         if cell.font.strike == True:
#             d1.append(cell.value)
#             # 如果该单元格带有删除线，那么把判断变成0，也就是说该行需要写入
#             judge = 0
#         else:
#             # 这行数据没有删除线，数据为空
#             d1.append('')
#     # 判断，该行有带有删除线的数据，写入到excel_date
#     if judge == 0:
#         excel_date.append(d1)
# print(excel_date)
# # print(output_excel_columns[0][0])

import pandas as pd

a = pd.read_excel(r'C:\Users\admin-geng\Desktop\data.xlsx')
print(a)
for i in a:
    print(i)