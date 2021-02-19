import pandas as pd
from pandas.core.frame import DataFrame

a = pd.read_excel(r'C:\Users\admin-geng\Desktop\data.xlsx')
# print(a)
b = a.values.tolist()
# # for i in a:
# #     print(i)
print(b)
del b[-1][2]
b[-1].append('')
print(type(b[-1][-1]))
# for i in b:
#     print(i)
data = DataFrame(b, index=None, columns=['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h'])
print(data)
# data.to_excel(r'C:\Users\admin-geng\Desktop\data2.xlsx', index=['a', 'b', 'c', 'd', 'e', 'f', 'g'])
data.to_excel(r'C:\Users\admin-geng\Desktop\data2.xlsx',columns=['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h'], index=None)
