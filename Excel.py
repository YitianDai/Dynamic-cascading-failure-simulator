# import pandas as pd
# import numpy as np
# import matplotlib.pyplot as plt
# df = pd.read_csv('C:\\Yitian\\FS_CE_NETS_Py\\1.csv')
# #df=df.to_excel('C:\Yitian\FS_Cascading_Events_NETS\Frequency Response.xlsx', 'Line_02-03_Outage', index=False)
# x=df.ix[2:2000,'Results']

# #Titles = np.array(df.loc[-1, :])
# #x = np.array(df.loc[3:, 1])
# #print(Titles)
# for i in range(1,34):

#     a='bus',i
#     print(a)
#     y = df.loc[2:2000, a]
#     plt.plot(x, y, label = a, linewidth = 0.2)

# plt.xlabel('Time')
# plt.ylabel('Frequency Response')

# plt.title('Frequency Responses')
# plt.legend()
# plt.xticks(range(0, 20, 5))
# #plt.yticks(range(0, 2, 0.01))
# plt.show()


def read_excel(file, sheet, row1, row2, col1, col2, sf):
    import xlrd
    import numpy as np
    # file = 'ACTIVSg200.xlsx'
    list_data = []
    wb = xlrd.open_workbook(filename = file)
    sheet1 = wb.sheet_by_name(sheet)
    rows = sheet1.row_values(2)
    for j in range(col1 - 1, col2) :
        data = []
        for i in range(row1 - 1, row2) :
            if sheet1.cell(i,j).ctype == 2 :
                data.append(float(sheet1.cell(i,j).value) * sf)
            else:
                data.append(sheet1.cell(i,j).value)
        list_data.append(data)
    datamatrix = np.array(list_data)    
    return datamatrix 

# # test
# import itertools
# import numpy as np
# data = read_excel('Line Parameters', 3, 36, 1, 3, 1.0)
# Lines = [1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 23, 24, 25, 26, 27, 28, 29, 30, 31, 35, 36, 38, 40, 42, 43, 44, 45]
# p1= []
# p2 =[]
# for lines in itertools.combinations(Lines, 2):
#         p1.append(lines[0])
#         p1.append(lines[1])
#         p2.append(p1)
#         p1= []
# P= np.array(p2)
# print(p2)
# print(len(p2))




