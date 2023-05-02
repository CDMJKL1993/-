import openpyxl

# 打开“总表”文件
wb = openpyxl.load_workbook('C:\\Users\\jkl05\\Desktop\\工作量04月份 - 副本\\总表.xlsx')
ws = wb.active

# 定义需要读取的表格名称列表
table_names = ['张永铭', '王启连','张玉萍','侯国君','张宁','于萌萌','孙龙','崔慧敏','李晓璐','吴彤','崔鑫培']

# 循环读取每个表格的第37列数据，并复制到“总表”的对应列中
for i, name in enumerate(table_names):
    # 打开当前表格
    file_path = 'C:\\Users\\jkl05\\Desktop\\工作量04月份 - 副本\\{}.xlsx'.format(name)
    file = openpyxl.load_workbook(file_path)
    sheet = file.active

    # 复制第37列数据,从第4行到第382行
    for j, row in enumerate(sheet.iter_rows(min_row=4, min_col=37, max_row=382)):
        value = row[0].value
        if value is None:
            value = 0
        ws.cell(row=j + 4, column=i + 6, value=value)

    # 复制第5列第1行的单元格内容到“总表”的第6列第3行
    value = sheet.cell(row=1, column=5).value
    ws.cell(row=3, column=6+i, value=value)

    if name == '张永铭':
        # 复制第2-5列第387行开始的单元格内容到“总表”的对应列中
        for j, row in enumerate(sheet.iter_rows(min_row=387, min_col=2, max_col=5)):
            for k, cell in enumerate(row):
                value = cell.value
                if value is None:
                    value = 0
                ws.cell(row=j+387, column=k+2, value=value)

    file.close()

# 保存并关闭“总表”文件
wb.save('C:\\Users\\jkl05\\Desktop\\工作量04月份 - 副本\\总表.xlsx')
wb.close()
