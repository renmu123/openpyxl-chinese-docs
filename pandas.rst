与 Pandas 和 NumPy 一起使用
=============================

openpyxl 可以与流行的 `Pandas <http://pandas.pydata.org>`_ 和 `NumPy <http://numpy.org>`_ 一起使用


NumPy 支持
-------------

openpyxl 内置支持 NumPy 的float，integer 和 boolean 类型。
DateTimes are supported using the Pandas' Timestamp type.


和 Pandas Dataframes 一起使用
------------------------------

:func:`openpyxl.utils.dataframe.dataframe_to_rows` 提供了一种使用 Pandas Dataframes 的简单方法::

    from openpyxl.utils.dataframe import dataframe_to_rows
    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(df, index=True, header=True):
        ws.append(r)


虽然Pandas本身支持对Excel的转换，但这为客户端代码提供了更多的灵活性，包括直接将数据帧（stream dataframes）流传输到文件的能力。

将 dataframe 转换为工作簿时高亮表头和索引::

    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(df, index=True, header=True):
        ws.append(r)

    for cell in ws['A'] + ws[1]:
        cell.style = 'Pandas'

    wb.save("pandas_openpyxl.xlsx")

另外，如果你只想转换数据，你可以使用只写模式::

    from openpyxl.cell.cell import WriteOnlyCell
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()

    cell = WriteOnlyCell(ws)
    cell.style = 'Pandas'

     def format_first_row(row, cell):

        for c in row:
            cell.value = c
            yield cell

    rows = dataframe_to_rows(df)
    first_row = format_first_row(next(rows), cell)
    ws.append(first_row)

    for row in rows:
        row = list(row)
        cell.value = row[0]
        row[0] = cell
        ws.append(row)

    wb.save("openpyxl_stream.xlsx")


此代码和标准工作簿一起起作用。


将工作簿转换为 Dataframe（PS：样例文件可以参考df.to_excel()的文件）
-------------------------------------

如果工作簿没有表头和索引很容易用 `values` 属性将一个工作簿转换为 Dataframe::

    df = DataFrame(ws.values)

如果工作簿确实有表头和索引，例如 Pandas 创建的文件，那还要做更多的一些工作::

    from itertools import islice
    data = ws.values
    cols = next(data)[1:]
    data = list(data)
    idx = [r[0] for r in data]
    data = (islice(r, 1, None) for r in data)
    df = DataFrame(data, index=idx, columns=cols)
