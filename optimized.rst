优化模式
===============


只读模式
--------------

有时，你可能需要打开或写入极端大的 XLSX 文件，但通用的 openpyxl 程序无法处理这么大的负载。
幸运的是，有两种模式可以使你在（几乎）恒定的内存消耗下读写无限量的数据。

介绍 :class:`openpyxl.worksheet._read_only.ReadOnlyWorksheet`::

    from openpyxl import load_workbook
    wb = load_workbook(filename='large_file.xlsx', read_only=True)
    ws = wb['big_data']

    for row in ws.rows:
        for cell in row:
            print(cell.value)

.. warning::

    * :class:`openpyxl.worksheet._read_only.ReadOnlyWorksheet` 是只读的

单元格的返回值不是 :class:`openpyxl.cell.cell.Cell` 而是
:class:`openpyxl.cell._read_only.ReadOnlyCell`.


工作表尺寸（dimensions）
++++++++++++++++++++

只读模式依赖创建文件的应用以及库提供工作表的正确信息，尤其是文件的已使用部分，称之为尺寸（dimensions）。
一些应用汇进行设置错误。可以使用 `ws.calculate_dimension()` 函数来检查工作表的尺寸（dimensions）。
如果返回和范围和你知道的不一样，比如说 `A1:A1` ，你可以简单重置 `max_row` 和 `max_column` 属性，即可使用该文件::

    ws.reset_dimensions()


只写模式
---------------

常规的 :class:`openpyxl.worksheet.worksheet.Worksheet` 被替代成更快的 :class:`openpyxl.worksheet._write_only.WriteOnlyWorksheet` 。
当你想导出大量数据的时候请确保安装了 `lxml` 库.

.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook(write_only=True)
>>> ws = wb.create_sheet()
>>>
>>> # now we'll fill it with 100 rows x 200 columns
>>>
>>> for irow in range(100):
...     ws.append(['%d' % i for i in range(200)])
>>> # save the file
>>> wb.save('new_big_file.xlsx') # doctest: +SKIP

如果你想要带有样式或者注释的单元格可以使用 :func:`openpyxl.cell.WriteOnlyCell`

.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook(write_only = True)
>>> ws = wb.create_sheet()
>>> from openpyxl.cell import WriteOnlyCell
>>> from openpyxl.comments import Comment
>>> from openpyxl.styles import Font
>>> cell = WriteOnlyCell(ws, value="hello world")
>>> cell.font = Font(name='Courier', size=36)
>>> cell.comment = Comment(text="A comment", author="Author's Name")
>>> ws.append([cell, 3.14, None])
>>> wb.save('write_only_file.xlsx')


以上会创建只有一张工作表的只写工作簿，一行写入（append）三个单元格：一个带有自定义字体和注释的文字单元格，一个浮点数单元格和一个空单元格（一定会被丢弃）。

.. warning::

    * 和普通工作簿不同的是，新创建的只写工作簿没有任何工作表；工作表只能由 :func:`create_sheet()` 方法进行创建。

    * 在只读工作簿中，只能由 :func:`append()` 来添加行。无法使用 :func:`cell()` 或 :func:`iter_rows()` 对任意位置的单元进行读取或写入。

    * 可以导出不限量的数据（即使超过 Excel 的处理上限），同时内存使用量小于10Mb。

    * 一个只写工作簿只能保存一次。之后如果任何尝试保存和添加数据（append()）的操作都会会引发 :class:`openpyxl.utils.exceptions.WorkbookAlreadySaved` 错误。

    * Everything that appears in the file before the actual cell data must be created
      before cells are added because it must written to the file before then.
      For example, `freeze_panes` should be set before cells are added.
