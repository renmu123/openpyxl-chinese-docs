插入删除行或列, 移动范围单元格
===============================================================


插入行和列
--------------------------

你可以使用工作表相关的方法来插入行和列:

    * :func:`openpyxl.worksheet.worksheet.Worksheet.insert_rows`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.insert_cols`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.delete_rows`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.delete_cols`

默认是一行或一列。 例如在第七行插入一行 (存在第七行)::

    >>> ws.insert_rows(7)


删除多行或多列
--------------------------

删除 ``F:H`` 列::

    >>> ws.delete_cols(6, 3)


Moving ranges of cells
----------------------

你也可以在一个工作表内移动范围单元格::

    >>> ws.move_range("D4:F10", rows=-1, cols=2)

这会将 ``D4:F10`` 单元格向上移动一行向右移动两列，已存在的单元格将会被覆盖

如果单元格包含公式，你可以让 openpyxl 帮你进行translate，但也并非总是你想要的结果，因此默认是禁用的。
同时，只有单元格本身的公式将会被translate。其他单元格对该单元格的引用或defined name将不会被更新。你可以使用 :doc:`formula` 来做这件事::

    >>> ws.move_range("G4:H10", rows=1, cols=1, translate=True)

公式中的相对引用移动一行和一列