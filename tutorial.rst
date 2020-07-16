Tutorial
========

新建工作表
-----------------

无须再文件系统中创建文件即可开始使用openpyxl。只要导入 `Workbook` 类就可以开始工作了::

    >>> from openpyxl import Workbook
    >>> wb = Workbook()

一个工作表至少有一个工作簿. 你可以通过 `Workbook.active` 来获取这个属性::

    >>> ws = wb.active

.. note::

    This is set to 0 by default. Unless you modify its value, you will always
    get the first worksheet by using this method.

你可以使用 `Workbook.create_sheet` 方法来创建新的工作簿::

    >>> ws1 = wb.create_sheet("Mysheet") # insert at the end (default)
    # or
    >>> ws2 = wb.create_sheet("Mysheet", 0) # insert at first position
    # or
    >>> ws3 = wb.create_sheet("Mysheet", -1) # insert at the penultimate position

工作薄在创建时会自动生成一个名字，以(Sheet, Sheet1, Sheet2, ...)来进行命名。你也可以通过 `Worksheet.title` 属性来改名::

    ws.title = "New Title"

默认情况下，包含该标题的选项卡的背景颜色为白色。你也可以使用 :code:`RRGGBB` 颜色来改变 `Worksheet.sheet_properties.tabColor` 属性::

    ws.sheet_properties.tabColor = "1072BA"

给工作表命名后，就可以将其作为工作簿的键::

    >>> ws3 = wb["New Title"]

你可以使用 `Workbook.sheetname` 属性查看工作簿中所有工作表的名称::

    >>> print(wb.sheetnames)
    ['Sheet2', 'New Title', 'Sheet1']

你可以遍历工作表::

    >>> for sheet in wb:
    ...     print(sheet.title)

You can create copies of worksheets **within a single workbook**:

`Workbook.copy_worksheet` method::

    >>> source = wb.active
    >>> target = wb.copy_worksheet(source)

.. note::

    Only cells (including values, styles, hyperlinks and comments) and
    certain worksheet attribues (including dimensions, format and
    properties) are copied. All other workbook / worksheet attributes
    are not copied - e.g. Images, Charts.

    You also **cannot** copy worksheets between workbooks. You cannot copy
    a worksheet if the workbook is open in `read-only` or `write-only`
    mode.


Playing with data
------------------

Accessing one cell
++++++++++++++++++

现在我们已经知道如何创建工作表，我们可以开始修改单元格内容了。
可以直接通过工作表的键来访问单元格::

    >>> c = ws['A4']

此处会返回 A4 单元格，如果不存在不将会进行创建
可以直接分配值::

    >>> ws['A4'] = 4

这里是 `Worksheet.cell` 方法.

也可以通过行列符号访问单元格::

    >>> d = ws.cell(row=4, column=2, value=10)

.. note::

    When a worksheet is created in memory, it contains no `cells`. They are
    created when first accessed.
    当工作薄在内存中被创建之后，其中没有单元格，只有在被第一次访问的时候才会创建

.. warning::

    Because of this feature, scrolling through cells instead of accessing them
    directly will create them all in memory, even if you don't assign them a value.

    Something like ::

        >>> for x in range(1,101):
        ...        for y in range(1,101):
        ...            ws.cell(row=x, column=y)

    will create 100x100 cells in memory, for nothing.


访问大量单元格
++++++++++++++++++++

可以使用切片来访问一系列单元格::

    >>> cell_range = ws['A1':'C2']


一系列的行和列也可以通过类似的方法获取::

    >>> colC = ws['C']
    >>> col_range = ws['C:D']
    >>> row10 = ws[10]
    >>> row_range = ws[5:10]

你也使用 `Worksheet.iter_rows` 方法::

    >>> for row in ws.iter_rows(min_row=1, max_col=3, max_row=2):
    ...    for cell in row:
    ...        print(cell)
    <Cell Sheet1.A1>
    <Cell Sheet1.B1>
    <Cell Sheet1.C1>
    <Cell Sheet1.A2>
    <Cell Sheet1.B2>
    <Cell Sheet1.C2>

同样 `Worksheet.iter_cols` 方法会返回列::

    >>> for col in ws.iter_cols(min_row=1, max_col=3, max_row=2):
    ...     for cell in col:
    ...         print(cell)
    <Cell Sheet1.A1>
    <Cell Sheet1.A2>
    <Cell Sheet1.B1>
    <Cell Sheet1.B2>
    <Cell Sheet1.C1>
    <Cell Sheet1.C2>

.. note::

  由于性能原因 `Worksheet.iter_cols()` 方法在只读模式下不可用。

如果你需要遍历文件中的所有行和列，你可以使用 `Worksheet.rows` 属性 ::

    >>> ws = wb.active
    >>> ws['C9'] = 'hello world'
    >>> tuple(ws.rows)
    ((<Cell Sheet.A1>, <Cell Sheet.B1>, <Cell Sheet.C1>),
    (<Cell Sheet.A2>, <Cell Sheet.B2>, <Cell Sheet.C2>),
    (<Cell Sheet.A3>, <Cell Sheet.B3>, <Cell Sheet.C3>),
    (<Cell Sheet.A4>, <Cell Sheet.B4>, <Cell Sheet.C4>),
    (<Cell Sheet.A5>, <Cell Sheet.B5>, <Cell Sheet.C5>),
    (<Cell Sheet.A6>, <Cell Sheet.B6>, <Cell Sheet.C6>),
    (<Cell Sheet.A7>, <Cell Sheet.B7>, <Cell Sheet.C7>),
    (<Cell Sheet.A8>, <Cell Sheet.B8>, <Cell Sheet.C8>),
    (<Cell Sheet.A9>, <Cell Sheet.B9>, <Cell Sheet.C9>))

或者 `Worksheet.columns` 属性::

    >>> tuple(ws.columns)
    ((<Cell Sheet.A1>,
    <Cell Sheet.A2>,
    <Cell Sheet.A3>,
    <Cell Sheet.A4>,
    <Cell Sheet.A5>,
    <Cell Sheet.A6>,
    ...
    <Cell Sheet.B7>,
    <Cell Sheet.B8>,
    <Cell Sheet.B9>),
    (<Cell Sheet.C1>,
    <Cell Sheet.C2>,
    <Cell Sheet.C3>,
    <Cell Sheet.C4>,
    <Cell Sheet.C5>,
    <Cell Sheet.C6>,
    <Cell Sheet.C7>,
    <Cell Sheet.C8>,
    <Cell Sheet.C9>))

.. note::

  由于性能原因 `Worksheet.columns` 方法在只读模式下不可用。


Values only
+++++++++++

如果你只想要工作薄的值，你可以使用 `Worksheet.values` 属性。
这会遍历工作簿中所有的行但只返回单元格值::

    for row in ws.values:
       for value in row:
         print(value)

`Worksheet.iter_rows` 和 `Worksheet.iter_cols` 可以用 :code:`values_only` 参数来近返回单元格值::

  >>> for row in ws.iter_rows(min_row=1, max_col=3, max_row=2, values_only=True):
  ...   print(row)

  (None, None, None)
  (None, None, None)


数据存储
------------

一旦有了 :class:`Cell`, 我们可以为其分配一个值::

    >>> c.value = 'hello, world'
    >>> print(c.value)
    'hello, world'

    >>> d.value = 3.14
    >>> print(d.value)
    3.14


保存至文件
++++++++++++++++

保存工作表最简单和安全的方法就是使用 :class:`Workbook` 类的 :func:`Workbook.save` 方法::

    >>> wb = Workbook()
    >>> wb.save('balances.xlsx')

.. warning::

   这个操作将会没有警告的覆盖已存在的we年

.. note::

    The filename extension is not forced to be xlsx or xlsm, although you might have
    some trouble opening it directly with another application if you don't
    use an official extension.

    As OOXML files are basically ZIP files, you can also  open it with your
    favourite ZIP archive manager.


保存成流
++++++++++++++++++

If you want to save the file to a stream, e.g. when using a web application
such as Pyramid, Flask or Django then you can simply provide a
:func:`NamedTemporaryFile`::


    >>> from tempfile import NamedTemporaryFile
    >>> from openpyxl import Workbook
    >>> wb = Workbook()
    >>> with NamedTemporaryFile() as tmp:
            wb.save(tmp.name)
            tmp.seek(0)
            stream = tmp.read()


你可以指定属性 `template=True` 将工作表保存为模板::

    >>> wb = load_workbook('document.xlsx')
    >>> wb.template = True
    >>> wb.save('document_template.xltx')

或者设置属性为 `False` (默认) 将其保存为一个文档::

    >>> wb = load_workbook('document_template.xltx')
    >>> wb.template = False
    >>> wb.save('document.xlsx', as_template=False)

.. warning::

    You should monitor the data attributes and document extensions
    for saving documents in the document templates and vice versa,
    otherwise the result table engine can not open the document.

.. note::

    以下操作将会失败::

    >>> wb = load_workbook('document.xlsx')
    >>> # Need to save with the extension *.xlsx
    >>> wb.save('new_document.xlsm')
    >>> # MS Excel can't open the document
    >>>
    >>> # or
    >>>
    >>> # Need specify attribute keep_vba=True
    >>> wb = load_workbook('document.xlsm')
    >>> wb.save('new_document.xlsm')
    >>> # MS Excel will not open the document
    >>>
    >>> # or
    >>>
    >>> wb = load_workbook('document.xltm', keep_vba=True)
    >>> # If we need a template document, then we must specify extension as *.xltm.
    >>> wb.save('new_document.xlsm')
    >>> # MS Excel will not open the document


从文件加载
-------------------

你可以使用 :func:`openpyxl.load_workbook` 方法来打开一个已存在的工作表::

    >>> from openpyxl import load_workbook
    >>> wb2 = load_workbook('test.xlsx')
    >>> print wb2.sheetnames
    ['Sheet2', 'New Title', 'Sheet1']

This ends the tutorial for now, you can proceed to the :doc:`usage` section
