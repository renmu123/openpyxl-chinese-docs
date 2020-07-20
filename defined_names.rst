定义名称
=============


该规范对定义的名称有以下说法:

    "定义名称是用于表示单元格，区域，公式或常量值的描述性文本。"

这意味着它们的定义是非常宽松的。它们可能包含一个常数，一个公式，一个单元格引用，一个区域或跨不同工作表的多个区域。
它们在工作簿全局定义并可以通过 `defined_names` 属性进行访问。


区域的使用示例
---------------------

访问名为 "my_range" 的区域::

    my_range = wb.defined_names['my_range']
    # if this contains a range of cells then the destinations attribute is not None
    dests = my_range.destinations # returns a generator of (worksheet title, cell range) tuples

    cells = []
    for title, coord in dests:
        ws = wb[title]
        cells.append(ws[coord])


创建新的命名区域
-------------------------

.. testcode::

    import openpyxl
    wb = openpyxl.Workbook()
    new_range = openpyxl.workbook.defined_name.DefinedName('newrange', attr_text='Sheet!$A$1:$A$5')
    wb.defined_names.append(new_range)

    # create a local named range (only valid for a specific sheet)
    sheetid = wb.sheetnames.index('Sheet')
    private_range = openpyxl.workbook.defined_name.DefinedName('privaterange', attr_text='Sheet!$A$6', localSheetId=sheetid)
    wb.defined_names.append(private_range)
    # this local range can't be retrieved from the global defined names
    assert('privaterange' not in wb.defined_names)

    # the scope has to be supplied to retrieve local ranges:
    print(wb.defined_names.localnames(sheetid))
    print(wb.defined_names.get('privaterange', sheetid).attr_text)

.. testoutput::

    ['privaterange']
    Sheet!$A$6
