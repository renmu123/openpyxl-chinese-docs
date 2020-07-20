工作簿表格
================


工作表表格是对单元格组的引用。这使得某些操作（例如，对表格中的单元格进行样式设置）更加容易。


创建表格
----------------

.. literalinclude:: table.py

在一个工作簿中表格名称必须是唯一的。默认情况下，表是从第一行的标题开始创建的，并且所有列的筛选以及表标题和列标题必须始终包含字符串。

.. warning::

  在只写模式下，您必须手动将列标题添加到表格中，并且值必须始终与相应单元格的值相同（有关如何执行此操作的示例，请参见下面的例子），否则 Excel 可能会认为该文件无效并删除表格。

通过 `TableStyleInfo` 来管理样式。这允许你对行和列设置条纹以及应用不同的颜色主题。


使用表格
-------------------

``ws.tables`` 是特定工作簿下所有表格的 dictionary-like 对象::

  >>> ws.tables
  {"Table1",  <openpyxl.worksheet.table.Table object>}


通过范围或者名称获取表格
++++++++++++++++++++++++++

.. code::

  >>> ws.tables["Table1"]
  or
  >>> ws.tables["A1:D10"]


遍历工作簿下所有的表格
+++++++++++++++++++++++++++++++++++++++++

.. code::

  >>> for table in ws.tables.values():
  >>>    print(table)


获取表名以及工作簿内所有表格的范围
+++++++++++++++++++++++++++++++++++++++++++++++++++++

返回表格名和范围的列表

.. code::

  >>> ws.tables.items()
  >>> [("Table1", "A1:D10")]


删除表格
++++++++++++++

.. code::

  >>> del ws.tables["Table1"]


工作簿中的表格数量
+++++++++++++++++++++++++++++++++++

.. code::

  >>> len(ws.tables)
  >>> 1


手动添加表格表头
-------------------------------

在只写模式下你可以添加没有表头的表格::

  >>> table.headerRowCount = False

或者手动初始化表头::

  >>> headings = ["Fruit", "2011", "2012", "2013", "2014"] # all values must be strings
  >>> table._initialise_columns()
  >>> for column, value in zip(table.tableColumns, headings):
      column.name = value