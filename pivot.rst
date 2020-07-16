数据透视表
============

openpyxl 为数据透视表提供读取支持以便于可以保留在现有的文件中。
数据透视表的规范虽然很广泛，但不是很清楚，也不意味着客户端代码应该能够创建数据透视表。（The specification for pivot tables, while
extensive, is not very clear and it is not intended that client code should
be able to create pivot tables.）
但是，应该可以编辑和操作现有的数据透视表，例如。更改其范围或是能自动更新设置。

和图表、图片、表格一样，数据透视表没有专门管理的 API，因此客户端代码不得不遍历工作表 ``_pivots`` 列表


例子
-------

.. code::

    from openpyxl import load_workbook
    wb = load_workbook("campaign.xlsx")
    ws = wb["Results"]
    pivot = ws._pivots[0] # any will do as they share the same cache
    pivot.cache.refreshOnLoad = True


更多信息 请查询 :class:`openpyxl.pivot.cache.CacheDefinition`
