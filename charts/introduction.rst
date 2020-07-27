图表
======


图标类型
-----------

以下图表是可用的:

.. toctree::

    area
    bar
    bubble
    line
    scatter
    pie
    doughnut
    radar
    stock
    surface


创建图表
----------------

图表由至少一个系列的一个或多个数据点组成。系列由单元格范围的引用组成。

.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> for i in range(10):
...     ws.append([i])
>>>
>>> from openpyxl.chart import BarChart, Reference, Series
>>> values = Reference(ws, min_col=1, min_row=1, max_col=1, max_row=10)
>>> chart = BarChart()
>>> chart.add_data(values)
>>> ws.add_chart(chart, "E15")
>>> wb.save("SampleChart.xlsx")


默认情况下，图表的左上角固定在单元格 E15 上，大小为 15 x 7.5厘米（大约5列乘14行）。可以通过设置图表的 `anchor`，`width` 和 `height` 属性来更改。
实际大小将取决于操作系统和设备。其他锚点（anchors ）也是有可能的。更多资料请参考 `openpyxl.drawing.spreadsheet_drawing`。


使用轴
-----------------

.. toctree::

    limits_and_scaling
    secondary


更改图表布局
-----------------------

.. toctree::

    chart_layout


图表样式
--------------

.. toctree::

    pattern


高级图表
---------------

图表能合并生成新的图表：

.. toctree::

    gauge


使用 chartsheets
-----------------

图表能被加入到一个称为 chartsheets 特殊工作簿中：

.. toctree::

    chartsheet
