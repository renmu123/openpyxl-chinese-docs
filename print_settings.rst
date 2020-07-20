打印设置
==============

openpyxl 为打印设置提供合理的全面支持


编辑打印设置
-------------------
.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.print_options.horizontalCentered = True
>>> ws.print_options.verticalCentered = True


页眉页脚
-------------------

页眉和页脚使用自己的格式语言。在编辑的时候完全可以支持但是由于于复杂和嵌套的可能性，在读取它们时仅部分支持。
支持字体，大小和颜色，居左，居中或居右元素。粒度控制（突出显示单个单词）需要手动应用控制代码（ Granular control
(highlighting individuals words) will require applying control codes
manually）


.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.oddHeader.left.text = "Page &[Page] of &N"
>>> ws.oddHeader.left.size = 14
>>> ws.oddHeader.left.font = "Tahoma,Bold"
>>> ws.oddHeader.left.color = "CC3366"


也支持 `evenHeader` 和 `evenFooter` 以及 `firstHeader` 和 `firstFooter`.


增加打印标题
----------------

您可以在每页上打印标题，以确保正确标记数据。

.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.print_title_cols = 'A:B' # the first two cols
>>> ws.print_title_rows = '1:1' # the first row


增加打印区域
----------------

你可以只选择工作簿的一部分来作为打印区域

.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.print_area = 'A1:F10'
