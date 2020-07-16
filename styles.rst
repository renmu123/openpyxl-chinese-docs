Working with styles
===================

介绍
------------

Styles are used to change the look of your data while displayed on screen.
They are also used to determine the formatting for numbers.

样式可以应用于以下方面:

   * Font：设置字体大小、颜色、下划线等等
   * PatternFill：设置图案或者颜色渐变
   * Border：设置单元格的边框
   * Alignment：单元格对齐
   * Protection：保护工作表

以下是默认值

.. :: doctest

>>> from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
>>> font = Font(name='Calibri',
...                 size=11,
...                 bold=False,
...                 italic=False,
...                 vertAlign=None,
...                 underline='none',
...                 strike=False,
...                 color='FF000000')
>>> fill = PatternFill(fill_type=None,
...                 start_color='FFFFFFFF',
...                 end_color='FF000000')
>>> border = Border(left=Side(border_style=None,
...                           color='FF000000'),
...                 right=Side(border_style=None,
...                            color='FF000000'),
...                 top=Side(border_style=None,
...                          color='FF000000'),
...                 bottom=Side(border_style=None,
...                             color='FF000000'),
...                 diagonal=Side(border_style=None,
...                               color='FF000000'),
...                 diagonal_direction=0,
...                 outline=Side(border_style=None,
...                              color='FF000000'),
...                 vertical=Side(border_style=None,
...                               color='FF000000'),
...                 horizontal=Side(border_style=None,
...                                color='FF000000')
...                )
>>> alignment=Alignment(horizontal='general',
...                     vertical='bottom',
...                     text_rotation=0,
...                     wrap_text=False,
...                     shrink_to_fit=False,
...                     indent=0)
>>> number_format = 'General'
>>> protection = Protection(locked=True,
...                         hidden=False)
>>>

单元格样式和命名样式
----------------------------

有两种不同的样式: 单元格样式和命名样式, 也被成为样式模板。

单元格样式
+++++++++++

单元格样式在对象之间共享，一旦被分配之后就无法更改。这样可以避免不必要的副作用，例如仅更改一个单元格时就更改许多单元格的样式。

.. :: doctest

>>> from openpyxl.styles import colors
>>> from openpyxl.styles import Font, Color
>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> a1 = ws['A1']
>>> d4 = ws['D4']
>>> ft = Font(color="FF0000")
>>> a1.font = ft
>>> d4.font = ft
>>>
>>> a1.font.italic = True # is not allowed # doctest: +SKIP
>>>
>>> # If you want to change the color of a Font, you need to reassign it
>>>
>>> a1.font = Font(color="FF0000", italic=True) # the change only affects A1


复制样式
--------------

样式也可以被复制

.. :: doctest

>>> from openpyxl.styles import Font
>>> from copy import copy
>>>
>>> ft1 = Font(name='Arial', size=14)
>>> ft2 = copy(ft1)
>>> ft2.name = "Tahoma"
>>> ft1.name
'Arial'
>>> ft2.name
'Tahoma'
>>> ft2.size # copied from the
14.0


颜色
-------

可以通过三种方式：indexed, aRGB or theme 来设置字体、背景、边框等的颜色。
索引颜色（indexed colours）是旧版实现，颜色本身取决于工作簿或应用程序默认提供的索引。主题颜色可用于互补色，但也取决于工作簿中存在的主题。因此，建议使用aRGB颜色。

.. :: doctest

aRGB 颜色
++++++++++++

使用红色，绿色和蓝色的十六进制值设置 RGB 颜色。

>>> from openpyxl.styles import Font
>>> font = Font(color="FF0000")

理论上，alpha值是指颜色的透明度，但这与单元格样式无关。默认值00将前置任何简单的RGB值：

>>> from openpyxl.styles import Font
>>> font = Font(color="00FF00")
>>> font.color.rgb
'0000FF00'

还支持传统索引颜色以及主题和色彩（ themes and tints）。

>>> from openpyxl.styles.colors import Color
>>> c = Color(indexed=32)
>>> c = Color(theme=6, tint=0.5)

Indexed Colours
+++++++++++++++

.. raw:: html
   :file: colours.html

索引64和65不能设置，并且分别保留给系统前景色和背景色。

应用样式
---------------
样式被直接应用到单元格

.. :: doctest

>>> from openpyxl.workbook import Workbook
>>> from openpyxl.styles import Font, Fill
>>> wb = Workbook()
>>> ws = wb.active
>>> c = ws['A1']
>>> c.font = Font(size=12)

样式也可以应用于列和行，但是请注意，这仅适用于关闭文件后创建的单元格（在Excel）。如果要对整个行和列应用样式，则必须自己将样式应用于每个单元格。这是文件格式的限制::
Styles can also applied to columns and rows but note that this applies only
to cells created (in Excel) after the file is closed. If you want to apply
styles to entire rows and columns then you must apply the style to each cell
yourself. This is a restriction of the file format::

>>> col = ws.column_dimensions['A']
>>> col.font = Font(bold=True)
>>> row = ws.row_dimensions[1]
>>> row.font = Font(underline="single")

.. _styling-merged-cells:

合并单元格的样式
--------------------

合并单元格和其他单元格对象的行为相似，通过左上单元格来定义值和样式。可以改变左上单元格的边框来改变整个合并单元格的边框。
这种格式是出于编辑目的才被生成（The formatting is generated for the purpose of writing.）

.. :: doctest

>>> from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
>>> from openpyxl import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>> ws.merge_cells('B2:F4')
>>>
>>> top_left_cell = ws['B2']
>>> top_left_cell.value = "My Cell"
>>>
>>> thin = Side(border_style="thin", color="000000")
>>> double = Side(border_style="double", color="ff0000")
>>>
>>> top_left_cell.border = Border(top=double, left=thin, right=thin, bottom=double)
>>> top_left_cell.fill = PatternFill("solid", fgColor="DDDDDD")
>>> top_left_cell.fill = fill = GradientFill(stop=("000000", "FFFFFF"))
>>> top_left_cell.font  = Font(b=True, color="FF0000")
>>> top_left_cell.alignment = Alignment(horizontal="center", vertical="center")
>>>
>>> wb.save("styled.xlsx")


编辑页面设置
-------------------
.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
>>> ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
>>> ws.page_setup.fitToHeight = 0
>>> ws.page_setup.fitToWidth = 1


命名样式
++++++++++++

与单元格样式相反，命名样式是可变的。当您想一次将格式应用于许多不同的单元格时，它们很有意义。注意一旦将命名样式分配给单元格后，对该样式的更改将**不会**影响单元格。

一旦命名样式被注册到工作簿，就可以简单的通过名字来进行引用


创建命名样式
----------------------

.. :: doctest

>>> from openpyxl.styles import NamedStyle, Font, Border, Side
>>> highlight = NamedStyle(name="highlight")
>>> highlight.font = Font(bold=True, size=20)
>>> bd = Side(style='thick', color="000000")
>>> highlight.border = Border(left=bd, top=bd, right=bd, bottom=bd)

创建命名样式后，即可将其注册到工作簿中：

>>> wb.add_named_style(highlight)

命名样式在首次分配给单元格时也会自动注册：

>>> ws['A1'].style = highlight

注册后，仅使用名称分配样式：

>>> ws['D5'].style = 'highlight'


使用内置样式（Ps：以下注释由译者根据office365中文版进行添加）
--------------------

该规范（specification）包括一些可以使用的内置样式。不幸的是，这些样式的名称以其本地化形式存储。openpyxl 仅会识别英文名称，并且只能与此处的文字完全一样。

* 'Normal' # 无样式

数字格式
++++++++++++++

* 'Comma' # 千位分隔，保留两位小数‘Warning Text’
* 'Comma [0]' # 千位分隔，不保留小数
* 'Currency' # 货币，保留两位小数
* 'Currency [0]' # 货币，不保留小数
* 'Percent' # 百分比

Informative
+++++++++++

* 'Calculation' # 计算
* 'Total' # 汇总
* 'Note' # 注释
* 'Warning Text' # 警告文本
* 'Explanatory Text' # 解释性文本

文字样式
+++++++++++

* 'Title' # 标题
* 'Headline 1' # 标题1
* 'Headline 2' # 标题2
* 'Headline 3' # 标题3
* 'Headline 4' # 标题4
* 'Hyperlink' # 超链接
* 'Followed Hyperlink' # 已访问的超链接
* 'Linked Cell' # 链接单元格

Comparisons
+++++++++++

* 'Input' # 输入
* 'Output' # 输出
* 'Check Cell' # 检查单元格
* 'Good' # 好
* 'Bad' # 坏
* 'Neutral' # 始终

高亮
++++++++++

* 'Accent1' # 着色1
* '20 % - Accent1'
* '40 % - Accent1'
* '60 % - Accent1'
* 'Accent2'  # 着色2
* '20 % - Accent2'
* '40 % - Accent2'
* '60 % - Accent2'
* 'Accent3' # 着色3
* '20 % - Accent3'
* '40 % - Accent3'
* '60 % - Accent3'
* 'Accent4' # 着色4
* '20 % - Accent4'
* '40 % - Accent4'
* '60 % - Accent4'
* 'Accent5' # 着色5
* '20 % - Accent5'
* '40 % - Accent5'
* '60 % - Accent5'
* 'Accent6' # 着色6
* '20 % - Accent6'
* '40 % - Accent6'
* '60 % - Accent6'
* 'Pandas' # 好像是自定义的

有关内置样式的更多信息，请参阅 :mod:`openpyxl.styles.builtins`
