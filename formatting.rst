条件格式
======================

Excel 支持三种类型的条件格式：内置、标准和自定义
内建条件格式将特定规则与预定义样式结合在一起。
标准条件格式将特定规则与自定义格式结合在一起。
In additional it is possible to define custom formulae for applying custom formats using differential styles.

.. note::

  不同规则的语法差异很大，以至于 openpyxl 不知道规则是否有意义。


创建条件格式规则的基本语法为：

.. doctest

>>> from openpyxl.formatting import Rule
>>> from openpyxl.styles import Font, PatternFill, Border
>>> from openpyxl.styles.differential import DifferentialStyle
>>> dxf = DifferentialStyle(font=Font(bold=True), fill=PatternFill(start_color='EE1111', end_color='EE1111'))
>>> rule = Rule(type='cellIs', dxf=dxf, formula=["10"])

由于某些规则的签名可能非常冗长，因此也有一些方便的工厂（factories）来创建它们。

内置格式
---------------

内置格式有：

  * 色阶（ColorScale）
  * 图表集（IconSet）
  * 数据条（DataBar）

Builtin formats contain a sequence of formatting settings which combine a type with an integer for comparison.
可能的类型有：`'num', 'percent', 'max', 'min', 'formula', 'percentile'`。


色阶
++++++++++

你可以使用 2 或 3 种颜色的色阶。2 种色阶产生一种颜色到另一种颜色的渐变；3 种颜色色阶会将 1 种颜色用于 2 个颜色的渐变。

创建色阶的完整规则为:

.. doctest

>>> from openpyxl.formatting.rule import ColorScale, FormatObject
>>> from openpyxl.styles import Color
>>> first = FormatObject(type='min')
>>> last = FormatObject(type='max')
>>> # colors match the format objects:
>>> colors = [Color('AA0000'), Color('00AA00')]
>>> cs2 = ColorScale(cfvo=[first, last], color=colors)
>>> # a three color scale would extend the sequences
>>> mid = FormatObject(type='num', val=40)
>>> colors.insert(1, Color('00AA00'))
>>> cs3 = ColorScale(cfvo=[first, mid, last], color=colors)
>>> # create a rule with the color scale
>>> from openpyxl.formatting.rule import Rule
>>> rule = Rule(type='colorScale', colorScale=cs3)

有一个方便创建色阶规则的函数：

.. doctest

>>> from openpyxl.formatting.rule import ColorScaleRule
>>> rule = ColorScaleRule(start_type='percentile', start_value=10, start_color='FFAA0000',
...                       mid_type='percentile', mid_value=50, mid_color='FF0000AA',
...                       end_type='percentile', end_value=90, end_color='FF00AA00')


图标集
+++++++

从以下图标中进行选择: `'3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters'`

创建图表集完整规则为：

.. doctest

>>> from openpyxl.formatting.rule import IconSet, FormatObject
>>> first = FormatObject(type='percent', val=0)
>>> second = FormatObject(type='percent', val=33)
>>> third = FormatObject(type='percent', val=67)
>>> iconset = IconSet(iconSet='3TrafficLights1', cfvo=[first, second, third], showValue=None, percent=None, reverse=None)
>>> # assign the icon set to a rule
>>> from openpyxl.formatting.rule import Rule
>>> rule = Rule(type='iconSet', iconSet=iconset)

有一个方便创建色阶图表集规则的函数：

.. doctest

>>> from openpyxl.formatting.rule import IconSetRule
>>> rule = IconSetRule('5Arrows', 'percent', [10, 20, 30, 40, 50], showValue=None, percent=None, reverse=None)


数据条
+++++++

目前，openpyxl 支持原始规范中定义的数据条。之后的扩展中添加了边框和方向。

完整创建数据条的规则为：

.. doctest

>>> from openpyxl.formatting.rule import DataBar, FormatObject
>>> first = FormatObject(type='min')
>>> second = FormatObject(type='max')
>>> data_bar = DataBar(cfvo=[first, second], color="638EC6", showValue=None, minLength=None, maxLength=None)
>>> # assign the data bar to a rule
>>> from openpyxl.formatting.rule import Rule
>>> rule = Rule(type='dataBar', dataBar=data_bar)

有一个方便创建数据条规则的函数：

.. doctest

>>> from openpyxl.formatting.rule import DataBarRule
>>> rule = DataBarRule(start_type='percentile', start_value=10, end_type='percentile', end_value='90',
...                    color="FF638EC6", showValue="None", minLength=None, maxLength=None)


标准条件格式
----------------------------

标准条件格式为：

  * 平均值（Average）
  * 百分比（Percent）
  * 唯一值或重复值（Unique or duplicate）
  * 值（Value）
  * 排名（Rank）

.. doctest

>>> from openpyxl import Workbook
>>> from openpyxl.styles import Color, PatternFill, Font, Border
>>> from openpyxl.styles.differential import DifferentialStyle
>>> from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> # Create fill
>>> redFill = PatternFill(start_color='EE1111',
...                end_color='EE1111',
...                fill_type='solid')
>>>
>>> # Add a two-color scale
>>> # Takes colors in excel 'RRGGBB' style.
>>> ws.conditional_formatting.add('A1:A10',
...             ColorScaleRule(start_type='min', start_color='AA0000',
...                           end_type='max', end_color='00AA00')
...                           )
>>>
>>> # Add a three-color scale
>>> ws.conditional_formatting.add('B1:B10',
...                ColorScaleRule(start_type='percentile', start_value=10, start_color='AA0000',
...                            mid_type='percentile', mid_value=50, mid_color='0000AA',
...                            end_type='percentile', end_value=90, end_color='00AA00')
...                              )
>>>
>>> # Add a conditional formatting based on a cell comparison
>>> # addCellIs(range_string, operator, formula, stopIfTrue, wb, font, border, fill)
>>> # Format if cell is less than 'formula'
>>> ws.conditional_formatting.add('C2:C10',
...             CellIsRule(operator='lessThan', formula=['C$1'], stopIfTrue=True, fill=redFill))
>>>
>>> # Format if cell is between 'formula'
>>> ws.conditional_formatting.add('D2:D10',
...             CellIsRule(operator='between', formula=['1','5'], stopIfTrue=True, fill=redFill))
>>>
>>> # Format using a formula
>>> ws.conditional_formatting.add('E1:E10',
...             FormulaRule(formula=['ISBLANK(E1)'], stopIfTrue=True, fill=redFill))
>>>
>>> # Aside from the 2-color and 3-color scales, format rules take fonts, borders and fills for styling:
>>> myFont = Font()
>>> myBorder = Border()
>>> ws.conditional_formatting.add('E1:E10',
...             FormulaRule(formula=['E1=0'], font=myFont, border=myBorder, fill=redFill))
>>>
>>> # Highlight cells that contain particular text by using a special formula
>>> red_text = Font(color="9C0006")
>>> red_fill = PatternFill(bgColor="FFC7CE")
>>> dxf = DifferentialStyle(font=red_text, fill=red_fill)
>>> rule = Rule(type="containsText", operator="containsText", text="highlight", dxf=dxf)
>>> rule.formula = ['NOT(ISERROR(SEARCH("highlight",A1)))']
>>> ws.conditional_formatting.add('A1:F40', rule)
>>> wb.save("test.xlsx")


条件格式应用在全部行
----------------------

有时你想将条件格式应用于多个单元格，例如一行包含特定值的一些单元格。

>>> ws.append(['Software', 'Developer', 'Version'])
>>> ws.append(['Excel', 'Microsoft', '2016'])
>>> ws.append(['openpyxl', 'Open source', '2.6'])
>>> ws.append(['OpenOffice', 'Apache', '4.1.4'])
>>> ws.append(['Word', 'Microsoft', '2010'])

我们要突出开发人员是 Microsoft 的行。我们通过创建表达式规则并使用公式来识别哪些行包含了 Microsoft 开发的 Software。

>>> red_fill = PatternFill(bgColor="FFC7CE")
>>> dxf = DifferentialStyle(fill=red_fill)
>>> r = Rule(type="expression", dxf=dxf, stopIfTrue=True)
>>> r.formula = ['$A2="Microsoft"']
>>> ws.conditional_formatting.add("A1:C10", r)

.. note::
    在这种情况下，该公式使用**绝对引用** B 列，以及**相对引用**行号，在这种情况下, ``1`` 是行号相对于应用格式的范围。
    做到这一点可能很棘手，但是即使已将规则添加到工作表的条件格式集合中，也可以对其进行调整。
