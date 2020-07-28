其他工作表属性
===============================

有一些特定行为的高级属性，最常用的是页面设置参数（page setup property） `fitTopage` 和 定义工作表选项卡颜色的`tabColor`。

工作表可用属性
-----------------------------------

* "enableFormatConditionsCalculation"
* "filterMode"
* "published"
* "syncHorizontal"
* "syncRef"
* "syncVertical"
* "transitionEvaluation"
* "transitionEntry"
* "tabColor"

页面设置属性的可用字段
------------------------------------------

"autoPageBreaks"
"fitToPage"

outlines的可用字段
-----------------------------

* "applyStyles"
* "summaryBelow"
* "summaryRight"
* "showOutlineSymbols"

更多信息请查询 http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sheetproperties%28v=office.14%29.aspx_

.. note::
        默认情况下，会对 `outline` 属性进行初始化，因此您可以直接修改它们的 4 个属性，而页面设置属性不一样。
        如果要修改后者，首先要必要的参数初始化对 `openpyxl.worksheet.properties.PageSetupProperties` 对象进行初始化。
        一旦完成，可以在以后需要时通过例程直接对其进行修改。


.. :: doctest

>>> from openpyxl.workbook import Workbook
>>> from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> wsprops = ws.sheet_properties
>>> wsprops.tabColor = "1072BA"
>>> wsprops.filterMode = False
>>> wsprops.pageSetUpPr = PageSetupProperties(fitToPage=True, autoPageBreaks=False)
>>> wsprops.outlinePr.summaryBelow = False
>>> wsprops.outlinePr.applyStyles = True
>>> wsprops.pageSetUpPr.autoPageBreaks = True
