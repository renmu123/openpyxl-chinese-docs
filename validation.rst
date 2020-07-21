数据验证
================

数据验证器可以应用于范围单元格，但也不是强制和evaluated。范围不必是连续的：例如 “ A1 B2：B5”包含A1和单元格B2至B5，但不包含A2或B2。


例子
--------

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.worksheet.datavalidation import DataValidation
>>>
>>> # Create the workbook and worksheet we'll be working with
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> # Create a data-validation object with list validation
>>> dv = DataValidation(type="list", formula1='"Dog,Cat,Bat"', allow_blank=True)
>>>
>>> # Optionally set a custom error message
>>> dv.error ='Your entry is not in the list'
>>> dv.errorTitle = 'Invalid Entry'
>>>
>>> # Optionally set a custom prompt message
>>> dv.prompt = 'Please select from the list'
>>> dv.promptTitle = 'List Selection'
>>>
>>> # Add the data-validation object to the worksheet
>>> ws.add_data_validation(dv)

>>> # Create some cells, and add them to the data-validation object
>>> c1 = ws["A1"]
>>> c1.value = "Dog"
>>> dv.add(c1)
>>> c2 = ws["A2"]
>>> c2.value = "An invalid value"
>>> dv.add(c2)
>>>
>>> # Or, apply the validation to a range of cells
>>> dv.add('B1:B1048576') # This is the same as for the whole of column B
>>>
>>> # Check with a cell is in the validator
>>> "B4" in dv
True


.. note ::

    没有在任何单元格应用的验证将会在保存的时候被忽略。

其他验证的例子
-------------------------

任何证书:
::

    dv = DataValidation(type="whole")

任何大于100的整数:
::

    dv = DataValidation(type="whole",
                        operator="greaterThan",
                        formula1=100)

任何小数:
::

    dv = DataValidation(type="decimal")

任何在0至1之间的小数:
::

    dv = DataValidation(type="decimal",
                        operator="between",
                        formula1=0,
                        formula2=1)

任何日期:
::

    dv = DataValidation(type="date")

时间:
::

    dv = DataValidation(type="time")

15长度以下的文本:
::

    dv = DataValidation(type="textLength",
                        operator="lessThanOrEqual"),
                        formula1=15)

序列:
::

    from openpyxl.utils import quote_sheetname
    dv = DataValidation(type="list",
                        formula1="{0}!$B$1:$B$10".format(quote_sheetname(sheetname))
                        )

自定义规则:
::

    dv = DataValidation(type="custom",
                        formula1"=SOMEFORMULA")

.. note::
    See http://www.contextures.com/xlDataVal07.html for custom rules
