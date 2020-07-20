筛选和排序
=======================


在工作簿中添加筛选是可能的

.. note::

  筛选和排序只能通过 openpyxl 进行设置，但是只有在 Excel 这样的程序中才会被应用。这是由于他们会在范围内重新排列或者格式化单元格或行。（This is because they actually rearranges or format cells or rows in the range）


定义一个范围后，你可以对一列添加筛选或者添加排序条件:（To add a filter you define a range and then add columns and sort conditions:）

.. literalinclude:: filters.py


这会将相关指令添加到文件中，但实际上不会 **过滤或排序** 。（PS：译者使用上诉代码在 Excel 中试了一下这个功能，其中已经出现了筛选栏控件，但是未生效，点击“确认”即可生效，排序功能点了“确认也没办法使用”）

.. image:: filters.png
   :alt: "Filter and sort prepared but not executed for a range of cells"
