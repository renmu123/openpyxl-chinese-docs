注释
========

.. warning::

    openpyxl 目前只支持读写文字注释。格式信息会丢失。在读取时，注释尺寸也会丢失，但是可以重新写入。注释目前不支持 `read_only=True` 模式下使用。


为单元格添加注释
--------------------------

注释的 text 和 author 是必填属性

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> wb = Workbook()
>>> ws = wb.active
>>> comment = ws["A1"].comment
>>> comment = Comment('This is the comment text', 'Comment Author')
>>> comment.text
'This is the comment text'
>>> comment.author
'Comment Author'

如果你为不同的单元格设置了相同的注释，那么 openpyxl 会自动进行复制

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> wb=Workbook()
>>> ws=wb.active
>>> comment = Comment("Text", "Author")
>>> ws["A1"].comment = comment
>>> ws["B2"].comment = comment
>>> ws["A1"].comment is comment
True
>>> ws["B2"].comment is comment
False


加载和保存注释
----------------------------

加载时工作簿中存在的注释会自动存储在其相应单元格的 comment 属性中。格式信息（如字体大小，粗体和斜体）以及注释的容器框的原始尺寸和位置都将丢失。

保存工作簿时保留在工作簿中的注释会自动保存到工作簿文件中。

注释尺寸可以设定成只写。注释尺寸以像素为单位。

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> from openpyxl.utils import units
>>>
>>> wb=Workbook()
>>> ws=wb.active
>>>
>>> comment = Comment("Text", "Author")
>>> comment.width = 300
>>> comment.height = 50
>>>
>>> ws["A1"].comment = comment
>>>
>>> wb.save('commented_book.xlsx')


如果有需要的话, ``openpyxl.utils.units`` 有将其他度量单位（mm，points）转换为像素的辅助函数:

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> from openpyxl.utils import units
>>>
>>> wb=Workbook()
>>> ws=wb.active
>>>
>>> comment = Comment("Text", "Author")
>>> comment.width = units.points_to_pixels(300)
>>> comment.height = units.points_to_pixels(50)
>>>
>>> ws["A1"].comment = comment
