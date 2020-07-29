保护
==========

.. warning::

    工作簿或工作表的密码保护仅提供了十分基础的安全。数据未进行加密，所以可以使用各种免费工具进行修改。
    实际上，规范指出：工作表或工作簿的保护不应该与文件安全性混淆。这是为了保护你的工作簿免受意外修改的影响，并不能保护免受恶意修改的影响。

Openpyxl 支持保护工作簿和工作表不被修改。除非指定明确算法，否则将使用 Open XML "Legacy Password Hash Algorithm" 来生成哈希密码值。

工作簿保护
-------------------

为防止其他用户查看隐藏的工作表、添加、移动、删除或隐藏工作表以及重命名工作表，可以使用密码保护工作簿的结构。
可以使用 ``openpyxl.workbook.protection.WorkbookProtection.workbookPassword`` 属性设置密码::

    >>> wb.security.workbookPassword = '...'
    >>> wb.security.lockStructure = True


同样，可以通过设置另一个密码来防止从共享工作簿中删除更改跟踪和更改历史记录。
可以使用 ``openpyxl.workbook.protection.WorkbookProtection.revisionsPassword`` 属性设置密码::

    >>> wb.security.revisionsPassword = '...'

 :class:`openpyxl.workbook.protection.WorkbookProtection` 对象上的其他属性可以精确控制所设置的限制（restrictions are in place），但是只有设置密码后，这些属性才能生效。


如果需要设置原始密码值而非使用默认哈希算法，我们也提供特定的设置函数-例如::

    hashed_password = ...
    wb.security.set_workbook_password(hashed_password, already_hashed=True)


工作表保护
--------------------

也可以通过在 :class:`openpyxl.worksheet.protection.SheetProtection` 对象上设置属性来锁定工作表。
与工作簿保护不同，可以使用或不使用密码来启用工作表保护。使用 :class:`openpyxl.worksheet.protection.SheetProtection.sheet` 属性或调用 enable() 或disable() 俩启用工作表保护::

    >>> ws = wb.active
    >>> ws.protection.sheet = True
    >>> ws.protection.enable()
    >>> ws.protection.disable()


如果未设置密码，那么用户不需要密码即可禁用工作表保护。否则，他们必要提供密码才能修改保护配置。
使用 :func:`openpxyl.worksheet.protection.SheetProtection.password` 设置密码::

    >>> ws = wb.active
    >>> ws.protection.password = '...'

