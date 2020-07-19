openpyxl - 读写Excel 2010 xlsx/xlsm 文件的Python库
===========================================================================


:Author: Eric Gazoni, Charlie Clark
:Source code: http://bitbucket.org/openpyxl/openpyxl/src
:Issues: http://bitbucket.org/openpyxl/openpyxl/issues
:Generated: |today|
:License: MIT/Expat
:Version: |release|


.. include:: ../README.rst


支持
-------

这是一个由志愿者在业余时间维护的开源项目。这很可能意味着会缺少你想要的特定的功能。
But things don't have to stay that way.
你可以对这个项目进行贡献 :doc:`development` 或者和开发者联系来开发特定的功能。

可以向`Clark Consulting & Research <http://www.clark-consulting.eu/>`_ 和
`Adimian <http://www.adimian.com>`_ 寻求专业支持。欢迎为该项目捐款以支持进一步的开发和维护。


错误报告和功能请求可以使用 `issue tracker <https://bitbucket.org/openpyxl/openpyxl/issues>`_ 来提交。
请提供错误的完整最终，并尽可能提交示例文件。如果出于保密原因您无法公开提供文件，请与开发人员联系。


如何贡献
-----------------

只要遵从了以下步骤，我们欢迎任何帮助:

    1.
    为了每一个独立的功能开了新的fork (https://bitbucket.org/openpyxl/openpyxl/fork) ,也不要想着同时解决所有的问题，这也能使为review和merge你的changes的人更加方便 ;-)

    2.
    Hack hack hack

    3.
    不要忘了为你的修改添加单元测试！（是的，即使只有一行代码，没有单元测试也是不会被接受的哦。）如果不知道怎么做，可以参考源代码中大量的例子

    4.
    如果添加了一个完整的功能或者对某个功能做出了改进，你可以自豪地把自己加入作者文件中:-)

    5.
    为了让大家知道你刚提交的功能是多么的棒，务必更新一下文档！

    6.
    当以上步骤都完成之后，提一个 pull request（在**你**的 repository 页点击大大的 "pull request"按钮）然后等你的代码被review。如果以上步骤都完成了，那么就会合并到主 repository 。


更多信息请查询 :doc:`development`


其他提供帮助的方式
++++++++++++++++++

即使你不会写代码（或者代码写得不是很好），也有多种方式来作出贡献

    * 为 bug 追踪器（bug tracker）进行分流: 关闭已经解决的，无关的，不能复现的bug

    * 对几乎每个方面的文档进行更新: 增加了大量大型的特性（主要是图表和图像）但是没有文档，因此很难用新特性来做点什么

    * proposing compatibility fixes for different versions of Python: 我们支持
      2.7, 3.4, 3.5, 3.6 和 3.7


安装
------------

使用 pip 安装 openpyxl。建议在不带系统软件包的 Python virtualenv 中执行此操作::

    $ pip install openpyxl

.. note::

    支持流行的 `lxml`_ 库， 在创建大量文件的时候特别有用。

.. _lxml: http://lxml.de

.. warning::

    为了在 openpyxl 文件中包含（jpeg, png, bmp,...）等图片，你还需要安装 `pillow`::

    $ pip install pillow

    或者你也可以浏览 https://pypi.python.org/pypi/Pillow/, 选择最新版本或下拉到页面最后选择 Windows 二进制版


Working with a checkout
+++++++++++++++++++++++

Sometimes you might want to work with the checkout of a particular version.
This may be the case if bugs have been fixed but a release has not yet been
made.

.. parsed-literal::
    $ pip install -e hg+https://bitbucket.org/openpyxl/openpyxl@\ |version|\ #egg=openpyxl


Usage examples
--------------


教程
++++++++

.. toctree::

    tutorial


Cookbook
++++++++

.. toctree::

    usage


性能
-----------

.. toctree::

    performance


其他主题
------------

    .. toctree::
        :maxdepth: 2

        optimized


    .. toctree::
        :maxdepth: 1

        editing_worksheets

    .. toctree::
        :maxdepth: 1

        pandas

    .. toctree::
        :maxdepth: 1

        charts/introduction

    .. toctree::
        :maxdepth: 1

        comments

    .. toctree::
        :maxdepth: 1

        styles

    .. toctree::
        :maxdepth: 1

        worksheet_properties

    .. toctree::
        :maxdepth: 1

        formatting

    .. toctree::
        :maxdepth: 1

        pivot

    .. toctree::
        :maxdepth: 1

        print_settings

    .. toctree::
        :maxdepth: 1

        filters

    .. toctree::
        :maxdepth: 1

        validation


    .. toctree::
        :maxdepth: 1

        defined_names

    .. toctree::
        :maxdepth: 1

        worksheet_tables

    .. toctree::
        :maxdepth: 1

        formula

    .. toctree::
        :maxdepth: 1

        protection


开发者信息
--------------------------

    .. toctree::
        :maxdepth: 1

        development


API 文档
------------------

关键类
+++++++++++

* :class:`openpyxl.workbook.workbook.Workbook`
* :class:`openpyxl.worksheet.worksheet.Worksheet`
* :class:`openpyxl.cell.cell.Cell`


完整 API
++++++++

.. toctree::
    :maxdepth: 2

    api/openpyxl


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`


发布说明
=============

.. toctree::
    :maxdepth: 1

    changes
