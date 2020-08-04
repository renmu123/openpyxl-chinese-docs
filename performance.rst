性能
===========

openpyxl 尝试来平衡功能与性能。如果有疑问，我们把重点放在功能而非性能上：一旦建立了 API ，性能调整将变得更简单。
与其他库和应用程序相比，内存使用率很高，约为原始文件大小的50倍，例如 50 MB 的 Excel 文件为内存使用约为 2.5 GB。
由于许多用例只涉及读取或写入文件，the :doc:`optimized` modes mean this is less of a problem.


基准测试
----------

所有基准都是综合性的，并且高度依赖于硬件，但是它们仍然可以提供说明（indication）。


写入性能
+++++++++++++++++

`benchmark code
<https://bitbucket.org/snippets/charlie_x/5edaGE/write-performance-benchmark>`_
可以调整使用更多的工作表以及数据中字符串的比例。由于不同版本的 Python 也会对性能有着显著影响，所以使用了 `driver script
<https://bitbucket.org/snippets/charlie_x/gebj7M/drive-a-script-through-different-python>`_ 对 tox 环境下不同的版本 Python 进行测试。

性能与出色的替代库 xlsxwriter 进行了比较

.. literalinclude:: write_performance.txt


读取性能
++++++++++++++++

读取性能测试使用了 `bug report <https://bitbucket.org/openpyxl/openpyxl/issues/494/>`_ 提供的文件，和早期的 xlrd 库进行比较。
xlrd 主要用于 .XLS 文件较旧的 BIFF 文件格式，它对 XLSX 文件支持有限。

`基准测试 <https://bitbucket.org/snippets/openpyxl/Ee9zqo>`_ 代码显示了处理文件时正确选项的重要性。
在这种情况下，禁用外部链接将让 openpyxl 停止打开链接工作表的缓存副本。

两个库的一个主要区别是 openpyxl 的只读模式可以快速打开工作簿，使其适用于多进程，这也大大减少了内存的使用。
xlrd 也不会自动将日期和时间转换为 Python 的 datetime，尽管它会相应地注释单元格（annotate cells），但是在客户端代码中这样做会大大降低性能。


.. literalinclude:: read_performance.txt


并行
+++++++++++++++

读取工作表会占用大量 CPU 从而限制了从并行中获取好处。但是，如果你主要对 dump 工作表内容感兴趣，你可以使用 openpyxl 的只读模式打开复数工作表来利用多核 CPU。

`Sample code <https://bitbucket.org/snippets/openpyxl/AexG8E>`_ using the
same source file as for read performance shows that performance scales
reasonably with only a slight overhead due to creating additional Python
processes.

.. code-block::

    Parallised Read
        Workbook loaded 1.12s
        >>DATA>> 2.27s
        Output Model 2.30s
        Store days 100% 37.18s
        OptimizationData 44.09s
        Store days 0% 45.60s
        Total time 46.76s
