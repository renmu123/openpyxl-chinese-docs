条形图和柱状图
=====================

在条形图中，值被绘制为水平条或垂直列。（In bar charts values are plotted as either horizontal bars or vertical columns.）

垂直水平和堆叠条形图
-------------------------------------------

.. note::

   以下设置会影响不同的图表类型。

   通过分别将 `type` 设置为 `col` 或 `bar`，可以柱状和水平条形图之间切换。

   使用堆叠图表时，需要将 `overlap` 属性设置为100。

   如果条是水平的，则 x 和 y 轴将反转。


.. image:: bar.png
   :alt: "Sample bar charts"


.. literalinclude:: bar.py

以上创建了四个图表，展示了各种可能性。


三维条形图
-------------

你也能创建三维条形图

.. literalinclude:: bar3d.py


这样能创建一个简单的三维条形图

.. image:: bar3D.png
   :alt: "Sample 3D bar chart"
