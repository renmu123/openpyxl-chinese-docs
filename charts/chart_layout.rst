更改绘图区和图例的布局
===========================================


可以通过使用 layout 类实例的 layout 属性来设置图表的布局。


表格布局
------------

位置和大小
+++++++++++++++++

图表可以放置在容器中。可以通过 ``x`` 和 ``y`` 调整位置。``w`` 和 ``h`` 调整大小。单位是容器的比例。
图表不能放置在容器的外部，并且宽度和高度是主要限制：如果 x+w>1，则 x=1-w。


| x是从左侧开始的水平位置
| y是从顶部开始的垂直位置
| h是图表相对于其容器的高度
| w是盒子（box）的宽度


模式
++++

除了大小和位置外，相关属性的模式也可以设置为 `factor` 或 `edge`。默认值为 `factor`:

.. code::

  layout.xMode = edge


目标（Target）
++++++

~layoutTarget` 属性可以设置成 ``outer`` 或者 ``inner``. 默认值为 ``outer``:

.. code::

  layout.layoutTarget = inner


图例布局
-------------

图例的位置可以通过设置位置参数来进行改变：
``r``、 ``l``、 ``t`、, ``b`` 和 ``tr``,分别代表右、左、上、下以及右上。默认值为 ``r``.

.. code::

  legend.position = 'tr'

或者应用手动布局：

.. code::

  legend.layout = ManualLayout()


.. literalinclude:: chart_layout.py


以上会创建四个图表并展示了各种可能性：

.. image:: chart_layout.png
   :alt: "Different chart and legend layouts"
