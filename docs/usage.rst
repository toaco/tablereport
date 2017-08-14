使用说明
=======

基本结构
-------
TableReport的4个基本部分包括：Cell，Table，Style，Writer

1. Cell 是一个单元格，包含了值，同时也可以包含样式（Style）。
2. Table 是表示一张表格的类，每个元素都是一个Cell。
3. Writer 用于将Table写入到Excel。
4. Style 表示单元格的样式，支持的设置包括：字体（颜色，大小，类型），对齐方式（水平对齐方式，垂直对齐方式），背景色

一个最简单的例子::

   def test_excel_writer():
       table = Table(header=[['header1', 'header2', 'header3']],
                     body=[[1, 2, 3], [1, 2, 4], [1, 3, 5]])

       write_to_excel('1.xlsx', table)

传入的Header和body会自动被包装为Cell对象，你可以通过 ``[]`` 访问Cell，比如::

   def test_each_elem_in_table_is_encapsulated_as_cell():
       table = Table(body=[[1, 2, ], [4, 5, ]])

       assert table.data[0][0] == Cell(1)
       assert table.data[0][1] == Cell(2)
       assert table.data[1][0] == Cell(4)
       assert table.data[1][1] == Cell(5)

同时，Cell重载了 ``__eq__`` 方法，你可以直接比较Cell和他的值。比如::

    assert table.data[0][0] == 1
    assert table.data[0] == [1,2]

如果你需要使用openpyxl对表格进行更多的控制，那么你可以使用WorkSheetWriter类带来导出表格。例子如下: ::

    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5]])

    wb = Workbook()
    ws = wb.active
    ws.title = 'report'

    WorkSheetWriter.write(ws, table, (1, 1))

    # 使用openpyxl进行其他处理
    # ...

    wb.save('1.xlsx')

除了这几个基本部分之外，TableReport提供了更多的类型帮助我们制作报表。

区域
----
区域相关的类型有3个：Area，Areas，Row

Area 表示Table的一个矩形区域，它不会存放数据和样式，但需要和Table关联。类似于SQL中的视图。对于一个Area我们可以进行：合并，合计，设置样式等操作，这些操作都会在Table上面体现出来::

    table = Table(body=[[1, 2, 3, ], [4, 5, 6], [7, 8, 9]])
    area = Area(table, 3, 3, (0, 0))

    assert area[0][0] = Cell(1)

我们可以通过Table的select方法选择区域，另外Table的header和body属性分别表示的就是Table的头部区域和内容区域。 ::

   def test_column_selector_select_right_area_of_table():
       table = Table(header=[['header1', 'header2', 'header3']],
                     body=[[1, 2, 3], [4, 5, 6], [7, 8, 9]])
       sub_area = table.body.select(ColumnSelector(column=2, group=False)).one()

       assert sub_area.height == 3
       assert sub_area.width == 1
       assert sub_area.position == (1, 1)

Areas 是一个Area的集合，拥有和Area相同的API，对Areas的操作会转发到所有内部的Area上去。

Row 表示的是区域中的一行，可以在Area上进行 ``[]`` 操作获得，作用类似于Areas，会转发Row的操作到内部的Cell中去。

选择器
-----

选择器Selector是用于从Table中选择Area或者从Area中选择子Area的类型.可以作为Table上select方法的参数。

目前提供了ColumnSelector，可以根据列号选择一个区域，ColumnSelector提供了一个group参数，如果该参数设置为True，将会根据该列中单元格的值进行分组，得到的结果将是一个Areas，其中每个Area的只都是相等的。

合并单元格
---------

可以对Area进行合并操作，合并后的值为该区域中第一个单元格的值。::

   def test_merge_areas_1():
       table = Table(header=[['header1', 'header2', 'header3']],
                     body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4]])
       areas = table.body.select(ColumnSelector(column=1, group=True))
       areas.merge()

       assert table.data == [['header1', 'header2', 'header3'],
                             [1, 2, 3], [None, 2, 4], [None, 3, 5], [2, 3, 4]]

``areas.merge()`` 将会合并areas包含的所有区域，合并后的区域只保留单元格。其余部分都会设置为None。

合计
----

对于ColumnSelector选择出来的区域，我们可以进行合计。目前合计将会对该区域右边的所有列进行求和运算。新的列可以放在该区域的右边，也可以放在该区域的下边。::

   def test_add_nested_summary_will_modify_table():
       table = Table(header=[['header1', 'header2', 'header3']],
                     body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])

       areas = table.body.select(ColumnSelector(column=1, group=True))
       areas.merge()
       areas.add_summary(text_span=1, text='total', location='left')

       area = Area(table, 3, 7, (1, 0))
       area.add_summary(text_span=2, text='total', location='bottom')

       assert table.data == [['header1', 'header2', 'header3'],
                             [1, 2, 3], [None, 2, 4], [None, 3, 5],
                             [None, 'total', 12],
                             [2, 3, 4], [None, 4, 5], [None, 'total', 9],
                             ['total', None, 21]]

       assert table.width == 3
       assert table.height == 9
