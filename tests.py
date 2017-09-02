#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from openpyxl import Workbook

from tablereport import *
from tablereport.shortcut import write_to_excel


def test_table_initialize():
    table = Table(
        header=[['test', None, None], ['header1', 'header2', 'header3']],
        body=[[1, 2, 3], [4, 5, 6], [7, 8, 9]]
    )

    assert table.width == 3
    assert table.height == 5
    assert list(table.header) == [['test', None, None],
                                  ['header1', 'header2', 'header3']]
    assert list(table.body) == [[1, 2, 3], [4, 5, 6], [7, 8, 9]]


def test_initialize_table_with_empty_headers_and_empty_body():
    table = Table(header=[],
                  body=[])

    assert table.width == 0
    assert table.height == 0


def test_column_selector_select_right_area_of_area():
    table = Table(body=[[1, 2, 3, ], [4, 5, 6], [7, 8, 9]])
    area = Area(table, 3, 3, (0, 0))
    sub_area = area.select(ColumnSelector(lambda col: col == 2)).one()

    assert sub_area.height == 3
    assert sub_area.width == 1
    assert sub_area.position == (0, 1)


def test_column_selector_select_right_area_of_area_2():
    table = Table(body=[[1, 2, 3, ], [4, 5, 6], [7, 8, 9]])
    area = Area(table, 3, 3, (0, 0))
    areas = area.select(ColumnSelector(lambda col: col % 2))

    assert len(areas) == 2
    assert areas[0].position == (0, 0)
    assert areas[1].position == (0, 2)


def test_column_selector_select_right_area_of_table():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [4, 5, 6], [7, 8, 9]])
    sub_area = table.body.select(ColumnSelector(lambda col: col == 2)).one()

    assert sub_area.height == 3
    assert sub_area.width == 1
    assert sub_area.position == (1, 1)


def test_values_in_area_of_table_selected_by_column_selector():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [4, 5, 6], [7, 8, 9]])
    area = table.body.select(ColumnSelector(lambda col: col == 2)).one()

    assert area.data == [[2], [5], [8]]


def test_column_selector_could_group_selected_areas():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4]])
    areas = table.body.select(ColumnSelector(lambda col: col == 1))

    areas.group()

    assert len(areas) == 2

    assert areas[0].width == 1
    assert areas[0].height == 3
    assert areas[0].position == (1, 0)

    assert areas[1].width == 1
    assert areas[1].height == 1
    assert areas[1].position == (4, 0)


def test_modify_area_is_equivalent_to_modify_table():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [4, 5, 6], [7, 8, 9]])
    area = table.body.select(ColumnSelector(lambda col: col == 2)).one()

    # area.data == [[2], [5], [8]]
    area.data[0][0] = 3
    area.data[1][0] = 6
    area.data[2][0] = 9

    assert area.data == [[3], [6], [9]]


def test_areas_is_list_like_object():
    areas = Areas()
    areas.append(1)
    areas.append(2)

    assert len(areas) == 2
    assert areas[0] == 1
    assert areas[1] == 2


def test_get_area_data():
    area = Table(body=[[1, 2, 3, ], [4, 5, 6], [7, 8, 9]])
    sub_area = area.select(ColumnSelector(lambda col: col == 2)).one()

    assert sub_area.data[0][0] == 2
    assert sub_area.data[1][0] == 5
    assert sub_area.data[2][0] == 8

    assert sub_area.data == [[2], [5], [8]]


def test_set_area_data():
    area = Table(body=[[1, 2, 3, ], [4, 5, 6], [7, 8, 9]])
    sub_area = area.select(ColumnSelector(lambda col: col == 2)).one()

    sub_area.data[0][0] = 1
    sub_area.data[1][0] = 2
    sub_area.data[2][0] = 3

    assert sub_area.data[0][0] == 1
    assert sub_area.data[1][0] == 2
    assert sub_area.data[2][0] == 3

    assert sub_area.data == [[1], [2], [3]]


def test_merge_areas_1():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4]])
    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()

    assert table.data == [['header1', 'header2', 'header3'],
                          [1, 2, 3], [None, 2, 4], [None, 3, 5], [2, 3, 4]]


def test_merge_areas_2():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])
    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()

    assert table.data == [['header1', 'header2', 'header3'],
                          [1, 2, 3], [None, 2, 4], [None, 3, 5], [2, 3, 4],
                          [None, 4, 5]]


def test_add_summary_at_left_side_will_modify_table():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])
    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()

    areas.add_summary(text_span=1, text='total', location='left')

    assert table.data == [['header1', 'header2', 'header3'],
                          [1, 2, 3], [None, 2, 4], [None, 3, 5],
                          [None, 'total', 12],
                          [2, 3, 4], [None, 4, 5], [None, 'total', 9]]


def test_add_summary_at_left_side_will_modify_areas():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])
    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()

    areas.add_summary(text_span=2, text='total', location='bottom')

    area1 = areas[0]
    assert area1.width == 1
    assert area1.height == 4
    assert area1.position == (1, 0)

    area2 = areas[1]
    assert area2.width == 1
    assert area2.height == 3
    assert area2.position == (5, 0)


def test_add_summary_below_will_modify_table():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])
    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()

    areas.add_summary(text_span=2, text='total', location='bottom')

    assert table.data == [['header1', 'header2', 'header3'],
                          [1, 2, 3], [None, 2, 4], [None, 3, 5],
                          ['total', None, 12],
                          [2, 3, 4], [None, 4, 5], ['total', None, 9]]


def test_add_summary_only_below_entire_table_will_modify_area():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])
    area = Area(table, 3, 5, (1, 0))
    area.add_summary(text_span=2, text='total', location='bottom')

    assert area.width == 3
    assert area.height == 6
    assert area.position == (1, 0)


def test_add_summary_only_below_entire_table_will_modify_table():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])
    area = Area(table, 3, 5, (1, 0))
    area.add_summary(text_span=2, text='total', location='bottom')

    assert area.data == [[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5],
                         ['total', None, 21]]


def test_add_nested_summary_will_modify_table():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])

    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()
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


def test_each_elem_in_table_is_encapsulated_as_cell():
    table = Table(body=[[1, 2, ], [4, 5, ]])

    assert table.data[0][0] == Cell(1)
    assert table.data[0][1] == Cell(2)
    assert table.data[1][0] == Cell(4)
    assert table.data[1][1] == Cell(5)


def test_set_cell_width_when_initialize_table():
    table = Table(header=[['test', None], ['header1', 'header2']],
                  body=[[1, 2], ])
    assert list(table) == [[Cell('test', width=2), None],
                           [Cell('header1'), Cell('header2')],
                           [Cell(1), Cell(2)]]


def test_iter_table_will_get_cell_list():
    table = Table(body=[[1, 2, ], [4, 5, ]])
    cells = [[Cell(1), Cell(2)], [Cell(4), Cell(5)]]

    assert cells == list(table)


def test_merge_area_will_modify_cell():
    table = Table(header=[['header1', 'header2']],
                  body=[[1, 2], [1, 3], [2, 3]])
    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()

    cells = [[Cell('header1'), Cell('header2')], [Cell(1, height=2), Cell(2)],
             [None, Cell(3)],
             [Cell(2), Cell(3)]]
    assert cells == list(table)


def test_merge_areas_of_three_columns():
    table = Table(header=[['header1', 'header2', 'header3', 'header4']],
                  body=[[1, 2, 3, 4], [1, 2, 3, 5], [1, 2, 3, 6]])
    areas = table.body.select(ColumnSelector(lambda col: col <= 3))
    areas.group().merge()

    cells = [
        [Cell('header1'), Cell('header2'), Cell('header3'), Cell('header4')],
        [Cell(1, height=3), Cell(2, height=3), Cell(3, height=3), Cell(4)],
        [None, None, None, Cell(5)],
        [None, None, None, Cell(6)]]
    assert cells == list(table)


def test_merge_areas_of_three_columns_2():
    table = Table(header=[['header1', 'header2', 'header3', 'header4']],
                  body=[[1, 2, 3, 5], [1, 2, 3, 9], [1, 2, 33, 6],
                        [1, 2, 33, 1],
                        [1, 22, 3, 2], [1, 22, 3, 4], [1, 22, 33, 3],
                        [1, 22, 33, 2]])
    areas = table.body.select(ColumnSelector(lambda col: col <= 3))
    areas.group().merge()

    cells = [
        [Cell('header1'), Cell('header2'), Cell('header3'), Cell('header4')],
        [Cell(1, height=8), Cell(2, height=4), Cell(3, height=2), Cell(5)],
        [None, None, None, Cell(9)],
        [None, None, Cell(33, height=2), Cell(6)],
        [None, None, None, Cell(1)],
        [None, Cell(22, height=4), Cell(3, height=2), Cell(2)],
        [None, None, None, Cell(4)],
        [None, None, Cell(33, height=2), Cell(3)],
        [None, None, None, Cell(2)]]
    assert cells == list(table)


def test_merge_areas_of_three_columns_without_headers():
    table = Table(header=[],
                  body=[[1, 2, 3, 4], [1, 2, 3, 5], [1, 2, 3, 6]])
    areas = table.body.select(ColumnSelector(lambda col: col <= 3))
    areas.group().merge()

    cells = [[Cell(1, height=3), Cell(2, height=3), Cell(3, height=3), Cell(4)],
             [None, None, None, Cell(5)],
             [None, None, None, Cell(6)]]
    assert cells == list(table)


def test_add_nested_summary_will_modify_cell():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])

    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()
    areas.add_summary(text_span=1, text='total', location='left')

    area = Area(table, 3, 7, (1, 0))
    area.add_summary(text_span=2, text='total', location='bottom')

    assert list(table) == [[Cell('header1'), Cell('header2'), Cell('header3')],
                           [Cell(1, height=4), Cell(2), Cell(3)],
                           [None, Cell(2), Cell(4)],
                           [None, Cell(3), Cell(5)],
                           [None, Cell('total'), Cell(12)],
                           [Cell(2, height=3), Cell(3), Cell(4)],
                           [None, Cell(4), Cell(5)],
                           [None, Cell('total'), Cell(9)],
                           [Cell('total', width=2), None, Cell(21)]]


def test_add_nested_summary_will_modify_cell_2():
    table = Table(
        header=[
            ['燃气销售报表', None, None, None, None, None, None],
            ['用气区域', '用气性质', '单价', '表具类型', '地址数', '发行气量', '发行应收']
        ],
        body=[
            ['歆茗', '商业用气', 1.3, '普表', 10, 12, 34],
            ['歆茗', '居民用气', 1.2, 'IC卡表', 11, 12, 12],
            ['歆茗', '居民用气', 1.5, 'IC卡表', 13, 12, 64],
            ['授保', '商业用气', 1.6, '普表', 23, 18, 25],
            ['授保', '居民用气', 1.7, 'IC卡表', 26, 10, 52],
            ['授保', '居民用气', 1.8, 'IC卡表', 16, 25, 12],
        ]
    )

    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()
    areas.add_summary(text_span=3, text='区域合计', location='left')

    table.add_summary(text_span=4, text='总合计', location='bottom')

    assert list(table) == [
        [Cell(value="燃气销售报表", width=7, ), None, None, None, None, None, None],
        [Cell(value="用气区域", ), Cell(value="用气性质", ), Cell(value="单价", ),
         Cell(value="表具类型", ),
         Cell(value="地址数", ), Cell(value="发行气量", ), Cell(value="发行应收", )],
        [Cell(value="歆茗", height=4), Cell(value="商业用气", ), Cell(value=1.3, ),
         Cell(value="普表", ),
         Cell(value=10, ), Cell(value=12, ), Cell(value=34, )],
        [None, Cell(value="居民用气", ), Cell(value=1.2, ), Cell(value="IC卡表", ),
         Cell(value=11, ),
         Cell(value=12, ), Cell(value=12, )],
        [None, Cell(value="居民用气", ), Cell(value=1.5, ), Cell(value="IC卡表", ),
         Cell(value=13, ),
         Cell(value=12, ), Cell(value=64, )],
        [None, Cell(value="区域合计", width=3, ), None, None, Cell(value=34, ),
         Cell(value=36, ),
         Cell(value=110, )],
        [Cell(value="授保", height=4), Cell(value="商业用气", ), Cell(value=1.6, ),
         Cell(value="普表", ),
         Cell(value=23, ), Cell(value=18, ), Cell(value=25, )],
        [None, Cell(value="居民用气", ), Cell(value=1.7, ), Cell(value="IC卡表", ),
         Cell(value=26, ),
         Cell(value=10, ), Cell(value=52, )],
        [None, Cell(value="居民用气", ), Cell(value=1.8, ), Cell(value="IC卡表", ),
         Cell(value=16, ),
         Cell(value=25, ), Cell(value=12, )],
        [None, Cell(value="区域合计", width=3, ), None, None, Cell(value=65, ),
         Cell(value=53, ),
         Cell(value=89, )],
        [Cell(value="总合计", width=4, ), None, None, None, Cell(value=99, ),
         Cell(value=89, ),
         Cell(value=199, )]
    ]


def test_set_global_style_on_table():
    style = {}
    table = Table(
        header=[['test', None, None], ['header1', 'header2', 'header3']],
        body=[[1, 2, 3], [4, 5, 6], [7, 8, 9]],
        style=style
    )

    for row in table:
        for cell in row:
            if cell:
                assert id(cell.style) == id(style)


def test_set_style_of_headers():
    table_style = Style()
    title_style = Style()
    header_style = Style()
    table = Table(header=[[('test', title_style), None, None],
                          [('header1', header_style),
                           ('header2', header_style),
                           ('header3', header_style)]],
                  body=[[1, 2, 3], [4, 5, 6], [7, 8, 9]], style=table_style)

    assert id(table[0][0].style) == id(title_style)
    assert id(table[1][0].style) == id(header_style)
    assert id(table[1][1].style) == id(header_style)
    assert id(table[1][2].style) == id(header_style)
    assert id(table[2][0].style) == id(table_style)


def test_set_style_of_headers_by_method():
    table_style = Style()
    title_style = Style()
    header_style = Style()
    table = Table(header=[['test', None, None],
                          ['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [4, 5, 6], [7, 8, 9]], style=table_style)
    table.header[0].set_style(title_style)
    table.header[1].set_style(header_style)

    assert id(table[0][0].style) == id(title_style)
    assert id(table[1][0].style) == id(header_style)
    assert id(table[1][1].style) == id(header_style)
    assert id(table[1][2].style) == id(header_style)
    assert id(table[2][0].style) == id(table_style)


def test_set_merged_areas_style():
    style = {}
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])
    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge(style=style)

    assert id(table[1][0].style) == id(style)
    assert id(table[4][0].style) == id(style)
    assert table[0][0].style is None


def test_set_summary_style():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])

    label_style = Style()
    value_style = Style()
    label_style2 = Style()
    value_style2 = Style()

    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()
    areas.add_summary(text_span=1, text='total', location='left',
                      label_style=label_style,
                      value_style=value_style)

    area = Area(table, 3, 7, (1, 0))
    area.add_summary(text_span=2, text='total', location='bottom',
                     label_style=label_style2,
                     value_style=value_style2)

    assert table.data == [['header1', 'header2', 'header3'],
                          [1, 2, 3], [None, 2, 4], [None, 3, 5],
                          [None, 'total', 12],
                          [2, 3, 4], [None, 4, 5], [None, 'total', 9],
                          ['total', None, 21]]

    assert id(table[4][1].style) == id(label_style)
    assert id(table[4][2].style) == id(value_style)
    assert id(table[7][1].style) == id(label_style)
    assert id(table[8][0].style) == id(label_style2)
    assert id(table[8][2].style) == id(value_style2)
    assert table[0][0].style is None


def test_excel_writer():
    table = Table(header=[['header1', 'header2', 'header3']],
                  body=[[1, 2, 3], [1, 2, 4], [1, 3, 5], [2, 3, 4], [2, 4, 5]])

    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()
    areas.add_summary(text_span=1, text='total', location='left')

    area = Area(table, 3, 7, (1, 0))
    area.add_summary(text_span=2, text='total', location='bottom')

    wb = Workbook()
    ws = wb.active
    # must be unicode
    ws.title = '报表'
    ws.sheet_properties.tabColor = "1072BA"

    WorkSheetWriter.write(ws, table, (1, 1))

    wb.save('1.xlsx')


# todo: dictnary pool,cell pool etc.
def test_write_excel_with_style():
    table_style = Style({
        'horizontal_align': 'center',
        'vertical_align': 'center',
        'font_size': 12,
        'height': 'auto',
        'width': 'auto',
    })

    title_style = Style({
        'font_size': 15,
        'background_color': 'FF87CEFA',
        'font_weight': 'blod'
    }, extend=table_style)

    header_style = Style({
        'background_color': 'FF87CEFA',
    }, extend=table_style)

    left_total_label_style = Style({
        'background_color': 'fff0f0f0',
    }, extend=table_style)

    left_total_value_style = Style({
        'background_color': 'fff0f0f0',
    }, extend=table_style)

    bottom_total_label_style = Style({
        'background_color': 'ffe6e6e6',
    }, extend=table_style)

    bottom_total_value_style = Style({
        'background_color': 'ffe6e6e6',
    }, extend=table_style)

    table = Table(header=[[('TEST', title_style), None, None, None],
                          [('HEADER1', header_style),
                           ('HEADER2', header_style),
                           ('HEADER3', header_style),
                           ('HEADER4', header_style)]],
                  body=[['One', 'A', 1, 2], ['One', 'A', 2, 3],
                        ['One', 'B', 3, 4], ['Two', 'A', 1, 2],
                        ['Two', 'B', 2, 3]],
                  style=table_style)
    table.header[0].set_style(title_style)
    table.header[1].set_style(header_style)

    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()
    areas.add_summary(text_span=1, text='Total', location='left',
                      label_style=left_total_label_style,
                      value_style=left_total_value_style)

    table.add_summary(text_span=2, text='Total', location='bottom',
                      label_style=bottom_total_label_style,
                      value_style=bottom_total_value_style)

    write_to_excel('2.xlsx', table)


def test_write_non_ascii_chracter_into_excel_with_style():
    table_style = Style({
        'horizontal_align': 'center',
        'vertical_align': 'center',
        'font_size': 12,
        'height': 'auto',
        'width': 'auto',
    })

    title_style = Style({
        'font_size': 15,
        'background_color': 'FF87CEFA',
    }, extend=table_style)

    header_style = Style({
        'background_color': 'FF87CEFA',
    }, extend=table_style)

    left_total_label_style = Style({
        'background_color': 'fff0f0f0',
    }, extend=table_style)

    left_total_value_style = Style({
        'background_color': 'fff0f0f0',
    }, extend=table_style)

    bottom_total_label_style = Style({
        'background_color': 'ffe6e6e6',
    }, extend=table_style)

    bottom_total_value_style = Style({
        'background_color': 'ffe6e6e6',
    }, extend=table_style)

    table = Table(header=[
        ['燃气销售报表', None, None, None, None, None, None],
        ['用气区域', '用气性质', '单价', '表具类型', '地址数', '发行气量', '发行应收']
    ],
        body=[
            ['歆茗', '商业用气', 1.3, '普表', 10, 12, 34],
            ['歆茗', '居民用气', 1.2, 'IC卡表', 11, 12, 12],
            ['歆茗', '居民用气', 1.5, 'IC卡表', 13, 12, 64],
            ['授保', '商业用气', 1.6, '普表', 23, 18, 25],
            ['授保', '居民用气', 1.7, 'IC卡表', 26, 10, 52],
            ['授保', '居民用气', 1.8, 'IC卡表', 16, 25, 12],
        ], style=table_style)
    table.header[0].set_style(title_style)
    table.header[1].set_style(header_style)

    areas = table.body.select(ColumnSelector(lambda col: col == 1))
    areas.group().merge()
    areas.add_summary(text_span=3, text='区域合计', location='left',
                      label_style=left_total_label_style,
                      value_style=left_total_value_style)

    table.add_summary(text_span=4, text='总合计', location='bottom',
                      label_style=bottom_total_label_style,
                      value_style=bottom_total_value_style)

    wb = Workbook()
    ws = wb.active
    # must be unicode
    ws.title = '报表'
    ws.sheet_properties.tabColor = "1072BA"

    WorkSheetWriter.write(ws, table, (0, 0))

    wb.save('3.xlsx')
