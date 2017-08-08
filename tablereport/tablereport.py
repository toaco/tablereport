#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import unicode_literals


class Areas(list):
    def __init__(self, areas=None):
        if areas is None:
            areas = []
        super(Areas, self).__init__(areas)

    def merge(self, style=None):
        for area in self:
            area.merge(style)

    def add_summary(self, text_span, text, location, label_style=None,
                    value_style=None):
        for area in self:
            area.add_summary(text_span, text, location, label_style,
                             value_style)

    def one(self):
        """assert Areas contain only one Area and return it"""
        assert len(self) == 1
        return self[0]


class Area(object):
    def __init__(self, table, width, height, position, style=None):
        self.table = table
        self.table.areas.append(self)

        self.width = width
        self.height = height

        self._x, self._y = position

        self.style = style

    @property
    def position(self):
        return self._x, self._y

    @position.setter
    def position(self, value):
        self._x, self._y = value

    @property
    def data(self):
        rows = []
        for row_num in xrange(self.height):
            position = self._x + row_num, self._y
            row = Row(self.table, position, self.width)
            rows.append(row)
        return rows

    def select(self, selector):
        # select an area in self
        area = selector.select(self)
        return area

    def merge(self, style=None):
        cell = self.data[0][0]

        for row in xrange(len(self.data)):
            for col in xrange(len(self.data[0])):
                self.data[row][col] = None
        cell.height = len(self.data)
        if style is not None:
            cell.style = style

        self.data[0][0] = cell

    def add_summary(self, text_span, text, location, label_style=None,
                    value_style=None):
        if location == 'left':
            cell = self.data[0][0]
            cell.height += 1
            new_col_num = self._add_row_at_bottom(
                label_style, text, text_span, value_style, self.width)
        elif location == 'down':
            new_col_num = self._add_row_at_bottom(
                label_style, text, text_span, value_style, offset=0)
        else:
            raise ValueError
        self._update_existed_areas(new_col_num)

        self.table.height += 1

    def set_style(self, style):
        for row in self.data:
            for cell in row:
                if cell:
                    cell.style = style

    def _update_existed_areas(self, new_col_num):
        for area in self.table.areas:
            x, y = area.position
            if new_col_num > x + area.height:
                continue
            elif x + area.height >= new_col_num > x:
                area.height += 1
            else:
                x += 1
                area.position = x, y

    def _add_row_at_bottom(self, label_style, text, text_span, value_style,
                           offset=0):
        new_col_num = self._x + self.height
        self.table.data.insert(new_col_num,
                               [None] * self.table.width)
        appended_row = self.table[new_col_num]

        # set summary cell
        if text_span != 0:
            appended_row[self._y + offset] = Cell(text, width=text_span)
            if label_style is not None:
                appended_row[self._y + offset].style = label_style

        # summarize columns need to be summarized
        for col_num in xrange(self._y + offset + text_span,
                              self.table.width):
            total = 0
            for row_num in xrange(self._x, self._x + self.height):
                if row_num in self.table.total_row_nums:
                    continue
                total += self.table[row_num][col_num].value
            appended_row[col_num] = Cell(total)
            if value_style is not None:
                appended_row[col_num].style = value_style
            else:
                appended_row[col_num].style = self.table.style
        self.table.total_row_nums.add(new_col_num)
        return new_col_num

    def __getitem__(self, item):
        if item == self.height:
            raise IndexError
        return Row(self.table, (self._x + item, self._y), self.width)

    def __setitem__(self, key, value):
        if key == self.height:
            raise IndexError
        self.table.data[key + self._x] = value


class Table(object):
    def __init__(self, header=None, body=None, style=None):

        if header is None:
            header = []

        if body is None:
            body = []

        self._header_data = header
        self._body_data = body
        self._data = self._header_data + self._body_data

        for row_num in xrange(len(self._data)):
            for col_num in xrange(len(self._data[0])):
                cell = self._data[row_num][col_num]
                if cell is not None:
                    if isinstance(cell, tuple):
                        self._data[row_num][col_num] = Cell(cell[0],
                                                            style=cell[1])
                    else:
                        self._data[row_num][col_num] = Cell(
                            self._data[row_num][col_num],
                            style=style)
                    self._auto_merge(self._data, row_num, col_num)
        self.areas = []
        self.total_row_nums = set()
        try:
            width = len(self._data[0])
        except IndexError:
            width = 0

        self.width = width
        self.height = len(self._data)
        self.style = style
        self.header = Area(table=self, width=self.width,
                           height=len(self._header_data),
                           position=(0, 0))
        self.body = Area(table=self, width=self.width,
                         height=len(self._body_data),
                         position=(len(self._header_data), 0))

    @property
    def data(self):
        return self._data

    def __getitem__(self, item):
        return self._data[item]

    def __setitem__(self, key, value):
        self._data[key] = value

    def select(self, selector):
        # select an area in self
        table = Area(table=self, width=self.width, height=self.height,
                     position=(0, 0))
        areas = selector.select(table)
        return areas

    def add_summary(self, text_span, text, location, label_style=None,
                    value_style=None):
        self.body.add_summary(text_span, text, location, label_style,
                              value_style)

    @staticmethod
    def _auto_merge(data, row_num, col_num):
        # todo: range judge
        for i in xrange(row_num + 1, len(data)):
            if data[i][col_num] is None:
                data[row_num][col_num].height += 1
            else:
                break

        for j in xrange(col_num + 1, len(data[0])):
            if data[row_num][j] is None:
                data[row_num][col_num].width += 1
            else:
                break


class Row(object):
    def __init__(self, table, position, width):
        self.table = table
        self.x, self.y = position
        self.width = width

    def __getitem__(self, col):
        assert col < self.width
        return self.table[self.x][self.y + col]

    def __setitem__(self, col, value):
        assert col < self.width
        self.table[self.x][self.y + col] = value

    def __iter__(self):
        for i in xrange(self.width):
            yield self.table[self.x][self.y + i]

    def __eq__(self, iterable):
        return all(self[i] == iterable[i] for i in xrange(self.width))

    def __len__(self):
        return self.width

    def __repr__(self):
        return str([self[i] for i in xrange(self.width)])

    def set_style(self, style):
        for cell in self:
            if cell:
                cell.style = style


class Cell(object):
    def __init__(self, value, style=None, width=1, height=1):
        self.value = value
        self.width = width
        self.height = height
        self.style = style

    def __eq__(self, other):
        if type(other) == Cell:
            return self.__dict__ == other.__dict__
        else:
            assert type(self.value) == type(other)
            return self.value == other

    def __repr__(self):
        return ('Cell(value="{}", style={}, width={}, height={})'
                .format(repr(self.value), self.style, self.width, self.height)
                .encode('utf-8'))

    __str__ = __repr__


class ColumnSelector:
    def __init__(self, column, group):
        self.column = column
        self.group = group

    def select(self, area):
        assert self.column <= area.width
        x, y = area.position

        y += self.column - 1
        area = Area(table=area.table, width=1, height=area.height,
                    position=(x, y))

        areas = Areas()
        if not self.group:
            areas.append(area)
        else:
            start_value = area.data[0][0]
            start_index = 0
            for row_num in xrange(1, len(area.data)):
                if area.data[row_num][0] == start_value:
                    continue
                else:
                    sub_area = Area(table=area.table, width=1,
                                    height=row_num - start_index,
                                    position=(x + start_index, y))
                    start_value = area.data[row_num][0]
                    start_index = row_num
                    areas.append(sub_area)
            else:
                sub_area = Area(table=area.table, width=1,
                                height=len(area.data) - start_index,
                                position=(x + start_index, y))
                areas.append(sub_area)
        return areas


class Style(object):
    """
    Style and style check is here
    """

    def __new__(cls, dict_1=None, extend=None):
        if dict_1 is None:
            dict_1 = {}
        else:
            assert isinstance(dict_1, dict)

        if extend is not None:
            assert isinstance(extend, dict)
            extend = extend.copy()
            extend.update(dict_1)
            return extend
        return dict_1
