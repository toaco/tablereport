#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from .style import _default_style


class Table(object):
    """
    Table is the core class of TableReport.
     
    A table looks like a nested list. Each inner list represents a row of the 
    table, and each row contains some cells::

        [
            [Cell('header1'), Cell('header2')],
            [Cell(1), Cell(2)],
            [Cell(3), Cell(4)]
        ]
    
    We can create this by the following ways, and each element in a row will be 
    auto wrapped into a cell::
        
        table = Table(
            header=[['header1', 'header2']],
            body=[[1, 2], [3, 4]]
        )

    All Cells in a table has a style attribute, the default value of which can 
    be set by ``Style`` argument, and we can also separately set the style 
    attribute of a cell as below::
    
        table = Table(
            header=[[('header1',style), 'header2']],
            body=[[1, 2], [3, 4]]
        )
    
    An import thing is  that ``None`` in a row will be specially handled. 
    ``None`` will be used to auto merge cells. A sample could explain this::
    
        table = Table(header=[['test', None], ['header1', 'header2']],
                      body=[[1, 2], ])
                      
    This will create a table as below. This feature is usually used for custom 
    table header::

        [[Cell('test', width=2), None],
        [Cell('header1'), Cell('header2')],
        [Cell(1), Cell(2)]]
    """

    def __init__(self, header=None, body=None, style=None):
        if header is None:
            header = []

        if body is None:
            body = []

        if style is None:
            style = _default_style
        self._flag = []
        self._header_data = header
        self._body_data = body
        self._data = self._header_data + self._body_data

        for row_num in range(len(self._data)):
            for col_num in range(len(self._data[0])):
                cell = self._data[row_num][col_num]
                if cell is not None:
                    if isinstance(cell, tuple):
                        self._data[row_num][col_num] = Cell(cell[0],
                                                            style=cell[1])
                    else:
                        self._data[row_num][col_num] = Cell(
                            self._data[row_num][col_num],
                            style=style)
                    self._auto_merge(self._data, row_num, col_num, self._flag)
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

    def summary(self, label, label_span, location='bottom', label_style=None,
                value_style=None):
        self.body.summary(label, label_span, location, label_style, value_style)

    @staticmethod
    def _auto_merge(data, row_num, col_num, flag):

        for i in range(row_num + 1, len(data)):
            if data[i][col_num] is None:
                data[row_num][col_num].height += 1
                flag.append((i, col_num))
            else:
                break

        for j in range(col_num + 1, len(data[0])):
            if data[row_num][j] is None and (row_num, j) not in flag:
                data[row_num][col_num].width += 1
            else:
                break


class Cell(object):
    def __init__(self, value, style=None, width=1, height=1):
        self.value = value
        self.width = width
        self.height = height
        if style is None:
            style = _default_style
        self.style = style

    def __eq__(self, other):
        if type(other) == Cell:
            return self.__dict__ == other.__dict__
        else:
            assert type(self.value) == type(other)
            return self.value == other

    def __repr__(self):
        return ('Cell(value="{}", style={}, width={}, height={})'
                .format(repr(self.value), self.style, self.width, self.height))

    __str__ = __repr__


class Cells(list):
    def __init__(self, areas=None):
        if areas is None:
            areas = []
        super(Cells, self).__init__(areas)

    def set_style(self, style):
        for cell in self:
            cell.style = style


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
        for row_num in range(self.height):
            position = self._x + row_num, self._y
            row = Row(self.table, position, self.width)
            rows.append(row)
        return rows

    @property
    def left(self):
        """left side area"""
        x, y = self.position
        position = x, y + self.width
        width = self.table.width - self.width - 1
        area = Area(self.table, width, self.height, position, style=None)
        return area

    def select(self, selector):
        # select an area in self
        area = selector.select(self)
        return area

    def group(self):
        """group a area, now only support group a col"""
        if not self.width == 1:
            return

        start_index = 0
        start_value = self.data[0][0]
        start_x, start_y = self.position

        areas = Areas()
        for row_num in range(1, len(self.data)):
            if self.data[row_num][0] == start_value:
                continue
            else:
                area = Area(table=self.table, width=1,
                            height=row_num - start_index,
                            position=(start_x + start_index, start_y))
                start_value = self.data[row_num][0]
                start_index = row_num
                areas.append(area)
        else:
            sub_area = Area(table=self.table, width=1,
                            height=len(self.data) - start_index,
                            position=(start_x + start_index, start_y))
            areas.append(sub_area)
        return areas

    def merge(self, style=None):
        cell = self.data[0][0]

        for row in range(len(self.data)):
            for col in range(len(self.data[0])):
                self.data[row][col] = None
        cell.height = len(self.data)
        if style is not None:
            cell.style = style

        self.data[0][0] = cell

    def summary(self, label=None, label_span=0, location='bottom',
                label_style=None,
                value_style=None):
        if location == 'bottom':
            new_row_num = self._add_row_at_bottom(label_style, label,
                                                  label_span, value_style)

            self._update_existed_areas(new_row_num)
            self.table.height += 1
        elif location == 'right':
            self._add_col_at_right(label_style, label, label_span, value_style)
            # todo: update existed areas
            self.table.width += 1
        else:
            raise NotImplemented

    def set_style(self, style):
        for row in self.data:
            for cell in row:
                if cell:
                    cell.style = style

    def _update_existed_areas(self, new_row_num):
        self_y = self.position[1]
        self_width = self.width
        for area in self.table.areas:
            x, y = area.position
            if new_row_num > x + area.height:
                continue
            elif x + area.height >= new_row_num > x:
                # handle merged cell
                cell = area[0][0]
                if cell.width == area.width and cell.height == area.height:
                    # todo
                    if self_y <= y + area.width - 1 \
                            and self_y + self_width - 1 >= y:
                        pass
                    else:
                        cell.height += 1

                area.height += 1

            else:
                x += 1
                area.position = x, y

    def _add_row_at_bottom(self, label_style, text, label_span, value_style):
        new_row_num = self._x + self.height
        self.table.data.insert(new_row_num,
                               [None] * self.table.width)
        appended_row = self.table[new_row_num]

        # add label cell
        if label_span != 0:
            appended_row[self._y] = Cell(text, width=label_span)
            if label_style is not None:
                appended_row[self._y].style = label_style
            else:
                appended_row[self._y].style = self.table.style

        # add summarized cells
        # todo: not iterate to self.table.width
        for col_num in range(self._y + label_span,
                             self.table.width):
            total = 0
            for row_num in range(self._x, self._x + self.height):
                if row_num in self.table.total_row_nums:
                    continue
                total += self.table[row_num][col_num].value
            appended_row[col_num] = Cell(total)
            if value_style is not None:
                appended_row[col_num].style = value_style
            else:
                appended_row[col_num].style = self.table.style
        self.table.total_row_nums.add(new_row_num)
        return new_row_num

    def _add_col_at_right(self, label_style, text, label_span, value_style):
        new_col_num = self._y + self.width
        for row in self.table.data:
            row.insert(new_col_num, None)

        appended_col = Column(table=self.table, position=(0, new_col_num),
                              height=self.height)

        # add label cell
        if label_span != 0:
            appended_col[self._x] = Cell(text, height=label_span)
            if label_style is not None:
                appended_col[self._x].style = label_style
            else:
                appended_col[self._x].style = self.table.style

        # add summarized cells
        for row_num in range(self._x + label_span, self._x + self.height):
            total = 0
            for col_num in range(self._y, self._y + self.width):
                total += self.table[row_num][col_num].value
            appended_col[row_num] = Cell(total)
            if value_style is not None:
                appended_col[row_num].style = value_style
            else:
                appended_col[row_num].style = self.table.style
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


class Areas(list):
    def __init__(self, areas=None):
        if areas is None:
            areas = []
        super(Areas, self).__init__(areas)

    @property
    def left(self):
        areas = Areas()
        for area in self:
            areas.append(area.left)

        return areas

    def group(self):
        areas = Areas()
        for area in self:
            areas.extend(area.group())

        return areas

    def merge(self, style=None):
        areas = Areas()
        for area in self:
            area.merge(style)
            areas.append(area)

        return areas

    def summary(self, label=None, label_span=0, location='bottom',
                label_style=None,
                value_style=None):
        for area in self:
            area.summary(label, label_span, location, label_style, value_style)

    def set_style(self, style):
        for area in self:
            area.set_style(style)

    def one(self):
        """assert Areas contain only one Area and return it"""
        assert len(self) == 1
        return self[0]


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
        for i in range(self.width):
            yield self.table[self.x][self.y + i]

    def __eq__(self, iterable):
        return all(self[i] == iterable[i] for i in range(self.width))

    def __len__(self):
        return self.width

    def __repr__(self):
        return str([self[i] for i in range(self.width)])

    def set_style(self, style):
        for cell in self:
            if cell:
                cell.style = style


class Column(object):
    def __init__(self, table, position, height):
        self.table = table
        self.x, self.y = position
        self.height = height

    def __getitem__(self, row):
        assert row < self.height
        return self.table[self.x + row][self.y]

    def __setitem__(self, row, value):
        assert row < self.height
        self.table[self.x + row][self.y] = value

    def __iter__(self):
        for i in range(self.height):
            yield self.table[self.x + i][self.y]

    def __eq__(self, iterable):
        return all(self[i] == iterable[i] for i in range(self.height))

    def __len__(self):
        return self.height

    def __repr__(self):
        return str([self[i] for i in range(self.height)])

    def set_style(self, style):
        for cell in self:
            if cell:
                cell.style = style
