#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from .tablereport import Areas, Area, Cells


class ColumnSelector(object):
    def __init__(self, func, width=1):
        """
        :param func: eg: ``lambda col:col==1``
        """
        self.func = func
        self.width = width

    def select(self, area):
        x, y = area.position
        width = area.width

        areas = Areas()
        for col in range(width):
            if self.func(col + 1):
                area = Area(table=area.table, width=self.width,
                            height=area.height,
                            position=(x, y + col))
                areas.append(area)
        return areas


class RowSelector(object):
    def __init__(self, func, height=1):
        """
        :param func: eg: ``lambda row:row==1``
        """
        self.func = func
        self.height = height

    def select(self, area):
        x, y = area.position
        height = area.height

        areas = Areas()
        for row in range(height):
            if self.func(row + 1):
                area = Area(table=area.table, width=area.width,
                            height=self.height,
                            position=(x + row, y))
                areas.append(area)
        return areas


class CellSelector(object):
    def __init__(self, func):
        self.func = func

    def select(self, area):
        cells = Cells()
        for row in area.data:
            for cell in row:
                if self.func(cell):
                    cells.append(cell)

        return cells
