#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import unicode_literals


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
        else:
            extend = {
                'font_size': 12,
                'height': 'auto',
                'width': 'auto',
                'horizontal_align': 'center',
                'vertical_align': 'center',
            }
        extend = extend.copy()
        extend.update(dict_1)
        return extend


_default_style = Style()
