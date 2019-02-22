#!/usr/bin/env python
# coding: utf-8
#
# Licensed under MIT
#

import setuptools
setuptools.setup(
    name = "tablereport",
    version = "0.1",
    packages = ['tablereport','tablereport/writer'],
    install_requires = [
    'openpyxl',
    'pytest-cov',
    'six',
    'pytest',
    ]
    )
