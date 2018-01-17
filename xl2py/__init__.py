# -*- coding: utf-8 -*-
from __future__ import division, absolute_import, print_function

__author__ = {'Gabriel S. Gusmao' : 'gusmaogabriels@gmail.com'}
__version__ = '2.0b'

"""

By Gabriel S. Gusmão (Gabriel Sabença Gusmão)
Oct, 2015

    xl2Py version 2.0b

    ~~~~

    "An Excel 2 Python I/O Structure reShaping"

    :copyright: (c) 2015 Gabriel S. Gusmão
    :license: MIT, see LICENSE for more details.

"""

from .core import processor
from .core.xlref_base import xlref
from .core.constructor import builder

__all__ = ['processor','builder','xlref','__author__','__version__']
