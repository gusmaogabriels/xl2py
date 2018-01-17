# -*- coding: utf-8 -*-
from __future__ import division, absolute_import, print_function

from . import dc
from .. import __version__, __author__

class xlref(object):

    def __init__(self,Workbook,Worksheet,Range):
        """ xlref
        Creates a XL reference handlers:
        Params(3):
            Workbook as string or integer -> XL Workbook name or equivalent number in .data
            Workbook as string or integer -> XL Worksheet name or equivalent number in .data
            Range as string -> XL R1C1- or A1- type ranges
        """
        self.__reference = [[Workbook, Worksheet, Range]]
        self.__type__ = 'SingleCell'
        super(xlref,self).__init__()

    def __add__(self,obj):
        if isinstance(obj,xlref):
            oobj = dc(self)
            for i in obj.__reference:
               oobj.__reference.append(i)
               oobj.__type__ = 'MultipleCell'
               return oobj
        else:
            print('Object type-mismatch')

    def __sub__(self,obj):
        if isinstance(obj,xlref):
            oobj = dc(self)
            for i in obj.__reference:
                if i in oobj.__reference:
                    oobj.__reference.remove(i)
                else:
                    pass
            if len(oobj.__reference) == 1 and oobj.__type__ is not 'SingleCell':
                oobj.__type__ = 'SingleCell'
            else:
                pass
            return oobj
        else:
            raise Exception('Object type-mismatch.')

    def __iadd__(self,obj):
        if isinstance(obj,xlref):
            for i in obj.__reference:
                self.__reference.append(i)
            self.__type__ = 'MultipleCell'
            return self
        else:
            raise Exception('Object type-mismatch.')

    def __isub__(self,obj):
        if isinstance(obj,xlref):
            for i in obj.__reference:
                if i in self.__reference:
                    self.__reference.remove(i)
                else:
                    pass
            if len(self.__reference) == 1 and self.__type__ is not 'SingleCell':
                self.__type__ = 'SingleCell'
            else:
                pass
            return self
        else:
            raise Exception('Object type-mismatch.')

    def __call__(self):
        if len(self.__reference)==1:
            return tuple(self.__reference[0])
        else:
            return tuple([tuple(i) for i in self.__reference])
