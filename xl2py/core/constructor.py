# -*- coding: utf-8 -*-
from __future__ import division, absolute_import, print_function

from . import processor, np, xlref_base
from .. import __author__, __version__
from ..com_handlers.handlers import xlcom

__all__ = ['builder']

class builder(object):

    def __init__(self):
        # structure and nodes are attached to this class once Processor generated them
        self.__author__ = __author__
        self.__version__ = __version__
        self.structure = []
        self.nodes = []
        self.__Processor = processor.Processor()
        self.__COM = xlcom()
        self.__COM_status__ = self.__COM.__status__
        self.__data_status__ = self.__Processor.__status__
        self.__status__ = False
        self.ofcell = []
        self.inputcells = []
        self.iolib = self.__Processor.iolib
        self.__path = []
        self.__password = []
        super(builder,self).__init__()

    def connect_com(self,path,password=''):
        """ openxl
        Class initializer:
            Create COM handler object via handlers.XLcom
        Params (2):
            path as string -> XL Workbook path
            password as string -> XL Workbook password
        """
        self.__path = path
        self.__password = password
        try:
            self.__COM.xlopen(path, password)
            if self.__COM.__status__:
                self.__COM_status__ = True
            else:
                print('Invalid XL-file path  and/or password.')
        except Exception as e:
            print(e)

    def disconnect_com(self):
        if self.__COM_status__:
            self.__COM.xlclose()
        else:
            pass

    def set_structure(self,inputs, ofcell, reset = False):
        """ set_structure
        xl2py I/O structure constructor
        Params (4):
            inputs as xlref -> inputs Workbook, Worksheer and A1- or R1C1-type address as xlref object
            ofcell as xlref -> xlref object for the objective function or output eval (__type__ should be 'SingleCell')
            reset -> Boolean for reseting (if True) the structure before defining it.
        """
        if not all([isinstance(x,xlref_base.xlref) for x in [inputs,ofcell]]):
            raise Exception('All inputs must be of type xlref.')
        elif ofcell.__type__ is not 'SingleCell':
            raise Exception("Objective function xlref reference should be 'SingleCell.'")
        else:
            pass
        if reset:
            self.structure = []
            self.nodes = []
            self.__Processor.iolib = {}
            self.__data_status__ = False
            self.__status__ = False
        else:
            pass
        if self.__COM_status__:
            self.__Processor.attach_com_obj(self.__COM) # couple the COM interface to the xl2py processor
            try:
                self.__Processor.set_io(inputs(),ofcell())
                self.nodes = self.__Processor.pynodes
                self.structure = self.__Processor.pydata
                self.inputcells = inputs
                self.ofcell = ofcell
                self.__status__ = True
            except Exception as e:
                print('set_io:', e)
        else:
            raise Exception('XLcom object is not connected. Use connect_com to connect to a XL file.')

    def set_input_values(self,values):
        """ set_input_values
        Change the already set input values.
        The shape of values must agree with the size of the input already parsed. (Checked within the Processor object)
        """
        if self.__status__:
            if not isinstance(values,(np.ndarray,list,tuple)) or np.shape(values)[0] != len(self.iolib['inputsref']) :
                raise Exception('values should match the size/shape of the parsed input ranges.')
            else:
                try:
                    self.__Processor.set_pyranges(self.iolib['inputsref'],values)
                except Exception as e:
                    print('Input values setting error: ',e)
        else:
            raise Exception('Builder object not properly initialized. I/O structure not available.')

    def get_output_value(self):
        """ get_output_value
        Evaluates the objective function cell and returns its value.
        """
        if self.__status__:
            self.__Processor.evalstructure()
            return self.__Processor.get_pyranges(self.iolib['ofcell'])[0]
        else:
            raise Exception('Builder object not properly initialized. I/O structure not available.')

    def test_nodes(self):
        """ test_nodes
        Returns a list of nodes that failed validation at given relative threshold
        """
        return self.__Processor.validatenodes()
