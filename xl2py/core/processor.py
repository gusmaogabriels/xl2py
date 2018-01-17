# -*- coding: utf-8 -*-
from __future__ import division, absolute_import, print_function

from .. import __author__, __version__
from ..core import np, re, dc, time
from ..com_handlers.handlers import xlcom
from ..conversion_lib.funcs_lib import Funlib_obj, FunObj, RefObj, CalcBlock, NumObj, CalcHandler

class Processor(object):

    def __init__(self):
        self.__author__ = __author__
        self.__version__ = __version__
        self.__COM = []
        self.__status__ = False
        self.Funlib = Funlib_obj
        self.__convref = self.Funlib.fun_database
        self.pydata = {'Workbooks' : []}
        self.__CalcHandler = CalcHandler(self.pydata,1e-10)
        self.pyformulas = []
        self.pynodes = {0 : []}
        self.datalevel = 0
        self.bkp = []
        self.buffer = {}
        self.calcstruct = []
        self.intranode = []
        self.circularrefs = []
        self.tracer = 0
        self.hascircularreferences = False
        self.iolib = {'inputsref' : [], 'ofcell' : [], 'algorithm' : {}}
        super(Processor,self).__init__()

    def attach_com_obj(self,xlcom_obj):
        if isinstance(xlcom_obj,xlcom):
            self.__COM = xlcom_obj
            if not self.__COM.__status__:
                raise Exception('xlcom object is not connected to a XL file.')
            else:
                pass
        else:
            raise Exception('xlcom_obj must be of type xlcom.')

    def set_pyranges(self,pylist,values):
        """ set_pyranges
        parameters must have been converted by the internal function listconnect
        or be already shapped in the way references point to internal structure (pydata)
        params(2):
            pylist as iinteger list/tuple
                lists/tuple of [[Workbook as integer],[Worksheet as integer],[[Row as integer],[Column as integer]]]
                e.g. [[1,2,[[2],[4,6]]]] -> Internal nodal reference WB = 1, WS = 2, Column = 2, Rows = 4 to 6
            values as numeric of the same size of total references in pylist
        """
        for i in range(0,len(pylist)):
            if not all((np.diff(pylist[i][2])+1).flatten() == np.shape(values[i])):
                print('pyranges ({}) and values ({}) shapes should be the same.'.format((np.diff(pylist[i][2])+1).flatten(),np.shape(values[i])))
                raise Exception
            else:
                pass
        for i in range(0,len(pylist)):
            val = np.array(values[i]).flatten().tolist()
            for r in range(min(pylist[i][2][0]),max(pylist[i][2][0])+1):
                for c in range(min(pylist[i][2][1]),max(pylist[i][2][1])+1):
                    self.pydata[pylist[i][0]][pylist[i][1]][r][c] = val[0]
                    val.pop(0)

    def get_pyranges(self,pylist):
        """ get_pyranges
        retrieves stored values from the internal structure (pydata)
        if they are not stored, they are retrieved from the XL process and then stored
        params (1): pylist as in set_pyranges
            e.g. [[1,2,[[2],[4,6]]]] -> Internal nodal reference WB = 1, WS = 1, Column = 2, Rows = 4 to 6
        returns pylist corresponding value from pydata
        """
        pyvalues = []
        flag = False
        for item in pylist:
            for r in range(min(item[2][0]),max(item[2][0])+1):
                for c in range(min(item[2][1]),max(item[2][1])+1):
                    try:
                        nWB = [item[0] if type(item[0]) is int else self.pydata['Workbooks'].index(item[0])+1][0]
                        nWS = [item[1] if type(item[1]) is int else self.pydata[nWB]['Worksheets'].index(item[1])+1][0]
                        pyvalues.append(self.pydata[nWB][nWS][r][c])
                    except:
                        if not flag:
                            print('Range was not read during py strucutre creation.')
                            print('Values being now attached to the structure...')
                        flag = True
                        pyvalues = []
                        pylist0 = dc(pylist)
                        for xlitem in pylist:
                            xlitem = self.processxlbuffer(xlitem)
                            for subitem in xlitem:
                                if subitem != []:
                                    self.processxlitem(subitem)
                                else:
                                    pass
                        return self.get_pyranges(pylist0)[0], flag
        return np.array(pyvalues), flag

    def __xlformula_excavator(self,formula):
        """ __xlformula_excavator
        param (1): formula as string
        returns a structure of cell blocks in the xl arithmetic formula
        """
        class0 = ['^']
        class1 = ['*','/']
        class2 = ['>','<','<=','>=','<>','=']
        class3 = ['+','-']
        classglobal = class0+class1+class2+class3
        subtrack = 0
        calc_block = []
        strbuffer = ''
        formula0 = ''
        for s in formula:
            formula0 += s
        # Let the arithmetic splitting (diggestion) beggin
        while len(formula)>0:
            if formula[0] in classglobal:
                if len(strbuffer)>0:
                    if all([re.compile(r'[0-9\.]').match(x) for x in strbuffer]):
                        calc_block.append(NumObj(strbuffer))
                    else:
                        calc_block.append(self.link_xlranges(self.datalevel, len(self.pynodes[self.datalevel])-1,strbuffer))
                else:
                    pass
                if formula[1] in classglobal:
                    calc_block.append(formula[0:2])
                    formula = formula[2:]
                else:
                    calc_block.append(formula[0])
                    formula = formula[1:]
                strbuffer = ''
            elif formula[0] == '(':
                if len(strbuffer) == 0: # then it MUST BE an internal block and, therefore, find the end and call the excavator recursively
                    subtrack = 1
                    formula = formula[1:]
                    while subtrack != 0:
                        if formula[0] == '(':
                            subtrack += 1
                        elif formula[0] == ')':
                            subtrack -= 1
                        else:
                            pass
                        if subtrack != 0:
                            strbuffer += formula[0]
                        else:
                            pass
                        formula = formula[1:]
                    calc_block.append(self.__xlformula_excavator(strbuffer))
                    strbuffer = ''
                elif len(strbuffer) > 0: # therefore HAS TO BE a function (otherwise would have been an operand)
                    # get function
                    fun = strbuffer
                    strbuffer = ''
                    parameters = [] # parameter holder
                    subtrack = 1
                    formula = formula[1:]
                    while subtrack != 0:
                        if formula[0] == '(':
                            subtrack += 1
                            strbuffer += formula[0]
                        elif formula[0] == ')':
                            subtrack -= 1
                            strbuffer += formula[0]
                        elif formula[0] == ',' and subtrack == 1:
                            parameters.append(self.__xlformula_excavator(strbuffer))
                            strbuffer = ''
                        else:
                            strbuffer += formula[0]
                        formula = formula[1:]
                    parameters.append(self.__xlformula_excavator(strbuffer[0:-1]))
                    strbuffer = ''
                    calc_block.append(FunObj(fun,parameters))
                else:
                    pass
            else:
                strbuffer += formula[0]
                formula = formula[1:]
        if len(strbuffer)>0:
            if all([re.compile(r'[0-9\.]').match(x) for x in strbuffer]):
                calc_block.append(NumObj(strbuffer))
            else:
                calc_block.append(self.link_xlranges(self.datalevel, len(self.pynodes[self.datalevel])-1,strbuffer))
        else:
            pass
        if len(calc_block)==1:
            return calc_block[0]
        else:
            return CalcBlock(calc_block,formula0)

    def xlformula2py(self,formula,output_ref):
        """ xlformula2py
        params (1): formula as string containing R1C1-type xl formulas
        returns a callable object (type : NodePy), which evaluates the cell formula and modify it as suitable
        """
        formula = self.__xlformula_excavator(formula[1:])
        formula.set_output(*output_ref)
        return formula

    def __get_WBWS(self,WB,WS):
        """ __get_WBWS
        params(2) -> WB and WS as string or number related to a Workbook or Worksheet, respectively
        returns a the a list containing a pairs of number and string references for them
            ... if WB or WS have not ever been seen, they are appended to the buffer and data
        """
        vals = [[],[],[],[],[False,False]]
        if isinstance(WB,(unicode,unicode,raw_input)):
            vals[0] = WB
            if WB in self.pydata['Workbooks']:
                vals[1] = self.pydata['Workbooks'].index(WB)+1
            else:
                self.pydata['Workbooks'].append(WB)
                vals[1] = len(self.pydata)
                self.pydata.__setitem__(vals[1],{'Worksheets' : []})
                vals[4][0] = True
        else:
            vals[1] = WB
            vals[0] = self.pydata['Workbooks'][WB-1]
        if isinstance(WS,(unicode,unicode,raw_input)):
            vals[2] = WS
            if WS in self.pydata[vals[1]]['Worksheets']:
                vals[3] = self.pydata[vals[1]]['Worksheets'].index(WS)+1
            else:
                self.pydata[vals[1]]['Worksheets'].append(WS)
                vals[3] = len(self.pydata[vals[1]])
                self.pydata[vals[1]].__setitem__(vals[3], {})
                vals[4][0] = True
        else:
            vals[3] = WS
            vals[2] = self.pydata[vals[1]]['Worksheets'][WS-1]
        if vals[0] not in self.buffer:
            self.buffer.__setitem__(vals[0],{vals[2]:[]})
            vals[4][1] = True
        elif WS not in self.buffer[WB]:
            self.buffer[vals[0]].__setitem__(vals[2],[])
            vals[4][1] = True
        else:
            pass
        return vals

    def link_xlranges(self,level,node,reference):
        """ inner method link_xlranges
        creates formula refs to pydata assgined positions
        this basically takes a R1C1 formula and change its R1C1 reference
        to an evaluable string that points to the pydata structure
         e.g. WB, WS, R1C1 -> pydata[nWB][nWS][R=1][C=1]
        """
        # p as range grabber
        reference = reference.replace("'",'')
        p = re.compile(r'([\w\s\.]+(?=\]))?.?([\w\s\.]+(?=\!))?.?(R[\d]+C[\d]+[[:R]*[\d]*[C]*[\d]*]?)',re.M | re.UNICODE)
        dependence = self.pynodes[level][node]['dependence']
        reference = p.findall(reference)[0]
        # colect worbook and worksheet names WB/WS and indices nWB/nWS that matches those in py data from formula
        WB = [self.pynodes[level][node]['filename'] if reference[0]=='' else reference[0]][0]
        WS = [self.pynodes[level][node]['sheet'] if reference[1]=='' else reference[1]][0]
        WB, nWB, WS, nWS, flags = self.__get_WBWS(WB,WS)
        # collect rows and columns from formula
        R = [int(pos) for pos in re.compile(r'(?<=R)\d+').findall(reference[2])]
        C = [int(pos) for pos in re.compile(r'(?<=C)\d+').findall(reference[2])]
        # define WS, WB, rows and columns on which the formula depends
        if dependence.has_key(nWB):
            if dependence[nWB].has_key(nWS):
                pass
            else:
                dependence[nWB].__setitem__(nWS,[])
        else:
            dependence.__setitem__(nWB,{nWS : []})
        checker = False
        for refs in dependence[nWB][nWS]:
            if min(refs[0])<=min(R)<=max(refs[0]) and min(refs[0])<=max(R)<=max(refs[0])\
            and min(refs[1])<=min(C)<=max(refs[1]) and min(refs[1])<=max(C)<=max(refs[1]):
                checker = True
            else:
                pass
        if checker == False:
            dependence[nWB][nWS].append([R,C])
        else:
            pass
        # return a reference object
        return RefObj(self.pydata,nWB,nWS,R,C)

    def creatxlnode(self,WB,WS,address,formula):
        """ inner method/function) creatxlnode
        if node already exists -> returns False
        if node has not been created -> creates node in pynodes
            converts formula to Python evaluable representation  with xlformula2py and store it in pyformulas
            indexes pyformulas position onto node
            links R1C1 addresses in formula to evaluable references to pydata
            returns True
        params (4):
            WB -> Workbook name as string
            WS -> Worksheet name as string
            address -> R1C1-type address
            formula -> R1C1-type formula
        """
        struct = {'filename': WB, 'sheet' : WS, 'row' : [], 'column' : [], 'dim' : [], 'formulaindex' : [], 'dependence' : dict()}
        struct['row'] = [int(R) for R in re.compile(r'(?<=R)[\d]+',re.M).findall(address)]
        struct['column'] = [int(C) for C in re.compile(r'(?<=C)[\d]+',re.M).findall(address)]
        for level in self.pynodes:
            for datastruct in self.pynodes[level]:
                if datastruct['filename'] == struct['filename'] and datastruct['sheet'] == struct['sheet'] and \
                datastruct['row'] == struct['row'] and datastruct['column'] == struct['column']:
                    return False
                else:
                    pass
        self.pynodes[self.datalevel].append(struct)
        WB, nWB, WS, nWS, flags = self.__get_WBWS(WB,WS)
        formula = self.xlformula2py(formula,[nWB,nWS,struct['row'],struct['column']])
        self.pyformulas.append(formula)
        formulaindex = len(self.pyformulas)-1
        struct['formulaindex'] = formulaindex
        dim = self.__COM.dim_ranges(address)
        struct['dim'] = [int(dim[0][0]), int(dim[0][1])]
        return True

    def storedata(self,rangeobj,nWS,nWB):
        """ inner method storedata
        Utilized for deploying XL COM range object values into pydata structure,
        references by their related WS and WB numbers.
        params (3):
            rangeobj as XL COM range object
            nWS and nWB as references to Workbook and Worksheet, respectively, as integers
        """
        values = []
        values = rangeobj.Value
        rows = [rangeobj.Cells.Row, rangeobj.Cells.Row+rangeobj.Cells.Rows.Count-1]
        columns = [rangeobj.Cells.Column, rangeobj.Cells.Column+rangeobj.Cells.Columns.Count-1]
        if not isinstance(values,(tuple,list)):
            flag = True
        else:
            flag = False
        for Row in range(rows[0],rows[1]+1):
            for Column in range(columns[0],columns[1]+1):
                if flag:
                    Value = values
                else:
                    Value = values[Row-rows[0]][Column-columns[0]]
                if not isinstance(Value,float):
                    try:
                        Value = float(Value)
                    except:
                        pass
                else:
                    pass
                if not(self.pydata[nWB][nWS].has_key(Row)):
                    self.pydata[nWB][nWS].__setitem__(Row, {Column : [Value if Value != None else 0.0][0]})
                elif not(self.pydata[nWB][nWS][Row].has_key(Column)):
                    self.pydata[nWB][nWS][Row].__setitem__(Column, [Value if Value != None else 0.0][0])
                else:
                    pass

    def processxlitem(self,item): # item as range
        """ inner method processxlitem
            check and add (if not already in) item into object buffer
            creates node for pynodes if item not yet in buffer (save time)
            ... called by xlstruct_constructor in recursive way in case XL object derived from
            item have dependents (dependents are then processed in tree-branch inner loops)
            params (1):
                item as range reference -> list or tuple of indexed WB, WS, rows and columns
        """
        WB = item[0]
        WS = item[1]
        R = item[2][0]
        C = item[2][1]
        if type(WS) is not int:
            if WB != self.__COM.Workbook.Name or WS != self.__COM.Worksheet.Name: # Change WB and WS if necessary
                self.__COM.change_path(WB,WS)
        else:
            if WS != self.pydata[WB]['Worksheets'].index(self.__COM.Worksheet.Name)+1:
                self.__COM.change_path(WB,WS)
        xladdress = [[r,c] for r in range(min(R),max(R)+1) for c in range(min(C),max(C)+1)]
        formula_array = []
        nWB = [WB if type(WB) is int else self.pydata['Workbooks'].index(WB)+1][0]
        WB = [self.pydata['Workbooks'][nWB-1] if type(WB) is int else WB][0]
        nWS = [WS if type(WS)  is int else self.pydata[nWB]['Worksheets'].index(WS)+1][0]
        WS = [self.pydata[nWB]['Worksheets'][nWS-1] if type(WS) is int else WS][0]
        if re.search(r'[A-Z][0-9]',repr(self.__COM.get_com_ranges_r1c1(R,C).Formula)):
            hasprecedents = True
        else:
            hasprecedents = False
        if hasprecedents:
            # find precedents
            while len(xladdress)>0:
                rangeobj = self.__COM.get_com_ranges_r1c1([xladdress[0][0]],[xladdress[0][1]])
                if rangeobj.HasArray:
                    xlobj = rangeobj.CurrentArray
                    arrayaddress = self.__COM.convert_r1c1A1(xlobj.Address)[0]
                    r = [int(r) for r in re.compile(r'(?<=R)\d+').findall(arrayaddress)]
                    c = [int(c) for c in re.compile(r'(?<=C)\d+').findall(arrayaddress)]
                    if [r,c] not in self.buffer[WB][WS]:
                        self.buffer[WB][WS].append([r,c])
                    arrayaddress = [r,c]
                else:
                    xlobj = rangeobj
                self.storedata(xlobj,nWS,nWB)
                xladdress.__delitem__(0)
                if xlobj.HasArray and xlobj.Cells.Count>1: # checks for array
                    # process in tree-branch inner loop precedents in object
                    for r in range(arrayaddress[0][0],arrayaddress[0][1]+1):
                        for c in range(arrayaddress[1][0],arrayaddress[1][1]+1):
                            if xladdress.count([r,c]):
                                xladdress.remove([r,c])
                    address = self.__COM.convert_r1c1A1(xlobj.Address)[0]
                    formula = self.__COM.convert_r1c1A1(xlobj.FormulaArray)[0]
                    if formula[0] == '=' and not formula.__contains__('ATG'):
                        # create node and append precedents to be processed
                        if self.creatxlnode(WB,WS,address,formula):
                            formula_array.append([formula,xlobj.Parent.Parent.Name,xlobj.Parent.Name])
                    else:
                        pass
                else:
                    # single precedent (no array)
                    address = self.__COM.convert_r1c1A1(xlobj.Address)[0]
                    formula = self.__COM.convert_r1c1A1(xlobj.Formula)[0]
                    if xlobj.HasFormula and formula[0] == '=' and not formula.__contains__('ATG'):
                        # create node and append precedents to be processed
                        if self.creatxlnode(WB,WS,address,formula):
                            formula_array.append([formula,xlobj.Parent.Parent.Name,xlobj.Parent.Name])
                    else:
                        pass
            if len(formula_array)>0:
                self.datalevel += 1 # go down one level
                for formula in formula_array:
                    self.xlstruct_constructor(*formula) # loop into precedents
                self.datalevel -= 1 # move back to main level
            else:
                pass
        else:
            self.storedata(self.__COM.get_com_ranges_r1c1(R,C),nWS,nWB)

    def processxlbuffer(self,item):
        """ inner method/function processxlbuffer
            processes new item from xlstruct_constructor into buffer
            if item
                is in buffer -> return []
                overalps buffered item -> return complement of item and buffered item, store complement into buffer
                is not in buffer -> return full item, store full item into buffer
        """
        item = [item]
        WB = item[0][0]; WS = item[0][1];
        WB, nWB, WS, nWS, flags = self.__get_WBWS(WB,WS)
        if flags[1]:
            self.buffer[WB][WS].append(item[0][2])
        else:
            numitem = len(item)-1
            while numitem <= len(item)-1:
                new_item = []
                sub_item = item[numitem]
                buffer0 = [buffered for buffered in self.buffer[WB][WS]]
                while new_item != sub_item and sub_item != []:
                    R = sub_item[2][0]; C = sub_item[2][1]
                    new_item = sub_item
                    while len(buffer0)>0:
                        rc = buffer0[0]
                        if min(R)>=min(rc[0]) and max(R)<=max(rc[0]) and min(C)>=min(rc[1]) and max(C)<=max(rc[1]):
                            sub_item = []
                            item[numitem] = sub_item
                            break
                        elif min(C)>=min(rc[1]) and max(C)<=max(rc[1]): # range withing buffered columns
                            if max(rc[0])>=min(R)>=min(rc[0]) and max(R)>max(rc[0]):
                                R[R.index(min(R))] = max(rc[0])+1
                                buffer0 = buffer0[1:len(buffer0)]+[buffer0[0]]
                            elif min(R)<min(rc[0]) and min(rc[0])<=max(R)<=max(rc[0]):
                                R[R.index(max(R))] = min(rc[0])-1
                                buffer0 = buffer0[1:len(buffer0)]+[buffer0[0]]
                            else:
                                buffer0.remove(rc)
                        elif min(R)>=min(rc[0]) and max(R)<=max(rc[0]): # range withing buffered lines
                            if max(rc[1])>=min(C)>=min(rc[1]) and max(C)>max(rc[1]):
                                C[C.index(min(C))] = max(rc[1])+1
                                buffer0 = buffer0[1:len(buffer0)]+[buffer0[0]]
                            elif min(C)<min(rc[1]) and min(rc[1])<=max(C)<=max(rc[1]):
                                C[C.index(max(C))] = min(rc[1])-1
                                buffer0 = buffer0[1:len(buffer0)]+[buffer0[0]]
                            else:
                                buffer0.remove(rc)
                        # item-breaker conditions below
                        elif max(rc[0])>=max(R)>=min(rc[0]) and max(rc[1])>=max(C)>=min(rc[1]): # range overalap rc first quadrant
                            item.append([WB,WS,[[min(rc[0]),R[R==max(R)]],[C[C==min(C)],min(rc[1])-1]]])
                            R[R.index(max(R))] = min(rc[0])-1
                        elif max(rc[0])>=max(R)>=min(rc[0]) and min(rc[1])<=min(C)<=max(rc[1]): # range overalap rc second quadrant
                            item.append([WB,WS,[[min(rc[0]),R[R==max(R)]],[max(rc[1])+1,C[C==max(C)]]]])
                            R[R.index(max(R))] = min(rc[0])-1
                        elif min(rc[0])<=min(R)<=max(rc[0]) and max(rc[1])>=max(C)>=min(rc[1]): # range overalap rc third quadrant
                            item.append([WB,WS,[[R[R==min(R)],max(rc[0])],[C[C==min(C)],min(rc[1])-1]]])
                            R[R.index(min(R))] = max(rc[0])+1
                        elif min(rc[0])<=min(R)<=max(rc[0]) and min(rc[1])<=min(C)<=max(rc[1]): # range overalap rc foruth quadrant
                            item.append([WB,WS,[[R[R==min(R)],max(rc[0])],[max(rc[1])+1,C[C==max(C)]]]])
                            R[R.index(min(R))] = max(rc[0])+1
                        else:
                            buffer0.remove(rc)
                        sub_item = [WB, WS, [R,C]]
                        item[numitem] = sub_item
                numitem += 1
            for sub_item in item:
                if sub_item != []:
                    self.buffer[WB][WS].append(sub_item[2])
        return item

    def xlstruct_constructor(self,formula,WB0,WS0):
        """ inner method/function (called recursivelly)
        Base function / method for scrapping the XL spreadsheet
            - Extracts R1C1 cell references from formula
            - Checks whether R1C1 references have been processed with processxlbuffer
            - Process items that have not been processed with processxlitem
        """
        if self.pynodes.__contains__(self.datalevel) is False:
            self.pynodes.__setitem__(self.datalevel,[])
        formula = formula.replace("'",'')
        p = re.compile(r'([\w\s\.]+(?=\]))?.?([\w\s\.]+(?=\!))?.?(R[\d]+C[\d]+[[:R]*[\d]*[C]*[\d]*]?)',re.M | re.UNICODE)
        itemlist = [list(item) for item in p.findall(formula)]
        while len(itemlist) > 0:
            item = []
            for iteritem in itemlist:
                if iteritem[1] == '':
                    iteritem[0] = WB0
                    iteritem[1] = WS0
                    item = iteritem
                    break
                elif iteritem[0] == '':
                    iteritem[0] = WB0
                    item = iteritem
                    break
                elif iteritem[1] == WS0 and iteritem[0] == WB0:
                    item = iteritem
                    break
                else:
                    pass
            if item == []:
                item = iteritem
            else:
                pass
            while itemlist.count(item)>0:
                itemlist.remove(item)
            R = [int(R) for R in re.compile(r'(?<=R)[\d]+',re.M).findall(item[2])]
            C = [int(C) for C in re.compile(r'(?<=C)[\d]+',re.M).findall(item[2])]
            item[2] = [R,C]
            item = self.processxlbuffer(item)
            if len(item) != 0:
                for sub_item in item:
                    if sub_item != []:
                        self.processxlitem(sub_item)        # break the item into those who do and do not belong to arryas
            else:
                pass

    def createpynodes(self,ofcell):
        """ START inner method createpynodes (ofcell as tree root)
            - Sets the pydata structure framework
            - Sets the pynodes tree structure root (createxlnode)
            - Calls the recursive method xlstruct_constructor as of the ofcell formula
            ... recursive process begins
        """
        # first derive actual WB-WS-Range from ofcell
        self.__COM.change_path(ofcell[0],ofcell[1])
        if re.compile('R\d+C\d+').search(unicode(ofcell[2])):
            ofcell[2] = self.__COM.convert_r1c1A1(ofcell[2])
        ofcellobj = self.__COM.Worksheet.Range(ofcell[2])
        print('Converting xl structure to py...')
        t0 = time.time()
        WB = self.__COM.Workbook.Name
        WS = self.__COM.Worksheet.Name
        bkp = self.__COM.change_path(WB,WS)
        formulaof = self.__COM.get_formulas_r1c1(ofcell[2]) # get formula cells/ranges
        # structure and dependences initializations
        self.pydata['Workbooks'].append(WB)
        self.pydata.__setitem__(1,{'Worksheets' : [WS]})
        self.pydata[1].__setitem__(1,{ofcellobj.Cells.Row : {ofcellobj.Cells.Column : ofcellobj.Cells.Value}})
        # preallocate OF cell in buffer
        self.buffer.__setitem__(WB,{WS:[[[ofcellobj.Cells.Row], [ofcellobj.Cells.Column]]]})
        self.creatxlnode(WB,WS,self.__COM.convert_r1c1A1(ofcell[2])[0],self.__COM.convert_r1c1A1(ofcellobj.Formula)[0])
        self.datalevel += 1
        self.xlstruct_constructor(formulaof[0],WB,WS)
        if self.__COM.Workbook.Name != WB or self.__COM.Worksheet.Name != WS:
            self.__COM.change_path([],[],bkp)
        self.__status__ = True
        print('Completed (elapsed time:{}s)'.format(time.time()-t0))

    def findpynodes(self,nWB,nWS,row,column):
        """ findpynodes
        params (4):
            nWB -> Workbook number in pydata as integer
            nWS -> Worksheet number in pydata[nWB] as integer
            row -> Row number as integer
            column -> Column number as integer
        returns a list of [level, nodes] that entails the parsed reference
        """
        nodes = []
        for level in self.pynodes:
            for node in self.pynodes[level]:
                if node['filename'] == self.pydata['Workbooks'][nWB-1]:
                    if node['sheet'] == self.pydata[nWB]['Worksheets'][nWS-1]:
                        if min(node['row'])<= row and max(node['row'])>= row:
                            if min(node['column'])<= column and max(node['column'])>= column:
                                nodes.append([level,self.pynodes[level].index(node)])
        return nodes

    def evalpynodes(self,level,node):
        """ method evalpynodes
        Updates the cell/array in  pydata to which the node makes reference
        by evaluating their converted formula using eval() and updates
        params(2):
            level as int
            node as int
        """
        self.__CalcHandler.execute(self.pyformulas[self.pynodes[level][node]['formulaindex']])

    def validatenodes(self):
        """ validatenodes
            evaluate all nodes and compare the difference between the evaluation results
            and actually pydata cell value. If the relative change is greater than 1e-10,
            the createcalcstruct inner method failed to bind node dependencies, returning False
            Otherwise, return True
            """
        flag = True
        print('Starting py structure validation...')
        t0 = time.time()
        outputs = []
        for level in self.pynodes:
            for node in range(0,len(self.pynodes[level])):
                flag_in, output = self.__CalcHandler.diagnose(self.pyformulas[self.pynodes[level][node]['formulaindex']])
                if not flag_in:
                    flag = False
                    outputs.append(output)
                    print('Validation failure - Level {} : Node {}'.format(level,node))
                    print(output)
        if flag is True:
            print('Validation successfully completed (elapsed time: {}s)'.format(time.time()-t0))
        else:
            print('Validation failed. Check parsed input.')
        return flag, outputs

    def circularrefwalker(self,index,index0,nodebuff):
        """ inner recursive method/function circularrefwalker
            recusively walk into all intranodes, hopping onto the next reference
            to see whether it gets to the starting point, thus comprising a circular reference
            circular references pairs are added to circularrefs list
            params(2):
                index as integer
                index0 as integer
                nodebuff as list (updated recursively)
        """
        ref0 = self.intranode[index][1]
        base = self.intranode[index0][0]
        for ref in ref0:
            if ref == base:
                if any([pair in self.circularrefs for pair in [[base, self.intranode[index][0]], [self.intranode[index][0], base]]]):
                    pass
                else:
                    self.circularrefs.append([base,self.intranode[index][0]])
            else:
                vec = [i[0] for i in self.intranode]
                if ref in vec:
                    nxtindex = vec.index(ref)
                    if nxtindex not in nodebuff:
                        nodebuff.append(nxtindex)
                        self.circularrefwalker(nxtindex,index0,nodebuff)
                    else:
                        pass

    def createintranodes(self):
        """ inner method createintranodes
            creates intranode list of dependencies between nodes from node dependencies
        """
        for level in self.pynodes:
            for node in range(0,len(self.pynodes[level])):
                for nWB in self.pynodes[level][node]['dependence']:
                    for nWS in self.pynodes[level][node]['dependence'][nWB]:
                        for item in self.pynodes[level][node]['dependence'][nWB][nWS]:
                            for innerlevel in self.pynodes:
                                for innernode in range(0,len(self.pynodes[innerlevel])):
                                    WB = self.pynodes[innerlevel][innernode]['filename']
                                    WS = self.pynodes[innerlevel][innernode]['sheet']
                                    nWBin = self.pydata['Workbooks'].index(WB)+1
                                    nWSin = self.pydata[nWBin]['Worksheets'].index(WS)+1
                                    if nWBin == nWB and nWSin == nWS:
                                        R = self.pynodes[innerlevel][innernode]['row']
                                        C = self.pynodes[innerlevel][innernode]['column']
                                        if (sum([min(R)<=min(item[0])<=max(R),min(R)<=max(item[0])<=max(R)])>0 and\
                                        sum([min(C)<=min(item[1])<=max(C),min(C)<=max(item[1])<=max(C)])>0) or\
                                        (sum([min(item[0])<=min(R)<=max(item[0]),min(item[0])<=max(R)<=max(item[0])])>0 and\
                                        sum([min(item[1])<=min(C)<=max(item[1]),min(item[1])<=max(C)<=max(item[1])])>0):
                                            vec = [i[0] == [level,node] for i in self.intranode]
                                            if not any(vec):
                                                self.intranode.append([[level,node],[[innerlevel,innernode]]])
                                            elif [level,node] not in self.intranode[vec.index(True)][1] and [level, node] != [innerlevel, innernode]:
                                                self.intranode[vec.index(True)][1].append([innerlevel,innernode])
                                            else:
                                                pass
                                        else:
                                            pass
                                    else:
                                        pass

    def hascircularref(self):
        """ inner method hascircularref
            Base method for node/intranode dependency generator
            creates intranode (createintranodes)
            checks for circular references (circularrefwalker)
        """
        print('Starting py circular reference verification...')
        self.circularrefs = []
        t0 = time.time()
        self.createintranodes()
        for index in range(0,len(self.intranode)):
            self.circularrefwalker(index,index,[index])
        if len(self.circularrefs)>0:
            print('{} circular references have been found (elapsed time: {}s)'.format(len(self.circularrefs),time.time()-t0))
            return True
        else:
            print('No circular references have been found (elapsed time: {}s)'.format(time.time()-t0))
            pass
        return False

    def nodeactivator(self,item):
        """ inner method/function nodeactivator
            for node activation
            - activates node
            - checks whether any other node depends on the activated one
            - activates dependent nodes (recursively)
            params (1) : item as pair of integer in list [level,node]
        """
        if item not in self.calcstruct:
            self.calcstruct.append(item)
            vec = [True if item in depedence[1] else False for depedence in self.intranode]
            while any(vec):
                index = vec.index(True)
                self.nodeactivator(dc(self.intranode[index][0]))
                vec[index] = False
        else:
            pass

    def createcalcstruct(self,inputsref,ref=True):
        """ inner method createcalcstruct

        """
        if ref:
            self.calcstruct = []
            self.intranode =[]
            self.circularrefs = []
            self.validatenodes()
            self.hascircularreferences = self.hascircularref()
            print('Beggining of py calculation structure creation...')
            t0 = time.time()
        else:
            pass
        for varyrange in inputsref:
            WB = varyrange[0]
            WS = varyrange[1]
            R = varyrange[2][0]
            C = varyrange[2][1]
            for level in self.pynodes:
               for node in range(0,len(self.pynodes[level])):
                    affected_node = []
                    flag = False
                    nodeobj = self.pynodes[level][node]
                    if nodeobj['dependence'].has_key(WB):
                        if nodeobj['dependence'][WB].has_key(WS):
                            for item in nodeobj['dependence'][WB][WS]:
                                if (sum([min(R)<=min(item[0])<=max(R),min(R)<=max(item[0])<=max(R)])>0 and\
                                sum([min(C)<=min(item[1])<=max(C),min(C)<=max(item[1])<=max(C)])>0) or\
                                (sum([min(item[0])<=min(R)<=max(item[0]),min(item[0])<=max(R)<=max(item[0])])>0 and\
                                sum([min(item[1])<=min(C)<=max(item[1]),min(item[1])<=max(C)<=max(item[1])])>0):
                                    self.nodeactivator([level,node])
                                    nWB = self.pydata['Workbooks'].index(nodeobj['filename'])+1
                                    nWS = self.pydata[nWB]['Worksheets'].index(nodeobj['sheet'])+1
                                    affected_node.append([nWB,nWS,[nodeobj['row'],nodeobj['column']]])
                                    flag = True
                                    break
            self.createcalcstruct(dc(affected_node),False)
        if ref:
            calcstruct = dc(self.calcstruct)
            intranode = dc(self.intranode)
            dependence = []
            vec = [j[0] for j in intranode]
            for i in calcstruct:
                if i in vec:
                    dependence.append(intranode[vec.index(i)][1])
                else:
                    dependence.append([])
            self.calcstruct = []
            i = 0
            while len(calcstruct)>0:
                flag = False
                flag_circ = False
                for j in range(0,len(calcstruct)):
                    if calcstruct[j] in dependence[i]:
                        if any([calcstruct[i] in [k[0] for k in self.circularrefs]]):
                            flag_circ = True
                            pass
                        else:
                            flag = True
                            break
                    else:
                        pass
                if flag:
                    i += 1
                elif not flag and flag_circ:
                    for k in self.circularrefs:
                        if calcstruct[i] == k[0]:
                            self.calcstruct.append(k[0])
                            self.calcstruct.append(k[1])
                            calcstruct.pop(i)
                            dependence.pop(i)
                            index = calcstruct.index(k[1])
                            calcstruct.pop(index)
                            dependence.pop(index)
                            break
                    i = 0
                else:
                    self.calcstruct.append(calcstruct[i])
                    calcstruct.pop(i)
                    dependence.pop(i)
                    i = 0
            print('Calculation structure created (elapsed time: {}s)'.format(time.time()-t0))
        else:
            pass

    def listconnect(self,itemlist,nWB,nWS):
        """ internal function listconnect
        params (3):
            itemlist as tuple or tuples of tuples
            nWB -> Workbook number as int
            nWS -> Worksheet number as int
        returns a modified version of itemlist with references converted from the
                original XL format to pointing to the converted structure
        """
        if type(itemlist) is tuple or type(itemlist[0]) is tuple:
            if type(itemlist[0]) is tuple:
                itemlist =  [[j for j in i] for i in itemlist]
            else:
                itemlist =  [[i for i in itemlist]]
        else:
            pass
        for i in range(0,len(itemlist)):
            itemlist[i][0] = [self.pydata['Workbooks'].index(itemlist[i][0])+1 if itemlist[i][0] != ([] or '') else nWB][0]
            itemlist[i][1] = [[self.pydata[itemlist[i][0]]['Worksheets'].index(self.__COM.Workbook.Sheets(itemlist[i][1]).Name)+1 if itemlist[i][1] is not int else itemlist[i][1]][0] if itemlist[i][1] != ([] or '') else nWS][0]
            if not re.compile('R\d+C\d+').search(unicode(itemlist[i][2])):
                itemlist[i][2] = self.__COM.convert_r1c1A1(itemlist[i][2])[0]
            itemlist[i][2] = [[int(r) for r in re.compile(r'(?<=R)\d+').findall(itemlist[i][2])],[int(c) for c in re.compile(r'(?<=C)\d+').findall(itemlist[i][2])]]
        return itemlist

    def evalstructure(self):
        """ evalstructure
            Evaluates the pydata structure based on the calcstruct defined by createcalcstruct.
            It also takes into consideration the existance of circular references in circularrefs.
                - nodes in circular loops are reevaluated until the absolute difference between calculated cells are lesser than 1e-3
        """
        for node in self.calcstruct:
            self.evalpynodes(node[0],node[1])
        if self.hascircularreferences:
            circularrefs = dc(self.circularrefs)
            base = []
            for i in range(0,len(circularrefs)):
                if circularrefs[i][0] not in base:
                    base.append(circularrefs[i][0])
                else:
                    pass
                if circularrefs[i][1] not in base:
                    base.append(circularrefs[i][1])
                else:
                    pass
            circdata = []; dataref = []; delta = []
            for i in base:
                nodeobj = self.pynodes[i[0]][i[1]]
                nWB = self.pydata['Workbooks'].index(nodeobj['filename'])+1
                nWS = self.pydata[nWB]['Worksheets'].index(nodeobj['sheet'])+1
                for R in range(min(nodeobj['row']),max(nodeobj['row'])+1):
                    for C in range(min(nodeobj['column']),max(nodeobj['column'])+1):
                        circdata.append(self.pydata[nWB][nWS][R][C])
                        dataref.append([nWB, nWS, R, C])
            for i in base:
                self.evalpynodes(i[0],i[1])
            for i in range(0,len(dataref)):
                nxtval = self.pydata[dataref[i][0]][dataref[i][1]][dataref[i][2]][dataref[i][3]]
                delta.append(np.divide(2*abs(circdata[i]-nxtval),(circdata[i]+nxtval)).tolist())
                circdata[i] = nxtval
            while max(delta)>0.001:
                print(max(delta))
                for i in base:
                    self.evalpynodes(i[0],i[1])
                for i in range(0,len(dataref)):
                    nxtval = self.pydata[dataref[i][0]][dataref[i][1]][dataref[i][2]][dataref[i][3]]
                    delta[i] = np.divide(2*abs(circdata[i]-nxtval),(circdata[i]+nxtval)).tolist()
                    circdata[i] = nxtval
        else:
            pass

    def set_io(self,inputsref,ofcell):
        """ set_io
        py solver handler:

        So far only single cell optimization has been included.
        Algorithm parameter allocation is still under development.

        Version 1.0 alpha
            - Single cell O.F.
        """
        if not self.__status__:
            self.createpynodes(ofcell) # create py structure
        else:
            pass
        # convert/link inputsref and ofcell xl references to py structure
        WB = self.__COM.Workbook.Name
        WS = self.__COM.Worksheet.Name
        nWB = self.pydata['Workbooks'].index(WB)+1
        nWS = self.pydata[nWB]['Worksheets'].index(WS)+1
        # link references to mapped xl cells in pydata
        self.iolib['ofcell'] = self.listconnect(ofcell,nWB,nWS)
        self.iolib['inputsref'] = self.listconnect(inputsref,nWB,nWS)
        # create calculation structure/sequence
        self.createcalcstruct(self.iolib['inputsref'])

    def Solver(self,solverparams):
        """ Solver
        This is a anectodal example on how to grab the Objective Function cell ('ofcell') value, modify variables in 'inputsref', recalculate the pydata structure and reevaluate 'ofcell'
        """
        # Algorithm Start
        # Flags are True when any addition to the pydata structure was mapped, when references have not been included during the object initialozation.
        ofvalue, flag_ofvalue = self.get_pyranges(self.iolib['ofcell'])
        seed, flag_seed = self.get_pyranges(self.iolib['inputs0ref'])
        stddev, flag_stddev = self.get_pyranges(self.iolib['boundsref'])
        # Checks wheter variables- and constraints-references are of the same length
        if inputsref_len.tolist() == inputs0ref_len.tolist() == boundsref_len.tolist():
            # if any flag is True, the following reevaluates the calculation structure
            if any([flag_seed, flag_ofvalue, flag_stddev]):
                self.createcalcstruct(self.iolib['inputsref'])
            else:
                pass
            ofvalue0 = self.get_pyranges(self.iolib['ofcell'])[0] # Gets the objective function value from 'ofcell'
            newval = 0 # newval for variables should be parsed from minimization algorithms constrained by 'boundsref'
            self.set_pyranges(self.iolib['inputsref'],newval)  # Varuables are update with the newval array
            self.evalstructure() # Structure is recalculated
            ofvalue = self.get_pyranges(self.iolib['ofcell'])[0] # new objective function values is assessed
            # Here a loop should proceed to minimize the 'ofcell' by varying 'inputsref' bounded within the 'boundsref' boundaries
        return
