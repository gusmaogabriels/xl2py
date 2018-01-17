# -*- coding: utf-8 -*-
from __future__ import division, absolute_import, print_function

from .. import __author__, __version__
from ..com_handlers import x32, pythoncom
from ..core import np, re

class xlcom(object):

    def __init__(self):
        self.__author__ = __author__
        self.__version__ = __version__
        self.path = ''
        self.password = ''
        self.ExcelObject = []
        self.__status__ = False
        self.Workbook = []
        self.Worksheet = []

    def xlclose(self):
        """ xlclose function
        xlcloses XL process attached to XL2py object
        """
        if self.__status__ is True:
            self.ExcelObject.Workbooks(self.Workbook.Name).Close(SaveChanges=False)
            #self.Workbook.xlclose(SaveChanges=False)
            self.Workbook = []
            del self.ExcelObject
            self.ExcelObject = []
        else:
            print('excel object: excel file not loaded.')

    def xlopen(self,path,password):
        """ xlopen function
        Pulls up the XL process from which to retrieve its structure
        """
        try:
            """
            This was used as option to always xlopen another instance of a file
            in case it had been already open
            ---
            try:
                self.ExcelObject = x32.GetActiveObject('Excel.Application')
                for WB in self.ExcelObject.Workbooks:
                    if re.search(WB.Name, self.path):
                        self.Workbook = WB
            except:
            """
            self.path = path
            self.password = password
            self.ExcelObject = x32.Dispatch('Excel.Application')
            self.Workbook = self.ExcelObject.Workbooks.Open(self.path,Password=self.password)
            self.Worksheet = self.Workbook.Sheets(1)
            self.Worksheet.Activate()
            self.Worksheet.Unprotect(self.password)
            self.ExcelObject.Visible = False
            self.ExcelObject.ScreenUpdating = False
            self.__status__ = True
            for i in self.Workbook.Sheets:
                i.Unprotect(self.password)
        except pythoncom.com_error as e:
            print(e)

    def set_screen_updating(self):
        """ set_screen_updating
        Toggles the screen update option on/off
        When on, each cell modification on the XL process is shown (slow!)
        """
        if self.ExcelObject.ScreenUpdating is True:
            self.ExcelObject.ScreenUpdating = False
            print('ScreenUpdating deactivated.')
        else:
            self.ExcelObject.ScreenUpdating = True
            print('ScreenUpdating activated.')

    def set_sheet(self,sheetnum):
        """ set_sheet
        params (1) -> sheetnum as numeric
        changes the XL process active sheet
        """
        self.Worksheet = self.Workbook.Sheets(sheetnum)
        self.Workbook.Sheets(sheetnum).Activate()

    def xlrange(self,xlrange): # shall be deprecated
            """ xlrange
            params (1) -> xlrange as string
            returns a XL object of the parsed xlrange
            """
            self.Range = self.Worksheet.Range(xlrange)
            return self.Worksheet.Range(xlrange)

    def set_ranges(self,xlranges,values):
        """ set_ranges
        changes XL process cells referenced by xlranges in the current Workbook-Worksheet
        to values specified (xlranges references and values must be of the same size)
        params (2):
            xlranges as string reference (either R1C1 or A1)
            values as numeric/string
        """
        datalength = self.rangelength(xlranges)
        if  datalength != len(values):
            print('xlranges ({}) and values ({}) length should be the same.'.format(len(xlranges),len(values)))
        else:
            for i in range(0,len(xlranges)):
                datalength = self.rangelength(xlranges[i])
                self.Worksheet.Range(xlranges[i]).Value = np.reshape(values[range(0,datalength)],[len(range(0,datalength)),1])
                values = values[range(datalength,len(values))]

    def get_ranges(self,xlranges,sheetnum=[]):
        """ get_ranges
        retrieves cell values referenced by xlranges from the XL process sheet = sheetnum
        params (2):
            xlranges as single string or list of string R1C1- or A1-tye XL references
            sheenum (optional): XL sheet from which data is retrieved. If none, current is taken.
        returns list of tuples of (ranges, retrieved values)
        """
        if len(sheetnum)==0:
            sheetnum = self.Worksheet.Index
        if type(xlranges) is not list:
            xlranges = [xlranges]
        ranges = np.array([],ndmin=2).transpose()
        sheet0 = self.Worksheet.Index
        self.set_sheet(sheetnum)
        for i in range(0,len(xlranges)):
            if re.compile('R\d+C\d+').search(xlranges[i]):
                xlranges[i] = self.convert_r1c1A1(xlranges[i])[0]
            val = np.array([self.Worksheet.Range(xlranges[i]).Value]).flatten()
            val = val.reshape([len(val),1])
            ranges = np.concatenate((ranges,val), axis=0)
        self.set_sheet(sheet0)
        return ranges

    def rangelength(self,xlranges):
        """ rangelength
        retrieves cell values referenced by xlranges from the XL process sheet = sheetnum
        params (1): xlranges as single string or list of string R1C1- or A1-tye XL references
        returns the total length of xlranges
        """
        if type(xlranges) is not list:
            xlranges = [xlranges]
        length = 0
        for i in range(0,len(xlranges)):
            if re.compile('R\d+C\d+').search(xlranges[i]):
                xlranges[i] = self.convert_r1c1A1(xlranges[i])[0]
            length += len(self.Worksheet.Range(xlranges[i]))
        return length

    def get_formulas(self,xlranges):
        """ get_formulas
        params (1): xlranges as string or list of strings in R1C1 or A1-type XL reference
        returns formula, formulas (if more than one ecxists) or array formula
                for each reference in xlranges
        """
        if type(xlranges) is not list:
            xlranges = [xlranges]
        formulas = []
        for i in range(0,len(xlranges)):
            if re.compile('R\d+C\d+').search(xlranges[i]):
                xlranges[i] = self.convert_r1c1A1(xlranges[i])[0]
            if self.Worksheet.Range(xlranges[i]).HasArray:
                formulas.append([unicode(self.Worksheet.Range(xlranges[i]).FormulaArray)])
            else:
                innerformulas = []
                if type(xlranges[i]) is not list:
                    xlranges[i] = [xlranges[i]]
                for cell in xlranges[i]:
                    innerformulas.append(unicode(self.Worksheet.Range(cell).Formula))
                formulas.append(innerformulas)
        return formulas

    def get_types(self,xlranges):
        """ get_types
        params (1): xlranges as string or list of strings in R1C1 or A1-type XL reference
        """
        if type(xlranges) is not list:
            xlranges = [xlranges]
        types = np.array([],ndmin=2).transpose()
        for i in range(0,len(xlranges)):
            if re.compile('R\d+C\d+').search(xlranges[i]):
                xlranges[i] = self.convert_r1c1A1(xlranges[i])[0]
            types = np.concatenate((types,np.array(self.get_ranges(xlranges[i]).dtype.kind,ndmin=2).transpose()),axis=0)
        return types

    def convert_r1c1A1(self,xlranges):
        """ convert_r1c1A1
        converts xlranges from R1C1 to A1 or A1 to R1C1 reference style having the first reference
        in xlranges (if it be a list) as base type if mixed-type xlranges be parsed
        params (1): xlranges as string or list of strings in R1C1 or A1-type XL reference
        return A1 from R1C1 or R1C1 from A1 references or formulas parsed as xlranges
        """
        if type(xlranges) is not list:
            xlranges = [xlranges]
        for i in range(0,len(xlranges)):
            if re.compile('R\d+C\d+').search(unicode(xlranges[0])):
                try:
                    xlranges[i] = self.ExcelObject.ConvertFormula(Formula=unicode(xlranges[i].encode('utf-8'),'utf-8'),FromReferenceStyle=0,ToReferenceStyle=1,ToAbsolute=1)
                except:
                    pass
            else:
                try:
                    xlranges[i] = self.ExcelObject.ConvertFormula(Formula=unicode(xlranges[i].encode('utf-8'),'utf-8'),FromReferenceStyle=1,ToReferenceStyle=0,ToAbsolute=1)
                except:
                    pass
        return xlranges

    def get_formulas_r1c1(self,xlranges):
        """ get_formulas_r1c1
        converts formulas in xlranges from R1C1 to A1 or A1 to R1C1 reference style having the first formulas
        from the reference in xlranges (if it be a list) as base type if mixed-type be parsed
        params (1): xlranges as string or list of strings in R1C1 or A1-type XL reference
        return formulas in xlranges converted from A1 to R1C1 or R1C1 from A1 styles.
        """
        formulas = self.get_formulas(xlranges)
        for i in range(0,len(formulas)):
            formulas[i] = self.convert_r1c1A1(formulas[i])[0]
        return formulas

    def dim_ranges(self,xlranges):
        """ dim_ranges
        params (1): xlranges as string or list of strings in R1C1 or A1-type XL reference
        returns a list of pairs number of rows, number of columns for each reference in xlranges
            e.g. [[[2],[4]],[[3],[1]]] -> two references, the first with 2 rows and 4 columns and 3 rows and 1 column
        """
        if type(xlranges) is not list:
            xlranges = [xlranges]
        dim = np.array([[],[]],ndmin=2).transpose()
        for i in range(0,len(xlranges)):
            if re.compile('R\d+C\d+').search(xlranges[i]):
                xlranges[i] = self.convert_r1c1A1(xlranges[i])[0]
            dim = np.concatenate((dim,np.transpose(np.array([[int(self.Worksheet.Range(xlranges[i]).Rows.Count)],[int(self.Worksheet.Range(xlranges[i]).Columns.Count)]],ndmin=2))),axis=0)
        dim = np.reshape(dim,[len(dim),len(dim[0])])
        return dim

    def get_com_ranges_r1c1(self,rows,columns):
        """ get_com_ranges_r1c1
        Gets the COM range defined by rows and columns in the actual Workbook-Worksheet
        params (2):
            rows as list containing initial and final rows -> e.g. [1,3] row 1 to 3
            columns as list in the same ways as rows
        returns a COM object of the range over which rows and columns span
        """
        if len(rows)==1:
            rows.append(rows[0])
        if len(columns)==1:
            columns.append(columns[0])
        return self.Worksheet.Range(self.Worksheet.Cells(rows[0],columns[0]),self.Worksheet.Cells(rows[1],columns[1]))

    def change_path(self,WB=[],WS=[],bkp=[]):
        """ change_path
        create bkp of the current XL objects and change the active Workbook and/or Worksheet if not the current (returns bkp)
        if WB = []  and WS = [] while bkp is parsed, XL reference objects are replaced (returns nothing)
        by those included in bkp
        params (3):
            WB -> Workbook name as string
            WS -> Worksheet name as string
            bkp -> tuple/list of XL objects (e.g. [ExcelObject, Workbook, Worksheet])
        """
        if WS != ([] and ''):
            bkp = [self.ExcelObject, self.Workbook, self.Worksheet]
            if all([WB != i for i in (self.Workbook.Name, [], '')]):
                ExcelObject = x32.gencache.EnsureDispatch('Excel.Application')
                Workbook = ExcelObject.Workbooks.Open(WB)
                Worksheet = Workbook.Sheets(WS)
                Workbook.Sheets(WS).Activate()
                self.ExcelObject = ExcelObject
                self.ExcelObject.ScreenUpdating = False
                self.Workbook = Workbook
                self.Worksheet = Worksheet
                self.Worksheet.Unprotect(self.password)
            if all([WS != i for i in (self.Worksheet.Name, self.Worksheet.Index)]):
                self.Worksheet = self.Workbook.Sheets(WS)
                self.Workbook.Sheets(WS).Activate()
                self.Worksheet.Unprotect(self.password)
            return bkp
        elif bkp != ([] and ''):
            try:
                self.ExcelObject = bkp[0]
                self.Workbook = bkp[1]
                self.Worksheet = bkp[2]
            except:
                pass
        else:
            pass
