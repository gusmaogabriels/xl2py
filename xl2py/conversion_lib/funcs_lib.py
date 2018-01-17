# -*- coding: utf-8 -*-
from __future__ import division, absolute_import, print_function

from .. import __author__, __version__
from ..core import np as np
import operator as op

class Funlib(object):
    """ Funlib
    This class must be loaded by xl2py.core.processor as Funlib to parse
    simple lambda-conversion references or customized functions (e.g. pyxl_error)
    """

    def __init__(self):
        # xl to py formulas conversion for eval()
        self.__author__ = __author__
        self.__version__ = __version__

        # xl to py formula conversion
        self.fun_database = {
                    'IF' : lambda args : [args[0]*args[1]+(abs(args[0]-1)*args[2])][0],\
                    'AVERAGE' : lambda args : np.average(args[0]),\
                    'STDEV.P' : lambda args : np.std(args[0]),\
                    'TRANSPOSE' : lambda args : np.transpose(args[0]),\
                    'ABS' : lambda args : np.abs(args[0]),\
                    'MMULT' : lambda args : np.dot(*args),\
                    'IFERROR' : lambda args : self.pyxl_error(*args),\
                    'SUM' : lambda args : np.sum(args[0]),\
                    'COUNT' : lambda args : np.size(args[0]),\
                    'SQRT' : lambda args : np.sqrt(args[0]),\
                    '^' : lambda args : np.power(*args),\
                    '<' : lambda args : np.float64(op.lt(*args)),\
                    '>' : lambda args : np.float64(op.gt(*args)),\
                    '<=' : lambda args : np.float64(op.le(*args)),\
                    '>=' : lambda args : np.float64(op.ge(*args)),\
                    '<>' : lambda args : np.float64(op.ne(*args)),\
                    '=' : lambda args : np.float64(op.eq(*args)),\
                    '+' : lambda args : np.add(*args),\
                    '-' : lambda args : np.subtract(*args),\
                    '/' : lambda args : np.divide(*args),\
                    '*' : lambda args : np.multiply(*args)
                    }

    # Further go all user-defined functions.

    # Which can be those that do not have a corresponding function in numpy and, therefore, has to be shaped verisimilarly.
    def pyxl_error(self,x,y):
        """ pyxl_error (substitute for XL fun IFERR)
        params (2):
            x as numeric or numeric array
            y as numeric
        returns x with nan and inf values converted to y
        """
        if any(np.isnan(x)+np.isinf(x)):
            x[np.isnan(x)+np.isinf(x)] = y[0][0]
        return x

Funlib_obj = Funlib()

class NumObj(object):

    def __init__(self,numeric):
        self.val = float(numeric)
        self.__hasoutput__ = False # True if sequence is to output values to the structure
        self.output = {'nWB':[],'nWS':[],'R':[],'C':[]}

    def  set_output(self,nWB,nWS,R,C):
        self.output['nWB'] = nWB
        self.output['nWS'] = nWS
        self.output['R'] = R
        self.output['C'] = C
        self.__hasoutput__ = True

    def __call__(self):
        val = np.array(self.val, ndmin=2)
        return val

class RefObj(object):

    def __init__(self,struct_ref,nWB,nWS,R,C):
        self.struct_ref = struct_ref
        self.shape = [np.ptp(R)+1,np.ptp(C)+1]
        self.ref = [nWB,nWS,R,C]
        self.__hasoutput__ = False # True if sequence is to output values to the structure
        self.output = {'nWB':[],'nWS':[],'R':[],'C':[]}

    def  set_output(self,nWB,nWS,R,C):
        self.output['nWB'] = nWB
        self.output['nWS'] = nWS
        self.output['R'] = R
        self.output['C'] = C
        self.__hasoutput__ = True

    def __call__(self):
        val = np.reshape([self.struct_ref[self.ref[0]][self.ref[1]][r][c] for r in range(min(self.ref[2]),max(self.ref[2])+1)\
        for c in range(min(self.ref[3]),max(self.ref[3])+1)],self.shape)
        return val

class FunObj(object):

    def __init__(self,funstr,params):
        self.funstr = funstr
        self.params = params
        self.__hasoutput__ = False # True if sequence is to output values to the structure
        self.output = {'nWB':[],'nWS':[],'R':[],'C':[]}

    def  set_output(self,nWB,nWS,R,C):
        self.output['nWB'] = nWB
        self.output['nWS'] = nWS
        self.output['R'] = R
        self.output['C'] = C
        self.__hasoutput__ = True

    def __call__(self):
        val = Funlib_obj.fun_database[self.funstr]([p() for p in self.params])
        return np.array(val,ndmin=2)

class CalcBlock(object):

    def __init__(self,calc_block,formula0):
        self.calc_block = calc_block
        # in case it starts with a sign operand
        if self.calc_block[0] in ['-','+']:
              self.calc_block[1].val = Funlib_obj.fun_database[self.calc_block[0]]([0,self.calc_block[1].val])
              self.calc_block = self.calc_block[1:]
        else:
              pass
        self.formula = formula0
        self.sequence = []
        self.__hasoutput__ = False # True if sequence is to output values to the structure
        self.__get_sequence()
        self.output = {'nWB':[],'nWS':[],'R':[],'C':[]}

    def __lambdify(self,operator,params):
        return Funlib_obj.fun_database[operator](params)

    def __rectify(self,vals,index0):
        return vals[0:index0+1]+vals[index0+3:]

    def __get_sequence(self):
        classes = [['^'],['/','*'],['-','+'],['>','<','<=','>=','<>','=']]
        reference = range(len(self.calc_block))
        operators = [self.calc_block[i] if i%2==1 else '' for i in range(len(self.calc_block))]
        for ops in classes:
            for o in ops:
                while operators.__contains__(o):
                    ref = operators.index(o)
                    references = [reference[ref-1],reference[ref+1]]
                    self.sequence += [[o, references, ref-1]]
                    for i in range(ref+2,len(reference)):
                        reference[i] -= 2
                    reference = reference[0:ref] + reference[ref+2:]
                    operators = operators[0:ref] + operators[ref+2:]

    def  set_output(self,nWB,nWS,R,C):
        self.output['nWB'] = nWB
        self.output['nWS'] = nWS
        self.output['R'] = R
        self.output['C'] = C
        self.__hasoutput__ = True

    def __call__(self):
        if len(self.sequence)>0:
            vals = [c for c in self.calc_block] # create a calc-block copy to place evaulated objects
            for seq in self.sequence:
                params = []
                for s in seq[1]:
                    if hasattr(vals[s],'__call__'):
                        params.append(vals[s]())
                    else:
                        params.append(vals[s])
                vals[seq[2]] = self.__lambdify(seq[0],params)
                vals = self.__rectify(vals,seq[2])
        else:
            vals = self.calc_block
        vals = vals[0]
        return vals

class CalcHandler(object):

    def __init__(self,struct_ref,diagnose_threshold = 1e-10):
        self.struct_ref = struct_ref
        self.diagnose_threshold = diagnose_threshold

    def diagnose(self,obj):
        if obj.__hasoutput__:
            value = obj.__call__()
            output = []
            flag = True
            for r in range(min(obj.output['R']),max(obj.output['R'])+1):
                for c in range(min(obj.output['C']),max(obj.output['C'])+1):
                    v0 = self.struct_ref[obj.output['nWB']][obj.output['nWS']][r][c]
                    vf = value[r-min(obj.output['R'])][c-min(obj.output['C'])]
                    if np.divide(abs(v0 - vf),v0)>self.diagnose_threshold:
                        flag = False
                        print('Original:', v0, 'Calculated:', vf)
                        output.append([obj.output['nWB'],obj.output['nWS'],r,c,v0,vf,np.divide(abs(v0 - vf),v0)])
            return flag, output
        else:
            raise Exception("Object is not output'able'.")

    def execute(self,obj):
        if obj.__hasoutput__:
            value = obj.__call__()
            for r in range(min(obj.output['R']),max(obj.output['R'])+1):
                for c in range(min(obj.output['C']),max(obj.output['C'])+1):
                    self.struct_ref[obj.output['nWB']][obj.output['nWS']][r][c] = float(value[r-min(obj.output['R'])][c-min(obj.output['C'])])
        else:
            raise Exception("Object is not output'able'.")







