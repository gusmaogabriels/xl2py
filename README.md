**xl2py**
=========

Copyright © 2015 - Gabriel Sabença Gusmão

[![linkedin](https://static.licdn.com/scds/common/u/img/webpromo/btn_viewmy_160x25.png)](https://br.linkedin.com/pub/gabriel-saben%C3%A7a-gusm%C3%A3o/115/aa6/aa8)

[![license](https://img.shields.io/pypi/l/xl2py.svg)](./LICENSE.md)
[![pypi version](https://img.shields.io/pypi/v/xl2py.svg)](https://pypi.python.org/pypi/xl2py)

An Excel (XL) 2 Python (Py) structure retriever for optimization. Convert the I/O of XL files into Python.

----------------
**Description**
----------------

*Convert an XL structure to Py and use any minimization algorithm of your will*

*Now, with object-oriented formulas.*

The current project makes use of the XL COM interface (win32com library) to:

  1. Read an objective function cell
  2. Recursively build its dependent structure as of its formula
      * The XL structure is represented in Py as a dict() object
      * The structure is referenced to as: 
          * `dictobj[Workbook number as int][Worksheet number as int][Row as int][Column as int]`
      * Whereby it handles:
          * multi-XL workbook/worksheet references
          * single worksheet multirange retrieval
  3. XL cell formulas are translated to **object oriented** calculation blocks (no more `evals` as of this update).
  4. The calculation structure is determinded by cell-dependency trees, which have been already stored during the conversion (2)
      - Handling of circular references
  
  Ongoing development: A simple evolutionary algorithm that runs based off the abovementioned structure.

----------------
**Features**
----------------

  - **Conversion Library**

    The following XL functions can be currently handled by xl2py.
    xl2py is capable of undertaking **single-cells**, **arrays** and **array/matrix operations**

        1. Standard operators: \+, \-, \/, \*, \^
        2. Logical operators: \<, \>, \<=, \>=, \<>, \=
        3. IF
        4. AVERAGE
        5. STDEV.P
        6. TRANSPOSE
        7. ABS
        8. MMULT
        9. IFERROR
        10. SUM
        11. COUNT
        12. SQRT

----------------------------------
**Tackled in the latest update**
----------------------------------

   1. No more `evals` -> formulas are object oriented (Calculation-, Formula- and Reference- and Numeric-Blocks)
   2. by-operand handling
    *Over the latest update development, by-operand handling of formulas took place of RPN (reverse-polish notation). For additional details, viz. github repository*

----------------
**On the way**
----------------

  1. Object serialization
  2. CVS outputs

----------------
**Instructions**
----------------

  - **Installation**

       `pip install xl2py==version_no`

  - **Example**: I/O object creation
  
        import xl2py 
        
        Builder = xl2py.builder() # creates a xl2py builder object
        # place the path of your XL file 
        path = r'C:\\User\\DEFAULT\\WHATEVER\\...' 
        # define your XL file password (if it exists)
        pwd = 'password' 
        # opens up a XL COM interface and attach it to the Builder object
        Builder.connect_com(path,pwd) 
        # declare your input cell/range references
        inputs = xl2py.xlref(<Workbook str or int>, \
            <Worksheet int>, <A1- or R1C1-type XL references>)
        # inputs include other inputs to the xlref object
        inputs += xl2py.xlref(<str or int>, <int>, <str>)
        # output must be a single cell reference
        output = xl2py.xlref(<str or int>, <int>, <str>) 
        # Now you are all set. You shall translate the XL structure to python.
        Builder.set_structure(inputs,output)
        # If you want to change the input cell/range values...
        # vals must be of the shape of the inputs 
        # and must be parsed as a list of lists or numpy arrays
        Builder.set_input_values(vals) 
        # grab the output (objective fun) value as numpy array
        output_val = Builder.get_output_value() # Grab the new output value
