"""
gfiXLSX.py

Class & supporting functions to genearate XLSX files based on GFI data.

This software uses an xlsx external library:
    xlsxwriter: http://xlsxwriter.readthedocs.org/

"""


import xlsxwriter
import types



def generateSumFunction(*args,**kwargs):
    _range = xlsxwriter.utility.xl_range(kwargs.get('startRow'),kwargs.get('col'),
            kwargs.get('row') -1,kwargs.get('col'))
    return '=SUM(' + _range + ')'



def generatePercentageFunction(*args,**kwargs):

    # dividend is uncl_r, divisor is curr_r
    _dividend = xlsxwriter.utility.xl_rowcol_to_cell(
            kwargs.get('row'), kwargs.get('col')-1 )
    _divisor = xlsxwriter.utility.xl_rowcol_to_cell(
            kwargs.get('row'), kwargs.get('col')-7 )

    return '=IF(%s=0,0,%s/%s)' % (_divisor,_dividend,_divisor)
    


class gfiSpreadsheet:
    filename = None
    sheetTitle='Exception Report'
    workbook = None
    worksheet = None
    formats = None
    workbookFormats = {}
    data = None
    header = None
    fieldOutline = None
    columnWidth = 8.5
    summaryRow=False
    zebraFormatting = False
    zebraField = False

    def __init__(self,**kwargs):
        if kwargs.get('filename'): self.filename = kwargs.get('filename')
        if kwargs.get('sheetTitle'): self.sheetTitle = kwargs.get('sheetTitle')
        if kwargs.get('data'): self.data = kwargs.get('data')
        if kwargs.get('header'): self.header = kwargs.get('header')
        if kwargs.get('fieldOutline'): self.fieldOutline = kwargs.get('fieldOutline')
        if kwargs.get('columnWidth'): self.columnWidth = kwargs.get('columnWidth')
        if kwargs.get('formats'): self.formats = kwargs.get('formats')
        if kwargs.get('summaryRow'): self.summaryRow= kwargs.get('summaryRow')
        if kwargs.get('zebraFormatting'): self.zebraFormatting = kwargs.get('zebraFormatting')
        if kwargs.get('zebraField'): self.zebraField = kwargs.get('zebraField')

        self.workbook = xlsxwriter.Workbook(self.filename)
        self.worksheet = self.workbook.add_worksheet(self.sheetTitle)


    def setCell(self,cell,data,style):
        pass

    def addFormats( self,formats ):
        [self.workbookFormats.update( {k:self.workbook.add_format( formats[k])}) for k in formats.keys()]

    def generateXLSX(self):
        row,col = 0,0
        if self.formats: self.addFormats(self.formats)
        else: return

        # output report title
        for name,format in self.header:
            if isinstance(name,(tuple,list)):
                self.worksheet.write_row(row,col,name,self.workbookFormats[ format ])
            else:
                self.worksheet.write(row,col,name,self.workbookFormats[ format ])
            row +=1
        
        row +=1
        _dataRowStart = row +1
        _numDataRows = len(self.data[self.data.keys()[0]])

        # output column titles
        col = 0
        for field,name,format,headerFormat,formula,highlightField,highlightValue,highlightFormat,zebraFormat in self.fieldOutline:
            self.worksheet.set_column(col,col,self.columnWidth) 
            self.worksheet.write(row,col,name,self.workbookFormats[headerFormat])
            col += 1

        # otuput data - row by row
        row += 1
        zebraFieldValue = None
        zebraOn = False  # flag for zebra formatting

        for r in range(0,_numDataRows):
            col = 0

            # alternate zebra formatting based on zebra field
            if self.zebraFormatting and (self.data[self.zebraField][r] != zebraFieldValue):
                zebraFieldValue = self.data[self.zebraField][r]
                zebraOn = True
            else:
                zebraOn = False

            for field,name,format,headerFormat,formula,highlightField,highlightValue,highlightFormat,zebraFormat in self.fieldOutline:

                # test type of formatting required for cell
                if highlightField and (highlightValue in self.data[highlightField][r]):
                    _formatName = highlightFormat
                else:
                    _formatName = format

                if zebraOn: _formatName += zebraFormat

                # test if data from query or empty field
                if type(field) == types.FunctionType:
                    self.worksheet.write_formula(row,col, field(row=row, col=col),self.workbookFormats[_formatName])
                else:
                    if field:
                        _data = self.data[field][r]
                    else: 
                        _data = ''
                    self.worksheet.write(row,col,_data,self.workbookFormats[_formatName])

                col += 1
            row += 1

        # output summary row
        if self.summaryRow:
            col = 0
            for field,name,format,headerFormat,formula,highlightField,highlightValue,highlightFormat,zebraFormat in self.fieldOutline:
                if formula:
                    self.worksheet.write_formula(row,col,formula(row=row,col=col,startRow=row-_numDataRows),
                            self.workbookFormats[headerFormat])
                else:
                    self.worksheet.write(row,col,'',self.workbookFormats[headerFormat])
                col += 1


    def close(self):
        self.workbook.close()






