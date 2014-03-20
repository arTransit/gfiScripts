"""
generateMSR.py

Generate GFI Monthly Route Summary Report
Command line usage: 
    generateMRSR.py -l locid(s) -y year -m month -c oracleCred -f xlsxName

This software uses two external libraries:
    cx_Oracle: http://cx-oracle.sourceforge.net/html/
    xlsxwriter: http://xlsxwriter.readthedocs.org/

"""


import sys
import cx_Oracle
import xlsxwriter
import argparse
import datetime
import calendar
import types

systemList = {
        1:"Victoria",
        2:"Langford",
        3:"Whistler",
        4:"Squamish",
        5:"Nanaimo",
        6:"Abbotsford",
        7:"Kelowna",
        8:"Kamloops",
        9:"Prince George",
        10:"Cowichan Valley",
        11:"Trail",
        12:"Comox",
        13:"Port Alberni",
        14:"Campbell River",
        15:"Powell River",
        16:"Sunshine",
        17:"Vernon",
        18:"Penticton",
        19:"Chilliwack",
        20:"Cranbrook",
        21:"Nelson",
        22:"Terrace",
        23:"Prince Rupert",
        24:"Kitimat",
        25:"Fort St. John"
    }


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
    

class GFIquery:
    """
    Manage GFI Oracle queries.

    Input: sql to be executed, and Oracle credentials in the form
    user/pass@db

    Output: class data and header arrays set if query successful.
    """

    sql = None
    credentials = None
    status = False
    headers = None
    data = {}

    def __init__(self,location,year,month,credentials):
        """
        Create new query object given query and connection string,
        execute, and store result in object variables if successful.
        """

        self.credentials = credentials

        try:
            connection = cx_Oracle.connect(credentials)
        except cx_Oracle.DatabaseError:
            connection.close()
            status = False
            return

        try:
            cursor = connection.cursor()
            print self.generateSQL(location,year,month)
            cursor.execute(self.generateSQL(location,year,month))
        except cx_Oracle.DatabaseError:
            connection.close()
            status = False
            return

        # get names in position 0 of description array
        self.headers = [i[0].lower() for i in cursor.description]
        for field in self.headers: self.data[field] = []

        for r in cursor:
            for field,value in zip(self.headers,r):
                self.data[field].append(value)
        connection.close()

        self.status = True

    def generateSQL(self,location,year,month):
        """
        Return SQL query using given location, year, and month attributes.
        """
        
        try:
            _location = ','.join([str(s) for s in location])
        except TypeError:
            _location = str(location)


        return (
            "SELECT 'Unknown' route,SUM(curr_r) curr_r,SUM(rdr_c) rdr_c,SUM(token_c) token_c, "
                "SUM(ticket_c) ticket_c, SUM(pass_c) pass_c,SUM(bill_c) bill_c, "
                "SUM(uncl_r) uncl_r,SUM(dump_c) dump_c, "
                "SUM(ttp1) ttp1, SUM(ttp2) ttp2, SUM(ttp3) ttp3, SUM(ttp4) ttp4, SUM(ttp5) ttp5, "
                "SUM(ttp6) ttp6, SUM(ttp7) ttp7, SUM(ttp8) ttp8, SUM(ttp9) ttp9, SUM(ttp10) ttp10, "
                "SUM(ttp11) ttp11, SUM(ttp12) ttp12, SUM(ttp13) ttp13, SUM(ttp14) ttp14, SUM(ttp15) ttp15, "
                "SUM(ttp16) ttp16, SUM(ttp17) ttp17, SUM(ttp18) ttp18, SUM(ttp19) ttp19, SUM(ttp20) ttp20, "
                "SUM(ttp21) ttp21, SUM(ttp22) ttp22, SUM(ttp23) ttp23, SUM(ttp24) ttp24, SUM(ttp25) ttp25, "
                "SUM(ttp26) ttp26, SUM(ttp27) ttp27, SUM(ttp28) ttp28, SUM(ttp29) ttp29, SUM(ttp30) ttp30, "
                "SUM(ttp31) ttp31, SUM(ttp32) ttp32, SUM(ttp33) ttp33, SUM(ttp34) ttp34, SUM(ttp35) ttp35, "
                "SUM(ttp36) ttp36, SUM(ttp37) ttp37, SUM(ttp38) ttp38, SUM(ttp39) ttp39, SUM(ttp40) ttp40, "
                "SUM(ttp41) ttp41, SUM(ttp42) ttp42, SUM(ttp43) ttp43, SUM(ttp44) ttp44, SUM(ttp45) ttp45, "
                "SUM(ttp46) ttp46, SUM(ttp47) ttp47, SUM(ttp48) ttp48, "
                "SUM(fare_c) fare_c, "
                "SUM(key1) key1, SUM(key2) key2, SUM(key3) key3, SUM(key4) key4, SUM(key5) key5, "
                "SUM(key6) key6, SUM(key7) key7, SUM(key8) key8, SUM(key9) key9, "
                "SUM(keyast) keyast, SUM(keya) keya, SUM(keyb) keyb, "
                "SUM(keyc) keyc, SUM(keyd) keyd "
            "FROM mrtesum "
            "WHERE mrtesum.loc_n in (%s) "
                "AND route =-3 "
                "AND mrtesum.mday = TO_DATE('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') "
            "UNION "
            "SELECT 'Other' route,SUM(curr_r) curr_r,SUM(rdr_c) rdr_c,SUM(token_c) token_c, "
                "SUM(ticket_c) ticket_c, SUM(pass_c) pass_c,SUM(bill_c) bill_c, "
                "SUM(uncl_r) uncl_r,SUM(dump_c) dump_c, "
                "SUM(ttp1) ttp1, SUM(ttp2) ttp2, SUM(ttp3) ttp3, SUM(ttp4) ttp4, SUM(ttp5) ttp5, "
                "SUM(ttp6) ttp6, SUM(ttp7) ttp7, SUM(ttp8) ttp8, SUM(ttp9) ttp9, SUM(ttp10) ttp10, "
                "SUM(ttp11) ttp11, SUM(ttp12) ttp12, SUM(ttp13) ttp13, SUM(ttp14) ttp14, SUM(ttp15) ttp15, "
                "SUM(ttp16) ttp16, SUM(ttp17) ttp17, SUM(ttp18) ttp18, SUM(ttp19) ttp19, SUM(ttp20) ttp20, "
                "SUM(ttp21) ttp21, SUM(ttp22) ttp22, SUM(ttp23) ttp23, SUM(ttp24) ttp24, SUM(ttp25) ttp25, "
                "SUM(ttp26) ttp26, SUM(ttp27) ttp27, SUM(ttp28) ttp28, SUM(ttp29) ttp29, SUM(ttp30) ttp30, "
                "SUM(ttp31) ttp31, SUM(ttp32) ttp32, SUM(ttp33) ttp33, SUM(ttp34) ttp34, SUM(ttp35) ttp35, "
                "SUM(ttp36) ttp36, SUM(ttp37) ttp37, SUM(ttp38) ttp38, SUM(ttp39) ttp39, SUM(ttp40) ttp40, "
                "SUM(ttp41) ttp41, SUM(ttp42) ttp42, SUM(ttp43) ttp43, SUM(ttp44) ttp44, SUM(ttp45) ttp45, "
                "SUM(ttp46) ttp46, SUM(ttp47) ttp47, SUM(ttp48) ttp48, "
                "SUM(fare_c) fare_c, "
                "SUM(key1) key1, SUM(key2) key2, SUM(key3) key3, SUM(key4) key4, SUM(key5) key5, "
                "SUM(key6) key6, SUM(key7) key7, SUM(key8) key8, SUM(key9) key9, "
                "SUM(keyast) keyast, SUM(keya) keya, SUM(keyb) keyb, "
                "SUM(keyc) keyc, SUM(keyd) keyd "
            "FROM mrtesum "
            "WHERE mrtesum.loc_n in (%s) "
                "AND route =-2 "
                "AND mrtesum.mday = TO_DATE('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') "
            "UNION "
            "SELECT LPAD(TO_CHAR(route),4,'0000') route,curr_r,rdr_c,token_c,ticket_c,pass_c, "
                "bill_c,uncl_r,dump_c, "
                "ttp1, ttp2, ttp3, ttp4, ttp5, ttp6, ttp7, ttp8, ttp9, ttp10, "
                "ttp11, ttp12, ttp13, ttp14, ttp15, ttp16, ttp17, ttp18, ttp19, ttp20, "
                "ttp21, ttp22, ttp23, ttp24, ttp25, ttp26, ttp27, ttp28, ttp29, ttp30, "
                "ttp31, ttp32, ttp33, ttp34, ttp35, ttp36, ttp37, ttp38, ttp39, ttp40, "
                "ttp41, ttp42, ttp43, ttp44, ttp45, ttp46, ttp47, ttp48, "
                "fare_c, "
                "key1, key2, key3, key4, key5, key6, key7, key8, key9, "
                "keyast, keya, keyb, keyc, keyd "
            "FROM mrtesum "
            "WHERE mrtesum.loc_n in (%s) "
                "AND route >=0 "
                "AND mrtesum.mday = TO_DATE('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') "
            "ORDER BY route"
            ) % (_location,str(year),str(month),_location,str(year),str(month),_location,str(year),str(month))
    

class gfiSpreadsheet:
    filename = None
    workbook = None
    worksheet = None
    formats = {}
    data = None
    header = None
    fieldOutline = None
    columnWidth = None

    def __init__(self,*args,**kwargs):
        if kwargs.get('filename'): self.filename = kwargs.get('filename')
        if kwargs.get('data'): self.data = kwargs.get('data')
        if kwargs.get('header'): self.header = kwargs.get('header')
        if kwargs.get('fieldOutline'): self.fieldOutline = kwargs.get('fieldOutline')
        if kwargs.get('columnWidth'): self.columnWidth = kwargs.get('columnWidth')

        self.workbook = xlsxwriter.Workbook(self.filename)
        self.worksheet = self.workbook.add_worksheet()
        self.addFormats( kwargs.get('formats') )


    def setCell(self,cell,data,style):
        pass

    def addFormats( self,_formats ):
        for f in _formats.keys():
            self.formats[ f ] = self.workbook.add_format( _formats[f] )

    def generateXLSX(self):
        row,col = 0,0

        # output report title
        for name,format in self.header:
            print "adding: %s" % name 
            self.worksheet.write(row,col,name,self.formats[ format ])
            row +=1

        # output column titles & data
        row +=1
        _dataRowStart = row +1
        _numDataRows = len(self.data[self.data.keys()[0]])
        for field,name,format,headerFormat,formula in self.fieldOutline:
            self.worksheet.set_column(col,col,self.columnWidth) 
            self.worksheet.write(row,col,name,self.formats['colTitle'])
            if type(field) == types.FunctionType:
                for r in range( row +1,row +_numDataRows +1):
                    self.worksheet.write_formula(r,col,
                            field(row=r, col=col), self.formats[format])
            else: 
                self.worksheet.write_column(row+1,col,self.data[field],self.formats[format])
            col += 1


        # output summary totals
        row += 1 + _numDataRows
        col = 0
        for field,name,format,headerFormat,formula in self.fieldOutline:
            if formula:
                _formula = formula(col=col,row=row,startRow=_dataRowStart)
                self.worksheet.write_formula(row,col,
                        formula(col=col,row=row,startRow=_dataRowStart),
                        self.formats[headerFormat])
            else:
                self.worksheet.write(row,col,'',self.formats[headerFormat])
            col +=1

    def close(self):
        self.workbook.close()



def getArgs():
    argsPsr = argparse.ArgumentParser(description='Generate GFI Monthly Summary Report')
    argsPsr.add_argument('-l','--location',required=True,nargs='+',type=int,help='id(s) of system/location')
    argsPsr.add_argument('-y','--year',required=True,type=int,help='eg 2014')
    argsPsr.add_argument('-m','--month',required=True,type=int,help='eg 12')
    argsPsr.add_argument('-f','--file',required=True,help='filename')
    argsPsr.add_argument('-c','--connection',required=True,help='eg user/pass@GFI')
    args = argsPsr.parse_args()
    args.error = False
    if (args.year > datetime.date.today().year) or (args.year < 2000):
        print "ERROR: year out of range (2000 - %d)" % datetime.date.today().year
        args.error = True
    if (args.month > 12) or (args.month < 1):
        print "ERROR: month out of range (1 - 12)" % datetime.date.today().year
        args.error = True
    return args


if __name__ == '__main__':
    args = getArgs()
    if args.error:
        print "Not completed."
        sys.exit(1)

    gfiQuery = GFIquery(args.location,args.year,args.month,args.connection)
    if not gfiQuery.status:
        print "DB error"
        sys.exit(1)

    _locationString = ''
    for l in args.location: _locationString += " / " + systemList[l]
    _locationString = _locationString[3:]

    reportHeader = [
            ['Monthly Route Summary Report','header'],
            [calendar.month_name[args.month]+" "+str(args.year),'subHeader'],
            [_locationString,'subHeader'] ] 

    cellFormats = {
        'header':{'bold':True,'font_size':14,'align':'left','align':'vcenter'},
        'subHeader':{'bold':True,'font_size':11,'align':'left','valign':'vcenter'},
        'colTitle':{'bold':True,'font_size':9,'align':'center',
            'valign':'vcenter','top':True,'bottom':True, 'bg_color':'#EEEEEE','text_wrap':True,
            'num_format':'#,###,##0'},
        'dataDecimal':{'font_size':9,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00'},
        'dataDecimalTitle':{'bold':True,'font_size':9,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00','top':True,'bottom':True, 'bg_color':'#EEEEEE'},
        'dataPercent':{'font_size':9,'align':'center','valign':'vcenter',
            'num_format':'0.00%'},
        'dataPercentTitle':{'bold':True,'font_size':9,'align':'center','valign':'vcenter',
            'num_format':'0.00%','top':True,'bottom':True, 'bg_color':'#EEEEEE'},
        'data':{'font_size':9,'align':'center','valign':'vcenter','num_format':0}
        }

    # structure of fields/columns:
    #   1 field name from SQL query
    #   2 col title used in worksheet
    #   3 format for data
    #   4 format for bottom summary
    #   5 function to generate summary function (sum, calculation, etc)

    fieldOutline = [
        ['route','Route','data','colTitle',None],
        ['curr_r','Current Revenue','dataDecimal','dataDecimalTitle',generateSumFunction],
        ['rdr_c','Ridership','data','colTitle',generateSumFunction],
        ['token_c','Token Count','data','colTitle',generateSumFunction],
        ['ticket_c','Ticket Count','data','colTitle',generateSumFunction],
        ['pass_c','Pass Count','data','colTitle',generateSumFunction],
        ['bill_c','Bill Count','data','colTitle',generateSumFunction],
        ['uncl_r','Unclassified Revenue','dataDecimal','dataDecimalTitle',generateSumFunction],
        [generatePercentageFunction,'%','dataPercent','dataPercentTitle',generatePercentageFunction],
        ['dump_c','Dump Count','data','colTitle',generateSumFunction],
        ['ttp1','TTP 1','data','colTitle',generateSumFunction],
        ['ttp2','TTP 2','data','colTitle',generateSumFunction],
        ['ttp3','TTP 3','data','colTitle',generateSumFunction],
        ['ttp4','TTP 4','data','colTitle',generateSumFunction],
        ['ttp5','TTP 5','data','colTitle',generateSumFunction],
        ['ttp6','TTP 6','data','colTitle',generateSumFunction],
        ['ttp7','TTP 7','data','colTitle',generateSumFunction],
        ['ttp8','TTP 8','data','colTitle',generateSumFunction],
        ['ttp9','TTP 9','data','colTitle',generateSumFunction],
        ['ttp10','TTP 10','data','colTitle',generateSumFunction],
        ['ttp11','TTP 11','data','colTitle',generateSumFunction],
        ['ttp12','TTP 12','data','colTitle',generateSumFunction],
        ['ttp13','TTP 13','data','colTitle',generateSumFunction],
        ['ttp14','TTP 14','data','colTitle',generateSumFunction],
        ['ttp15','TTP 15','data','colTitle',generateSumFunction],
        ['ttp16','TTP 16','data','colTitle',generateSumFunction],
        ['ttp17','TTP 17','data','colTitle',generateSumFunction],
        ['ttp18','TTP 18','data','colTitle',generateSumFunction],
        ['ttp19','TTP 19','data','colTitle',generateSumFunction],
        ['ttp20','TTP 20','data','colTitle',generateSumFunction],
        ['ttp21','TTP 21','data','colTitle',generateSumFunction],
        ['ttp22','TTP 22','data','colTitle',generateSumFunction],
        ['ttp23','TTP 23','data','colTitle',generateSumFunction],
        ['ttp24','TTP 24','data','colTitle',generateSumFunction],
        ['ttp25','TTP 25','data','colTitle',generateSumFunction],
        ['ttp26','TTP 26','data','colTitle',generateSumFunction],
        ['ttp27','TTP 27','data','colTitle',generateSumFunction],
        ['ttp28','TTP 28','data','colTitle',generateSumFunction],
        ['ttp29','TTP 29','data','colTitle',generateSumFunction],
        ['ttp30','TTP 30','data','colTitle',generateSumFunction],
        ['ttp31','TTP 31','data','colTitle',generateSumFunction],
        ['ttp32','TTP 32','data','colTitle',generateSumFunction],
        ['ttp33','TTP 33','data','colTitle',generateSumFunction],
        ['ttp34','TTP 34','data','colTitle',generateSumFunction],
        ['ttp35','TTP 35','data','colTitle',generateSumFunction],
        ['ttp36','TTP 36','data','colTitle',generateSumFunction],
        ['ttp37','TTP 37','data','colTitle',generateSumFunction],
        ['ttp38','TTP 38','data','colTitle',generateSumFunction],
        ['ttp39','TTP 39','data','colTitle',generateSumFunction],
        ['ttp40','TTP 40','data','colTitle',generateSumFunction],
        ['ttp41','TTP 41','data','colTitle',generateSumFunction],
        ['ttp42','TTP 42','data','colTitle',generateSumFunction],
        ['ttp43','TTP 43','data','colTitle',generateSumFunction],
        ['ttp44','TTP 44','data','colTitle',generateSumFunction],
        ['ttp45','TTP 45','data','colTitle',generateSumFunction],
        ['ttp46','TTP 46','data','colTitle',generateSumFunction],
        ['ttp47','TTP 47','data','colTitle',generateSumFunction],
        ['ttp48','TTP 48','data','colTitle',generateSumFunction],
        ['fare_c','Preset','data','colTitle',generateSumFunction],
        ['key1','KEY 1','data','colTitle',generateSumFunction],
        ['key2','KEY 2','data','colTitle',generateSumFunction],
        ['key3','KEY 3','data','colTitle',generateSumFunction],
        ['key4','KEY 4','data','colTitle',generateSumFunction],
        ['key5','KEY 5','data','colTitle',generateSumFunction],
        ['key6','KEY 6','data','colTitle',generateSumFunction],
        ['key7','KEY 7','data','colTitle',generateSumFunction],
        ['key8','KEY 8','data','colTitle',generateSumFunction],
        ['key9','KEY 9','data','colTitle',generateSumFunction],
        ['keyast','KEY *','data','colTitle',generateSumFunction],
        ['keya','KEY A','data','colTitle',generateSumFunction],
        ['keyb','KEY B','data','colTitle',generateSumFunction],
        ['keyc','KEY C','data','colTitle',generateSumFunction],
        ['keyd','KEY D','data','colTitle',generateSumFunction]
        ]

    xlsx = gfiSpreadsheet(filename=args.file,formats=cellFormats,
            header=reportHeader,columnWidth=8.5)
    xlsx.fieldOutline = fieldOutline
    xlsx.data = gfiQuery.data
    xlsx.generateXLSX()
    xlsx.close()

    print "Completed."
    sys.exit(0)




"""
Outstanding issues:

Add UNKNOWN data to table
build table from raw data/transactions
"""

