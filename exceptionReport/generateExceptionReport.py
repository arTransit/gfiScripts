"""
generateMSR.py

Generate GFI Monthly Exception report
Command line usage: 
    generateExceptionReport.py -l locid(s) -y year -m month -c oracleCred -f xlsxName

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
import gfiConfig



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

        Note that types refer to payments.  A driver or route change (type 104 & 106)
        will put the new route id or driver id in the N field of TR table.  The operator
        may enter one or more incorrect ids before entering the correct route & driver id.
        These should be ignored.  Only incorrect route and driver ids that are in effect
        when a payment is made are used.
        """
        
        try:
            _location = ','.join([str(s) for s in location])
        except TypeError:
            _location = str(location)


        return (
            "select bus,probetime,eventtime,route,drv,curr_r,rdr_c,wm_concat(issue) as issue "
            "from ( "
                "select  "
                    "ml.bus,  "
                    "TO_CHAR(ml.ts,'YYYY-MM-DD HH24:MI') probetime,  "
                    "TO_CHAR(ev.ts,'YYYY-MM-DD HH24:MI') eventtime,  "
                    "ev.route,ev.drv,ev.curr_r,ev.rdr_c,'route' issue  "
                "from ml left join ev on ml.loc_n=ev.loc_n and ml.id=ev.id  "
                "where   "
                    "ml.loc_n in ( %s ) and  "
                    "ev.ts between to_date('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') and last_day(to_date('%s-%s-01 23:59:59', 'YYYY-MM-DD HH24:MI:SS')) and  "
                    "ev.route not in (select route from rtelst where loc_n in (%s) ) and  "
                    "((ev.curr_r >0) or (ev.rdr_c >0)) "
                "union "
                "select  "
                    "ml.bus,  "
                    "TO_CHAR(ml.ts,'YYYY-MM-DD HH24:MI') probetime, "
                    "TO_CHAR(ev.ts,'YYYY-MM-DD HH24:MI') eventtime, "
                    "ev.route,ev.drv,ev.curr_r,ev.rdr_c,'driver' issue  "
                "from ml left join ev on ml.loc_n=ev.loc_n and ml.id=ev.id "
                "where  "
                    "ml.loc_n in ( %s ) and  "
                    "ev.ts between to_date('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') and last_day(to_date('%s-%s-01 23:59:59', 'YYYY-MM-DD HH24:MI:SS')) and  "
                    "ev.drv not in (select drv from drvlst where loc_n in (%s) ) and "
                    "((ev.curr_r >0) or (ev.rdr_c >0))  "
            ") "
            "group by bus,probetime,eventtime,route,drv,curr_r,rdr_c "
            "order by bus,eventtime "
            ) % (
                    _location,str(year),str(month),str(year),str(month),_location,
                    _location,str(year),str(month),str(year),str(month),_location )
    

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
        for field,name,format,headerFormat,formula,highlightField,highlightValue,highlightFormat in self.fieldOutline:
            self.worksheet.set_column(col,col,self.columnWidth) 
            self.worksheet.write(row,col,name,self.formats['colTitle'])
            if field: 
                for r in range(0,_numDataRows):
                    if highlightField:
                        if highlightValue in self.data[highlightField][r]:
                            self.worksheet.write(row +r +1,col,self.data[field][r],self.formats[highlightFormat])
                        else:
                            self.worksheet.write(row +r +1,col,self.data[field][r],self.formats[format])
                    else:
                        self.worksheet.write(row +r +1,col,self.data[field][r],self.formats[format])

            col += 1

        """
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
        """

    def close(self):
        self.workbook.close()



def getArgs():
    argsPsr = argparse.ArgumentParser(description='Generate GFI Monthly Exception Report')
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
    for l in args.location: _locationString += " / " + gfiConfig.systemList[l]
    _locationString = _locationString[3:]

    reportHeader = [
            ['Monthly Exception Report','header'],
            [calendar.month_name[args.month]+" "+str(args.year),'subHeader'],
            [_locationString,'subHeader'] ] 

    cellFormats = gfiConfig.cellFormats
    fieldOutline = gfiConfig.exceptionReportFieldOutline 

    xlsx = gfiSpreadsheet(filename=args.file,formats=cellFormats,
            header=reportHeader,columnWidth=12)
    xlsx.fieldOutline = fieldOutline
    xlsx.data = gfiQuery.data
    xlsx.generateXLSX()
    xlsx.close()

    print "Completed."
    sys.exit(0)



"""
Outstanding issues:


"""

