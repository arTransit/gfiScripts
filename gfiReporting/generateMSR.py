"""
generateMSR.py

Generate GFI Monthly Summary Report
Command line usage: 
    generateMSR.py -l locid(s) -y year -m month -c oracleCred -f xlsxName

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


cellFormats = {
    'header':{'bold':True,'font_size':14,'align':'left','align':'vcenter'},
    'subHeader':{'bold':True,'font_size':11,'align':'left','valign':'vcenter'},
    'rowTitles':{'bold':True,'font_size':9,'align':'center',
            'valign':'vcenter','top':True,'bottom':True,
            'bg_color':'#EEEEEE','text_wrap':True},
    'data':{'font_size':9,'align':'center','valign':'vcenter'} }

queryFields = (
    ['TDAY','Date'],
    ['BUS','Bus Probed'],
    ['CURR_R','Current Revenue'],
    ['RDR_C','Ridership'],
    ['TOKEN_C','Token Count'],
    ['TICKET_C','Ticket Count'],
    ['PASS_C','Pass Count'],
    ['BILL_C','Bill Count'],
    ['COIN_C','Coin Count'],
    ['UNCL_R','Unclassified Revenue'],
    ['CBXALM','Cahsbox Alarm'],
    ['BYPASS','Bypass Alarm']
    )

COLUMN_WIDTH=8.5

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
    data = []

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
            cursor.execute(self.generateSQL(location,year,month))
        except cx_Oracle.DatabaseError:
            connection.close()
            status = False
            return

        # get names in position 0 of description array
        self.headers = [i[0] for i in cursor.description]
        for r in cursor: self.data.append(r)
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
            "SELECT to_char(ml.tday, 'YYYY-MM-DD') serviceDate,"
            " count(ml.bus) bus_c,sum(ml.curr_r) curr_r,"
            " sum(ml.rdr_c) rdr_c, sum(ml.token_c) token_c,"
            " sum(ml.ticket_c) ticket_c,sum(ml.pass_c) pass_c,"
            " sum(ml.bill_c) bill_c,"
            " (sum(ml.nickel)+sum(ml.dime)+sum(ml.quarter)+sum(ml.half)+sum(ml.one)+sum(ml.two)) coin_c,"
            " sum(ml.uncl_r) uncl_r,sum(ml.cbxalm) cbxalm,"
            " sum(ml.bypass) bypass"
            " "
            "FROM ml"
            " "
            "WHERE ml.tday BETWEEN to_date('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS')"
            " AND last_day(to_date('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS'))"
            " AND ml.loc_n in ( %s )"
            " "
            "GROUP BY to_char(ml.tday, 'YYYY-MM-DD')"
            " "
            "ORDER BY to_char(ml.tday, 'YYYY-MM-DD')"
            ) % (str(year),str(month),str(year),str(month),_location)
    

class gfiSpreadsheet:
    filename = None
    workbook = None
    worksheet = None
    formats = {}
    data = None
    header = None
    rowTitles = None

    def __init__(self,*args,**kwargs):
        if kwargs.get('filename'): self.filename = kwargs.get('filename')
        if kwargs.get('data'): self.data = kwargs.get('data')
        if kwargs.get('header'): self.header = kwargs.get('header')
        if kwargs.get('rowTitles'): self.rowTitles = kwargs.get('rowTitles')

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
        for h in self.header:
            print "adding: %s" % h[0]
            self.worksheet.write(row,col,h[0],self.formats[ h[1] ])
            row +=1

        # output column titles
        row +=1
        for t in self.rowTitles:
            self.worksheet.set_column(col,col,COLUMN_WIDTH) 
            self.worksheet.write(row,col,t,self.formats['rowTitles'])
            col += 1

        # output data
        row += 1
        col = 0
        dataRowStart = row
        for d in self.data:
            for c in d:
                self.worksheet.write(row,col,c,self.formats['data'])
                col += 1
            row += 1
            col = 0


        # output summary totals
        self.worksheet.write(row,0,'',self.formats['rowTitles'])
        for col in range(1,len(self.rowTitles)):
            _formula='=SUM('+chr(col + ord('A')) + str(dataRowStart+1) + ':' + \
                    chr(col + ord('A')) + str(dataRowStart + len(self.data)) + ')'
            self.worksheet.write_formula(row,col,'{' + _formula + '}',self.formats['rowTitles'])

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
            ['Monthly Summary Report','header'],
            [calendar.month_name[args.month]+" "+str(args.year),'subHeader'],
            [_locationString,'subHeader'] ] 

    xlsx = gfiSpreadsheet(filename=args.file,formats=cellFormats,header=reportHeader)
    xlsx.rowTitles = [f[1] for f in queryFields]
    xlsx.data = gfiQuery.data
    xlsx.generateXLSX()
    xlsx.close()


    print "Completed."
    sys.exit(0)


