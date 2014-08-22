
import openpyxl
import sqlite3
import sys
import os
import argparse
import cx_Oracle
import gfiConfig
import xlsxwriter




"""
Read roster - get bus list per month

Data structures:
    Bus list per month
        associative array: busRoster[bus][month] = []
    Bus list per month from GFI last probe date VTC & LTC
        associative array: busRoster[bus][month] = [vtc,ltc]

read roster from sqlite
generate busRoster

read GFI data for vtc,ltc
update busRoster data

generate spreadsheet
"""




def getArgs():
    argsPsr = argparse.ArgumentParser(description='Generate bus roster/probe report')
    argsPsr.add_argument('-d','--database',required=True,help='input roster database')
    argsPsr.add_argument('-f','--file',required=True,help='output xlsx file')
    argsPsr.add_argument('-c','--connection',required=False,help='connection string')
    args = argsPsr.parse_args()
    return args

class rostertable:
    # self.rosterTable = {}
    #   rosterTable[bus][month] = [vtcProbe,ltcProbe]

    def __init__( self,database ):
        if os.path.isfile(database):
            self.database = database
            self.extractRoster()
        else:
            raise Exception('rostertable ERROR: db file does not exist')

    def extractRoster( self ):
        self.rosterTable = {}

        sqliteconnection = sqlite3.connect(self.database)
        sqliteconnection.row_factory = sqlite3.Row
        sqlitecursor = sqliteconnection.cursor()
        sqlitecursor.execute( (
                'select distinct busid,month '
                'from busroster '
                'left join systemmatch on busroster.accountingcode=systemmatch.accountingcode '
                'where systemmatch.gfiid in (1,2)'
                ))
        sqliterows = sqlitecursor.fetchall()
        for row in sqliterows:
            try:
                self.rosterTable[ int(row['busid']) ][ row['month']]=['not probed','not probed']
            except KeyError:
                self.rosterTable[ int(row['busid']) ]={row['month']:['not probed','not probed']}
        sqliteconnection.close()

    
    def retrieveProbes( self,credentials ):
        sql = (
            "select system,month,bus,to_char(max(ts),'YYYY-MM-DD HH24:MI') ts "
            "from ( "
                "select 'vtc' system, to_char(ts,'YYYY-MM') month,ts,bus "
                "from ml "
                "where loc_n=1 and ts >= to_date('2014-01-01','YYYY-MM-DD') "
                "union all "
                "select 'ltc' system, to_char(ts,'YYYY-MM') month,ts,bus "
                "from ml "
                "where loc_n=2 and ts >= to_date('2014-01-01','YYYY-MM-DD') "
            ") "
            "group by system,month,bus "
            "order by system,month,bus "
            )

        probeData = self.executeOracle( credentials,sql )
        for r in probeData:
            if r['bus'] in self.rosterTable.keys():  #test if bus on roster
                if r['month'] in self.rosterTable[ r['bus'] ].keys():  # test if month valid for bus
                    # update vtc/ltc probing timestamp
                    self.rosterTable[r['bus']][r['month']][{'vtc':0,'ltc':1}[r['system']]] = r['ts']
        

    def executeOracle( self, credentials,sql ):
        connection = cx_Oracle.connect(credentials)
        cursor = connection.cursor()
        cursor.execute(sql)
        headers = [i[0].lower() for i in cursor.description]
        return [dict(zip(headers, row)) for row in cursor]
        #connection.close()
        

    def generateSpreadsheet( self,filename,formats ):
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet('Probing History')
        workbookFormats ={}
        [workbookFormats.update( {k:workbook.add_format( formats[k])}) for k in formats.keys()]

        row,col = 0,0
        worksheet.write(row,col,'Conventional GFI Probing History',workbookFormats[ 'header' ])
        row +=1
        worksheet.write(row,col,'Victoria/Langford 2014',workbookFormats[ 'subHeader' ])
        row +=2
        worksheet.write(row,col,'This report compares the conventional roster with probe dates from the GFI database.',workbookFormats[ 'subHeader' ])
        row +=1
        worksheet.write(row,col,'Date/times indicate last probe data for each bus per month.',workbookFormats[ 'subHeader' ])
        row +=1
        worksheet.write(row,col,'Buses on the roster that have not been probed in that month.',workbookFormats[ 'headermed' ])
        row +=1
        worksheet.write(row,col,'Note blank cells indicate buses that are NOT on the roster for that month.',workbookFormats[ 'subHeader' ])
        row +=2

        worksheet.set_column(0,14,12.0) 
        worksheet.write(row,col,'Bus',workbookFormats[ 'colTitleRightB' ])
        col +=1
        for m in ['2014-'+('00'+str(m+1))[-2:] for m in range(0,12)]:
            worksheet.write(row,col, m,workbookFormats[ 'colTitle' ])
            worksheet.write(row,col+1, '',workbookFormats[ 'colTitleRightB' ])
            col +=2
        row +=1 
        col=1
        worksheet.write(row,0, '',workbookFormats[ 'colTitle' ])
        for m in range(0,12):
            worksheet.write(row,col, 'VTC',workbookFormats[ 'colTitle' ])
            worksheet.write(row,col+1, 'LTC',workbookFormats[ 'colTitleRightB' ])
            col +=2

        for bus in sorted(self.rosterTable.keys()):
            row +=1
            col =0
            worksheet.write(row,col,bus,workbookFormats[ 'dataRightB' ])
            col +=1
            for m in ['2014-'+('00'+str(m+1))[-2:] for m in range(0,12)]:
                if m in self.rosterTable[bus].keys():
                    if self.rosterTable[bus][m] == ['not probed','not probed']:
                        worksheet.write(row,col,self.rosterTable[bus][m][0],workbookFormats[ 'datamed' ])
                        worksheet.write(row,col+1,self.rosterTable[bus][m][1],workbookFormats[ 'datamedRightB' ])
                    else:
                        worksheet.write(row,col,self.rosterTable[bus][m][0],workbookFormats[ 'data' ])
                        worksheet.write(row,col+1,self.rosterTable[bus][m][1],workbookFormats[ 'dataRightB' ])
                else:
                    worksheet.write(row,col+1,'',workbookFormats[ 'dataRightB' ])
                col +=2

        workbook.close()

    def formatTestWrite(self, worksheet,row,col,value,formatFalse, formatTrue):
        if value == 'not probed':
            worksheet.write(row,col,value,formatTrue )
        else:
            worksheet.write(row,col,value,formatFalse )
    

if __name__ == '__main__':
    args = getArgs()
    
    rt = rostertable(args.database)
    rt.retrieveProbes(args.connection)
    rt.generateSpreadsheet(args.file, gfiConfig.cellFormats)


    print "Completed."
    sys.exit(0)


