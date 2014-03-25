"""
generateExceptionReport.py

Generate GFI Monthly Exception report
Command line usage: 
    generateExceptionReport.py -l locid(s) -y year -m month -c user/pass@db -f xlsxName

"""


import sys
import argparse
import datetime
import gfiConfig
import gfiQuery
import gfiXLSX



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
        print "ERROR: month out of range (1 - 12)"
        args.error = True
    return args


def createReport(location,year,month,filename,connection):
    gq = gfiQuery.GFIquery(connection, 
            gfiConfig.exceptionReportSQL(location,year,month) )
    gq.execute()
    if not gq.status:
        print "DB error"
        sys.exit(1)

    xlsx = gfiXLSX.gfiSpreadsheet(filename=filename,
            header=gfiConfig.exceptionReportHeader(location,year,month),
            columnWidth=gfiConfig.exceptionReportColumnWidth,
            zebraFormatting=True,
            zebraField='bus')
    xlsx.formats = gfiConfig.cellFormats
    xlsx.fieldOutline = gfiConfig.exceptionReportFieldOutline 
    xlsx.data = gq.data
    xlsx.generateXLSX()
    xlsx.close()



if __name__ == '__main__':
    args = getArgs()
    if args.error:
        print "Not completed."
        sys.exit(1)

    createReport(args.location,args.year,args.month,args.file,args.connection)

    print "Completed."
    sys.exit(0)




