"""
generateDriverReport.py

Generate GFI Monthly Driver report
Command line usage: 
    generateDriverReport.py -l locid(s) -y year -m month -c user/pass@db -f xlsxName

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
    # argsPsr.add_argument('-f','--file',required=True,help='filename')
    argsPsr.add_argument('-c','--credentials',required=True,help='user/pass@GFI')
    args = argsPsr.parse_args()
    args.error = False
    if (args.year > datetime.date.today().year) or (args.year < 2000):
        print "ERROR: year out of range (2000 - %d)" % datetime.date.today().year
        args.error = True
    if (args.month > 12) or (args.month < 1):
        print "ERROR: month out of range (1 - 12)"
        args.error = True
    return args


def createBestReport(location,year,month,filename,credentials):
    gq = gfiQuery.GFIquery(credentials, 
            gfiConfig.bestDriverReportSQL(location,year,month) )
    gq.execute()
    if not gq.status:
        print "DB error: Exception report"
        print 'Loc: %s, year: %s, month: %s, filename: %s' % (
                str(location),str(year),str(month),str(filename))
        print gfiConfig.bestDriverReportSQL(location,year,month) 
        sys.exit(1)

    xlsx = gfiXLSX.gfiSpreadsheet(
            filename=filename,
            sheetTitle='%s-%s' % (str(year),('00'+str(month))[-2:]),
            header=gfiConfig.bestDriverReportHeader(location,year,month),
            columnWidth=gfiConfig.driverReportColumnWidth )
    xlsx.formats = gfiConfig.cellFormats
    xlsx.fieldOutline = gfiConfig.driverReportFieldOutline 
    xlsx.data = gq.data
    xlsx.generateXLSX()
    xlsx.close()


def createWorstReport(location,year,month,filename,credentials):
    gq = gfiQuery.GFIquery(credentials, 
            gfiConfig.worstDriverReportSQL(location,year,month) )
    gq.execute()
    if not gq.status:
        print "DB error: Exception report"
        print 'Loc: %s, year: %s, month: %s, filename: %s' % (
                str(location),str(year),str(month),str(filename))
        print gfiConfig.worstbestDriverReportSQL(location,year,month) 
        sys.exit(1)

    xlsx = gfiXLSX.gfiSpreadsheet(
            filename=filename,
            sheetTitle='%s-%s' % (str(year),('00'+str(month))[-2:]),
            header=gfiConfig.worstDriverReportHeader(location,year,month),
            columnWidth=gfiConfig.driverReportColumnWidth )
    xlsx.formats = gfiConfig.cellFormats
    xlsx.fieldOutline = gfiConfig.driverReportFieldOutline 
    xlsx.data = gq.data
    xlsx.generateXLSX()
    xlsx.close()



if __name__ == '__main__':
    args = getArgs()
    if args.error:
        print "Not completed."
        sys.exit(1)

    _filename = '%s_GFImonthlyDriverReport_best_%s_%s.xlsx' % ( 
            gfiConfig.locationString( args.location ), str(args.year),
            ('000' + str(args.month) )[-2:] )
    createBestReport(args.location,args.year,args.month,_filename,args.credentials)

    _filename = '%s_GFImonthlyDriverReport_worst_%s_%s.xlsx' % ( 
            gfiConfig.locationString( args.location ), str(args.year),
            ('000'+str(args.month))[-2:] )
    createWorstReport(args.location,args.year,args.month,_filename,args.credentials)

    print "Completed."
    sys.exit(0)




