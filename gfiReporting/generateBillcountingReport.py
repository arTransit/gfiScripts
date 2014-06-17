"""
generateBillcountingReport.py

Generate GFI Bill Counting Error Report
Command line usage: 
    generateBillcountingReport.py -l locid(s) -y year -m month -c user/pass@db

"""


import sys
import argparse
import datetime
import gfiConfig
import gfiQuery
import gfiXLSX



def getArgs():
    argsPsr = argparse.ArgumentParser(description='Generate GFI Driver Key Report')
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


def createReport(location,year,month,filename,credentials):

    gq = gfiQuery.GFIquery(credentials, 
            gfiConfig.billcountingReportSQL(location,year,month) )
    gq.execute()
    if not gq.status:
        print "DB error: Bill Error Counting report"
        print 'Loc: %s, year: %s, month: %s, filename: %s' % (
                str(location),str(year),str(month),str(filename))
        print gfiConfig.billcountingReportSQL(location,year,month) 
        sys.exit(1)

    xlsx = gfiXLSX.gfiSpreadsheet(
            filename=filename,
            sheetTitle='%s-%s' % (str(year),('00'+str(month))[-2:]),
            header=gfiConfig.billcountingReportHeader(location,year,month),
            columnWidth=gfiConfig.billcountingReportColumnWidth,
            summaryRow=True,
            zebraFormatting=False,
            zebraField=None)
    xlsx.formats = gfiConfig.cellFormats
    xlsx.fieldOutline = gfiConfig.billcountingReportFieldOutline
    xlsx.data = gq.data
    xlsx.generateXLSX()
    xlsx.close()



if __name__ == '__main__':
    args = getArgs()
    if args.error:
        print "Not completed."
        sys.exit(1)

    _filename = '%s_GFIbillCountingErrorReport_%s_%s.xlsx' % (
            gfiConfig.locationString( args.location ), str(args.year),
            ('000' + str(args.month) )[-2:] )

    createReport(args.location,args.year,args.month,_filename,args.credentials)

    print "Completed."
    sys.exit(0)




