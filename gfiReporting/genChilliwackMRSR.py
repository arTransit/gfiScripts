"""
genChilliwackMRSR.py

Generate GFI Monthly Route Summary Report
Special report for Chilliwack with seperate 
Command line usage: 
    generateMRSR.py -l locid(s) -y year -m month -c oracleCred -f xlsxName

"""

import sys
import argparse
import datetime
import chilliwackConfig
import gfiQuery
import gfiXLSX


def getArgs():
    argsPsr = argparse.ArgumentParser(description='Generate GFI Monthly Summary Report')
    # argsPsr.add_argument('-l','--location',required=True,nargs='+',type=int,help='id(s) of system/location')
    argsPsr.add_argument('-y','--year',required=True,type=int,help='eg 2014')
    argsPsr.add_argument('-m','--month',required=True,type=int,help='eg 12')
    #argsPsr.add_argument('-f','--file',required=True,help='filename')
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



def createChilliwackReport(location,year,month,filename,connection):
    gq = gfiQuery.GFIquery(connection, 
            chilliwackConfig.mrsChilliwackReportSQL(location,year,month) )
    gq.execute()
    if not gq.status:
        print "DB error"
        sys.exit(1)

    xlsx = gfiXLSX.gfiSpreadsheet(
            filename=filename,
            sheetTitle='%s-%s' % (str(year),('00'+str(month))[-2:]),
            header=chilliwackConfig.mrsrChilliwackReportHeader(location,year,month),
            columnWidth=chilliwackConfig.mrsReportColumnWidth,
            summaryRow=True,
            zebraFormatting=False)
    xlsx.formats = chilliwackConfig.cellFormats
    xlsx.fieldOutline = chilliwackConfig.mrsrFieldOutline
    xlsx.data = gq.data
    xlsx.generateXLSX()
    xlsx.close()



def createAgassizReport(location,year,month,filename,connection):
    gq = gfiQuery.GFIquery(connection, 
            chilliwackConfig.mrsAgassizReportSQL(location,year,month) )
    gq.execute()
    if not gq.status:
        print "DB error"
        sys.exit(1)

    xlsx = gfiXLSX.gfiSpreadsheet(
            filename=filename,
            sheetTitle='%s-%s' % (str(year),('00'+str(month))[-2:]),
            header=chilliwackConfig.mrsrAgassizReportHeader(location,year,month),
            columnWidth=chilliwackConfig.mrsReportColumnWidth,
            summaryRow=True,
            zebraFormatting=False)
    xlsx.formats = chilliwackConfig.cellFormats
    xlsx.fieldOutline = chilliwackConfig.mrsrFieldOutline
    xlsx.data = gq.data
    xlsx.generateXLSX()
    xlsx.close()



if __name__ == '__main__':
    args = getArgs()
    if args.error:
        print "Not completed."
        sys.exit(1)

    _chilliwackFilename = 'Chilliwack_MonthlyRouteSummary_' +str(args.year)+'-'+('00'+str(args.month))[-2:] +'.xlsx'
    _agassizFilename = 'Agassiz_MonthlyRouteSummary_' +str(args.year)+'-'+('00'+str(args.month))[-2:] +'.xlsx'
    createChilliwackReport(19,args.year,args.month,_chilliwackFilename,args.connection)
    createAgassizReport(19,args.year,args.month,_agassizFilename,args.connection)

    print "Completed."
    sys.exit(0)



