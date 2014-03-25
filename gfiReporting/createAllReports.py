"""
createAllReports.py
Generate all reports for all systems, emailing and transferring to network drives as required

"""


import sys
import argparse
import datetime
import gfiConfig
import gfiQuery
import gfiXLSX


systemList = [
        [ids:[1,2],name:"Victoria & Landford"],
        [ids:[3],name:"Whistler"],
        [ids:[4],name:"Squamish"],
        [ids:[5],name:"Nanaimo"],
        [ids:[6],name:"Abbotsford"],
        [ids:[7],name:"Kelowna"],
        [ids:[8],name:"Kamloops"],
        [ids:[9],name:"Prince George"],
        [ids:[10],name:"Cowichan Valley"],
        [ids:[11],name:"Trail"],
        [ids:[12],name:"Comox"],
        [ids:[13],name:"Port Alberni"],
        [ids:[14],name:"Campbell River"],
        [ids:[15],name:"Powell River"],
        [ids:[16],name:"Sunshine Valley"],
        [ids:[17],name:"Vernon"],
        [ids:[18],name:"Penticton"],
        [ids:[19],name:"Chilliwack"],
        [ids:[20],name:"Cranbrook"],
        [ids:[21],name:"Nelson"],
        [ids:[22],name:"Terrace"],
        [ids:[23],name:"Prince Rupert"],
        [ids:[24],name:"Kitimat"],
        [ids:[25],name:"Fort St. John]"
    ]


def getArgs():
    argsPsr = argparse.ArgumentParser(description='Create GFI reports: Exception, MRSR, MSR')
    argsPsr.add_argument('-e','--email',required=True,type=int,help='')
    argsPsr.add_argument('-y','--year',required=True,type=int,help='eg 2014')
    argsPsr.add_argument('-m','--month',required=True,type=int,help='eg 12')
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


if __name__ == '__main__':
    args = getArgs()
    if args.error:
        print "Arguement error"
        sys.exit(1)

    gq = gfiQuery.GFIquery(args.connection, 
            gfiConfig.exceptionReportSQL(args.location,args.year,args.month) )
    gq.execute()
    if not gq.status:
        print "DB error"
        sys.exit(1)

    xlsx = gfiXLSX.gfiSpreadsheet(filename=args.file,
            header=gfiConfig.exceptionReportHeader(args.location,args.year,args.month),
            columnWidth=12,
            zebraFormatting=True,
            zebraField='bus')
    xlsx.formats = gfiConfig.cellFormats
    xlsx.fieldOutline = gfiConfig.exceptionReportFieldOutline 
    xlsx.data = gq.data
    xlsx.generateXLSX()
    xlsx.close()

    print "Completed."
    sys.exit(0)

