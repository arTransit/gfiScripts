""" createAllReports.py
Generate all reports for all systems, emailing and transferring to network drives as required

"""


import sys
import os
import errno
import argparse
import datetime
import calendar
import generateMSR
import generateMRSR
import generateExceptionReport
import smtplib
import email, email.encoders,email.mime.text,email.mime.base


SMTPSERVER = '10.170.3.119'
FROM_EMAIL = 'gfiReporting@bctransit.com'


REPORT_BASE_DIRECTORY='G:/BusinessIntelligence/Temp'
#REPORT_BASE_DIRECTORY='C:/Temp/GFIreporting'
#REPORTBASEDIRECTORY='G:/Public/GFI/GFIreporting'


reportingSystemList = [
        {'ids':[1,2],'name':'Victoria_Langford','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[3],'name':'Whistler','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[4],'name':'Squamish','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[5],'name':'Nanaimo','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[6],'name':'Abbotsford','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[7],'name':'Kelowna','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[8],'name':'Kamloops','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[9],'name':'Prince George','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[10],'name':'Cowichan Valley','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[11],'name':'Trail','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[12],'name':'Comox','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[13],'name':'Port Alberni','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[14],'name':'Campbell River','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[15],'name':'Powell River','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[16],'name':'Sunshine Valley','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[17],'name':'Vernon','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[18],'name':'Penticton','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[19],'name':'Chilliwack','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[20],'name':'Cranbrook','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[21],'name':'Nelson','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[22],'name':'Terrace','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[23],'name':'Prince Rupert','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[24],'name':'Kitimat','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']},
        {'ids':[25],'name':'Fort StJohn','email':['andrew_ross@bctransit.com','andrew_miller@bctransit.com']}
    ]


def getArgs():
    argsPsr = argparse.ArgumentParser(description='Create GFI reports: Exception, MRSR, MSR')
    argsPsr.add_argument('-e','--email',action='store_true',default=False,help='flag to email reports')
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



def makePath(path):
    try:
        os.makedirs(path)
    except OSError as exception:
        if exception.errno != errno.EEXIST:
            raise


def emailReport(emailTo,emailFrom,emailSubject,emailBody,filepath):
    emailMsg = email.MIMEMultipart.MIMEMultipart('alternative')
    emailMsg['Subject'] = emailSubject
    emailMsg['From'] = emailFrom
    emailMsg['To'] = ', '.join(emailTo)
    emailMsg.attach(email.mime.text.MIMEText(emailBody,'html'))

    # attach file
    filename = os.path.basename(filepath)
    fileMsg = email.mime.base.MIMEBase('application','octet-stream')
    fileMsg.set_payload(open(filepath,'rb').read() )
    email.encoders.encode_base64(fileMsg)
    fileMsg.add_header('Content-Disposition','attachment;filename=%s' % filename)
    emailMsg.attach(fileMsg)

    # send email
    server = smtplib.SMTP(SMTPSERVER)
    server.sendmail(emailFrom,emailTo,emailMsg.as_string())
    server.quit()


if __name__ == '__main__':
    args = getArgs()
    if args.error:
        print "Arguement error"
        sys.exit(1)



    # make directories
    for s in reportingSystemList:
        makePath(REPORT_BASE_DIRECTORY + '/' + s['name'])



    # Exception Reports

    print "Generating Monthly Exception Reports"
    print "  Year:%s\n  Month:%s" % (str(args.year),str(args.month))

    for s in reportingSystemList:
        sys.stdout.write('. ')
        sys.stdout.flush()
        _filename = '%s/%s/MonthlyException_%s_%s.xlsx' % (
                REPORT_BASE_DIRECTORY,s['name'],str(args.year),''.join(['000',str(args.month)])[-2:])
        generateExceptionReport.createReport(s['ids'],args.year,args.month,_filename,args.connection)

        emailBody = (
                '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" '
                '"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml">'
                '<body style="font-size:12px;font-family:Tahoma">'
                '<p>Please find attached the <b>%s GFI Monthly Summary Report, %s %s</b>.</p>'
                '<p>All reports are located here: <i>G:\BusinessIntelligence\Temp</i></p>'
                '<p>The GFI Reporting Team &lt;gfiReporting@bctransit.com&gt;</p>'
                '</body></html>' ) % (s['name'],calendar.month_name[args.month],str(args.year))
        emailReport(
                s['email'],
                FROM_EMAIL,
                'GFI Monthly Exception Report - %s, %s %d' % (s['name'],calendar.month_name[args.month],args.year),
                emailBody,
                _filename)
    print '\n\n'


    # Monthly Summary Reports (MSR)

    print "Generating Monthly Summary Reports"
    print "  Year:%s\n  Month:%s" % (str(args.year),str(args.month))

    for s in reportingSystemList:
        sys.stdout.write('. ')
        sys.stdout.flush()
        _filename = '%s/%s/MonthlySummary_%s_%s.xlsx' % (
                REPORT_BASE_DIRECTORY,s['name'],str(args.year),''.join(['000',str(args.month)])[-2:])
        generateMSR.createReport(s['ids'],args.year,args.month,_filename,args.connection)
    print '\n\n'



    # Monthly Route Summary Reports (MRSR)
    print "Generating Monthly Route Summary Reports"
    print "  Year:%s\n  Month:%s" % (str(args.year),str(args.month))

    for s in reportingSystemList:
        sys.stdout.write('. ')
        sys.stdout.flush()
        _filename = '%s/%s/MonthlyRouteSummary_%s_%s.xlsx' % (
                REPORT_BASE_DIRECTORY,s['name'],str(args.year),''.join(['000',str(args.month)])[-2:])
        generateMRSR.createReport(s['ids'],args.year,args.month,_filename,args.connection)
    print '\n\n'



    print "Completed."
    sys.exit(0)

