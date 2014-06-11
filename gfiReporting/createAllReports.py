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
#REPORT_BASE_DIRECTORY='G:/BusinessIntelligence/Temp/GFIreporting/'
#REPORT_BASE_DIRECTORY='C:/Temp/GFIreporting'
#REPORT_BASE_DIRECTORY='G:/Public/GFI/GFIreporting'
REPORT_BASE_DIRECTORY='G:/Public/GFI/x'


reportingSystemList = [
        {'ids':[1,2],'name':'Victoria_Langford','email':['gfiReporting@bctransit.com']},
        {'ids':[6],'name':'Abbotsford','email':['Gabe Colusso <gabe.colusso@firstgroup.com>']},
        {'ids':[14],'name':'Campbell River','email':['Bill Richards <crtransit@shaw.ca>']},
        {'ids':[19],'name':'Chilliwack','email':['Gabe Colusso <gabe.colusso@firstgroup.com>']},
        {'ids':[12],'name':'Comox','email':['Darren Richards <watsonandash@shaw.ca>']},
        {'ids':[10],'name':'Duncan','email':['Colin Oakes <colin.oakes@firstgroup.com>']},
        {'ids':[20],'name':'Cranbrook','email':['Lynda Lawrence <lynda@suncity.bc.ca>','John Darula <john.darula@suncity.bc.ca>']},
        {'ids':[25],'name':'FSJ','email':['Shelley Lindaas <shelleyl@peacetransit.pwt.ca>']},
        {'ids':[8],'name':'Kamloops','email':['Bart Carrigan <bart.carrigan@firstgroup.com>']},
        {'ids':[7],'name':'Kelowna','email':['Alanna Zaharko <alanna.zaharko@firstgroup.com>']},
        {'ids':[24],'name':'Kitimat','email':['Phil Malnis <phil.malnis@firstgroup.com>','Crylstal Colongard <crystal.colongard@firstgroup.com>']},
        {'ids':[5],'name':'Nanaimo','email':['Darren Marshall <dmarshall@rdn.bc.ca>','David Stowell-Smith <dstowell-smith@rdn.bc.ca>','Dave Sakai <dsakai@rdn.bc.ca>','Jamie Logan <JLogan@rdn.bc.ca>']},
        {'ids':[21],'name':'Nelson','email':['Gerry Tennant <GTennant@nelson.ca>']},
        {'ids':[18],'name':'Penticton','email':['Mike Palosky <mikepalosky@berryandsmith.com>']},
        {'ids':[13],'name':'Port Alberni','email':['Phil Atkinson <phil@patransit.pwt.ca>']},
        {'ids':[15],'name':'Powell River','email':['Gerry Woods <gwoods@cdpr.bc.ca>']},
        {'ids':[9],'name':'Prince George','email':['Erik Madsen <erikm@pgtransit.pwt.ca>']},
        {'ids':[23],'name':'Prince Rupert','email':['Darby Minhas <darbara.minhas@firstgroup.com>']},
        {'ids':[4],'name':'Squamish','email':['Christine Darling <christined@squamishtransit.pwt.ca>']},
        {'ids':[16],'name':'Sunshine','email':['Amanda Walkley <amanda.walkey@scrd.ca>']},
        {'ids':[22],'name':'Terrace','email':['Marilyn Ouellet <marilyn.ouellet@firstgroup.com>']},
        {'ids':[11],'name':'Trail','email':['Sharman Thomas <sharman.trailtransit@shawlink.ca>']},
        {'ids':[17],'name':'Vernon','email':['Cindy Laidlaw <cindy.laidlaw@firstgroup.com>','Doreen Stanton <doreen.stanton@firstgroup.com>']},
        {'ids':[3],'name':'Whistler','email':['Steve Antil <steve@whistlertransit.ca>']}
    ]


def getArgs():
    argsPsr = argparse.ArgumentParser(description='Create GFI reports: Exception, MRSR, MSR')
    argsPsr.add_argument('-e','--email',action='store_true',default=False,help='flag to email reports')
    argsPsr.add_argument('-a','--all',action='store_true',default=False,help='create all reports')
    argsPsr.add_argument('-x','--exception',action='store_true',default=False,help='create exception reports')
    argsPsr.add_argument('-r','--mrsr',action='store_true',default=False,help='create month route summary reports')
    argsPsr.add_argument('-s','--msr',action='store_true',default=False,help='create month summary reports')
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
    emailMsg['Bcc'] = 'gfireporting@bctransit.com'
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

    if args.all or args.exception:
        print "Generating Monthly Exception Reports"
        print "  Year:%s\n  Month:%s" % (str(args.year),str(args.month))
        if args.email:
            print "  email is ON"
        else:
            print "  no email"

        for s in reportingSystemList:
            sys.stdout.write('. ')
            sys.stdout.flush()
            _filename = '%s/%s/MonthlyException_%s_%s.xlsx' % (
                    REPORT_BASE_DIRECTORY,s['name'],str(args.year),''.join(['000',str(args.month)])[-2:])
            print '_filename: %s' % _filename

            generateExceptionReport.createReport(s['ids'],args.year,args.month,_filename,args.connection)

            if args.email:
                emailBody = (
                        '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" '
                        '"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml">'
                        '<body style="font-size:12px;font-family:Tahoma">'
                        '<p>Please find attached the <b>%s GFI Monthly Exception Report, %s %s</b>.</p>'
                        '<p>The GFI Reporting Team &lt;gfiReporting@bctransit.com&gt;</p>'
                        '</body></html>' ) % (s['name'],calendar.month_name[args.month],str(args.year))
                emailReport(
                        s['email'],
                        FROM_EMAIL,
                        'GFI Monthly Exception Report - %s, %s %d' % (s['name'],calendar.month_name[args.month],args.year),
                        emailBody,
                        _filename)
    else:
        print "No Monthly Exception Reports"
    print '\n\n'


    # Monthly Summary Reports (MSR)

    if args.all or args.msr:
        print "Generating Monthly Summary Reports"
        print "  Year:%s\n  Month:%s" % (str(args.year),str(args.month))

        for s in reportingSystemList:
            sys.stdout.write('. ')
            sys.stdout.flush()
            _filename = '%s/%s/MonthlySummaryReport_%s_%s.xlsx' % (
                    REPORT_BASE_DIRECTORY,s['name'],str(args.year),''.join(['000',str(args.month)])[-2:])
            generateMSR.createReport(s['ids'],args.year,args.month,_filename,args.connection)
    else:
        print "No Monthly Summary Reports"
    print '\n\n'



    # Monthly Route Summary Reports (MRSR)
    if args.all or args.mrsr:
        print "Generating Monthly Route Summary Reports"
        print "  Year:%s\n  Month:%s" % (str(args.year),str(args.month))

        for s in reportingSystemList:
            sys.stdout.write('. ')
            sys.stdout.flush()
            _filename = '%s/%s/MonthlyRouteSummaryReport_%s_%s.xlsx' % (
                    REPORT_BASE_DIRECTORY,s['name'],str(args.year),''.join(['000',str(args.month)])[-2:])
            generateMRSR.createReport(s['ids'],args.year,args.month,_filename,args.connection)

    else:
        print "No Monthly Route Summary Reports"
    print '\n\n'

    print "Completed."
    sys.exit(0)

