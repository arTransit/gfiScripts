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
import sqlite3


SMTPSERVER = '10.170.3.119'
FROM_EMAIL = 'gfiReporting@bctransit.com'
#REPORT_BASE_DIRECTORY='G:/BusinessIntelligence/Temp/GFIreporting/'
#REPORT_BASE_DIRECTORY='C:/Temp/GFIreporting'
#REPORT_BASE_DIRECTORY='G:/Public/GFI/GFIreporting'
REPORT_BASE_DIRECTORY='G:/Public/GFI'
EXCEPTIONDB='./exceptionReport.db'



def getArgs():
    argsPsr = argparse.ArgumentParser(description='Create GFI reports: Exception, MRSR, MSR')
    argsPsr.add_argument('-e','--email',action='store_true',default=False,help='flag to email reports')
    argsPsr.add_argument('-a','--all',action='store_true',default=False,help='create all reports')
    argsPsr.add_argument('-x','--exception',action='store_true',default=False,help='create exception reports')
    argsPsr.add_argument('--reminder',action='store_true',default=False,help='send exception report reminder')
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


def emailReport(emailTo,emailFrom,emailSubject,emailBody,filepaths):

    _bcc = ['gfiReporting@bctransit.com']
    
    emailMsg = email.MIMEMultipart.MIMEMultipart('alternative')
    emailMsg['Subject'] = emailSubject
    emailMsg['From'] = emailFrom
    emailMsg['To'] = ', '.join(emailTo)
    emailMsg['Bcc'] = ', '.join(_bcc)
    emailMsg.attach(email.mime.text.MIMEText(emailBody,'html'))

    # attach file
    for f in filepaths:
        filename = os.path.basename(f)
        fileMsg = email.mime.base.MIMEBase('application','octet-stream')
        fileMsg.set_payload(open(f,'rb').read() )
        email.encoders.encode_base64(fileMsg)
        fileMsg.add_header('Content-Disposition','attachment;filename=%s' % filename)
        emailMsg.attach(fileMsg)

    # send email
    server = smtplib.SMTP(SMTPSERVER)
    server.sendmail(emailFrom,emailTo + _bcc,emailMsg.as_string())
    server.quit()


def dict_factory(cursor, row):
    d = {}
    for idx, col in enumerate(cursor.description):
        d[col[0]] = row[idx]
    return d



if __name__ == '__main__':
    args = getArgs()
    if args.error:
        print "Arguement error"
        sys.exit(1)


    # create locationList
    con = sqlite3.connect(EXCEPTIONDB)
    con.row_factory = dict_factory
    cur = con.cursor()
    cur.execute('select locid,locname,directoryname,email from v_contacts_local')


    systemList = {}
    for l in cur.fetchall():
        systemList[l['locid']] = {
                'locname':l['locname'],
                'directoryname':l['directoryname'],
                'email':l['email'].split(';')}

    # make directories
    for l in systemList.keys():
        makePath(REPORT_BASE_DIRECTORY + '/' + systemList[l]['directoryname'])

    # Exception Reports
    if args.all or args.exception:
        print "Generating Monthly Exception Reports"
        print "  Year:%s\n  Month:%s" % (str(args.year),str(args.month))
        if args.email:
            print "  email is ON"
        else:
            print "  no email"

        for l in systemList.keys():
            #sys.stdout.write('. ')
            sys.stdout.write(systemList[l]['locname'])
            sys.stdout.write('\n')
            sys.stdout.flush()
            _filename = '%s/%s/%s_GFImonthlyExceptionReport_%s_%s.xlsx' % (
                    REPORT_BASE_DIRECTORY,
                    systemList[l]['locname'],
                    systemList[l]['locname'],
                    str(args.year),
                    ''.join(['000',str(args.month)])[-2:])

            generateExceptionReport.createReport( [l],args.year,args.month,_filename,args.connection)

            if args.email:
                emailBody = (
                        '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" '
                        '"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml">'
                        '<body style="font-size:12px;font-family:Tahoma">'
                        '<p>Please find attached the <b>%s GFI Monthly Exception Report, %s %s</b>.'
                        'Please ensure it is returned by 14 April 2015.</p>'
                        '<p>Exception reports are distributed on the 3rd of each month (or the next business day), and should be returned by the 10th of each month (or the next business day).</p>'
                        '<p>In addition, please use the following guidelines when reporting GFI-related issues and changes:</p>'
                        '<ul>'
                        '<li><strong>New Driver/Operator IDs</strong>: email GFIreporting@bctransit.com as soon as possible so they are not reported as errors on the exception report;</li>'
                        '<li><strong>Old Driver/Operator IDs</strong>: email GFIreporting@bctransit.com to have them removed;</li>'
                        '<li><strong>New Route Numbers</strong>: email GFIreporting@bctransit.com as soon as possible so they are not considered errors on the exception report and are included in the Monthly Route Summary Report;</li>'
                        '<li><strong>Old Route Numbers</strong>: email GFIreporting@bctransit.com as soon as possible so they are not included in the Monthly Route Summary Report;</li>'
                        '<li><strong>New Bus numbers</strong>: no action required - they are automatically added to your system when probed;</li>'
                        '<li><strong>Old Bus numbers</strong>: email GFIreporting@bctransit.com as soon as possible so these busses are removed from future reports;</li>'
                        '<li><strong>Unknown Bus numbers</strong>: sometimes the farebox is not reporting the correct bus id (for example, many fareboxes report they are on bus number 0).  These fareboxes need to be reprogrammed by your maintenance staff, and we can help identify the misreporting bus.</li>'
                        '</ul>'
                        '<p>If you need any GFI reports for specific routes, times of day, periods, etc., please email GFIreporting@bctransit.com</p>'
                        '<p>Thanks.</p>'
                        '<p>The GFI Reporting Team.</p>'
                        '</body></html>' ) % (systemList[l]['locname'],calendar.month_name[args.month],str(args.year))
                emailReport(
                        systemList[l]['email'],
                        FROM_EMAIL,
                        'GFI: %s exception report' % systemList[l]['locname'],
                        emailBody,
                        [_filename])
    else:
        print "No Monthly Exception Reports"
    print '\n\n'


    # Monthly Summary Reports (MSR)

    if args.all or args.msr:
        print "Generating Monthly Summary Reports"
        print "  Year:%s\n  Month:%s" % (str(args.year),str(args.month))

        for l in systemList.keys():
            sys.stdout.write('. ')
            sys.stdout.flush()
            _filename = '%s/%s/%s_GFImonthlySummaryReport%s_%s.xlsx' % (
                    REPORT_BASE_DIRECTORY,
                    systemList[l]['locname'],
                    systemList[l]['locname'],
                    str(args.year),
                    ''.join(['000',str(args.month)])[-2:])
            generateMSR.createReport([l],args.year,args.month,_filename,args.connection)
    else:
        print "No Monthly Summary Reports"
    print '\n\n'



    # Monthly Route Summary Reports (MRSR)
    if args.all or args.mrsr:
        print "Generating Monthly Route Summary Reports"
        print "  Year:%s\n  Month:%s" % (str(args.year),str(args.month))

        for l in systemList.keys():
            sys.stdout.write('. ')
            sys.stdout.flush()
            _filename = '%s/%s/%s_GFImonthlyRouteSummaryReport_%s_%s.xlsx' % (
                    REPORT_BASE_DIRECTORY,
                    systemList[l]['locname'],
                    systemList[l]['locname'],
                    str(args.year),
                    ''.join(['000',str(args.month)])[-2:])
            generateMRSR.createReport([l],args.year,args.month,_filename,args.connection)

    else:
        print "No Monthly Route Summary Reports"
    print '\n\n'

    if args.reminder:
        print "Generating Reminders"
        # systemList = {x['ids'][0]:x for x in reportingSystemList}

        cur.execute('select locid,locname,directoryname,email from v_contacts_all')
        systemList = {}
        for l in cur.fetchall():
            systemList[l['locid']] = {
                    'locname':l['locname'],
                    'directoryname':l['directoryname'],
                    'email':l['email'].split(';')}
            
        cur.execute('select locid,year,month from v_exceptionreportsmissing')

        missingReports = {}
        for row in cur.fetchall():
            if row['locid'] in systemList.keys():
                try:
                    missingReports[row['locid']].append({'year':row['year'],'month':row['month']})
                except KeyError:
                    missingReports[row['locid']] = [{'year':row['year'],'month':row['month']}]
            else:
                print "ERROR: %d not in system list" % row['locid']

        for k in missingReports.keys():
            print systemList[k]['locname']

            _filenames=[]
            _reportList = ''
            for r in missingReports[k]:
                _filename = '%s/%s/%s_GFImonthlyExceptionReport_%s_%s.xlsx' % (
                        REPORT_BASE_DIRECTORY,
                        systemList[k]['locname'],
                        systemList[k]['locname'],
                        str(r['year']),
                        ''.join(['000',str(r['month'])])[-2:])
                if os.path.isfile(_filename):
                    _filenames.append(_filename)
                    _reportList+='<li>%s %d</li>' % ( 
                            ('January','February','March','April','May','June',
                            'July','August','September','October','November','December')[r['month'] -1],
                            r['year'])
                else:
                    print "ERROR: missing exception report %s" % _filename

            emailBody = (
                    '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" '
                    '"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml">'
                    '<body style="font-size:12px;font-family:Tahoma">'
                    '<p>The following GFI exception reports have not been received:'
                    '<ul>'
                    '%s'
                    '</ul>'
                    'If you have already sent them, would you please send them again, otherwise kindly complete them as soon as possible.'
                    '<p>Thanks,</p>'
                    '<p>The GFI Reporting Team.</p>'
                    '</body></html>' ) % _reportList
            emailReport(
                    systemList[k]['email'],
                    FROM_EMAIL,
                    'GFI: %s outstanding exception reports' % systemList[k]['locname'],
                    emailBody,
                    _filenames)

    else:
        print "No Reminders"
    print '\n\n'
        

    con.close()
    print "Completed."
    sys.exit(0)

