"""
Variables and config setting for managing GFI data and spreadhseets"
"""

import calendar
import xlsxwriter


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
        16:"Sunshine Valley",
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
        'header':{'bold':True,'font_size':14,'align':'left','valign':'vcenter'},
        'subHeader':{'bold':True,'font_size':11,'align':'left','valign':'vcenter'},
        'colTitle':{'bold':True,'font_size':9,'align':'center',
            'valign':'vcenter','top':True,'bottom':True, 'bg_color':'#EEEEEE','text_wrap':True,
            'num_format':'#,###,##0'},
        'dataDecimal':{'font_size':9,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00'},
        'dataDecimalzebra':{'font_size':9,'top':1,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00'},
        'dataDecimalTitle':{'bold':True,'font_size':9,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00','top':True,'bottom':True, 'bg_color':'#EEEEEE'},
        'dataDecimalGrey':{'bg_color':'E0E0E0','font_size':9,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00'},
        'dataPercent':{'font_size':9,'align':'center','valign':'vcenter',
            'num_format':'0.00%'},
        'dataPercentTitle':{'bold':True,'font_size':9,'align':'center','valign':'vcenter',
            'num_format':'0.00%','top':True,'bottom':True, 'bg_color':'#EEEEEE'},
        'data':{'font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'datazebra':{'font_size':9,'top':1,'align':'center','valign':'vcenter','num_format':0},
        'datagrey':{'bg_color':'E0E0E0','font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'datared':{'bg_color':'FF9E9E','font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'dataredzebra':{'bg_color':'FF9E9E','top':1,'font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'headerred':{'bg_color':'FF9E9E','font_size':9,'align':'left','valign':'vcenter','num_format':0},
        'headeryellow':{'bg_color':'FFFF80','font_size':9,'align':'left','valign':'vcenter','num_format':0},
        'datayellow':{'bg_color':'FFFF80','font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'datayellowzebra':{'bg_color':'FFFF80','top':1,'font_size':9,'align':'center','valign':'vcenter','num_format':0}
        }



def generateSumFunction(*args,**kwargs):
    _range = xlsxwriter.utility.xl_range(kwargs.get('startRow'),kwargs.get('col'),
            kwargs.get('row') -1,kwargs.get('col'))
    return '=SUM(' + _range + ')'



def generatePercentageFunction(*args,**kwargs):

    # dividend is uncl_r, divisor is curr_r
    _dividend = xlsxwriter.utility.xl_rowcol_to_cell(
            kwargs.get('row'), kwargs.get('col')-1 )
    _divisor = xlsxwriter.utility.xl_rowcol_to_cell(
            kwargs.get('row'), kwargs.get('col')-7 )

    return '=IF(%s=0,0,%s/%s)' % (_divisor,_dividend,_divisor)
    


def locationString( locationIds ):
    """
    Return string of location names specified by locationIds
    """

    _locationString = ''
    for l in locationIds: _locationString += " / " + systemList[l]
    return _locationString[3:]




"""
###########################################################

Chilliwack Route 11 Reports SQL, Header, Field outline

###########################################################
"""


chilliwackRoute11busList = [3005,3006,2401,2238,2319]

def chilliwackRoute11SQL(year,month):
    """
    Return SQL for exception reports using location, year, and month attributes.
    """
    _location = str(19)
    _route=11
    _busList= ','.join([str(s) for s in chilliwackRoute11busList])
    
    return (
        "select bus,probetime,eventtime,route,drv,curr_r,rdr_c,wm_concat(issue) as issue "
        "from ( "
            "select "
                "ml.bus, "
                "TO_CHAR(ml.ts,'YYYY-MM-DD HH24:MI') probetime, "
                "TO_CHAR(ev.ts,'YYYY-MM-DD HH24:MI') eventtime, "
                "ev.route, "
                "ev.drv, "
                "ev.curr_r, "
                "ev.rdr_c, "
                "'route' issue "
            "from ml left join ev on ml.loc_n=ev.loc_n and ml.id=ev.id "
            "where "
                "ml.loc_n in ( %s ) "
                "and ml.ts between to_date('%s-%s-01', 'YYYY-MM-DD') and last_day(to_date('%s-%s-01', 'YYYY-MM-DD')) "
                "and ml.bus in (%s) "
                "and ev.route <> %s "
                "and ((ev.curr_r >0) or (ev.rdr_c >0)) "
            "union "
            "select "
                "ml.bus, "
                "TO_CHAR(ml.ts,'YYYY-MM-DD HH24:MI') probetime, "
                "TO_CHAR(ev.ts,'YYYY-MM-DD HH24:MI') eventtime, "
                "ev.route, "
                "ev.drv, "
                "ev.curr_r, "
                "ev.rdr_c, "
                "'route' issue "
            "from ml left join ev on ml.loc_n=ev.loc_n and ml.id=ev.id "
            "where "
                "ml.loc_n in ( %s ) "
                "and  ml.ts between to_date('%s-%s-01', 'YYYY-MM-DD') and last_day(to_date('%s-%s-01', 'YYYY-MM-DD')) "
                "and ml.bus not in (%s) "
                "and ev.route = %s "
                "and ((ev.curr_r >0) or (ev.rdr_c >0)) "
        ") "
        "group by bus,probetime,eventtime,route,drv,curr_r,rdr_c "
        "order by bus,eventtime "
        ) % (
                _location,str(year),str(month),str(year),str(month),
                _busList,str(_route),
                _location,str(year),str(month),str(year),str(month),
                _busList,str(_route) )


chilliwackRoute11FieldOutline = [
        ['bus','Bus','data','colTitle',None,None,None,None,'zebra'],
        ['probetime','Probe Time','data','colTitle',None,None,None,None,'zebra'],
        ['eventtime','Event time','data','colTitle',None,None,None,None,'zebra'],
        ['route','Route','data','colTitle',None,'issue','route','datared','zebra'],
        [None,'Route Correction','data','colTitle',None,'issue','route','datayellow','zebra'],
        ['drv','Driver','data','colTitle',None,'issue','driver','datared','zebra'],
        [None,'Driver Correction','data','colTitle',None,'issue','driver','datayellow','zebra'],
        ['curr_r','Revenue','dataDecimal','colTitle',None,None,None,None,'zebra'],
        ['rdr_c','Ridership','data','colTitle',None,None,None,None,'zebra']
        ]



def chilliwackRoute11ReportHeader(year,month):
    location = [19]
    return [
        ['Monthly Exception Report','header'],
        [calendar.month_name[month]+" "+str(year),'subHeader'],
        [locationString( location ),'subHeader'],
        ['','subHeader'],
        [['Incorrect Route and Driver numbers are highlighted in RED','','',''],'headerred'], 
        [['Please enter correct values in YELLOW cells','','',''],'headeryellow'] ] 




"""
###########################################################

Exception Reports SQL, Header, Field outline

###########################################################
"""



exceptionReportColumnWidth=12


def exceptionReportSQL(location,year,month):
    """
    Return SQL for exception reports using location, year, and month attributes.
    """
    
    try:
        _location = ','.join([str(s) for s in location])
    except TypeError:
        _location = str(location)


    return (
        "select bus,probetime,eventtime,route,drv,curr_r,rdr_c,wm_concat(issue) as issue "
        "from ( "
            "select  "
                "ml.bus,  "
                "TO_CHAR(ml.ts,'YYYY-MM-DD HH24:MI') probetime,  "
                "TO_CHAR(ev.ts,'YYYY-MM-DD HH24:MI') eventtime,  "
                "ev.route,ev.drv,ev.curr_r,ev.rdr_c,'route' issue  "
            "from ml left join ev on ml.loc_n=ev.loc_n and ml.id=ev.id  "
            "where   "
                "ml.loc_n in ( %s ) and  "
                "ml.ts between to_date('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') and last_day(to_date('%s-%s-01 23:59:59', 'YYYY-MM-DD HH24:MI:SS')) and  "
                "ev.route not in (select route from rtelst where loc_n in (%s) ) and  "
                "((ev.curr_r >0) or (ev.rdr_c >0)) "
            "union "
            "select  "
                "ml.bus,  "
                "TO_CHAR(ml.ts,'YYYY-MM-DD HH24:MI') probetime, "
                "TO_CHAR(ev.ts,'YYYY-MM-DD HH24:MI') eventtime, "
                "ev.route,ev.drv,ev.curr_r,ev.rdr_c,'driver' issue  "
            "from ml left join ev on ml.loc_n=ev.loc_n and ml.id=ev.id "
            "where  "
                "ml.loc_n in ( %s ) and  "
                "ml.ts between to_date('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') and "
                        "last_day(to_date('%s-%s-01 23:59:59', 'YYYY-MM-DD HH24:MI:SS')) and  "
                "not ( "
                    "(%s in (select loc_n from gfi_range) and "
                    "( ev.drv between (select v1 from gfi_range where loc_n in (%s)) and "
                                     "(select v2 from gfi_range where loc_n in (%s)))) "
                    "or ev.drv in (select drv from drvlst where loc_n in (%s)) "
                ") and "
                "((ev.curr_r >0) or (ev.rdr_c >0))  "
        ") "
        "group by bus,probetime,eventtime,route,drv,curr_r,rdr_c "
        "order by bus,eventtime "
        ) % (
                _location,str(year),str(month),str(year),str(month),_location,
                _location,str(year),str(month),str(year),str(month),
                _location,_location,_location,_location )



# structure of fields/columns:
#   1 field name from SQL query
#   2 col title used in worksheet
#   3 format for data
#   4 format for bottom summary
#   5 function to generate summary function (sum, calculation, etc)
#   6 field used for highlight test
#   7 value to search for in highlight test field
#   8 format to use if highlight test TRUE
#   9 format to use for zebra formatting - note string appended to format name to get zebra format

exceptionReportFieldOutline = [
        ['bus','Bus','data','colTitle',None,None,None,None,'zebra'],
        ['probetime','Probe Time','data','colTitle',None,None,None,None,'zebra'],
        ['eventtime','Event time','data','colTitle',None,None,None,None,'zebra'],
        ['route','Route','data','colTitle',None,'issue','route','datared','zebra'],
        [None,'Route Correction','data','colTitle',None,'issue','route','datayellow','zebra'],
        ['drv','Driver','data','colTitle',None,'issue','driver','datared','zebra'],
        [None,'Driver Correction','data','colTitle',None,'issue','driver','datayellow','zebra'],
        ['curr_r','Revenue','dataDecimal','colTitle',None,None,None,None,'zebra'],
        ['rdr_c','Ridership','data','colTitle',None,None,None,None,'zebra']
        ]



def exceptionReportHeader(location,year,month):
    return [
        ['Monthly Exception Report','header'],
        [calendar.month_name[month]+" "+str(year),'subHeader'],
        [locationString( location ),'subHeader'],
        ['','subHeader'],
        [['Incorrect Route and Driver numbers are highlighted in RED','','',''],'headerred'], 
        [['Please enter correct values in YELLOW cells','','',''],'headeryellow'] ] 



"""
###########################################################

Monthly Route Summary Reports SQL, Header, Field outline

###########################################################
"""


mrsReportColumnWidth=8.5


def mrsreportSQL(location,year,month):
    """
    Return SQL query using given location, year, and month attributes.
    """
    
    try:
        _location = ','.join([str(s) for s in location])
    except TypeError:
        _location = str(location)


    return (
        "SELECT 'Unknown' route,SUM(curr_r) curr_r,SUM(rdr_c) rdr_c,SUM(token_c) token_c, "
            "SUM(ticket_c) ticket_c, SUM(pass_c) pass_c,SUM(bill_c) bill_c, "
            "SUM(uncl_r) uncl_r,SUM(dump_c) dump_c, "
            "SUM(ttp1) ttp1, SUM(ttp2) ttp2, SUM(ttp3) ttp3, SUM(ttp4) ttp4, SUM(ttp5) ttp5, "
            "SUM(ttp6) ttp6, SUM(ttp7) ttp7, SUM(ttp8) ttp8, SUM(ttp9) ttp9, SUM(ttp10) ttp10, "
            "SUM(ttp11) ttp11, SUM(ttp12) ttp12, SUM(ttp13) ttp13, SUM(ttp14) ttp14, SUM(ttp15) ttp15, "
            "SUM(ttp16) ttp16, SUM(ttp17) ttp17, SUM(ttp18) ttp18, SUM(ttp19) ttp19, SUM(ttp20) ttp20, "
            "SUM(ttp21) ttp21, SUM(ttp22) ttp22, SUM(ttp23) ttp23, SUM(ttp24) ttp24, SUM(ttp25) ttp25, "
            "SUM(ttp26) ttp26, SUM(ttp27) ttp27, SUM(ttp28) ttp28, SUM(ttp29) ttp29, SUM(ttp30) ttp30, "
            "SUM(ttp31) ttp31, SUM(ttp32) ttp32, SUM(ttp33) ttp33, SUM(ttp34) ttp34, SUM(ttp35) ttp35, "
            "SUM(ttp36) ttp36, SUM(ttp37) ttp37, SUM(ttp38) ttp38, SUM(ttp39) ttp39, SUM(ttp40) ttp40, "
            "SUM(ttp41) ttp41, SUM(ttp42) ttp42, SUM(ttp43) ttp43, SUM(ttp44) ttp44, SUM(ttp45) ttp45, "
            "SUM(ttp46) ttp46, SUM(ttp47) ttp47, SUM(ttp48) ttp48, "
            "SUM(fare_c) fare_c, "
            "SUM(key1) key1, SUM(key2) key2, SUM(key3) key3, SUM(key4) key4, SUM(key5) key5, "
            "SUM(key6) key6, SUM(key7) key7, SUM(key8) key8, SUM(key9) key9, "
            "SUM(keyast) keyast, SUM(keya) keya, SUM(keyb) keyb, "
            "SUM(keyc) keyc, SUM(keyd) keyd "
        "FROM mrtesum "
        "WHERE mrtesum.loc_n in (%s) "
            "AND route =-3 "
            "AND mrtesum.mday = TO_DATE('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') "
        "UNION "
        "SELECT 'Other' route,SUM(curr_r) curr_r,SUM(rdr_c) rdr_c,SUM(token_c) token_c, "
            "SUM(ticket_c) ticket_c, SUM(pass_c) pass_c,SUM(bill_c) bill_c, "
            "SUM(uncl_r) uncl_r,SUM(dump_c) dump_c, "
            "SUM(ttp1) ttp1, SUM(ttp2) ttp2, SUM(ttp3) ttp3, SUM(ttp4) ttp4, SUM(ttp5) ttp5, "
            "SUM(ttp6) ttp6, SUM(ttp7) ttp7, SUM(ttp8) ttp8, SUM(ttp9) ttp9, SUM(ttp10) ttp10, "
            "SUM(ttp11) ttp11, SUM(ttp12) ttp12, SUM(ttp13) ttp13, SUM(ttp14) ttp14, SUM(ttp15) ttp15, "
            "SUM(ttp16) ttp16, SUM(ttp17) ttp17, SUM(ttp18) ttp18, SUM(ttp19) ttp19, SUM(ttp20) ttp20, "
            "SUM(ttp21) ttp21, SUM(ttp22) ttp22, SUM(ttp23) ttp23, SUM(ttp24) ttp24, SUM(ttp25) ttp25, "
            "SUM(ttp26) ttp26, SUM(ttp27) ttp27, SUM(ttp28) ttp28, SUM(ttp29) ttp29, SUM(ttp30) ttp30, "
            "SUM(ttp31) ttp31, SUM(ttp32) ttp32, SUM(ttp33) ttp33, SUM(ttp34) ttp34, SUM(ttp35) ttp35, "
            "SUM(ttp36) ttp36, SUM(ttp37) ttp37, SUM(ttp38) ttp38, SUM(ttp39) ttp39, SUM(ttp40) ttp40, "
            "SUM(ttp41) ttp41, SUM(ttp42) ttp42, SUM(ttp43) ttp43, SUM(ttp44) ttp44, SUM(ttp45) ttp45, "
            "SUM(ttp46) ttp46, SUM(ttp47) ttp47, SUM(ttp48) ttp48, "
            "SUM(fare_c) fare_c, "
            "SUM(key1) key1, SUM(key2) key2, SUM(key3) key3, SUM(key4) key4, SUM(key5) key5, "
            "SUM(key6) key6, SUM(key7) key7, SUM(key8) key8, SUM(key9) key9, "
            "SUM(keyast) keyast, SUM(keya) keya, SUM(keyb) keyb, "
            "SUM(keyc) keyc, SUM(keyd) keyd "
        "FROM mrtesum "
        "WHERE mrtesum.loc_n in (%s) "
            "AND route =-2 "
            "AND mrtesum.mday = TO_DATE('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') "
        "UNION "
        "SELECT LPAD(TO_CHAR(route),4,'0000') route,curr_r,rdr_c,token_c,ticket_c,pass_c, "
            "bill_c,uncl_r,dump_c, "
            "ttp1, ttp2, ttp3, ttp4, ttp5, ttp6, ttp7, ttp8, ttp9, ttp10, "
            "ttp11, ttp12, ttp13, ttp14, ttp15, ttp16, ttp17, ttp18, ttp19, ttp20, "
            "ttp21, ttp22, ttp23, ttp24, ttp25, ttp26, ttp27, ttp28, ttp29, ttp30, "
            "ttp31, ttp32, ttp33, ttp34, ttp35, ttp36, ttp37, ttp38, ttp39, ttp40, "
            "ttp41, ttp42, ttp43, ttp44, ttp45, ttp46, ttp47, ttp48, "
            "fare_c, "
            "key1, key2, key3, key4, key5, key6, key7, key8, key9, "
            "keyast, keya, keyb, keyc, keyd "
        "FROM mrtesum "
        "WHERE mrtesum.loc_n in (%s) "
            "AND route >=0 "
            "AND mrtesum.mday = TO_DATE('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') "
        "ORDER BY route"
        ) % (_location,str(year),str(month),_location,str(year),str(month),_location,str(year),str(month))


# structure of fields/columns:
#   1 field name from SQL query
#   2 col title used in worksheet
#   3 format for data
#   4 format for bottom summary
#   5 function to generate summary function (sum, calculation, etc)

# structure of fields/columns:
#   1 field name from SQL query
#   2 col title used in worksheet
#   3 format for data
#   4 format for bottom summary
#   5 function to generate summary function (sum, calculation, etc)
#   6 field used for highlight test
#   7 value to search for in highlight test field
#   8 format to use if highlight test TRUE
#   9 format to use for zebra formatting

mrsrFieldOutline = [
    ['route','Route','data','colTitle',None,None,None,None,None],
    ['curr_r','Current Revenue','dataDecimal','dataDecimalTitle',generateSumFunction,None,None,None,None],
    ['rdr_c','Ridership','data','colTitle',generateSumFunction,None,None,None,None],
    ['token_c','Token Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['ticket_c','Ticket Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['pass_c','Pass Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['bill_c','Bill Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['uncl_r','Unclassified Revenue','dataDecimal','dataDecimalTitle',generateSumFunction,None,None,None,None],
    [generatePercentageFunction,'%','dataPercent','dataPercentTitle',generatePercentageFunction,None,None,None,None],
    ['dump_c','Dump Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp1','TTP 1','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp2','TTP 2','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp3','TTP 3','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp4','TTP 4','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp5','TTP 5','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp6','TTP 6','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp7','TTP 7','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp8','TTP 8','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp9','TTP 9','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp10','TTP 10','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp11','TTP 11','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp12','TTP 12','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp13','TTP 13','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp14','TTP 14','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp15','TTP 15','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp16','TTP 16','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp17','TTP 17','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp18','TTP 18','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp19','TTP 19','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp20','TTP 20','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp21','TTP 21','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp22','TTP 22','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp23','TTP 23','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp24','TTP 24','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp25','TTP 25','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp26','TTP 26','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp27','TTP 27','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp28','TTP 28','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp29','TTP 29','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp30','TTP 30','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp31','TTP 31','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp32','TTP 32','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp33','TTP 33','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp34','TTP 34','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp35','TTP 35','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp36','TTP 36','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp37','TTP 37','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp38','TTP 38','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp39','TTP 39','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp40','TTP 40','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp41','TTP 41','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp42','TTP 42','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp43','TTP 43','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp44','TTP 44','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp45','TTP 45','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp46','TTP 46','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp47','TTP 47','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp48','TTP 48','data','colTitle',generateSumFunction,None,None,None,None],
    ['fare_c','Preset','data','colTitle',generateSumFunction,None,None,None,None],
    ['key1','KEY 1','data','colTitle',generateSumFunction,None,None,None,None],
    ['key2','KEY 2','data','colTitle',generateSumFunction,None,None,None,None],
    ['key3','KEY 3','data','colTitle',generateSumFunction,None,None,None,None],
    ['key4','KEY 4','data','colTitle',generateSumFunction,None,None,None,None],
    ['key5','KEY 5','data','colTitle',generateSumFunction,None,None,None,None],
    ['key6','KEY 6','data','colTitle',generateSumFunction,None,None,None,None],
    ['key7','KEY 7','data','colTitle',generateSumFunction,None,None,None,None],
    ['key8','KEY 8','data','colTitle',generateSumFunction,None,None,None,None],
    ['key9','KEY 9','data','colTitle',generateSumFunction,None,None,None,None],
    ['keyast','KEY *','data','colTitle',generateSumFunction,None,None,None,None],
    ['keya','KEY A','data','colTitle',generateSumFunction,None,None,None,None],
    ['keyb','KEY B','data','colTitle',generateSumFunction,None,None,None,None],
    ['keyc','KEY C','data','colTitle',generateSumFunction,None,None,None,None],
    ['keyd','KEY D','data','colTitle',generateSumFunction,None,None,None,None]
    ]



def mrsrReportHeader(location,year,month):

    return [
        ['Monthly Route Summary Report','header'],
        [calendar.month_name[month]+" "+str(year),'subHeader'],
        [locationString( location ),'subHeader'],
        ['','subHeader' ] ]



"""
###########################################################

Monthly Summary Reports SQL, Header, Field outline

###########################################################
"""

msReportColumnWidth=8.5



# structure of fields/columns:
#   1 field name from SQL query
#   2 col title used in worksheet
#   3 format for data
#   4 format for bottom summary
#   5 function to generate summary function (sum, calculation, etc)
#   6 field used for highlight test
#   7 value to search for in highlight test field
#   8 format to use if highlight test TRUE
#   9 format to use for zebra formatting

msrFieldOutline = [
    ['servicedate','Date','data','colTitle',None,None,None,None,None],
    ['bus_c','Bus Probed','data','colTitle',generateSumFunction,None,None,None,None],
    ['curr_r','Current Revenue','dataDecimal','dataDecimalTitle',generateSumFunction,None,None,None,None],
    ['rdr_c','Ridership','data','colTitle',generateSumFunction,None,None,None,None],
    ['token_c','Token Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['ticket_c','Ticket Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['pass_c','Pass Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['bill_c','Bill Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['coin_c','Coin Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['uncl_r','Unclassified Revenue','dataDecimal','dataDecimalTitle',generateSumFunction,None,None,None,None],
    ['cbxalm','Cahsbox Alarm','data','colTitle',generateSumFunction,None,None,None,None],
    ['bypass','Bypass Alarm','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp1','TTP 1','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp2','TTP 2','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp3','TTP 3','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp4','TTP 4','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp5','TTP 5','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp6','TTP 6','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp7','TTP 7','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp8','TTP 8','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp9','TTP 9','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp10','TTP 10','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp11','TTP 11','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp12','TTP 12','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp13','TTP 13','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp14','TTP 14','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp15','TTP 15','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp16','TTP 16','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp17','TTP 17','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp18','TTP 18','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp19','TTP 19','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp20','TTP 20','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp21','TTP 21','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp22','TTP 22','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp23','TTP 23','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp24','TTP 24','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp25','TTP 25','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp26','TTP 26','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp27','TTP 27','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp28','TTP 28','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp29','TTP 29','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp30','TTP 30','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp31','TTP 31','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp32','TTP 32','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp33','TTP 33','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp34','TTP 34','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp35','TTP 35','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp36','TTP 36','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp37','TTP 37','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp38','TTP 38','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp39','TTP 39','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp40','TTP 40','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp41','TTP 41','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp42','TTP 42','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp43','TTP 43','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp44','TTP 44','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp45','TTP 45','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp46','TTP 46','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp47','TTP 47','data','colTitle',generateSumFunction,None,None,None,None],
    ['ttp48','TTP 48','data','colTitle',generateSumFunction,None,None,None,None],
    ['key1','KEY 1','data','colTitle',generateSumFunction,None,None,None,None],
    ['key2','KEY 2','data','colTitle',generateSumFunction,None,None,None,None],
    ['key3','KEY 3','data','colTitle',generateSumFunction,None,None,None,None],
    ['key4','KEY 4','data','colTitle',generateSumFunction,None,None,None,None],
    ['key5','KEY 5','data','colTitle',generateSumFunction,None,None,None,None],
    ['key6','KEY 6','data','colTitle',generateSumFunction,None,None,None,None],
    ['key7','KEY 7','data','colTitle',generateSumFunction,None,None,None,None],
    ['key8','KEY 8','data','colTitle',generateSumFunction,None,None,None,None],
    ['key9','KEY 9','data','colTitle',generateSumFunction,None,None,None,None],
    ['keyast','KEY *','data','colTitle',generateSumFunction,None,None,None,None],
    ['keya','KEY A','data','colTitle',generateSumFunction,None,None,None,None],
    ['keyb','KEY B','data','colTitle',generateSumFunction,None,None,None,None],
    ['keyc','KEY C','data','colTitle',generateSumFunction,None,None,None,None],
    ['keyd','KEY D','data','colTitle',generateSumFunction,None,None,None,None],
    ['preset','Preset','data','colTitle',generateSumFunction,None,None,None,None]
    ]

def msreportSQL(location,year,month):
    """
    Return SQL query using given location, year, and month attributes.
    """
    
    try:
        _location = ','.join([str(s) for s in location])
    except TypeError:
        _location = str(location)

    return (
        "SELECT to_char(ml.tday, 'YYYY-MM-DD') serviceDate,"
        " count(distinct ml.bus) bus_c,sum(ml.curr_r) curr_r,"
        " sum(ml.rdr_c) rdr_c, sum(ml.token_c) token_c,"
        " sum(ml.ticket_c) ticket_c, "
        " sum(ml.pass_c - gfi_ml.misread_c - gfi_ml.passback_c - gfi_ml.invalid_c - gfi_ml.expired_c - gfi_ml.badlist_c) pass_c, "
        " sum(ml.bill_c) bill_c,"
        " sum(ml.dime + ml.penny + ml.nickel + ml.quarter + ml.half + ml.sba) coin_c, "
        " sum(ml.uncl_r) uncl_r,sum(ml.cbxalm) cbxalm,"
        " sum(ml.bypass) bypass, "
        " sum(ml.ttp1) ttp1, sum(ml.ttp2) ttp2, sum(ml.ttp3) ttp3, "
        " sum(ml.ttp4) ttp4, sum(ml.ttp5) ttp5, sum(ml.ttp6) ttp6, "
        " sum(ml.ttp7) ttp7, sum(ml.ttp8) ttp8, sum(ml.ttp9) ttp9, "
        " sum(ml.ttp10) ttp10, sum(ml.ttp11) ttp11, sum(ml.ttp12) ttp12, "
        " sum(ml.ttp13) ttp13, sum(ml.ttp14) ttp14, sum(ml.ttp15) ttp15, "
        " sum(ml.ttp16) ttp16, sum(ml.ttp17) ttp17, sum(ml.ttp18) ttp18, "
        " sum(ml.ttp19) ttp19, sum(ml.ttp20) ttp20, sum(ml.ttp21) ttp21, "
        " sum(ml.ttp22) ttp22, sum(ml.ttp23) ttp23, sum(ml.ttp24) ttp24, "
        " sum(ml.ttp25) ttp25, sum(ml.ttp26) ttp26, sum(ml.ttp27) ttp27, "
        " sum(ml.ttp28) ttp28, sum(ml.ttp29) ttp29, sum(ml.ttp30) ttp30, "
        " sum(ml.ttp31) ttp31, sum(ml.ttp32) ttp32, sum(ml.ttp33) ttp33, "
        " sum(ml.ttp34) ttp34, sum(ml.ttp35) ttp35, sum(ml.ttp36) ttp36, "
        " sum(ml.ttp37) ttp37, sum(ml.ttp38) ttp38, sum(ml.ttp39) ttp39, "
        " sum(ml.ttp40) ttp40, sum(ml.ttp41) ttp41, sum(ml.ttp42) ttp42, "
        " sum(ml.ttp43) ttp43, sum(ml.ttp44) ttp44, sum(ml.ttp45) ttp45, "
        " sum(ml.ttp46) ttp46, sum(ml.ttp47) ttp47, sum(ml.ttp48) ttp48, "
        " sum(ml.key1) key1, sum(ml.key2) key2, sum(ml.key3) key3, "
        " sum(ml.key4) key4, sum(ml.key5) key5, sum(ml.key6) key6, "
        " sum(ml.key7) key7, sum(ml.key8) key8, sum(ml.key9) key9, "
        " sum(ml.keyast) keyast, sum(ml.keya) keya, sum(ml.keyb) keyb, "
        " sum(ml.keyc) keyc, sum(ml.keyd) keyd, sum(ml.fare_c) preset "
        " "
        "FROM ml left join gfi_ml on ml.loc_n = gfi_ml.loc_n and ml.id=gfi_ml.id "
        " "
        "WHERE "
        " ml.tday BETWEEN to_date('%s-%s-01', 'YYYY-MM-DD') "
        " AND last_day(to_date('%s-%s-01', 'YYYY-MM-DD')) "
        " AND ml.loc_n in ( %s ) "
        " "
        "GROUP BY to_char(ml.tday, 'YYYY-MM-DD') "
        " "
        "ORDER BY serviceDate"
        ) % (str(year),str(month),str(year),str(month),_location)



def msrReportHeader(location,year,month):
    return [
        ['Monthly Summary Report','header'],
        [calendar.month_name[month]+" "+str(year),'subHeader'],
        [locationString( location ),'subHeader'],
        ['','subHeader'] ]



