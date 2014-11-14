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
        16:"Sunshine Coast",
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
            'text_wrap':True,
            'num_format':'#,###,##0'},
        'colTitleRightB':{'bold':True,'font_size':9,'align':'center',
            'valign':'vcenter','top':True,'bottom':True, 'bg_color':'#EEEEEE','text_wrap':True,
            'text_wrap':True,
            'right':True,
            'num_format':'#,###,##0'},
        'dataDecimal':{'font_size':9,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00'},
        'dataDecimalhot':{'bg_color':'FF5757','font_size':9,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00'},
        'dataDecimalmed':{'bg_color':'FFCDCD','font_size':9,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00'},
        'dataDecimalzebra':{'font_size':9,'top':1,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00'},
        'dataDecimalTitle':{'bold':True,'font_size':9,'align':'right','valign':'vcenter',
            'text_wrap':True,
            'num_format':'#,###,##0.00','top':True,'bottom':True, 'bg_color':'#EEEEEE'},
        'dataDecimalGrey':{'bg_color':'E0E0E0','font_size':9,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00'},
        'dataPercent':{'font_size':9,'align':'center','valign':'vcenter',
            'num_format':'0.0%'},
        'dataPercentTitle':{'bold':True,'font_size':9,'align':'center','valign':'vcenter',
            'num_format':'0.00%','top':True,'bottom':True, 'bg_color':'#EEEEEE'},
        'data':{'font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'dataRightB':{'font_size':9,'align':'center','valign':'vcenter',
            'right':True,'num_format':0},
        'datazebra':{'font_size':9,'top':1,'align':'center','valign':'vcenter','num_format':0},
        'datagrey':{'bg_color':'E0E0E0','font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'datared':{'bg_color':'FF9E9E','font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'datahot':{'bg_color':'FF5757','font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'datamed':{'bg_color':'FFCDCD','font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'datamedRightB':{'bg_color':'FFCDCD','font_size':9,
            'right':True,
            'align':'center','valign':'vcenter','num_format':0},
        'dataredzebra':{'bg_color':'FF9E9E','top':1,'font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'headerred':{'bg_color':'FF9E9E','font_size':9,'align':'left','valign':'vcenter','num_format':0},
        'headerhot':{'bg_color':'FF5757','font_size':9,'align':'left','valign':'vcenter','num_format':0},
        'headermed':{'bg_color':'FFCDCD','font_size':9,'align':'left','valign':'vcenter','num_format':0},
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
    for l in locationIds: _locationString += " - " + systemList[l]
    return _locationString[3:]




"""
###########################################################

Bill Counting Error Report SQL, Header, Field outline

###########################################################
"""



billcountingReportColumnWidth=9


def billcountingReportSQL(location,year):
    """
    Return SQL for exception reports using location, year, and month attributes.
    """
    
    try:
        _location = ','.join([str(s) for s in location])
    except TypeError:
        _location = str(location)


    return (
        "select * from ("
        "select bus, fbx_n,transitmonth,sum(bill_c) bill_sum,avg(bill_c) bill_av "
        "from ( "
        "    select "
        "        bus, fbx_n,to_char(tday,'YYYY-MM') transitmonth,to_char(tday,'YYYY-MM-DD') transitday,sum(bill_c) bill_c "
        "    from ml "
        "    where  "
        "        ml.loc_n in (%s) and "
        "        ml.tday between to_date('%s-01-01','YYYY-MM-DD') and to_date('%s-01-01','YYYY-MM-DD') "
        "    group by bus, fbx_n, to_char(tday,'YYYY-MM'), to_char(tday,'YYYY-MM-DD') "
        ") "
        "group by bus, fbx_n,transitmonth "
        "order by bus,fbx_n,transitmonth "
        ") "
                ) % (
                _location,str(year),str(int(year) +1) )


billcountingReportFieldOutline = [
        ['fbx_n','Farebox ID','data','colTitle',None,None,None,None,None],
        ['bus','Bus ID','data','colTitle',None,None,None,None,None]
        ]

def billcountingReportHeader(location,year):
    return [
        ['GFI Bill Counting Report','header'],
        ['Total & average daily bills counted per farebox/bus per month','subHeader'],
        [locationString( location ),'subHeader'],
        [str(year),'subHeader'],
        ['','subHeader'],
        [['Average daily bill count > 10','','',''],'headermed'], 
        [['Average daily bill count > 20','','',''],'headerhot'] ] 




"""
###########################################################

Driver Key Report SQL, Header, Field outline

###########################################################
"""



driverkeyReportColumnWidth=9


def driverkeyReportSQL(location,year,month):
    """
    Return SQL for exception reports using location, year, and month attributes.
    """
    
    try:
        _location = ','.join([str(s) for s in location])
    except TypeError:
        _location = str(location)


    return (
        "select * from ( "
        "    select "
        "        ev.drv, "
        "        sum(ev.curr_r) curr_r, "
        "        sum(ev.uncl_r) uncl_r, "
        "        sum(ev.dump_c) dump_c, "
        "        sum(ev.fare_c) fare_c, "
        "        sum(ev.key1) key1, "
        "        sum(ev.key2) key2, "
        "        sum(ev.key3) key3, "
        "        sum(ev.key4) key4, "
        "        sum(ev.key5) key5, "
        "        sum(ev.key6) key6, "
        "        sum(ev.key7) key7, "
        "        sum(ev.key8) key8, "
        "        sum(ev.key9) key9, "
        "        sum(ev.keyast) keyast, "
        "        sum(ev.keya) keya, "
        "        sum(ev.keyb) keyb, "
        "        sum(ev.keyc) keyc, "
        "        sum(ev.keyd) keyd, "
        "        sum(ev.ttp1) ttp1, "
        "        sum(ev.ttp2) ttp2, "
        "        sum(ev.ttp3) ttp3, "
        "        sum(ev.ttp4) ttp4, "
        "        sum(ev.ttp5) ttp5, "
        "        sum(ev.ttp6) ttp6, "
        "        sum(ev.ttp7) ttp7, "
        "        sum(ev.ttp8) ttp8, "
        "        sum(ev.ttp9) ttp9, "
        "        sum(ev.ttp10) ttp10, "
        "        sum(ev.ttp11) ttp11, "
        "        sum(ev.ttp12) ttp12, "
        "        sum(ev.ttp13) ttp13, "
        "        sum(ev.ttp14) ttp14, "
        "        sum(ev.ttp15) ttp15, "
        "        sum(ev.ttp16) ttp16, "
        "        sum(ev.ttp17) ttp17, "
        "        sum(ev.ttp18) ttp18, "
        "        sum(ev.ttp19) ttp19, "
        "        sum(ev.ttp20) ttp20, "
        "        sum(ev.ttp21) ttp21, "
        "        sum(ev.ttp22) ttp22, "
        "        sum(ev.ttp23) ttp23, "
        "        sum(ev.ttp24) ttp24, "
        "        sum(ev.ttp25) ttp25, "
        "        sum(ev.ttp26) ttp26, "
        "        sum(ev.ttp27) ttp27, "
        "        sum(ev.ttp28) ttp28, "
        "        sum(ev.ttp29) ttp29, "
        "        sum(ev.ttp30) ttp30, "
        "        sum(ev.ttp31) ttp31, "
        "        sum(ev.ttp32) ttp32, "
        "        sum(ev.ttp33) ttp33, "
        "        sum(ev.ttp34) ttp34, "
        "        sum(ev.ttp35) ttp35, "
        "        sum(ev.ttp36) ttp36, "
        "        sum(ev.ttp37) ttp37, "
        "        sum(ev.ttp38) ttp38, "
        "        sum(ev.ttp39) ttp39, "
        "        sum(ev.ttp40) ttp40, "
        "        sum(ev.ttp41) ttp41, "
        "        sum(ev.ttp42) ttp42, "
        "        sum(ev.ttp43) ttp43, "
        "        sum(ev.ttp44) ttp44, "
        "        sum(ev.ttp45) ttp45, "
        "        sum(ev.ttp46) ttp46, "
        "        sum(ev.ttp47) ttp47, "
        "        sum(ev.ttp48) ttp48 "
        "    from ev "
        "    where "
        "        ev.loc_n in (%s) and "
        "        ev.ts between to_date('%s-%s-01','YYYY-MM-DD') and last_day(to_date('%s-%s-01','YYYY-MM-DD'))+0.99999 "
        "    group by ev.drv "
        ") "
        "where "
        "    (curr_r >0) or "
        "    (drv in (select drvlst.drv from drvlst where drvlst.loc_n in (%s))) "
        "order by drv "
        ) % (
                _location,str(year),str(month),str(year),str(month),
                _location )


driverkeyReportFieldOutline = [
        ['drv','Driver','data','colTitle',None,None,None,None,None],
        ['curr_r','Current Revenue','dataDecimal','dataDecimalTitle',generateSumFunction,None,None,None,None],
        ['uncl_r','Unclassified Revenue','dataDecimal','dataDecimalTitle',generateSumFunction,None,None,None,None],
        ['dump_c','Dump Count','data','colTitle',generateSumFunction,None,None,None,None],
        ['fare_c','Preset Count','data','colTitle',generateSumFunction,None,None,None,None],
        ['key1','Key 1','data','colTitle',generateSumFunction,None,None,None,None],
        ['key2','Key 2','data','colTitle',generateSumFunction,None,None,None,None],
        ['key3','Key 3','data','colTitle',generateSumFunction,None,None,None,None],
        ['key4','Key 4','data','colTitle',generateSumFunction,None,None,None,None],
        ['key5','Key 5','data','colTitle',generateSumFunction,None,None,None,None],
        ['key6','Key 6','data','colTitle',generateSumFunction,None,None,None,None],
        ['key7','Key 7','data','colTitle',generateSumFunction,None,None,None,None],
        ['key8','Key 8','data','colTitle',generateSumFunction,None,None,None,None],
        ['key9','Key 9','data','colTitle',generateSumFunction,None,None,None,None],
        ['keyast','Key *','data','colTitle',generateSumFunction,None,None,None,None],
        ['keya','Key A','data','colTitle',generateSumFunction,None,None,None,None],
        ['keyb','Key B','data','colTitle',generateSumFunction,None,None,None,None],
        ['keyc','Key C','data','colTitle',generateSumFunction,None,None,None,None],
        ['keyd','Key D','data','colTitle',generateSumFunction,None,None,None,None],
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
        ['ttp48','TTP 48','data','colTitle',generateSumFunction,None,None,None,None]
        ]

def driverkeyReportHeader(location,year,month):
    return [
        ['Driver Key Report','header'],
        [calendar.month_name[month]+" "+str(year),'subHeader'],
        [locationString( location ),'subHeader'],
        ['','subHeader'] ]



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
                "and ml.ts between to_date('%s-%s-01', 'YYYY-MM-DD') and last_day(to_date('%s-%s-01', 'YYYY-MM-DD'))+0.99999 "
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
                "and  ml.ts between to_date('%s-%s-01', 'YYYY-MM-DD') and last_day(to_date('%s-%s-01', 'YYYY-MM-DD'))+0.99999 "
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
        ['Route 11 Exception Report','header'],
        [calendar.month_name[month]+" "+str(year),'subHeader'],
        [locationString( location ),'subHeader'],
        ['','subHeader'],
        [['Incorrect Route and Driver numbers are highlighted in RED','','',''],'headerred'], 
        [['Please enter correct values in YELLOW cells','','',''],'headeryellow'] ] 




"""
###########################################################

Driver Reports SQL, Header, Field outline

###########################################################
"""



driverReportColumnWidth=14.4


def bestDriverReportSQL(location,year,month):
    """
    Return SQL for driver reports using location, year, and month attributes.
    """
    
    try:
        _location = ','.join([str(s) for s in location])
    except TypeError:
        _location = str(location)

    return (
        "select rownum,drv,per,uncl_r,curr_r "
        "from ( "
        "    select "
        "        drv, "
        "        (uncl_r/curr_r) per, "
        "        uncl_r, "
        "        curr_r "
        "    from ( "
        "        select  "
        "            ev.drv, "
        "            sum(ev.uncl_r) uncl_r, "
        "            sum(ev.curr_r) curr_r "
        "        from ev left join ml on ml.loc_n=ev.loc_n and ml.id=ev.id "
        "        where "
        "            ev.loc_n in (%s) and "
        "            ev.curr_r >0 and "
        "            ml.tday between to_date('%s-%s-01', 'YYYY-MM-DD') and last_day(to_date('%s-%s-01', 'YYYY-MM-DD'))+0.99999 "
        "        group by drv "
        "    ) "
        "    where curr_r >100 "
        "    order by per "
        ") "
        "where rownum <=10 "
        ) % ( _location,str(year),str(month),str(year),str(month) )

def worstDriverReportSQL(location,year,month):
    """
    Return SQL for driver reports using location, year, and month attributes.
    """
    
    try:
        _location = ','.join([str(s) for s in location])
    except TypeError:
        _location = str(location)

    return (
        "select rownum,drv,per,uncl_r,curr_r "
        "from ( "
        "    select "
        "        drv, "
        "        (uncl_r/curr_r) per, "
        "        uncl_r, "
        "        curr_r "
        "    from ( "
        "        select  "
        "            ev.drv, "
        "            sum(ev.uncl_r) uncl_r, "
        "            sum(ev.curr_r) curr_r "
        "        from ev left join ml on ml.loc_n=ev.loc_n and ml.id=ev.id "
        "        where "
        "            ev.loc_n in (%s) and "
        "            ev.curr_r >0 and "
        "            ml.tday between to_date('%s-%s-01', 'YYYY-MM-DD') and last_day(to_date('%s-%s-01', 'YYYY-MM-DD'))+0.99999 "
        "        group by drv "
        "    ) "
        "    where curr_r >100 "
        "    order by per desc "
        ") "
        "where rownum <=10 "
        ) % ( _location,str(year),str(month),str(year),str(month) )



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

driverReportFieldOutline = [
        ['rownum','Order','data','colTitle',None,None,None,None,None],
        ['drv','Driver','data','colTitle',None,None,None,None,None],
        ['per','Unclassified/Current Revenue (%)','dataPercent','colTitle',None,None,None,None,None],
        ['uncl_r','Unclassified Revenue','dataDecimal','colTitle',None,None,None,None,None],
        ['curr_r','Classified Revenue','dataDecimal','colTitle',None,None,None,None,None]
        ]


def bestDriverReportHeader(location,year,month):
    return [
        ['Driver Unclassified Statistics Report','header'],
        [locationString( location ),'subHeader'],
        ['10 Best Drivers','subHeader'],
        [calendar.month_name[month]+" "+str(year),'subHeader'],
        ['','subHeader'] ]


def worstDriverReportHeader(location,year,month):
    return [
        ['Driver Unclassified Statistics Report','header'],
        [locationString( location ),'subHeader'],
        ['10 Worst Drivers','subHeader'],
        [calendar.month_name[month]+" "+str(year),'subHeader'],
        ['','subHeader'] ]



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
                "ev.route <> 99999 and "
                "ml.loc_n in ( %s ) and  "
                "ml.ts between (to_date('%s-%s-01', 'YYYY-MM-DD') -15) and last_day(to_date('%s-%s-01', 'YYYY-MM-DD'))+0.99999 and  "
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
                "ev.drv <> 99999 and "
                "ml.loc_n in ( %s ) and  "
                "ml.ts between (to_date('%s-%s-01', 'YYYY-MM-DD') -15) and "
                        "last_day(to_date('%s-%s-01', 'YYYY-MM-DD'))+0.99999 and  "
                "ev.drv not in ( "
                    "select drv "
                    "from gfi_range left join "
                        "(select rownum drv from all_objects where rownum < 9999) dids "
                        "on dids.drv >= gfi_range.v1 and dids.drv <= gfi_range.v2 "
                    "where gfi_range.loc_n in ( %s ) "
                    "union "
                    "select drv from drvlst where loc_n in ( %s ) "
                ") and "
                "((ev.curr_r >0) or (ev.rdr_c >0))  "
        ") "
        "group by bus,probetime,eventtime,route,drv,curr_r,rdr_c "
        "order by bus,eventtime "
        ) % (
                _location,str(year),str(month),str(year),str(month),_location,
                _location,str(year),str(month),str(year),str(month),
                _location,_location )



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
        ['GFI Monthly Exception Report','header'],
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
            "AND mrtesum.mday between to_date('%s-%s-01', 'YYYY-MM-DD') and last_day(to_date('%s-%s-01', 'YYYY-MM-DD'))+0.99999 "
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
            "AND mrtesum.mday between to_date('%s-%s-01', 'YYYY-MM-DD') and last_day(to_date('%s-%s-01', 'YYYY-MM-DD'))+0.99999 "
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
            "AND mrtesum.mday between to_date('%s-%s-01', 'YYYY-MM-DD') and last_day(to_date('%s-%s-01', 'YYYY-MM-DD'))+0.99999 "
        "ORDER BY route"
        ) % (_location,str(year),str(month),str(year),str(month),
                _location,str(year),str(month),str(year),str(month),
                _location,str(year),str(month),str(year),str(month))


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
    ['token_t','Token Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['ticket_t','Ticket Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['pass_t','Pass Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['bill_t','Bill Count','data','colTitle',generateSumFunction,None,None,None,None],
    ['coin_t','Coin Count','data','colTitle',generateSumFunction,None,None,None,None],
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
        "select dates.servicedate,data.* "
        "from ( "
            "select to_char(to_date('%s-%s-01','YYYY-MM-DD') + rownum -1,'YYYY-MM-DD') servicedate "
            "from all_objects "
            "where rownum <= last_day(to_date('%s-%s-01','YYYY-MM-DD'))+0.99999 - to_date('%s-%s-01','YYYY-MM-DD') +1 "
        ") dates left join ( "
            "SELECT "
                "to_char(gs.tday, 'YYYY-MM-DD') x_servicedate, "
                "sum(bus_c) bus_c, sum(curr_r) curr_r, sum(rdr_c) rdr_c, "
                "sum(token_t) token_t, sum(ticket_t) ticket_t, sum(pass_t) pass_t, "
                "sum(bill_t) bill_t, sum(coin_t) coin_t, "
                "sum(uncl_r) uncl_r, sum(cbxalm) cbxalm, sum(bypass) bypass, "
                "sum(ttp1) ttp1, sum(ttp2) ttp2, sum(ttp3) ttp3, sum(ttp4) ttp4, sum(ttp5) ttp5, "
                "sum(ttp6) ttp6, sum(ttp7) ttp7, sum(ttp8) ttp8, sum(ttp9) ttp9, sum(ttp10) ttp10, "
                "sum(ttp11) ttp11, sum(ttp12) ttp12, sum(ttp13) ttp13, sum(ttp14) ttp14, sum(ttp15) ttp15, "
                "sum(ttp16) ttp16, sum(ttp17) ttp17, sum(ttp18) ttp18, sum(ttp19) ttp19, sum(ttp20) ttp20, "
                "sum(ttp21) ttp21, sum(ttp22) ttp22, sum(ttp23) ttp23, sum(ttp24) ttp24, sum(ttp25) ttp25, "
                "sum(ttp26) ttp26, sum(ttp27) ttp27, sum(ttp28) ttp28, sum(ttp29) ttp29, sum(ttp30) ttp30, "
                "sum(ttp31) ttp31, sum(ttp32) ttp32, sum(ttp33) ttp33, sum(ttp34) ttp34, sum(ttp35) ttp35, "
                "sum(ttp36) ttp36, sum(ttp37) ttp37, sum(ttp38) ttp38, sum(ttp39) ttp39, sum(ttp40) ttp40, "
                "sum(ttp41) ttp41, sum(ttp42) ttp42, sum(ttp43) ttp43, sum(ttp44) ttp44, sum(ttp45) ttp45, "
                "sum(ttp46) ttp46, sum(ttp47) ttp47, sum(ttp48) ttp48, "
                "sum(key1) key1, sum(key2) key2, sum(key3) key3, sum(key4) key4, sum(key5) key5, "
                "sum(key6) key6, sum(key7) key7, sum(key8) key8, sum(key9) key9, "
                "sum(keyast) keyast, "
                "sum(keya) keya, sum(keyb) keyb, sum(keyc) keyc, sum(keyd) keyd, "
                "sum(fare_c) preset  "
            "FROM gs "
            "WHERE "
                "gs.fs=0 and "
                "gs.tday BETWEEN to_date('%s-%s-01', 'YYYY-MM-DD') and last_day(to_date('%s-%s-01', 'YYYY-MM-DD'))+0.99999  AND "
                "gs.loc_n in ( %s ) "
            "GROUP BY to_char(gs.tday, 'YYYY-MM-DD') "
        ") data on dates.servicedate=data.x_servicedate "
        "order by dates.servicedate "
        ) % (str(year),str(month),str(year),str(month),str(year),str(month),
                str(year),str(month),str(year),str(month),_location)



def msrReportHeader(location,year,month):
    return [
        ['Monthly Summary Report','header'],
        [calendar.month_name[month]+" "+str(year),'subHeader'],
        [locationString( location ),'subHeader'],
        ['','subHeader'] ]



