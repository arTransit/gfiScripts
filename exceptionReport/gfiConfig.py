"""
Variables and config setting for managing GFI data and spreadhseets"
"""

import calendar


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
        'dataDecimalTitle':{'bold':True,'font_size':9,'align':'right','valign':'vcenter',
            'num_format':'#,###,##0.00','top':True,'bottom':True, 'bg_color':'#EEEEEE'},
        'dataPercent':{'font_size':9,'align':'center','valign':'vcenter',
            'num_format':'0.00%'},
        'dataPercentTitle':{'bold':True,'font_size':9,'align':'center','valign':'vcenter',
            'num_format':'0.00%','top':True,'bottom':True, 'bg_color':'#EEEEEE'},
        'data':{'font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'datagrey':{'bg_color':'E0E0E0','font_size':9,'align':'center','valign':'vcenter','num_format':0},
        'datared':{'bg_color':'FF0000','font_size':9,'align':'center','valign':'vcenter','num_format':0}
        }


def locationString( locationIds ):
    """
    Return string of location names specified by locationIds
    """

    _locationString = ''
    for l in locationIds: _locationString += " / " + systemList[l]
    return _locationString[3:]


"""
###########################################################

Exception Reports SQL, Header, Field outline

###########################################################
"""


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
                "ev.ts between to_date('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') and last_day(to_date('%s-%s-01 23:59:59', 'YYYY-MM-DD HH24:MI:SS')) and  "
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
                "ev.ts between to_date('%s-%s-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') and last_day(to_date('%s-%s-01 23:59:59', 'YYYY-MM-DD HH24:MI:SS')) and  "
                "ev.drv not in (select drv from drvlst where loc_n in (%s) ) and "
                "((ev.curr_r >0) or (ev.rdr_c >0))  "
        ") "
        "group by bus,probetime,eventtime,route,drv,curr_r,rdr_c "
        "order by bus,eventtime "
        ) % (
                _location,str(year),str(month),str(year),str(month),_location,
                _location,str(year),str(month),str(year),str(month),_location )



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

exceptionReportFieldOutline = [
        ['bus','Bus','data','colTitle',None,None,None,None,'datagrey'],
        ['probetime','Probe Time','data','colTitle',None,None,None,None,'datagrey'],
        ['eventtime','Event time','data','colTitle',None,None,None,None,'datagrey'],
        ['route','Route','data','colTitle',None,'issue','route','datared','datared'],
        [None,'Route Correction','data','colTitle',None,None,None,None,'datagrey'],
        ['drv','Driver','data','colTitle',None,'issue','driver','datared','datared'],
        [None,'Driver Correction','data','colTitle',None,None,None,None,'datagrey'],
        ['curr_r','Revenue','data','colTitle',None,None,None,None,'datagrey'],
        ['rdr_c','Ridership','data','colTitle',None,None,None,None,'datagrey']
        ]



def exceptionReportHeader(month,year,location):
    return [
            ['Monthly Exception Report','header'],
            [calendar.month_name[month]+" "+str(year),'subHeader'],
            [locationString( location ),'subHeader'] ] 



