"""
transitData

Variables and functions directly related to BC transit.
"""


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
        'header':{'bold':True,'font_size':14,'align':'left','align':'vcenter'},
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
        'datared':{'bg_color':'FF4D4D','font_size':9,'align':'center','valign':'vcenter','num_format':0}
        }

# structure of fields/columns:
#   1 field name from SQL query
#   2 col title used in worksheet
#   3 format for data
#   4 format for bottom summary
#   5 function to generate summary function (sum, calculation, etc)
#   6 field used for highlight test
#   7 value to search for in highlight test field
#   8 format to use if highlight test TRUE

exceptionReportFieldOutline = [
        ['bus','Bus','data','colTitle',None,None,None,None],
        ['probetime','Probe Time','data','colTitle',None,None,None,None],
        ['eventtime','Event time','data','colTitle',None,None,None,None],
        ['route','Route','data','colTitle',None,'issue','route','datared'],
        [None,'Route Correction','data','colTitle',None,None,None,None],
        ['drv','Driver','data','colTitle',None,'issue','driver','datared'],
        [None,'Driver Correction','data','colTitle',None,None,None,None],
        ['curr_r','Revenue','data','colTitle',None,None,None,None],
        ['rdr_c','Ridership','data','colTitle',None,None,None,None]
        ]

