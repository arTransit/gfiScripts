#!/usr/bin/bash

PUBLICBASEDIR="/cygdrive/g/Public/GFI"
GFILOG="/home/aross/gfi/gfiReporting/gfi.log"
QMONTH=$(date --date "last month" +'%m')
QYEAR=$(date --date "last month" +"%Y")
QLOC=1
QDIR="."
QCONNECTION="gfi/gfi@gfi"
MONTHNAMES=(xxx Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec)
EXCEPTIONREPORTDB="exceptionReport.db"


declare -A LOCLIST=( 
    [1 2]="Victoria_Langford" 
    [1]="Victoria" 
    [2]="Langford" 
    [6]="Abbotsford" 
    [14]="Campbell River" 
    [19]="Chilliwack" 
    [12]="Comox" 
    [10]="Duncan" 
    [20]="Cranbrook" 
    [25]="FSJ" 
    [8]="Kamloops" 
    [7]="Kelowna" 
    [24]="Kitimat" 
    [5]="Nanaimo" 
    [21]="Nelson" 
    [18]="Penticton" 
    [13]="Port Alberni" 
    [15]="Powell River" 
    [9]="Prince George" 
    [23]="Prince Rupert" 
    [4]="Squamish" 
    [16]="Sunshine" 
    [22]="Terrace" 
    [11]="Trail" 
    [17]="Vernon" 
    [3]="Whistler" 
    )

declare -A EXCEPTIONACTIONS=( 
    # use $'' notation for newlines
    [6]=$'Abbotsford:\nSend driver unclassified report to Gabe Colusso <gabe.colusso@firstgroup.com>' 
    [19]=$'Chilliwack:\nSend Chilliwack MRS to:\nRod Sanderson <SANDERSO@chilliwack.com>; Jennifer Kooistra <kooistra@chilliwack.com>; Johann VanSchaik <Johann_VanSchaik@BCTransit.Com>\n\nSubject:GFI Monthly Summary Report - Chilliwack ${MONTHNAMES[$QMONTH]} $QYEAR\n\nPlease find attached the GFI Monthly Summary Report for Chilliwack - ${MONTHNAMES[$QMONTH]} $QYEAR\n\nSend Agassiz (just route 11) MRS to:\nAlison Stewart <astewart@fvrd.bc.ca>; Barclay Pitkethly <bpitkethly@fvrd.bc.ca>;accountsreceivable@fvrd.bc.ca; Jennifer Kooistra <kooistra@chilliwack.com>; Mike Veenbaas <mveenbaas@fvrd.bc.ca>; Johann VanSchaik <Johann_VanSchaik@BCTransit.Com>; Leigh Kreitner <kreitner@chilliwack.com>; Michelle McGrath <mcgrath@chilliwack.com>\n\nSubject:GFI Monthly Summary Report - Agassiz ${MONTHNAMES[$QMONTH]} $QYEAR\n\nPlease find attached the GFI Monthly Summary Report for Agassiz - ${MONTHNAMES[$QMONTH]} $QYEAR\n' 
    [7]=$'Kelowna:\nSend driver unclassified report to Bill Harding <bill.harding@firstgroup.com>' 
    [4]=$'Squamish:\nSend driver key report to Christine Darling <christined@squamishtransit.pwt.ca>' 
    [17]=$'Vernon:\nSend MRS to Cindy Laidlaw <cindy.laidlaw@firstgroup.com>; Doreen Stanton <doreen.stanton@firstgroup.com>\nBody: Please find attached the GFI Monthly RouteSummary Report for Vernon ${MONTHNAMES[$QMONTH]} $QYEAR.' 
    [11]=$'Trail:\nUse West Kootenay Updater to update spreadsheet and forward to \nDaniel Pizarro <daniel_pizarro@bctransit.com>; Randall Matheson <RMatheson@rdck.bc.ca>; John MacLean <jmaclean@rdkb.com>\n\nSubject:Ridership and Revenue Reports - WKT ${MONTHNAMES[$QMONTH]} $QYEAR\n\nPlease find attached the WKT GFI Ridership and Cash Revenue Report for ${MONTHNAMES[$QMONTH]} $QYEAR\n' 
    [3]=$'Whistler:\nSend MRSR to  \nEmma DalSanto <EDalSanto@whistler.ca>; Peter DeJong <PDeJong@slrd.bc.ca> \n\nSubject:GFI Whistler: Monthly Route Summary Report ${MONTHNAMES[$QMONTH]} $QYEAR\n\nPlease find attached the GFI Whislter Monthly Route Summary Report for ${MONTHNAMES[$QMONTH]} $QYEAR\n' 
    )


function getLocation {
    echo -e "\nLocations:"
    for l in "${!LOCLIST[@]}" ; do
        echo "$l;${LOCLIST[$l]}"; 
    done | sort -t';' -k2 | column -t -s';' 
    echo -n "New location:"; read QLOC
}

function getDirectory {
    echo -n "New directory:"
    read d
    case $d in
        p) QDIR="$PUBLICBASEDIR/${LOCLIST[$QLOC]}";;
        *) QDIR=$d;;
    esac
}


function getFileVersion {
    local DIR="$1"
    local FILEBASE="$2"
    local EXTENSION="$3"
    local FILEVERSION

    FILEVERSION=$(ls "$DIR/$FILEBASE"*".$EXTENSION" | sed  -n "s/^.*_v\([0-9]*\)\.${EXTENSION}/\1/p" | sort -nr | head -1)
    if [[ -n $FILEVERSION && $FILEVERSION -gt "0" ]]; then
        echo $FILEVERSION
    elif [ -f "$DIR/$FILEBASE.$EXTENSION" ]; then
        echo 1
    else
        echo 0
    fi
}


function logThis {
    echo "$1" >> "$GFILOG"
}


function exceptionReportDB {
    # argumenets locid,year,month

    sqlite3 $EXCEPTIONREPORTDB "insert into exceptionreports(locid,year,month) values ($1,$2,$3)" || echo "DB insert failed"
}


function logException {
    local x
    local logline

    echo -e "\nLog Exception Update:"
    echo -n "Enter description -->"
    read x
    if [ -n "$x" ]; then
        logline=$(date +'%F %T')" Exception report update ${QYEAR}-${QMONTH} ${LOCLIST[$QLOC]} $x"
        logThis "$logline"
        echo "$logline"
        exceptionReportDB "$QLOC" "$QYEAR" "$QMONTH" 

        echo "Thanks - these updates have been entered into the GFI database."
        echo
        eval "echo \"${EXCEPTIONACTIONS[$QLOC]}\""
        echo
        echo "Generate MSR and MRSR?"
        read x
        if [ "$x" = 'y' ]; then
            QDIR="$PUBLICBASEDIR/${LOCLIST[$QLOC]}"
            monthlySummaryReport
            monthlyRouteSummaryReport
        fi
    else
        echo "Nothing logged"
    fi
    read x
}


function logExceptionReportsSent {
    echo "Exception reports sent: $QYEAR, $QMONTH"
    sqlite3 $EXCEPTIONREPORTDB "insert into exceptionreportsent(year,month) values ($QYEAR,$QMONTH)" || echo "DB insert failed"
    read x
}


function logEvent {
    local x
    local logline

    echo -e "\nLog Event:"
    echo -n "Enter description -->"
    read x
    if [ -n "$x" ]; then
        logline=$(date +'%F %T')" $x"
        logThis "$logline"
    fi
    echo "$logline"
    read x
}


function monthlySummaryReport {
    echo -e "\nMonthly summary report"
    QFILE="${LOCLIST[$QLOC]}_GFImonthlySummaryReport_$QYEAR-${QMONTH}"
    echo "Filename: $QFILE"
    FILEVERSION=$(getFileVersion "$QDIR" "$QFILE" "xlsx")

    if [ $FILEVERSION -gt "0" ]; then
        echo -n "Current version is $FILEVERSION -- Increment version number?"
        read v
        if [ $v = "y" ]; then
            ((FILEVERSION++))
            QFILE="${QFILE}_v${FILEVERSION}"
        fi
    fi
    echo "python generateMSR.py -y $QYEAR -m $QMONTH -l $QLOC -f $QDIR/$QFILE.xlsx -c $QCONNECTION "
    python generateMSR.py -y $QYEAR -m $QMONTH -l $QLOC -f "$QDIR/$QFILE.xlsx" -c $QCONNECTION
}

function monthlyRouteSummaryReport {
    echo -e "\nMonthly route summary report"
    QFILE="${LOCLIST[$QLOC]}_GFImonthlyRouteSummaryReport_$QYEAR-${QMONTH}"
    echo "Filename: $QFILE"
    FILEVERSION=$(getFileVersion "$QDIR" "$QFILE" "xlsx")

    if [ $FILEVERSION -gt "0" ]; then
        echo -n "Current version is $FILEVERSION -- Increment version number?"
        read v
        if [ $v = "y" ]; then
            ((FILEVERSION++))
            QFILE="${QFILE}_v${FILEVERSION}"
        fi
    fi
    echo "python generateMRSR.py -y $QYEAR -m $QMONTH -l $QLOC -f $QDIR/$QFILE.xlsx -c $QCONNECTION"
    python generateMRSR.py -y $QYEAR -m $QMONTH -l $QLOC -f "$QDIR/$QFILE.xlsx" -c $QCONNECTION
}

function exceptionReport {
    echo -e "\nException report"
    QFILE="${LOCLIST[$QLOC]}_GFIexceptionReport_$QYEAR-${QMONTH}"
    echo "Filename: $QFILE"
    FILEVERSION=$(getFileVersion "$QDIR" "$QFILE" "xlsx")

    if [ $FILEVERSION -gt "0" ]; then
        echo -n "Current version is $FILEVERSION -- Increment version number?"
        read v
        if [ $v = "y" ]; then
            ((FILEVERSION++))
            QFILE="${QFILE}_v${FILEVERSION}"
        fi
    fi
    echo "python generateExceptionReport.py -y $QYEAR -m $QMONTH -l $QLOC -f $QDIR/$QFILE.xlsx -c $QCONNECTION"
    python generateExceptionReport.py -y $QYEAR -m $QMONTH -l $QLOC -f "$QDIR/$QFILE.xlsx" -c $QCONNECTION
}


function driverUnclassified {
    echo -e "\nDriver unclassified report"
    echo "python generateDriverReport.py -y $QYEAR -m $QMONTH -l $QLOC -c $QCONNECTION"
    python generateDriverReport.py -y $QYEAR -m $QMONTH -l $QLOC -c $QCONNECTION
}


function driverKey {
    echo -e "\nDriver key report"
    echo "python generateDriverkeyReport.py -y $QYEAR -m $QMONTH -l $QLOC -c $QCONNECTION"
    python generateDriverkeyReport.py -y $QYEAR -m $QMONTH -l $QLOC -c $QCONNECTION
}


function chilliwackMRSR {
    echo -e "\nChilliwack/Agassiz monthly route summary report"
    echo "   will be generated in current directory"
    echo "python genChilliwackMRSR.py -y $QYEAR -m $QMONTH -c $QCONNECTION"
    python genChilliwackMRSR.py -y $QYEAR -m $QMONTH -c $QCONNECTION
    QDIR="."
    QLOC=19
}


function chilliwack11Exception {
    echo -e "\nChilliwack/Agassiz route 11 exception report"
    echo "   will be generated in current directory"
    echo "python genChilliwackRoute11exception.py -y $QYEAR -m $QMONTH -c $QCONNECTION"
    python genChilliwackRoute11exception.py -y $QYEAR -m $QMONTH -c $QCONNECTION
    QDIR="."
    QLOC=19

    echo -e "\nEmail"
    echo "To: Lanine Matthews <Lanine.Matthews@firstgroup.com>"
    echo "Subject: GFI route 11 exception report - $QMONTH $QYEAR"
    echo "Please find attached the GFI route 11 exception report for $QMONTH $QYEAR."
}


function updateWKTdata {
    echo -e "\nWest Kootenay (Trail) Ridership/Revenue Spreadsheet Updater"
    ls *.xlsx
    echo "Enter filename:"
    read f
    echo "python updateWKTdata.py -y $QYEAR -m $QMONTH -c $QCONNECTION -f \"$f\""
    python updateWKTdata.py -y $QYEAR -m $QMONTH -c $QCONNECTION -f "$f"
    QDIR="."
    QLOC=11
}


while : ; do
    #clear
    echo -e "\n\n\n"
    echo "----------------------------------------"
    echo "GFI Reporting"
    echo "----------------------------------------"
    echo "  Year:       $QYEAR"
    echo "  Month:      $QMONTH"
    echo "  Location:   ${LOCLIST[$QLOC]}"
    echo "  Directory:  $QDIR"
    echo "  Connection: $QCONNECTION"
    echo "----------------------------------------"
    echo "  [y] change year"
    echo "  [m] change month"
    echo "  [l] change location"
    echo "  [d] change directory"
    echo "  [v] view log"
    echo "  --------------------------------------"
    echo "  [1] log exception report"
    echo "  [e] exception reports sent"
    echo "  [2] log event"
    echo "  [3] monthly summary report"
    echo "  [4] monthly route summary report"
    echo "  [5] exception report"
    echo "  [11] Chilliwack route 11 exception report"
    echo "  [6] driver unclassified report (best & worst)"
    echo "  [7] driver key report"
    echo "  [8] Chilliwack/Agassiz monthly route summary report (MRSR)"
    echo "  [9] West Kootenay (Trail) Spreadsheet updater"
    echo -n "--> "
    read x

    case $x in
        y) echo -n "New year:"; 
           read QYEAR;;
        m) echo -n "New month:"; 
           read m; 
           m="000$m";
           QMONTH=${m:(-2)};;
        l) getLocation;;
        d) getDirectory;;
        v) less -i "$GFILOG";;
        1) logException;;
        e) logExceptionReportsSent;;
        2) logEvent;;
        3) monthlySummaryReport;read x;;
        4) monthlyRouteSummaryReport;read x;;
        5) exceptionReport;read x;; 
        6) driverUnclassified;read x;; 
        7) driverKey;read x;; 
        8) chilliwackMRSR;read x;; 
        9) updateWKTdata;read x;; 
        11) chilliwack11Exception;read x;; 
        *) echo "Huh?";
           echo "Press a key..."; read;;
    esac
done


