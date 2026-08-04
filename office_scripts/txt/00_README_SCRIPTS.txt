AMKAD AR REPORTING — OFFICE SCRIPTS  (HANDOVER)
================================================

This folder contains the four Office Scripts that run the AMKAD AR
reporting, as plain text files. Each .txt is the complete, current code.

If you have not worked with this automation before, read the whole of this
file first. It is short, and it will save you a lot of guessing.


-------------------------------------------------------------------------
1. WHAT AN OFFICE SCRIPT IS, AND WHERE TO FIND THESE
-------------------------------------------------------------------------
An Office Script is code that runs inside Excel Online. It is not a macro
and it is not VBA, and it does not live inside the workbook file — it is
stored in OneDrive and attached to a workbook when you run it.

To see them:
   1. Open the workbook in Excel ONLINE (in a browser — not desktop Excel,
      which does not have this feature)
   2. Automate tab → All Scripts
   3. The four scripts are listed there

Power Automate runs them through an action called "Run script", where you
pick the workbook and then the script by name.

IMPORTANT: if you rename a script in Excel, the "Run script" action in the
flow does not follow the rename. It keeps pointing at the old name and
fails. If you must rename one, open every flow that uses it and re-select
it afterwards.


-------------------------------------------------------------------------
2. THE ONE CONSTRAINT THAT EXPLAINS THE DESIGN
-------------------------------------------------------------------------
An Office Script can only read and write the ONE workbook it was launched
against. It cannot open a second file.

This is why the daily build needs two scripts instead of one. It has to
read the raw source file and write to a different workbook, which is
impossible in a single script. So one script reads and returns the data as
text, Power Automate holds that text, and a second script writes it into
the other workbook.

If that seems roundabout, that is why. It is not a stylistic choice.


-------------------------------------------------------------------------
3. THE THREE FLOWS, IN THE ORDER THEY RUN EACH DAY
-------------------------------------------------------------------------

   FLOW 1  Daily File Build
           Builds today's Daily_Performance_Report .xlsm
           Uses: custsol_read_daily  +  custsol_append_rawdata_eur

   FLOW 2  Trigger Daily Refresh
           Opens the input workbook so Power Query refreshes
           Uses: refresh_trigger

   FLOW 3  AMKAD AR Report
           Produces and sends the HTML report
           Uses: amkad_ar_report

They must run in this order, with gaps between them. Flow 3 reads data that
Flow 1 created and Flow 2 refreshed. Running them out of order produces a
report built on yesterday's numbers, with no error message.


=========================================================================
FLOW 1 — DAILY FILE BUILD
=========================================================================
PURPOSE
   Every day a raw file arrives from Cust Sol. This flow takes that file
   and produces the day's Daily_Performance_Report .xlsm, which is the
   file the rest of the reporting reads.

HOW THE HISTORY IS KEPT
   The flow does NOT build the daily file from scratch. It copies
   yesterday's file, renames the copy for today, and appends one day of
   rows to it. Every daily file therefore contains the full history.

   Consequence worth understanding: if a run fails, that day is missing
   from the copy made the next day, and from every file after that. The
   RawData query in the input workbook now repairs this automatically by
   reading the original source files, but the daily files themselves stay
   permanently short that day.


   -- SCRIPT: custsol_read_daily.txt --------------------------------

   RUNS ON      The SOURCE file — the raw report that arrives in
                …/Daily/Cust_Sol Daily File/
   GIVES BACK   Text (JSON) containing all the rows it read, plus counts
   GOES TO      The second script, passed along by Power Automate

   WHAT IT DOES
   Reads every data row from the raw file and returns it as text.

   It finds columns by looking up the HEADER NAME, not by column letter.
   So if the source file ever reorders its columns, nothing breaks. If a
   column is RENAMED, the script stops with a message naming the column it
   could not find — which tells you exactly what changed.

   The source file has a column called "Total Payments" TWICE — once in
   local currency, once in euros. The script deliberately takes the second
   one for the euro figure. Do not "fix" this.

   The count called missingEuroRows is rows where the euro columns were
   empty. This is normal for a day or two after a month closes. It is
   reported so you can see it, not treated as a failure.


   -- SCRIPT: custsol_append_rawdata_eur.txt ------------------------

   RUNS ON      The NEW daily file (the copy the flow just made)
   TAKES        The text produced by custsol_read_daily
   GIVES BACK   rowsAppended, skippedDuplicates, tableRowCount
   WRITES TO    The table named RawData inside that file

   WHAT IT DOES
   Adds the day's rows to the RawData table.

   SAFE TO RE-RUN. Before adding anything it compares Date + Country +
   Customer against the rows already there and skips any match. So if you
   re-run a day, you get skippedDuplicates instead of doubled figures.
   After any re-run, check that number — it should equal the row count for
   that day, and rowsAppended should be 0.

   TWO TRAPS IF YOU EDIT THIS SCRIPT
     - The date column in RawData is named "Month", not "Date", even
       though it holds a daily date. Looking for "Date" fails.
     - Excel throws an error if you try to add an empty list of rows, so
       there is a check for that before the append. Leave it in.


=========================================================================
FLOW 2 — TRIGGER DAILY REFRESH
=========================================================================

   -- SCRIPT: refresh_trigger.txt -----------------------------------

   RUNS ON      The input workbook (AMKAD_Current_Input.xlsm)
   GIVES BACK   Nothing

   WHAT IT DOES
   Almost nothing, on purpose. It writes one line to the log and stops.

   Its only job is to OPEN the workbook, because opening the workbook is
   what makes Power Query refresh the RawData query ("Refresh data when
   opening the file" is switched on in the query settings).

   WHY IT IS A SEPARATE FLOW
   So there is a gap between the refresh starting and the report reading
   the data. Nothing actually confirms the refresh has finished — the gap
   in the schedule is the only thing making it work.

   THEREFORE
     - This flow must stay switched on
     - It must stay scheduled comfortably BEFORE Flow 3
     - If the report ever shows yesterday's numbers, check this first

   It looks pointless. It is not. Do not delete it.


=========================================================================
FLOW 3 — AMKAD AR REPORT
=========================================================================

   -- SCRIPT: amkad_ar_report.txt -----------------------------------

   RUNS ON      The input workbook (AMKAD_Current_Input.xlsm)
   GIVES BACK   A large block of text containing ready-made HTML tables
                and the data the charts read
   GOES TO      The "Compose HTML Report" action in the flow, which drops
                it into the HTML template

   WHAT IT DOES
   This is where every number in the report is calculated. The HTML
   template only decides where things sit on the page. If a figure is
   wrong, it is wrong here, not in the template.

   IT READS FOUR THINGS FROM THE WORKBOOK
     RawData            The daily AR history
     MonthEndAR         The latest closed month, used by the SPR tab
     SPR_<CC> sheets    Collector commentary. Found by table name PREFIX
                        (SPRIssues_ and SPRForecast_), which means adding
                        a new country needs NO code change — just create
                        the sheet and name the tables correctly
     ReportConfigTbl    The risk thresholds

   THE MOST IMPORTANT RULE IN THE WHOLE REPORT
   For each month it keeps only the LATEST snapshot of that month.

   AR is a balance, not a total. You cannot add up daily snapshots. Twenty
   daily snapshots of a €4 million book would add up to €80 million, which
   is meaningless. This rule is why that never happens.

   FOUR THINGS THAT LOOK LIKE MISTAKES BUT ARE DELIBERATE
     1. DSO and TDSO are read from the data but left out of every
        calculated figure. They are ratios, and averaging them across
        customers treats a €2,000 account as equal to an €800,000 one.
        A real portfolio DSO has to be recalculated from the totals.

     2. The aging buckets are nested, not separate: over-90 sits inside
        over-60, which sits inside overdue, which sits inside total AR.
        Never stack them in a chart or put them in a pie — the same money
        would be counted more than once.

     3. Colours follow whether a movement is GOOD, not which direction it
        points. A fall in AR or overdue is green. A rise in payments is
        green. If you add a metric, you must state which direction is
        favourable for it.

     4. On the Action Required tab, the suggested steps are an
        approximation mapped from how old a balance is. The report does
        not know WHY a balance is old, which matters just as much. Every
        suggestion is worded as a potential step needing review, and
        nothing affecting a customer's service is ever suggested
        automatically. Keep it that way.


=========================================================================
IF YOU NEED TO CHANGE A SCRIPT
=========================================================================
   1. Open the .txt file here and copy all of it
   2. Excel Online → Automate → All Scripts → open the matching script
   3. Select all, paste, Save
   4. Run the flow once manually and check it succeeded
   5. Update the copy kept in the shared SharePoint folder next to the
      workbooks

Step 5 matters. There is no build process and no other backup — the copy
in Excel Online and the copy in SharePoint are the only two that exist,
and if they drift, nobody can tell which one is right.

Test on a copy of the workbook before changing anything that writes data.
custsol_append_rawdata_eur writes into RawData, and a bad version of it
can corrupt history that is tedious to rebuild.
