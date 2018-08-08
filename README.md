# utl_put_excel_sheetnames_into_sas_macro_variable
utl_put_excel_sheetnames_into_sas_macro_variable.  Keywords: sas sql join merge big data analytics macros oracle teradata mysql sas communities stackoverflow statistics artificial inteligence AI Python R Java Javascript WPS Matlab SPSS Scala Perl C C# Excel MS Access JSON graphics maps NLP natural language processing machine learning igraph DOSUBL DOW loop stackoverflow SAS community.
    Put excel sheetnames into sas macro variable

       Macro on end

       This will even get all sheet with 31 character names.
       Will work with IML/R

       WORKING CODE  (31 chars appears to be the max for XLConnect)

             wb <- loadWorkbook("d:/xls/longnames.xlsx");
             sheets<-getSheets(wb);


    see
    https://tinyurl.com/y6vmg5de
    https://communities.sas.com/t5/Base-SAS-Programming/Excel-sheet-name/m-p/483927

    see
    https://goo.gl/6ztMWo
    https://communities.sas.com/t5/Base-SAS-Programming/Query-excel-to-find-sheet-name/m-p/45477/highlight/true#M9403

    HAVE workbook d:/xls/longnames.xlsx with three sheets
    ======================================================

      THREE SHEETS

          A123456789012345678901234567890   ** 31 characters

          B123456789012345678901234567890   ** 31 characters

          B1234567890                       ** 11 characters



    WANT SAS macro variable 'sheets'
    ===============================

     %put &=sheets

      sheets= A123456789012345678901234567890
              B123456789012345678901234567890
              B1234567890


    *                _              _               _
     _ __ ___   __ _| | _____   ___| |__   ___  ___| |_ ___
    | '_ ` _ \ / _` | |/ / _ \ / __| '_ \ / _ \/ _ \ __/ __|
    | | | | | | (_| |   <  __/ \__ \ | | |  __/  __/ |_\__ \
    |_| |_| |_|\__,_|_|\_\___| |___/_| |_|\___|\___|\__|___/

    ;

    %utlfkil(d:/xls/longnames.xlsx);   * delete if exist;

    %utl_submit_r64('
       source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);
       library("XLConnect");
       wb <- loadWorkbook("d:/xls/longnames.xlsx",create=TRUE);
       createSheet(wb, name="A123456789012345678901234567890");
       writeWorksheet(wb,iris, sheet = "A123456789012345678901234567890");
       createSheet(wb, name="B123456789012345678901234567890");
       writeWorksheet(wb,iris, sheet = "B123456789012345678901234567890");
       createSheet(wb, name="B1234567890");
       writeWorksheet(wb,iris, sheet = "B1234567890");
       saveWorkbook(wb);
    ');

    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| '_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|

    ;

    %symdel sheets / nowarn;
    %utl_submit_r64('
       source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);
       library("XLConnect");
       wb <- loadWorkbook("d:/xls/longnames.xlsx");
       sheets<-getSheets(wb);
       sheets;
       writeClipboard(as.character(paste(sheets, collapse = " ")));
    ',returnVar=sheets)

     %put &=sheets;

    *      _   _               _               _ _             __   _  _
     _   _| |_| |    ___ _   _| |__  _ __ ___ (_) |_     _ __ / /_ | || |
    | | | | __| |   / __| | | | '_ \| '_ ` _ \| | __|   | '__| '_ \| || |_
    | |_| | |_| |   \__ \ |_| | |_) | | | | | | | |_    | |  | (_) |__   _|
     \__,_|\__|_|___|___/\__,_|_.__/|_| |_| |_|_|\__|___|_|   \___/   |_|
               |_____|                             |_____|
    ;

    %macro utl_submit_r64(
          pgmx
         ,returnVar=N           /* set to Y if you want a return SAS macro variable from python */
         )/des="Semi colon separated set of R commands - drop down to R";
      * write the program to a temporary file;
      %utlfkil(d:/txt/r_pgm.txt);
      filename r_pgm "d:/txt/r_pgm.txt" lrecl=32766 recfm=v;
      data _null_;
        length pgm $32756;
        file r_pgm;
        pgm=&pgmx;
        put pgm;
        putlog pgm;
      run;
      %let __loc=%sysfunc(pathname(r_pgm));
      * pipe file through R;
      filename rut pipe "c:\Progra~1\R\R-3.3.2\bin\x64\R.exe --vanilla --quiet --no-save < &__loc";
      data _null_;
        file print;
        infile rut recfm=v lrecl=32756;
        input;
        put _infile_;
        putlog _infile_;
      run;
      filename rut clear;
      filename r_pgm clear;

      * use the clipboard to create macro variable;
      %if %upcase(%substr(&returnVar.,1,1)) ne N %then %do;
        filename clp clipbrd ;
        data _null_;
         length txt $200;
         infile clp;
         input;
         putlog "macro variable &returnVar = " _infile_;
         call symputx("&returnVar.",_infile_,"G");
        run;quit;
      %end;

    %mend utl_submit_r64;


