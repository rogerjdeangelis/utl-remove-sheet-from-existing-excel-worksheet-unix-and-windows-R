# utl-remove-sheet-from-existing-excel-worksheet-unix-and-windows-R
Remove sheet from existing excel workbook R
    Remove sheet from existing excel workbook R

    SAS IML has an interface to R in Unix and Windows;

    Problem;
       Remove sheet "class' from workbook but keep sheet students


    SAS Forum
    https://tinyurl.com/y65w5gct
    https://communities.sas.com/t5/New-SAS-User/Deleting-an-Excel-sheet-with-SAS-on-a-unix-server/m-p/561080

    *_                   _
    (_)_ __  _ __  _   _| |_
    | | '_ \| '_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    ;

    * I assume youalready have a eorkbook and sheet you need to remove;
    * Hopefully SAS will eventuall deprecate 'proc import' and beef up the libname engines;

    * create workbook d:/xls/class.xlsx with sheets CLASS and STUDENTS;

    %utlfkil(d:/xls/class.xlsx); * delete if exist;

    libname xel "d:/xls/class.xlsx";
    data xel.class xel.students(keep=name);
       set sashelp.class(keep=name age);
    run;quit;
    libname xel clear;


    WORKBOOK d:/xls/class.xlsx

       d:/xls/class.xlsx
          +---------=-------+
          |     A   |  B    |
          +---------+-------+
       1  | NAME    | Age   |
          +---------+-------+
       2  | ALFRED  | 11    |
          +---------+-------+
       3  | BARBARA | 11    |
          +---------+-------+
       3  | DAVID   | 11    |
          +---------+-------+
           ...
          +---------+-------+
       20 | WILLIAM | 16    |
          +---------+-------+

       [CLASS]


    WORKBOOK d:/xls/class.xlsx

       d:/xls/class.xlsx

          +---------+
          |     A   |
          +---------+
       1  | NAME    |
          +---------+
       2  | ALFRED  |
          +---------+
       3  | BARBARA |
          +---------+
       3  | DAVID   |
          +---------+
           ...
          +---------+
       20 | WILLIAM |
          +---------+

       [STUDENTS]

    libname xel "d:/xls/class.xlsx";
    proc contents data=xel._all_;
    run;quit;
    libname xel clear;


    The CONTENTS Procedure

               Directory

    Libref         XEL
    Engine         EXCEL
    Physical Name  d:/xls/class.xlsx
    User           Admin

                          DBMS
                  Member  Member
    #  Name       Type    Type

    1  class$     DATA    TABLE
    2  students$  DATA    TABLE

    *            _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| '_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    ;

    The CONTENTS Procedure

               Directory

    Libref         XEL
    Engine         EXCEL
    Physical Name  d:/xls/class.xlsx
    User           Admin


                          DBMS
                  Member  Member
    #  Name       Type    Type

    1  students$  DATA    TABLE

    *
     _ __  _ __ ___   ___ ___  ___ ___
    | '_ \| '__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    ;


    %utl_submit_r64('
    library(XLConnect);
    wb <- loadWorkbook("d:/xls/class.xlsx");
    removeSheet(wb, sheet = "CLASS");
    saveWorkbook(wb);
    ');


    libname xel "d:/xls/class.xlsx";
    proc contents data=xel._all_;
    run;quit;
    libname xel clear;

    The CONTENTS Procedure

               Directory

    Libref         XEL
    Engine         EXCEL
    Physical Name  d:/xls/class.xlsx
    User           Admin

                          DBMS
                  Member  Member
    #  Name       Type    Type

    1  students$  DATA    TABLE


    FYI

     SAS also creates a named range for class (ODBS table)

     You can clear the contents of the class table (but keep column names using

    libname xel "d:/xls/class.xlsx";
    proc sql;
      drop table xel.class
    ;quit;

    I expect you can also do tthis in R with sqldf package?


