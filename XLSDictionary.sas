/**********************************************************************************************************************************
MACRO for creating Excel-data dictionary for a dataset
***********************************************************************************************************************************/

%macro XLSDictionary(
						outpath=,
						/*Path to the output location*/

						out=,
						/*Name of output xls-file*/

						data=,
						/*Input dataset*/

						categ=,
						/*numeric/date variables that should be summarized as categorical*/

						excl=
						/*variables that should be excluded*/

					);

proc format;
		picture mypct low-high='000,009.999%';
run;

/*Extracting library and dataset name for input data*/
%let dotpos=%index(&data,.)	;
%if dotpos^=0 %then %do;
	%let dsn=%substr(&data,%eval(&dotpos+1),%eval(%length(&data)-&dotpos));
	%let lib=%substr(&data,1,%eval(&dotpos-1));
%end;
%else %do;
	%let dsn=&data;
	%let lib=WORK ;
%end;

/*Adding quotes around variables that should be considered categorical (from &categ) or excluded
First, add quote in the start and end, then replacing blanks between values with quote-blank-quote
and then upcase-ing everything*/
%let categ_q=%upcase(%sysfunc(transtrn(%str(%')&categ%str(%'),%str( ),%str(' '))));
%let excl_q=%upcase(%sysfunc(transtrn(%str(%')&excl%str(%'),%str( ),%str(' '))));

/*Getting numeric, date and character variables separately*/
proc sql noprint;
	select name into :numvars separated by ' '
	from sashelp.vcolumn
	where libname=upcase("&lib") and memname=upcase("&dsn")
	and	type='num' and format not like 'DATE%' and upcase(name) not in (&categ_q &excl_q);

	select name into :datevars separated by ' '
	from sashelp.vcolumn
	where libname=upcase("&lib") and memname=upcase("&dsn")
	and	type='num' and format like 'DATE%' and upcase(name) not in (&categ_q &excl_q);

	select name into :charvars separated by ' '
	from sashelp.vcolumn
	where libname=upcase("&lib") and memname=upcase("&dsn")
	and	(    type='char' or upcase(name) in (&categ_q)  ) and upcase(name) not in (&excl_q);
quit;


/*creating Excel-file*/
ods tagsets.excelxp path="&outpath"
					file="&out._%sysfunc(left(%sysfunc(date(),date9.))).xls"
					style=Printer;
ods select position;

ods tagsets.ExcelXP options(embedded_titles='yes'
							embedded_footnotes='yes' 
 							sheet_name='Data Dictionary'
							sheet_interval='none');
title "Data Dictionary for SAS table &dsn";
proc contents data=&data order=varnum; run;

ods tagsets.ExcelXP options(embedded_titles='yes'
							embedded_footnotes='yes' 
 							sheet_name='Summary Stats'
							Width_Fudge='1'
							Autofit_height='yes'
							sheet_interval='none');
title "Summary Statistics for SAS table &dsn";

%if %symexist(datevars) %then %do;
title2 'Date variables';

options label=off;
proc tabulate data=&data;
	var &datevars;
	table	&datevars
			,n nmiss (min q1 mean median q3 max)*f=date9.;
run;
title;
%end;

%if %symexist(numvars) %then %do;
title2 'Numeric variables';
proc tabulate data=&data;
	var  &numvars;
	table &numvars
			,n nmiss (min q1 mean median q3 max)*f=best12.;
run;
%end;


%if %symexist(charvars) %then %do;
*formatting values in PROC FREQ output;
ODS PATH RESET;                              
ODS PATH (PREPEND) WORK.Templat(UPDATE) ;    
                                             
PROC TEMPLATE;                               
  EDIT Base.Freq.OneWayList;                 
    EDIT Frequency;                          
      FORMAT = COMMA6.;                      
    END;                                     
    EDIT Percent;                            
      FORMAT = mypct.;                          
    END;                                     
  END;                                       
RUN;        

title2 'Categorical variables';

proc freq data=&data;
	tables	&charvars/ nocum missing;
run;


ods tagsets.excelxp close;

title2;
options label=on;

PROC TEMPLATE;
delete Base.Freq.OneWayList;
run;
%end;
/******************************************
FORMATTING WITH DDE
*******************************************/

/*starting Excel*/
options noxsync noxwait;
filename sas2xl dde 'excel|system';

data _null_;
	length fid rc start stop time 8;
	fid=fopen('sas2xl','s');
	if (fid le 0) then do;
		rc=system('start excel');
		start=datetime();
		stop=start+10;
		do while (fid le 0);
			fid=fopen('sas2xl','s');
			time=datetime();
			if (time ge stop) then fid=1;
		end;
	end;
rc=fclose(fid);
run;


/*open workbook and format*/
data _null_;
	file sas2xl;
	put '[error(false)]';
	put "[open(""&outpath.\&out._%sysfunc(left(%sysfunc(date(),date9.))).xls"")]";
	put	'[column.width(0,"c1:c6",false,3)]';
	put	'[column.width(200,"c7")]';
	put '[select("r1c1")]';
	put '[format.font("Thorndale AMT",13,true,false,false,false,0,false,false)]';
	put '[workbook.activate("Summary Stats")]';
	put	'[column.width(0,"c1:c9",false,3)]';
	put '[select("r1c1")]';
	put '[format.font("Thorndale AMT",13,true,false,false,false,0,false,false)]';
	put '[workbook.activate("Data Dictionary")]';
	put '[save()]';
	put '[file.close(false)]';
	put '[error(true)]';
run;

%mend;

/*Example of using*/
/*%XLSDictionary(*/
/*						outpath=R:\Dmitry,*/
/**/
/*						out=test,*/
/**/
/*						data=test.req241_careretention_obj12,*/
/**/
/*						categ=ha aids,*/
/**/
/*						excl=moh_id*/
/*					)*/
