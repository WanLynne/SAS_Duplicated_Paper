
* 計算其他IPO首日與期初報酬 ;

*  Macro Function  計算其他IPO首日與期初報酬; 
%MACRO CalculateAverageReturnPriorIPO (Name= ,Time= );
	proc sql;
		create table test as select Name, IPODate, FirstDayReturn, InitialReturn, YearMonth
		from IPOWithOtherReturn
		where YearMonth =  &Time and Name ^= &Name;
	quit;

	proc means data = test noprint;
		var FirstDayReturn InitialReturn;
		output out = test1;
	quit;

data test2;
	retain Name FirstDayReturn InitialReturn;
	set test1;
	if _STAT_ = "MEAN";
	Name = &Name;
	keep Name FirstDayReturn InitialReturn;
	rename FirstDayReturn = FirstDayReturn_Mean InitialReturn = InitialReturn_Mean;
run;
	&test2
%MEND CalculateAverageReturnPriorIPO;

%let TimeSign = "2001/4";
%put &TimeSign;
%let NameSign = "1476 儒鴻";
proc sql;
	create table test as select Name, IPODate, FirstDayReturn, InitialReturn, YearMonth
	from IPOWithOtherReturn
	where YearMonth =  &TimeSign and Name ^= &NameSign;
quit;
proc means data = Test noprint;
	var FirstDayReturn InitialReturn;
	output out = Test1;
quit;
data test2;
	retain Name FirstDayReturn InitialReturn;
	set test1 ;
	if _STAT_ = "MEAN";
	Name = &NameSign;
	keep Name FirstDayReturn InitialReturn;
	rename FirstDayReturn = FirstDayReturn_Mean InitialReturn = InitialReturn_Mean;
run;

%CalculateAverageReturnPriorIPO (Name =  &NameSign, Time = &TimeSign);
%CalculateAverageReturnPriorIPO (Name =  &NameSign, Time = &TimeSign);
