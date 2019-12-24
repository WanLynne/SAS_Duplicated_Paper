%let path = %str(C:\Users\user\Desktop\F_Final);

data IPOApplyData;
	infile "&path\IPOInformation_1221.csv"  dlm=',' firstobs=2 dsd;
	input Name :$9. IPODate :yymmdd10. TejIndustry :$15. SettingDate :yymmdd10. OfferPrice NumberOfShares :comma9.  
			AnnounceDate :yymmdd10. UnderwritingDateStart  :yymmdd10. UnderwritingDateEnd :yymmdd10. WiningRate 
			Market $ IPODatePrice  IPODateAmount :comma9. Date1AfterIPO :yymmdd10. Day1Price Day2Price Day3Price Day4Price Day5Price;
	format IPODate yymmdd10. SettingDate  yymmdd10. AnnounceDate yymmdd10. UnderwritingDateStart yymmdd10. UnderwritingDateEnd yymmdd10. Date1AfterIPO yymmdd10.;
run;

* 計算首日與期初報酬;
data IPOApplyDateWithNewVar;
	set IPOApplyData;
		FirstDayReturn = (IPODatePrice/ OfferPrice - 1);
		Day1Return = (Day1Price/ IPODatePrice - 1);
		Day2Return = (Day2Price/ Day1Price - 1);
		Day3Return = (Day3Price/ Day2Price - 1);
		Day4Return = (Day4Price/ Day3Price - 1);
		Day5Return = (Day5Price/ Day4Price - 1); 
		InitialReturn = ((1+FirstDayReturn)*(1+Day1Return)*(1+Day2Return)*(1+Day3Return)*(1+Day4Return)*(1+Day5Return))**(1/5)-1;
		Year = year(IPODate);
		Month = month(IPODate);
		if Month = 1 then do; 
			YearMonthPrior1Month = cat(Year-1, "/", 12);
			end;
		else do; 
			YearMonthPrior1Month = cat(Year, "/", Month-1);
			end;
	keep Name IPODate FirstDayReturn InitialReturn YearMonthPrior1Month;
run;

* 建立DataToMerge 準備拿去merge;
data DataToMerge;
	retain Name YearMonthPrior1Month FirstDayReturn InitialReturn;
	set IPOApplyDateWithNewVar;
	keep Name YearMonthPrior1Month FirstDayReturn InitialReturn;
run;

* 建立merge 後的data;
proc sql;
	create table MergeData 
	as select a.*, b.Name as OtherFirm, b.FirstDayReturn as OtherIPOFirstDayR, b.InitialReturn as OtherIPOInitialDayReturn
	from IPOApplyDateWithNewVar a left join DataToMerge b
	on a.YearMonthPrior1Month = b.YearMonthPrior1Month
	order by Name, YearMonthPrior1Month;
quit;

data MergeDataWithDelete;
	set MergeData;
	if Name = OtherFirm then delete;
run;

proc sql;
	create table CalculateMean
	as select *, mean(OtherIPOFirstDayR) as OtherIPOFirstDayR_Mean, mean(OtherIPOInitialDayReturn) as OtherIPOInitialDayReturn_Mean
	from MergeDataWithDelete
	group by Name
	order by Name;
quit;

proc sort data = CalculateMean;
	by Name IPODate;
	run;

data Final;
	set CalculateMean;
	by Name IPODate;
	if first.Name;
	keep Name OtherIPOFirstDayR_Mean OtherIPOInitialDayReturn_Mean;
run;

*** Output the dataset as XLS file ***;
PROC EXPORT DATA= Final
            OUTFILE= "&path\OtherIPOPriorReturn.xlsx" 
            DBMS=EXCEL REPLACE;
     SHEET="sheet 1"; 
     NEWFILE=YES;
RUN;
	









	































