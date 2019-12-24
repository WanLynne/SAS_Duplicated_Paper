%let path = %str(C:\Users\lynne\Desktop\F_Final);

data IPOApplyData;
 infile "&path\IPOInformation_1221.csv"  dlm=',' firstobs=2 dsd;
 input Name :$9. IPODate :yymmdd10. TejIndustry :$15. SettingDate :yymmdd10. OfferPrice NumberOfShares :comma9.  
   AnnounceDate :yymmdd10. UnderwritingDateStart  :yymmdd10. UnderwritingDateEnd :yymmdd10. WiningRate 
   Market $ IPODatePrice  IPODateAmount :comma9. Date1AfterIPO :yymmdd10. Day1Price Day2Price Day3Price Day4Price Day5Price;
 format IPODate yymmdd10. SettingDate  yymmdd10. AnnounceDate yymmdd10. UnderwritingDateStart yymmdd10. UnderwritingDateEnd yymmdd10. Date1AfterIPO yymmdd10.;
 keep Name UnderwritingDateEnd Market;
run;

data MarketIndexReturn;
infile "&path\Market_Index_Return.csv"  dlm=',' firstobs=2 dsd;
input Date: yymmdd10. TSE_Index OTC_Index;
format Date yymmdd10.;
run;
*** 將TSE及OTC市場分為兩的資料 ***;
data TSE;
	set IPOApplyData;
	where Market = "TSE";
run;
data OTC;
	set IPOApplyData;
	where Market = "OTC";
run;

***** TSE *****;
*** 將該公司在上市日的市場報酬Merge ***;
proc sql;
	create table TSE_Market 
	as select a.Name, a.UnderwritingDateEnd, b.TSE_Index, b.date
	from TSE a left join MarketIndexReturn b
	on intnx('weekday',b.date,1)<=a.UnderwritingDateEnd<=intnx('weekday',b.date,20)      /* 挑選20天(不包含周休二日)，擴大選取範圍 */
	order by a.Name, b.date;
quit;
*** 抓取上市日當天前10天的資料 ***;
*** 先排序資料，日期的資料要倒推 ***;
proc sort data=TSE_Market;
	by Name descending Date;
run;
*** 計算每一家公司有幾筆資料 ***;
data TSE_Market_Count;
set TSE_Market;
by Name descending Date;
retain n 0;
n=n+1;
if first.Name then n=1;
run;
*** 想抓在n為1-10(上市日前10天)的資料，並將報酬從%調整為數字 ***;
proc sql;
	create table TSE_Market_Final (drop=n)
	as select *, TSE_Index/100 as MarketReturn
	from TSE_Market_Count
	where 1<=n<=10
	order by Name, n;
quit;

proc sort data=TSE_Market_Final;
by Name Date;
run;

*** 計算市場投資組合持有期間報酬 ***;
data TSE_Market_Return;
set TSE_Market_Final;
by Name ;
retain MarketReturnPrior10Days;
if first.Name then do;
	MarketReturnPrior10Days = (1+MarketReturn);
end;
else do;
	MarketReturnPrior10Days = MarketReturnPrior10Days * (1+MarketReturn);
end;
if last.Name;
keep Name MarketReturnPrior10Days;
run;

data TSE_Market_Final_Return;
	set TSE_Market_Return;
	MarketReturnPrior10Days = MarketReturnPrior10Days-1;
run;


***** OTC *****;
*** 將該公司在上市日的市場報酬Merge ***;
proc sql;
	create table OTC_Market 
	as select a.Name, a.UnderwritingDateEnd, b.OTC_Index, b.date
	from OTC a left join MarketIndexReturn b
	on intnx('weekday',b.date,1)<=a.UnderwritingDateEnd<=intnx('weekday',b.date,20)      /* 挑選20天(不包含周休二日)，擴大選取範圍 */
	order by a.Name, b.date;
quit;
*** 抓取上市日當天前10天的資料 ***;
*** 先排序資料，日期的資料要倒推 ***;
proc sort data=OTC_Market;
	by Name descending Date;
run;
*** 計算每一家公司有幾筆資料 ***;
data OTC_Market_Count;
set OTC_Market;
by Name descending Date;
retain n 0;
n=n+1;
if first.Name then n=1;
run;
*** 想抓在n為1-10(上市日前10天)的資料，並將報酬從%調整為數字 ***;
proc sql;
	create table OTC_Market_Final (drop=n)
	as select *, OTC_Index/100 as MarketReturn
	from OTC_Market_Count
	where 1<=n<=10
	order by Name, n;
quit;

proc sort data=OTC_Market_Final;
by Name Date;
run;

*** 計算市場投資組合持有期間報酬 ***;
data OTC_Market_Return;
set OTC_Market_Final;
by Name ;
retain MarketReturnPrior10Days;
if first.Name then do;
	MarketReturnPrior10Days = (1+MarketReturn);
end;
else do;
	MarketReturnPrior10Days = MarketReturnPrior10Days * (1+MarketReturn);
end;
if last.Name;
keep Name MarketReturnPrior10Days;
run;

data OTC_Market_Final_Return;
	set OTC_Market_Return;
	MarketReturnPrior10Days = MarketReturnPrior10Days-1;
run;

*** 將TSE和OTC資料垂直合併 ***;
data MarketReturnPrior10Days;
	set TSE_Market_Final_Return OTC_Market_Final_Return;
run;

*** 匯出資料 ***;
PROC EXPORT DATA= WORK.MarketReturnPrior10Days
            OUTFILE= "&path\MarketReturnPrior10Days.xlsx" 
            DBMS=XLSX REPLACE;
     SHEET="sheet 1"; 
     NEWFILE=YES;
RUN;
