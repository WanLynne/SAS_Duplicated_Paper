%let path = %str(C:\Users\user\Desktop\F_Final);
%put &path;

* 取得每一間公司屬於TSE還是OTC ;
data IPOApplyMarket_Data;
	infile "&path\IPOInformation_1221.csv"  dlm=',' firstobs=2 dsd;
	input Name :$9. IPODate :yymmdd10. TejIndustry :$15. SettingDate :yymmdd10. OfferPrice NumberOfShares :comma9.  
			AnnounceDate :yymmdd10. UnderwritingDateStart  :yymmdd10. UnderwritingDateEnd :yymmdd10. WiningRate 
			Market $ IPODatePrice  IPODateAmount :comma9. Date1AfterIPO :yymmdd10. Day1Price Day2Price Day3Price Day4Price Day5Price;
	format IPODate yymmdd10. SettingDate  yymmdd10. AnnounceDate yymmdd10. UnderwritingDateStart yymmdd10. UnderwritingDateEnd yymmdd10. Date1AfterIPO yymmdd10.;
	keep Name Market;
run;

* 取得每一間公司的每日報酬; 
data EachFirmReturn_Data;
	 infile "&path\EachFirmReturn_Data.csv"  dlm=',' firstobs=2 dsd;
	 input Name :$9. Date :yymmdd10. Return ;
	 format Date yymmdd10.;
	 Return = Return / 100;
run;

*報酬是用比率 不是%; 
data MarketReturnToRM_Data;
	 infile "&path\Market_Return_To_Calculate_RM.csv"  dlm=',' firstobs=2 dsd;
	 input Date :yymmdd10. TSEReturn OTCReturn ;
	 format Date yymmdd10.;
	 TSEReturn = TSEReturn / 100;
	 OTCReturn = OTCReturn / 100;
run;

proc sort data = MarketReturnToRM_Data;
	by Date;
run;

* 合併資料取得公司上市時的市場為何 (TSE or OTC); 
proc sql;
	create table EachFirmReturn
	as select a.*, b.*
	from EachFirmReturn_Data a left join IPOApplyMarket_Data b
	on a.Name = b.Name
	order by Name, Date ;
quit;

* 屬於 TSE的公司 接著合併每日的市場報酬;
data TSEFrimReturn;
	retain Name Date Market Return;
	set  EachFirmReturn;
	where Market = "TSE";
run;

proc sql;
	create table TSEFirmReturnWithMarketReturn
	as select a.*, b.Date, b.TSEReturn
	from TSEFrimReturn a left join MarketReturnToRM_Data b
	on a.Date = b.Date
	order by Name, Date ;
quit;

data TSEFirmReturnWithMarketReturn;
	set TSEFirmReturnWithMarketReturn;
	rename TSEReturn = MarketReturn;
run;

* 屬於 OTC的公司 接著合併每日的市場報酬;
data OTCFrimReturn;
	retain Name Date Market Return;
	set  EachFirmReturn;
	where Market = "OTC";
run;

proc sql;
	create table OTCFirmReturnWithMarketReturn
	as select a.*, b.Date, b.OTCReturn
	from OTCFrimReturn a left join MarketReturnToRM_Data b
	on a.Date = b.Date
	order by Name, Date ;
quit;

data OTCFirmReturnWithMarketReturn;
	set OTCFirmReturnWithMarketReturn;
	rename OTCReturn = MarketReturn;
run;

* 合併TSE與OTC公司 且此處已包含市場報酬;
data EachFrimReturnWithMarketReturn;
	set TSEFirmReturnWithMarketReturn OTCFirmReturnWithMarketReturn;
run;

proc sort data = EachFrimReturnWithMarketReturn;
	by Name Date;
run;

* 將每一筆資料依據公司名稱作編號; 
data Firm_MarketReturnWithNumber;
	set EachFrimReturnWithMarketReturn;
	by Name;
	retain n 0;
		n = n+1;
		if first.Name then n = 1;
run;

proc sort data = Firm_MarketReturnWithNumber;
	by Name Date n;
run;

* 計算每一間公司與市場各期的ＢＨＲ (但不包含20~21 20~ 750);
data CalculateFirm_MarketBHR;
	set Firm_MarketReturnWithNumber;
	by Name;
	retain BHR_NeedMinus1 MarketBHR_NeedMinus1;
	if first.Name then do;
		BHR_NeedMinus1 = (1 + Return);
		MarketBHR_NeedMinus1 = (1 + MarketReturn);
		end;
	else do ;
		 BHR_NeedMinus1 = BHR_NeedMinus1 * (1+Return);
		 MarketBHR_NeedMinus1 = MarketBHR_NeedMinus1 * (1 + MarketReturn);
		end;
	BHR = BHR_NeedMinus1 -1; 
	MarketBHR = MarketBHR_NeedMinus1 - 1;
	drop BHR_NeedMinus1 MarketBHR_NeedMinus1;
	keep Name n BHR MarketBHR;
run;

proc sql;
	create table Frim_MarketBHRForEachPeriod
	as select *
	from CalculateFirm_MarketBHR
	where n in (1, 2, 3, 4, 5, 10, 20, 125, 250, 375, 500, 625, 750)
	order by Name, n;
quit;

* 轉置資料 (各公司各期BHR); 
data FirmBHRForEachPeriod;
	set Frim_MarketBHRForEachPeriod;
	drop MarketBHR;
run;

proc transpose data = FirmBHRForEachPeriod 
						out = TransposeFirmBHRForEachPeriod 
						prefix = n;
						var BHR;
						by Name;
run;

* FirmBHR0_ (除了 20~21的BHR 與 20~750的BHR);
data FirmBHR0_;
	set TransposeFirmBHRForEachPeriod;
	rename n1 = BHR0_1 n2 = BHR0_2 n3 = BHR0_3 n4 = BHR0_4 n5 = BHR0_5 n6 = BHR0_10 n7 = BHR0_20 
				n8 = BHR0_125 n9 = BHR0_250 n10 = BHR0_375 n11 = BHR0_500 n12 = BHR0_625 n13 = BHR0_750;
	drop _NAME_;
run;

* 轉置資料 (市場各期BHR); 
data MarketBHRForEachPeriod;
	set Frim_MarketBHRForEachPeriod;
	drop BHR;
run;

proc transpose data = MarketBHRForEachPeriod 
						out = TransposeMarketBHRForEachPeriod 
						prefix = n;
						var MarketBHR;
						by Name;
run;

* MarketBHR0_ (除了 20~21的BHR 與 20~750的BHR;
data MarketBHR0_;
	set TransposeMarketBHRForEachPeriod;
	rename n1 = MarketBHR0_1 n2 = MarketBHR0_2 n3 = MarketBHR0_3 n4 = MarketBHR0_4 n5 = MarketBHR0_5 n6 = MarketBHR0_10 
				n7 = MarketBHR0_20 n8 = MarketBHR0_125 n9 = MarketBHR0_250 n10 = MarketBHR0_375 n11 = MarketBHR0_500 
				n12 = MarketBHR0_625 n13 = MarketBHR0_750;
	drop _NAME_;
run;

*計算Firm and Market 20~21的BHR;
data Firm_MarketBHR20_21_All;
	set Firm_MarketReturnWithNumber;
	by Name;
	where n in (20, 21);
	retain BHR_NeedMinus1 MarketBHR_NeedMinus1;
	if first.Name then do;
		BHR_NeedMinus1 = (1 + Return);
		MarketBHR_NeedMinus1 = (1 + MarketReturn);
		end;
	else do ;
		 BHR_NeedMinus1 = BHR_NeedMinus1 * (1+Return);
		 MarketBHR_NeedMinus1 = MarketBHR_NeedMinus1 * (1 + MarketReturn);
		end;
	BHR = BHR_NeedMinus1 -1; 
	MarketBHR = MarketBHR_NeedMinus1 - 1;
	drop BHR_NeedMinus1 MarketBHR_NeedMinus1;
	keep Name n BHR MarketBHR;
run;

data Firm_MarketBHR20_21;
	set Firm_MarketBHR20_21_All;
	where n = 21;
	rename BHR = BHR20_21 MarketBHR =  MarketBHR20_21;
	drop n;
run;

*計算Firm and Market 20~750 的BHR;
data Firm_MarketBHR20_750_All;
	set Firm_MarketReturnWithNumber;
	by Name;
	where 20 <= n <= 750;
	retain BHR_NeedMinus1 MarketBHR_NeedMinus1;
	if first.Name then do;
		BHR_NeedMinus1 = (1 + Return);
		MarketBHR_NeedMinus1 = (1 + MarketReturn);
		end;
	else do ;
		 BHR_NeedMinus1 = BHR_NeedMinus1 * (1+Return);
		 MarketBHR_NeedMinus1 = MarketBHR_NeedMinus1 * (1 + MarketReturn);
		end;
	BHR = BHR_NeedMinus1 -1; 
	MarketBHR = MarketBHR_NeedMinus1 - 1;
	drop BHR_NeedMinus1 MarketBHR_NeedMinus1;
	keep Name n BHR MarketBHR;
run;

data Firm_MarketBHR20_750;
	set Firm_MarketBHR20_750_All;
	where n = 750;
	rename BHR = BHR20_750  MarketBHR =  MarketBHR20_750;
	drop n;
run;

* Merge Data ;
proc sql;
	create table Firm_MarketBHRFinal
	as select a.*,  b.*,  c.*, d.*
	from FirmBHR0_ as a
	left join MarketBHR0_ as b
		on a.Name = b.Name
	left join Firm_MarketBHR20_21 as c 
		on a.Name = c.Name
	left join Firm_MarketBHR20_750 as d
		on a.Name = d.Name
	order by Name;
quit;

* 計算 AR;
data FirmBHR_AR;
	retain Name BHR0_1 - BHR0_750;
	set Firm_MarketBHRFinal;
		AR0_1 = BHR0_1 - MarketBHR0_1;
		AR0_2 = BHR0_2 - MarketBHR0_2;
		AR0_3 = BHR0_3 - MarketBHR0_3;
		AR0_4 = BHR0_4 - MarketBHR0_4;
		AR0_5 = BHR0_5 - MarketBHR0_5;
		AR0_10 = BHR0_10 - MarketBHR0_10;
		AR0_20 = BHR0_20 - MarketBHR0_20;
		AR0_125 = BHR0_125 - MarketBHR0_125;
		AR0_250 = BHR0_250 - MarketBHR0_250;
		AR0_375 = BHR0_375 - MarketBHR0_375;
		AR0_500 = BHR0_500 - MarketBHR0_500;
		AR0_625 = BHR0_625 - MarketBHR0_625;
		AR0_750 = BHR0_750 - MarketBHR0_750;
		AR20_21 = BHR20_21 - MarketBHR20_21;
		AR20_750 = BHR20_750 - MarketBHR20_750;
run;

data FirmBHRFinal;
	set FirmBHR_AR (keep = Name BHR:);
run;

data FirmARFinal;
	set FirmBHR_AR (keep = Name AR:);
run;

data MarketBHRFinal;
	set FirmBHR_AR (keep = Name Market:);
run;

proc sql;
	create table FirmBHR_ARFinal
	as select a.*, b.*, c.*
	from FirmBHRFinal as a
	left join FirmARFinal as b
		on a.Name = b.Name
	left join MarketBHRFinal as c
		on a.Name = c.Name
	order by Name;
quit;

*** Output the dataset as XLS file  此DATA包含各公司與市場的各期BHR***;
PROC EXPORT DATA= FirmBHR_ARFinal
            OUTFILE= "&path\FirmBHR_AR.xlsx" 
            DBMS=EXCEL REPLACE;
     SHEET="sheet 1"; 
     NEWFILE=YES;
RUN;








