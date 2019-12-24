%let path = %str(C:\Users\lynne\Desktop\F_Final);

* ���~ format;  
proc format;
	value $IndustryFormat 
							"M11" = "���d"
							"M12" = "���~"
							"M13" = "�ۤ�"
							"M14" = "��´�H��"
							"M15" = "���q"
							"M16" = "�q�u�q�l"
							"M17" = "�ƾ�"
							"M18" = "��������"
							"M19" = "�y��"
							"M20" = "���K����"
							"M21" = "�󽦽��L"
							"M22" = "�T��"
							"M23" = "��T�q�l"
							"M25" = "���"
							"M26" = "�B��"
							"M27" = "�[��"
							"M28" = "���īO�I"
							"M29" = "�ʳf"
							"M30" = "�Ҩ�"
							"M99" = "��L";

* �s�@�Ϥ��������C�n���ת�format ;
proc format;
	 value AggreFormat 
					LOW - 0.01620 = "���n����" 
	 				0.01620 -< 0.43875 = "���n����"
					0.43875 - HIGH = "�C�n����";

proc format;
 value BHRFormat 
       1 = "BHR0_1"
       3 = "BHR0_2"
       5 = "BHR0_3"
       7 = "BHR0_4"
       9 = "BHR0_5"
       11 = "BHR0_10"
       13 = "BHR0_20"
       15 = "BHR0_125"
       17 = "BHR0_250"
       19 = "BHR0_375"
       21 = "BHR0_500"
       23 = "BHR0_625"
       25 = "BHR0_750"
       27 = "BHR20_21"
       29 = "BHR20_750";

proc format;
 value ARFormat 
       1 = "AR0_1"
       3 = "AR0_2"
       5 = "AR0_3"
       7 = "AR0_4"
       9 = "AR0_5"
       11 = "AR0_10"
       13 = "AR0_20"
       15 = "AR0_125"
       17 = "AR0_250"
       19 = "AR0_375"
       21 = "AR0_500"
       23 = "AR0_625"
       25 = "AR0_750"
       27 = "AR20_21"
       29 = "AR20_750";

data IPOApplyData;
	infile "&path\IPOInformation_1221.csv"  dlm=',' firstobs=2 dsd;
	input Name :$9. IPODate :yymmdd10. TejIndustry :$15. SettingDate :yymmdd10. OfferPrice NumberOfShares :comma9.  
			AnnounceDate :yymmdd10. UnderwritingDateStart  :yymmdd10. UnderwritingDateEnd :yymmdd10. WiningRate 
			Market $ IPODatePrice  IPODateAmount :comma9. Date1AfterIPO :yymmdd10. Day1Price Day2Price Day3Price Day4Price Day5Price;
		OperatingYearPriorIPO = (IPODate - SettingDate) / 365;
		AskPeriod = UnderwritingDateEnd - UnderwritingDateStart;
		IndustrySymbol = substr(TEJIndustry, 1, 3);
		IPODateAmount = IPODateAmount / 1000;
	format IPODate yymmdd10. SettingDate  yymmdd10. AnnounceDate yymmdd10. UnderwritingDateStart yymmdd10. 
				UnderwritingDateEnd yymmdd10. Date1AfterIPO yymmdd10. IndustrySymbol $IndustryFormat.;
	keep Name IPODate IndustrySymbol TejIndustry OperatingYearPriorIPO Market IPODatePrice IPODateAmount AskPeriod;
run;

data OtherIPOPriorReturn;
	infile "&path\OtherIPOPriorReturn.csv" dlm=',' firstobs=2 dsd;
	input Name :$9. OtherIPOFirstDayR_Mean OtherIPOInitialDayR_Mean;
run;

data MarketReturnPrior10Days;
	infile "&path\MarketReturnPrior10Days.csv"  dlm=',' firstobs=2 dsd;
	input Name :$9. MarketReturnPrior10Days ;
run;

proc import
	datafile = "&path\FirmBHR_AR.csv"
	out = FirmBHR_AR
	dbms = CSV
	replace;
	getnames=Yes;
run;

*Qual �N���`�X���� Allot ���Ҳv Aggre �n���� ; 
data Final_Aggre;
	infile "&path\Final_Aggre.csv" dlm=',' firstobs=2 dsd;
	input Name  :$9. Ask OfferPrice TotalOfferValue Qual Allot Allot_1 Aggre_1;
		OfferAmount = (TotalOfferValue/ OfferPrice)/1000000;
		Aggre = log(1 / Allot);
		Ask = Ask / 1000; 
		Qual = Qual / 1000;
	drop Allot_1 Aggre_1;
run;

* Merge �Ҧ�Data ; 
proc sql;
	create table FinalData
	as select a.*, b.*, c.*, d.*, e.*
	from Final_Aggre as a
	left join OtherIPOPriorReturn as b
		on a.Name = b.Name
	left join MarketReturnPrior10Days as c
		on a.Name = c.Name
	left join IPOApplyData as d
		on a.Name = d.Name
	left join FirmBHR_AR as e
		on a.Name = e.Name
	order by Name;
quit;

*** Output the dataset as CSV file ***;
PROC EXPORT DATA = FinalData
            OUTFILE = "&path\FinalData.csv" 
            DBMS = CSV 
			REPLACE; 
RUN;

*** Table 1 IPOs ���~���G ***;
proc freq data = FinalData;
	table IndustrySymbol;
	title "Table 1 IPOs ���~���G";
run;

*** Table 2 IPOs �S�x���򥻲έp�q ***;
Data DataForTable2;
	retain Name OfferPrice OfferAmount Ask AskPeriod Qual Allot Aggre MarketReturnPrior10Days OtherIPOFirstDayR_Mean
			OtherIPOInitialDayR_Mean OperatingYearPriorIPO IPODateAmount;
	set FinalData;
	keep Name OfferPrice OfferAmount Ask AskPeriod Qual Allot Aggre MarketReturnPrior10Days OtherIPOFirstDayR_Mean
			OtherIPOInitialDayR_Mean OperatingYearPriorIPO IPODateAmount;
run;

proc sort data = DataForTable2;
	by Name;
run;

proc means data = DataForTable2 n mean median std  min max stackODS;
	var OfferPrice OfferAmount Ask AskPeriod Qual Allot Aggre MarketReturnPrior10Days OtherIPOFirstDayR_Mean
		  OtherIPOInitialDayR_Mean OperatingYearPriorIPO IPODateAmount;  
	output out = Table2_Summary mean()= median()= std()=/autoname;
	title "Table 2 IPOs �S�x���򥻲έp�q"; 
quit;

*** Table 3 IPOs �`�˥����������S ***;
Data DataForTable3;
	set FinalData;
	keep Name BHR: ;
run;

proc sort data = DataForTable3;
	by Name;
run;

proc means data = DataForTable3 n mean median std  min max stackODS;
	var BHR: ;  
	output out = Table3_Summary mean()= median()= std()=/autoname;
	title "Table 3 IPOs �`�˥������������W�B���S"; 
quit;


*** Table 4 IPOs ���ռ˥��������������S ***;
* �Hproc univariate ��X       25�ʤ����: 0.01620         75�ʤ���� : 0.43875;
proc univariate data = FinalData ;
	var Allot ;
run;

* �Ndata �Ϥ��� �����C�n���� ;
Data DataForTable4;
	set FinalData;
	keep Name Allot BHR: ;
	format Allot AggreFormat.;
run;

proc freq data = DataForTable4;
	table Allot;
	title "���ʿn���� ���Ƥ��t��";
run;

proc means data = DataForTable4  mean median std stackODS;
	var BHR: ;
	class Allot; 
	output out = Table4_Summary mean()= median()= std()=/autoname;
	title "Table 4 IPOs ���ռ˥��������������S"; 
quit;

*** Table 5 IPOs �`�˥����������S ***;
Data DataForTable5;
	set FinalData;
	keep Name AR: ;
run;

proc sort data = DataForTable5;
	by Name;
run;

proc means data = DataForTable5 n mean median std  min max stackODS;
	var AR: ;  
	output out = Table5_Summary mean()= median()= std()=/autoname;
	title "Table 5 IPOs �`�˥������������W�B���S"; 
quit;

*** Table 6 IPOs ���ռ˥������������W�B���S ***;
* �Ndata �Ϥ��� �����C�n���� ;
Data DataForTable6;
	set FinalData;
	keep Name Allot AR: ;
	format Allot AggreFormat.;
run;

proc means data = DataForTable6  mean median std stackODS;
	var AR: ;
	class Allot; 
	output out = Table6_Summary mean()= median()= std()=/autoname;
	title "Table 6 IPOs ���ռ˥������������W�B���S"; 
quit;

*** Table 7 IPOs �����������S����H���ʿn���ʪ��j�k���G ***;
data SetBHRSign;
input var $ text $9.;
call symput(var,text);
cards;
VAR1 BHR0_1
VAR2 BHR0_2
VAR3 BHR0_3
VAR4 BHR0_4
VAR5 BHR0_5
VAR6 BHR0_10
VAR7 BHR0_20
VAR8 BHR0_125
VAR9 BHR0_250
VAR10 BHR0_375
VAR11 BHR0_500
VAR12 BHR0_625
VAR13 BHR0_750
VAR14 BHR20_21
VAR15 BHR20_750
;
run;
%put _user_;

%macro RegAggreToBHR(x,var);
title "&&var&x";
proc surveyreg data=FinalData;
model &&var&x = Aggre:/adjrsq;
ods output ParameterEstimates=BHR&x FitStatistics=BHR_F&x;
run;

%mend;
      %macro loop;
             %do x=1 %to 15;
                  %RegAggreToBHR(&x);
             %end;
        %mend;
        %loop

/*--�X�֤��R���G*/
Data BHR_AlphaBeta_Data;
	Set BHR1-BHR15;
	retain BHRPeriod 0;
		BHRPeriod = BHRPeriod +1;
run;

data BHR_AlphaBeta;
	retain BHRPeriod Parameter Estimate;
	set BHR_AlphaBeta_Data;
	if mod(BHRPeriod, 2) = 0 then do; 
		BHRPeriod = BHRPeriod -1;
	end;
	else do;
		BHRPeriod = BHRPeriod;
	end;
	keep BHRPeriod Parameter Estimate;
run;

proc transpose data = BHR_AlphaBeta 
						out = BHR_AlphaBeta_Transpose 
						prefix = Parameter;
						var Estimate;
						by BHRPeriod;
run;

data BHR_AlphaBeta_ToJoin;
	set BHR_AlphaBeta_Transpose;
	drop _NAME_;
	rename Parameter1 = Intercept Parameter2 = Aggre;
run;

Data BHR_RSquared_Data;
	Set BHR_F1-BHR_F15 (firstobs = 1 obs = 2);
	drop cValue1;
	retain BHRPeriod 0;
		BHRPeriod = BHRPeriod +1;
run;

data BHR_RSquared;
	retain BHRPeriod Label1 nValue1;
	set BHR_RSquared_Data;
	if mod(BHRPeriod, 2) = 0 then do; 
		BHRPeriod = BHRPeriod -1;
	end;
	else do;
		BHRPeriod = BHRPeriod;
	end;
	keep BHRPeriod Label1 nValue1;
run;

proc transpose data = BHR_RSquared 
						out = BHR_RSquared_Transpose 
						prefix = Label1;
						var nValue1;
						by BHRPeriod;
run;

data BHR_RSquared_ToJoin;
	set BHR_RSquared_Transpose;
	drop _NAME_;
	rename Label11 = RSquared Label12 = AdjRSquared;
run;

proc sql;
	create table Table7_Data
	as select a.*, b.*
	from BHR_AlphaBeta_ToJoin as a
	left join BHR_RSquared_ToJoin as b
		on a.BHRPeriod = b.BHRPeriod
	order by BHRPeriod;
quit;

data Table7;
	set Table7_Data;
	format  BHRPeriod BHRFormat.;
run;

proc print data=Table7;
	title "Table 7 IPOs �����������S����H���ʿn���ʪ��j�k���G"; 
Run;

*** Table 8 IPO ���������W�B���S����H���ʿn���ʪ��j�k���G ***;
data SetARSign;
	input var $ text $9.;
	call symput(var,text);
	cards;
VAR1 AR0_1
VAR2 AR0_2
VAR3 AR0_3
VAR4 AR0_4
VAR5 AR0_5
VAR6 AR0_10
VAR7 AR0_20
VAR8 AR0_125
VAR9 AR0_250
VAR10 AR0_375
VAR11 AR0_500
VAR12 AR0_625
VAR13 AR0_750
VAR14 AR20_21
VAR15 AR20_750
;
run;
%put _user_;

%macro RegAggreToAR(x,var);
title "&&var&x";
proc surveyreg data=FinalData;
model &&var&x = Aggre:/adjrsq;
ods output ParameterEstimates=AR&x FitStatistics=AR_F&x;
run;

%mend;
 		%macro loop;
             %do x=1 %to 15;
                  %RegAggreToAR(&x);
             %end;
        %mend;
        %loop

/*--�X�֤��R���G*/
Data AR_AlphaBeta_Data;
	Set AR1-AR15;
	retain ARPeriod 0;
		ARPeriod = ARPeriod +1;
run;

data AR_AlphaBeta;
	retain ARPeriod Parameter Estimate;
	set AR_AlphaBeta_Data;
	if mod(ARPeriod, 2) = 0 then do; 
		ARPeriod = ARPeriod -1;
	end;
	else do;
		ARPeriod = ARPeriod;
	end;
	keep ARPeriod Parameter Estimate;
run;

proc transpose data = AR_AlphaBeta 
						out = AR_AlphaBeta_Transpose 
						prefix = Parameter;
						var Estimate;
						by ARPeriod;
run;

data AR_AlphaBeta_ToJoin;
	set AR_AlphaBeta_Transpose;
	drop _NAME_;
	rename Parameter1 = Intercept Parameter2 = Aggre;
run;

Data AR_RSquared_Data;
	Set AR_F1-AR_F15 (firstobs = 1 obs = 2);
	drop cValue1;
	retain ARPeriod 0;
		ARPeriod = ARPeriod +1;
run;

data AR_RSquared;
	retain ARPeriod Label1 nValue1;
	set AR_RSquared_Data;
	if mod(ARPeriod, 2) = 0 then do; 
		ARPeriod = ARPeriod -1;
	end;
	else do;
		ARPeriod = ARPeriod;
	end;
	keep ARPeriod Label1 nValue1;
run;

proc transpose data = AR_RSquared 
						out = AR_RSquared_Transpose 
						prefix = Label1;
						var nValue1;
						by ARPeriod;
run;

data AR_RSquared_ToJoin;
	set AR_RSquared_Transpose;
	drop _NAME_;
	rename Label11 = RSquared Label12 = AdjRSquared;
run;

proc sql;
	create table Table8_Data
	as select a.*, b.*
	from AR_AlphaBeta_ToJoin as a
	left join AR_RSquared_ToJoin as b
		on a.ARPeriod = b.ARPeriod
	order by ARPeriod;
quit;

data Table8;
	set Table8_Data;
	format  ARPeriod ARFormat.;
run;

proc print data=Table8;
	title "Table 8 IPOs ���������W�B���S����H���ʿn���ʪ��j�k���G"; 
Run;

*** Table 9 ���H���ʿn���ʹ�IPOs�S�x���j�k���G ***;
Data DataForTable9;
	set FinalData;
	if Market = "OTC" then catmk=1; else catmk=0;
	if substr(TejIndustry, 1, 3) = "M23" then catind=1; else catind=0;
run;
* �j�k�Ҧ��@;
proc surveyreg data=DataForTable9;
	model Aggre=OfferPrice MarketReturnPrior10Days OperatingYearPriorIPO catmk catind:/adjrsq;
	ods output ParameterEstimates=ParameterEstimates_1 FitStatistics=FitStatistics_1;
	title "Table 9 ���H���ʿn���ʹ�IPOs�S�x���j�k���G"; 
quit;
* �j�k�Ҧ��G;
proc surveyreg data=DataForTable9;
	model Aggre=OfferPrice OtherIPOFirstDayR_Mean OperatingYearPriorIPO catmk catind:/adjrsq;
	ods output ParameterEstimates=ParameterEstimates_2 FitStatistics=FitStatistics_2;
	title "Table 9 ���H���ʿn���ʹ�IPOs�S�x���j�k���G"; 
quit;
* �j�k�Ҧ��T;
proc surveyreg data=DataForTable9;
	model Aggre=OfferPrice OtherIPOInitialDayR_Mean OperatingYearPriorIPO catmk catind:/adjrsq;
	ods output ParameterEstimates=ParameterEstimates_3 FitStatistics=FitStatistics_3;
	title "Table 9 ���H���ʿn���ʹ�IPOs�S�x���j�k���G"; 
quit;
