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
*** �NTSE��OTC���������⪺��� ***;
data TSE;
	set IPOApplyData;
	where Market = "TSE";
run;
data OTC;
	set IPOApplyData;
	where Market = "OTC";
run;

***** TSE *****;
*** �N�Ӥ��q�b�W���骺�������SMerge ***;
proc sql;
	create table TSE_Market 
	as select a.Name, a.UnderwritingDateEnd, b.TSE_Index, b.date
	from TSE a left join MarketIndexReturn b
	on intnx('weekday',b.date,1)<=a.UnderwritingDateEnd<=intnx('weekday',b.date,20)      /* �D��20��(���]�t�P��G��)�A�X�j����d�� */
	order by a.Name, b.date;
quit;
*** ����W�����ѫe10�Ѫ���� ***;
*** ���ƧǸ�ơA�������ƭn�˱� ***;
proc sort data=TSE_Market;
	by Name descending Date;
run;
*** �p��C�@�a���q���X����� ***;
data TSE_Market_Count;
set TSE_Market;
by Name descending Date;
retain n 0;
n=n+1;
if first.Name then n=1;
run;
*** �Q��bn��1-10(�W����e10��)����ơA�ñN���S�q%�վ㬰�Ʀr ***;
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

*** �p�⥫�����զX�����������S ***;
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
*** �N�Ӥ��q�b�W���骺�������SMerge ***;
proc sql;
	create table OTC_Market 
	as select a.Name, a.UnderwritingDateEnd, b.OTC_Index, b.date
	from OTC a left join MarketIndexReturn b
	on intnx('weekday',b.date,1)<=a.UnderwritingDateEnd<=intnx('weekday',b.date,20)      /* �D��20��(���]�t�P��G��)�A�X�j����d�� */
	order by a.Name, b.date;
quit;
*** ����W�����ѫe10�Ѫ���� ***;
*** ���ƧǸ�ơA�������ƭn�˱� ***;
proc sort data=OTC_Market;
	by Name descending Date;
run;
*** �p��C�@�a���q���X����� ***;
data OTC_Market_Count;
set OTC_Market;
by Name descending Date;
retain n 0;
n=n+1;
if first.Name then n=1;
run;
*** �Q��bn��1-10(�W����e10��)����ơA�ñN���S�q%�վ㬰�Ʀr ***;
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

*** �p�⥫�����զX�����������S ***;
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

*** �NTSE�MOTC��ƫ����X�� ***;
data MarketReturnPrior10Days;
	set TSE_Market_Final_Return OTC_Market_Final_Return;
run;

*** �ץX��� ***;
PROC EXPORT DATA= WORK.MarketReturnPrior10Days
            OUTFILE= "&path\MarketReturnPrior10Days.xlsx" 
            DBMS=XLSX REPLACE;
     SHEET="sheet 1"; 
     NEWFILE=YES;
RUN;
