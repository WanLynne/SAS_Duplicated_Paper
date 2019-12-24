%let path = %str(C:\Users\lynne\Desktop\F_Final);

*** �פJ��� ***;
data Allot;
	infile "&path\Allot.csv"  dlm=',' firstobs=2 dsd;
	input WholeName :$20. Symbol Type :$20. Ask Price TotalPrice Qual Allot;
run;
*** �פJ���q��Symbol�A���F�ѨM2000�~�S��Symbol����� ***;
data CompanyWholeName;
	infile "&path\CompanyWholeName.csv"  dlm=',' firstobs=2 dsd;
	input Symbol WholeName :$20.;
run;
*** �N2000�~�S��Symbol����ƨ��X ***;
data Allot_NA;
	set Allot;
	where Symbol = .;
run;
*** Merge Symbol ***;
proc sql;
	create table Allot_Check 
	as select a.WholeName, a.Type, a.Ask, a.Price, a.TotalPrice, a.Qual, a.Allot, b.Symbol
	from Allot_NA a left join CompanyWholeName b
	on a.WholeName = b.WholeName
	order by a.WholeName;
quit;
*** ���ܸ�ƶ��� ***;
data Allot_Year2000;
	retain WholeName Symbol Type Ask Price TotalPrice Qual Allot;
	set Allot_Check;
run;
*** �t�~��X2001�~��2003�~����� ***;
data Allot_Year2001To2003;
	set Allot;
	where Symbol ^= .;
run;
*** �X�ֳ̲�Data ***;
data Allot_Final;
	set Allot_Year2000 Allot_Year2001To2003;
run;
*** �פJ�һݤ��q�˥� ***;
data CompanyList;
	infile "&path\Company_List.csv"  dlm=',' firstobs=2 dsd;
	input Symbol Name :$9. WholeName :$20.;
run;

*** �N�˥����q�P�Ҧ������Ҳv��ƪ����qMerge ***;
proc sql;
	create table AllotList 
	as select a.Symbol, a.Name, b.Type, b.Ask, b.Price, b.TotalPrice, b.Qual, b.Allot
	from CompanyList a left join Allot_Final b
	on a.Symbol = b.Symbol
	order by a.Symbol;
quit;
*** �ץX��ơA��ʧ� ***;
PROC EXPORT DATA= WORK.AllotList
            OUTFILE= "&path\Allot_List.xlsx" 
            DBMS=XLSX REPLACE;
     SHEET="sheet 1"; 
     NEWFILE=YES;
RUN;
