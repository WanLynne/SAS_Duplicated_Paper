%let path = %str(C:\Users\lynne\Desktop\F_Final);

*** 匯入資料 ***;
data Allot;
	infile "&path\Allot.csv"  dlm=',' firstobs=2 dsd;
	input WholeName :$20. Symbol Type :$20. Ask Price TotalPrice Qual Allot;
run;
*** 匯入公司及Symbol，為了解決2000年沒有Symbol的資料 ***;
data CompanyWholeName;
	infile "&path\CompanyWholeName.csv"  dlm=',' firstobs=2 dsd;
	input Symbol WholeName :$20.;
run;
*** 將2000年沒有Symbol的資料取出 ***;
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
*** 改變資料順序 ***;
data Allot_Year2000;
	retain WholeName Symbol Type Ask Price TotalPrice Qual Allot;
	set Allot_Check;
run;
*** 另外抓出2001年到2003年的資料 ***;
data Allot_Year2001To2003;
	set Allot;
	where Symbol ^= .;
run;
*** 合併最終Data ***;
data Allot_Final;
	set Allot_Year2000 Allot_Year2001To2003;
run;
*** 匯入所需公司樣本 ***;
data CompanyList;
	infile "&path\Company_List.csv"  dlm=',' firstobs=2 dsd;
	input Symbol Name :$9. WholeName :$20.;
run;

*** 將樣本公司與所有有中籤率資料的公司Merge ***;
proc sql;
	create table AllotList 
	as select a.Symbol, a.Name, b.Type, b.Ask, b.Price, b.TotalPrice, b.Qual, b.Allot
	from CompanyList a left join Allot_Final b
	on a.Symbol = b.Symbol
	order by a.Symbol;
quit;
*** 匯出資料，手動找 ***;
PROC EXPORT DATA= WORK.AllotList
            OUTFILE= "&path\Allot_List.xlsx" 
            DBMS=XLSX REPLACE;
     SHEET="sheet 1"; 
     NEWFILE=YES;
RUN;
