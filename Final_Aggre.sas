%let path = %str(C:\Users\lynne\Desktop\F_Final);

data AllotListFinalDelete;
infile "&path\Allot_List_Final_Delete.csv"  dlm=',' firstobs=2 dsd;
input Symbol Name :$9. Type :$9. Ask Price TotalPrice Qual Allot;
drop Type;
run;
*** �ۤv�p�⤤�Ҳv ***;
proc sql;
	create table AllotData
	as select *, (TotalPrice/Price)/(Qual*Ask) as Calculate_Allot
	from AllotListFinalDelete
	order by Symbol;
quit;
*** �p����ʿn���� ***;
proc sql;
	create table Final_Aggre
	as select *, log(1/Calculate_Allot) as Aggre
	from AllotData
	order by Symbol;
quit;

*** �ץX��� ***;
PROC EXPORT DATA= WORK.Final_Allot
            OUTFILE= "&path\Final_Aggre.xlsx" 
            DBMS=XLSX REPLACE;
     SHEET="sheet 1"; 
     NEWFILE=YES;
RUN;
