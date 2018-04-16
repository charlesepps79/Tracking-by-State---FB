/*When importing latest cleared check master, ensure that Amount (currency), 
Percent, and Number_of_Payments/pmt come in as a number; bring drop_date and datecleared in as date
Must modify CPP per Barbara's plan number. */


%let cpp=.6657;
%let dropdate='05mar2018'd;

*"\\server-fs02\Marketing\2018 Programs\FB_Cleared_2018.xlsx";
data ccfile;
length Fico_25pt $10;
set WORK.FB_Cleared_2018;
if datecleared = "" then Booked = 0;
else booked = 1;
if fico=0 then Fico_25pt= "0";
if 0<fico<26 then Fico_25pt="1-25";
if 25<=fico<=49	then Fico_25pt= "25-49";
if 50<=fico<=74	then Fico_25pt= "50-74";
if 75<=fico<=99	then Fico_25pt= "75-99";
if 100<=fico<=124 then Fico_25pt= "100-124";	
if 125<=fico<=149 then Fico_25pt= "125-149";
if 150<=fico<=174 then Fico_25pt= "150-174";
if 175<=fico<=199 then Fico_25pt= "175-199";
if 200<=fico<=224 then Fico_25pt= "200-224";
if 225<=fico<=249 then Fico_25pt= "225-249";
if 300<=fico<=324 then Fico_25pt= "300-324";
if 350<=fico<=374 then Fico_25pt= "350-374";
if 400<=fico<=424 then Fico_25pt= "400-424";
if 425<=fico<=449 then Fico_25pt= "425-449";
if 450<=fico<=474 then Fico_25pt= "450-474";
if 475<=fico<=499 then Fico_25pt= "475-499";
if 500<=fico<=524 then Fico_25pt= "500-524";
if 525<=fico<=549 then Fico_25pt= "525-549";
if 550<=fico<=574 then Fico_25pt= "550-574";
if 575<=fico<=599 then Fico_25pt= "575-599";
if 600<=fico<=624 then Fico_25pt= "600-624";
if 625<=fico<=649 then Fico_25pt= "625-649";
if 650<=fico<=674 then Fico_25pt= "650-674";
if 675<=fico<=699 then Fico_25pt= "675-699";
if 700<=fico<=724 then Fico_25pt= "700-724";
if 725<=fico<=749 then Fico_25pt= "725-749";
if 750<=fico<=774 then Fico_25pt= "750-774";
if 775<=fico<=799 then Fico_25pt= "775-799";
if 800<=fico<=824 then Fico_25pt= "800-824";
if 825<=fico<=849 then Fico_25pt= "825-849";
if 850<=fico<=874 then Fico_25pt= "850-874";
if 875<=fico<=899 then Fico_25pt= "875-899";
if 975<=fico<=999 then Fico_25pt= "975-999";
if fico="" then Fico_25pt= "";
Mailed=1;
CPP=&cpp;
Month_CL=month(datecleared);
Month_Drop=month(drop_date);
bookedtot=booked*100;
if booked=1 then Volume=amount;
Number_of_Payments=pmt;
Payment=pmt_amt;
run;


data ccfile2;
set ccfile;
if drop_date=&dropdate;
run;
 



data CCdaily;
set ccfile2;
if booked=1;
DaysCleared = intck('weekday',drop_date,datecleared);
run;
proc sort data=ccdaily;
by DaysCleared;
run;
proc format;
picture pctpic (round) low-high='09.00%';
run; 





ods excel options(rowbreaks_interval="OUTPUT" sheet_interval="NONE");
Title "Tracking by State FB 2017";
proc tabulate data=ccfile2 missing;
class amtid state;
var Amount mailed Percent Number_of_Payments booked payment cpp bookedtot volume;
table state*(amtid all), Mailed="Mail Qty"*f=10. booked*f=5.0
booked="Bkg Rate"*rowpctsum<Mailed>*f=pctpic.
Volume*f=dollar18.2 volume*mean*f=dollar18.2  cpp="Mktg Cost"*f=dollar18.2 cpp="CPA"*pctsum<bookedtot>*f=dollar18.2
amount*mean="Face Amt"*f=dollar18.2 Percent*mean= "wAPR"*f=pctpic.
Number_of_Payments*mean="Term"*f=5.0 Payment*mean*f=dollar18.2/nocellmerge;
label amount="AvgCk" state="State";
run; 
/*
proc tabulate data=ccfile2 missing;
class fico_25pt amtid  state;
var amount Percent Number_of_Payments booked payment cpp amount Mailed bookedtot volume;
table FICO_25pt all, Mailed="Mail Qty"*f=10. booked*f=5.0
booked="Bkg Rate"*rowpctsum<mailed>*f=pctpic.
Volume*f=dollar18.2 amount*mean*f=dollar18.2  cpp="Mktg Cost"*f=dollar18.2 cpp="CPA"*pctsum<bookedtot>*f=dollar18.2
amount*mean="Face Amt"*f=dollar18.2 Percent*mean= "wAPR"*f=pctpic.
Number_of_Payments*mean="Term"*f=5.0 Payment*mean*f=dollar18.2/nocellmerge;
label fico_25pt="FICO_Range" amount="AvgCk" state="State";
run; 
proc tabulate data=ccfile2 missing;
class FICO_25pt amtid  state;
var amount Percent Number_of_Payments booked payment cpp Mailed bookedtot volume;
table state all, Mailed="Mail Qty"*f=10. booked*f=5.0
booked="Bkg Rate"*rowpctsum<mailed>*f=pctpic.
Volume*f=dollar18.2 amount*mean*f=dollar18.2  cpp="Mktg Cost"*f=dollar18.2 cpp="CPA"*pctsum<bookedtot>*f=dollar18.2
amount*mean="Face Amt"*f=dollar18.2 Percent*mean= "wAPR"*f=pctpic.
Number_of_Payments*mean="Term"*f=5.0 Payment*mean*f=dollar18.2/nocellmerge;
label fico_25pt="FICO_Range" amount="AvgCk" state="State";
run;
*/ 
proc tabulate data=ccdaily;
class dayscleared;
var booked amount;
table DaysCleared all,booked amount*f=dollar18.2/nocellmerge;
run;
ods excel close;



