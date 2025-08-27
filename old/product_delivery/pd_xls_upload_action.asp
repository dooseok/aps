<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
dim CNT1
dim CNT2
dim UpLoad
dim objXLS
dim XLSConnection
dim temp
dim strProperties

dim RS1
dim SQL
dim Sheet_Name
dim arrXLS

dim strFile
dim arrFile
dim File_Name

dim Exist_YN

set UpLoad	= Server.CreateObject("Dext.FileUpLoad")
UpLoad.DefaultPath = DefaultPath_SCS_XLS_Reader

strFile 	= UpLoad("strFile")
arrFile		= split(strFile,"\")
File_Name	= lcase(arrFile(ubound(arrFile)))

temp = UpLoad("strFile").SaveAs(DefaultPath_SAGUP_XLS_Reader & File_Name, False)
temp = replace(temp,"\","/")

set objXLS = Server.CreateObject("ADODB.Connection") 
XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""

objXLS.Open XLSConnection    
set RS1 = objXLS.OpenSchema(20)
Sheet_Name	= RS1(2)
Sheet_Name	= "["&replace(Sheet_Name,"'","")&"]"
set RS1 = nothing

set RS1 = Server.CreateObject("ADODB.RecordSet") 
SQL  = " select * from "&Sheet_Name
RS1.Open SQL,objXLS 
arrXLS = RS1.getRows()
RS1.close
set RS1 = nothing

if instr(File_Name,"rcv10") > 0 then
	call XLS_Upload()
end if
%>
<form name="frmRedirect" action="pd_list.asp" method="post">
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
objXLS.close
set objXLS = nothing
set UpLoad = nothing
%>

<%
sub XLS_Upload()
	dim RS1
	dim RS2
	dim SQL
	
	dim min_Receiving_Date
	dim max_Receiving_Date
	
	dim	BOM_Sub_BS_D_No
	dim	PD_Receiving_Qty
	dim	PD_Currency
	dim	PD_Unit_Price
	dim	PD_Sum_Price
	dim	PD_Receiving_Date
	dim	PD_Departure_Date
	dim	PD_Market
	dim	PD_Work_Order
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
			
	min_Receiving_Date	= 29991231
	max_Receiving_Date	= 19000101
	
	for CNT1 = 2 to ubound(arrXLS, 2)
		if trim(arrXLS(1, CNT1)) <> "" then
			if min_Receiving_Date > trim(arrXLS(1, CNT1)) then
				min_Receiving_Date = trim(arrXLS(1, CNT1))
			end if
			
			if max_Receiving_Date < trim(arrXLS(1, CNT1)) then
				max_Receiving_Date = trim(arrXLS(1, CNT1))
			end if
		end if
	next
	
	SQL = "delete from tbProduct_Delivery where PD_Receiving_Date between '"&left(min_Receiving_Date,4) &"-"& mid(min_Receiving_Date,5,2) &"-"&right(min_Receiving_Date,2)&"' and '"&left(max_Receiving_Date,4) &"-"& mid(max_Receiving_Date,5,2) &"-"&right(max_Receiving_Date,2)&"'"
	sys_DBCon.execute(SQL)
	
	for CNT1 = 2 to ubound(arrXLS, 2)	
		BOM_Sub_BS_D_No		= trim(arrXLS(2, CNT1))
		PD_Receiving_Qty	= trim(arrXLS(6, CNT1))
		PD_Currency				= trim(arrXLS(7, CNT1))
		PD_Unit_Price			= trim(arrXLS(9, CNT1))
		PD_Sum_Price			= trim(arrXLS(11, CNT1))
		PD_Receiving_Date	= trim(arrXLS(1, CNT1))
		PD_Receiving_Date	= left(PD_Receiving_Date,4) &"-"& mid(PD_Receiving_Date,5,2) &"-"&right(PD_Receiving_Date,2)
		PD_Departure_Date	= trim(arrXLS(12, CNT1))
		if trim(PD_Departure_Date) <> "" then
			PD_Departure_Date	= left(PD_Departure_Date,4) &"-"& mid(PD_Departure_Date,5,2) &"-"&right(PD_Departure_Date,2)
		else
			PD_Departure_Date	= ""
		end if
		PD_Market			= trim(arrXLS(13, CNT1))
		PD_Work_Order		= trim(arrXLS(14, CNT1))

		SQL = "insert into tbProduct_Delivery (BOM_Sub_BS_D_No, PD_Receiving_Qty, PD_Currency, PD_Unit_Price,	PD_Sum_Price, PD_Receiving_Date, PD_Departure_Date, PD_Market, PD_Work_Order) values "
		SQL = SQL & "('"&BOM_Sub_BS_D_No&"',"&PD_Receiving_Qty&",'"&PD_Currency&"',"&PD_Unit_Price&","&PD_Sum_Price&",'"&PD_Receiving_Date&"','"&PD_Departure_Date&"','"&PD_Market&"','"&PD_Work_Order&"')"
		sys_DBCon.execute(SQL)

	next
	
	set RS1 = nothing
	set RS2 = nothing
end sub
%>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
