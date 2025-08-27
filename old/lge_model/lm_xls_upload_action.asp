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

if instr(File_Name,"map-action") > 0 then
	call XLS_Upload()
end if
%>
<form name="frmRedirect" action="lm_list.asp" method="post">
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
	
	dim BOM_Sub_BS_D_No
	dim PD_Location
	dim PD_In_Date
	dim PD_Qty
	dim PD_Price_KRW
	dim PD_Price_USD
	dim PD_Price_Sum
	dim PD_Work_Order
	dim PD_Delivery_Number
	dim PD_Start_Date
	dim PD_Order_Type	
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
	
	PD_In_Date			= trim(arrXLS(8, 1))
	PD_In_Date			= left(PD_In_Date,4) &"-"& mid(PD_In_Date,5,2) &"-"&right(PD_In_Date,2)
	
	SQL = "delete from tbProduct_Delivery where left(convert(varchar,PD_In_Date,121),7) = '"&left(PD_In_Date,7)&"'"
	sys_DBCon.execute(SQL)
				
	for CNT1 = 1 to ubound(arrXLS, 2)
			
		BOM_Sub_BS_D_No		= trim(arrXLS(0, CNT1))
		PD_Location			= trim(arrXLS(1, CNT1))
		PD_In_Date			= trim(arrXLS(8, CNT1))
		PD_In_Date			= left(PD_In_Date,4) &"-"& mid(PD_In_Date,5,2) &"-"&right(PD_In_Date,2)
		PD_Qty				= trim(arrXLS(9, CNT1))
		PD_Price_KRW		= trim(arrXLS(11, CNT1))
		PD_Price_USD		= trim(arrXLS(13, CNT1))
		PD_Price_Sum		= trim(arrXLS(12, CNT1))
		PD_Work_Order		= trim(arrXLS(15, CNT1))
		PD_Delivery_Number	= trim(arrXLS(16, CNT1))
		PD_Start_Date		= trim(arrXLS(17, CNT1))
		PD_Start_Date		= left(PD_Start_Date,4) &"-"& mid(PD_Start_Date,5,2) &"-"&right(PD_Start_Date,2)
		PD_Order_Type		= trim(arrXLS(20, CNT1))
		
		SQL = "insert into tbProduct_Delivery (BOM_Sub_BS_D_No,PD_Location,PD_In_Date,PD_Qty,PD_Price_KRW,PD_Price_USD,PD_Price_Sum,PD_Work_Order,PD_Delivery_Number,PD_Start_Date,PD_Order_Type) values "
		SQL = SQL & "('"&BOM_Sub_BS_D_No&"','"&PD_Location&"','"&PD_In_Date&"',"&PD_Qty&","&PD_Price_KRW&","&PD_Price_USD&","&PD_Price_Sum&",'"&PD_Work_Order&"','"&PD_Delivery_Number&"','"&PD_Start_Date&"','"&PD_Order_Type&"')"
		sys_DBCon.execute(SQL)

	next
	
	set RS1 = nothing
	set RS2 = nothing
end sub
%>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
