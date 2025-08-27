<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
'변수선언
dim CNT1
dim CNT2
dim UpLoad
dim objXLS
dim XLSConnection
dim temp
dim strProperties

dim BP_Division

dim RS1
dim SQL
dim Sheet_Name
dim arrXLS

dim strFile
dim arrFile
dim File_Name

dim strDesc
dim Exist_YN

dim strError
set UpLoad	= Server.CreateObject("Dext.FileUpLoad")
UpLoad.DefaultPath = DefaultPath_BOM_Update
strFile 	= UpLoad("strFile")
arrFile		= split(strFile,"\")
File_Name	= lcase(arrFile(ubound(arrFile)))

'파일명 체크
if instr("-DGZ-DMZ-","-"&left(ucase(File_Name),3)&"-") > 0 then
	BP_Division			= left(ucase(File_Name),3)
else
	strError			= "파일명을 사업부로 변경하여 주십시오. 예) DGZ.xls"
end if


temp = UpLoad("strFile").SaveAs(DefaultPath_BOM_Update & File_Name, False)
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
if strError = "" then
	call XLS_Upload()
%>
<form name="frmRedirect" action="lr_list.asp" method="post">
</form>
<script language="javascript">
alert("<%=left(ucase(File_Name),3)%> 입고내역 업데이트가 완료되었습니다.");
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="lr_list.asp" method=post>
</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
objXLS.close
set objXLS = nothing
set UpLoad = nothing
%>

<%
sub XLS_Upload()

	dim SQL
	
	dim BOM_Sub_BS_D_No
	dim LR_Date
	dim LR_Division
	dim LR_Qty
	dim LR_Currency
	dim LR_Price
	dim LR_PriceQty
	dim LR_MKT
	
	dim LR_Date_First
	dim LR_Date_Last
	
	LR_Date_First	= trim(arrXLS(1, 1))
	LR_Date_First	= left(LR_Date_First,4)&"-"&mid(LR_Date_First,5,2)&"-"&mid(LR_Date_First,7,2)
	
	LR_Date_Last	= trim(arrXLS(1, ubound(arrXLS, 2)))
	LR_Date_Last	= left(LR_Date_Last,4)&"-"&mid(LR_Date_Last,5,2)&"-"&mid(LR_Date_Last,7,2)
	
	SQL = "delete tbLG_Receiving where LR_Date between '"&LR_Date_First&"' and '"&LR_Date_Last&"'"
	sys_DBCon.execute(SQL) 
	
	for CNT1 = 1 to ubound(arrXLS, 2)

		BOM_Sub_BS_D_No = trim(arrXLS(2, CNT1))
		LR_Date			= trim(arrXLS(1, CNT1))
		LR_Date			= left(LR_Date,4)&"-"&mid(LR_Date,5,2)&"-"&mid(LR_Date,7,2)
		LR_Division		= trim(arrXLS(4, CNT1))
		LR_Division		= left(LR_Division,3)
		LR_Qty			= trim(arrXLS(6, CNT1))
		LR_Currency 	= trim(arrXLS(7, CNT1))
		LR_Price 		= trim(arrXLS(9, CNT1))
		LR_PriceQty 	= trim(arrXLS(11, CNT1))
		LR_MKT 			= trim(arrXLS(13, CNT1))
	
		SQL = "insert into tbLG_Receiving (BOM_Sub_BS_D_No,LR_Date,LR_Division,LR_Qty,LR_Currency,LR_Price,LR_PriceQty,LR_MKT) values "
		SQL = SQL & "('"&BOM_Sub_BS_D_No&"','"&LR_Date&"','"&LR_Division&"',"&LR_Qty&",'"&LR_Currency&"',"&LR_Price&","&LR_PriceQty&",'"&LR_MKT&"')"
		sys_DBCon.execute(SQL)
		
	next
	
	SQL = "insert into tbMacro_Log (ML_Item,ML_UploadDate) values ('Receiving_"&LR_Division&"',getdate())"
	sys_DBCon.execute(SQL)
end sub
%>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
