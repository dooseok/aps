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
<form name="frmRedirect" action="cosp_price_list.asp" method="post">
</form>
<script language="javascript">
alert("<%=left(ucase(File_Name),3)%> 사급가 업데이트가 완료되었습니다.");
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="cosp_price_list.asp" method=post>
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
	dim RS1
	dim SQL
	
	dim Material_M_P_No
	dim CP_Type
	dim CP_Price
	dim CP_StartDate
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	
	for CNT1 = 1 to ubound(arrXLS, 2)
		
		Material_M_P_No	= trim(arrXLS(1, CNT1))
		CP_Type			= trim(arrXLS(2, CNT1))
		CP_Price		= trim(arrXLS(6, CNT1))
		CP_StartDate	= trim(arrXLS(7, CNT1))	
	
		if isnumeric(CP_Price) then 
			CP_Price = replace(CP_Price,",","")
			SQL = "select top 1 CP_StartDate from tbCOSP_Price where Material_M_P_No = '"&Material_M_P_No&"'"
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				SQL = "insert into tbCOSP_Price (Material_M_P_No,CP_Type,CP_Price,CP_StartDate) values "
				SQL = SQL & "('"&Material_M_P_No&"','"&CP_Type&"',"&CP_Price&",'"&CP_StartDate&"')"
				sys_DBCon.execute(SQL)
			else
				if RS1("CP_StartDate") <= CP_StartDate then
					SQL = "update tbCOSP_Price set CP_Type = '"&CP_Type&"', CP_Price = "&CP_Price&", CP_StartDate = '"&CP_StartDate&"' "
					SQL = SQL & "where Material_M_P_No = '"&Material_M_P_No&"'"
					sys_DBCon.execute(SQL)
				end if
			end if
			RS1.Close
		end if
	next
	
	SQL = "insert into tbMacro_Log (ML_Item,ML_UploadDate) values ('COSP_Price_"&left(ucase(File_Name),3)&"',getdate())"
	sys_DBCon.execute(SQL)
	
	set RS1 = nothing
end sub
%>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
