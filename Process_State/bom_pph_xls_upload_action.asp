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
<form name="frmRedirect" action="bom_pph_list.asp" method="post">
</form>
<script language="javascript">
alert("업데이트가 완료되었습니다.");
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="bom_pph_list.asp" method=post>
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
	
	dim BOM_Sub_BS_D_No
	dim BP_PPH
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	for CNT1 = 1 to ubound(arrXLS, 2)
		BOM_Sub_BS_D_No	= trim(arrXLS(1, CNT1))
		BP_PPH			= trim(arrXLS(2, CNT1))
		if isnumeric(BP_PPH) then
		else
			BP_PPH = 0
		end if
		SQL = "select top 1 BP_Code from tbBOM_PPH where BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			SQL = "insert into tbBOM_PPH (BOM_Sub_BS_D_No,BP_PPH) values ('"&BOM_Sub_BS_D_No&"',"&BP_PPH&")"
			sys_DBCon.execute(SQL)
		else
			SQL = "update tbBOM_PPH set BP_PPH = "&BP_PPH&" where BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"'"
			sys_DBCon.execute(SQL)
		end if
		RS1.Close
	next
	
	set RS1 = nothing
end sub
%>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
