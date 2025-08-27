<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
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
<form name="frmRedirect" action="parts_desc_sort_list.asp" method="post">
</form>
<script language="javascript">
alert("업데이트가 완료되었습니다.");
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="parts_desc_sort_list.asp" method=post>
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
	
	dim CNT1
	
	dim BMD_Desc
	dim BMD_Sort

	SQL = "delete from tblBOM_Mask_Desc"
	sys_DBCon.execute(SQL)

	for CNT1 = 1 to ubound(arrXLS, 2)
		
		BMD_Desc		= replace(trim(arrXLS(0, CNT1)),"'","''")
		BMD_Sort		= trim(arrXLS(1, CNT1))

		SQL = "insert into tblBOM_Mask_Desc (BMD_Desc,BMD_Sort) values "
		SQL = SQL & "('"&BMD_Desc&"','"&BMD_Sort&"')"
		sys_DBCon.execute(SQL)

	next
end sub
%>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
