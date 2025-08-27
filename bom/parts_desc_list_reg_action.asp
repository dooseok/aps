<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem 변수선언
dim RS1
dim SQL
dim CNT1

dim strError
dim strError_Temp

dim BMDD_Desc_BOM
dim BOM_Mask_Desc_BMD_Desc

BMDD_Desc_BOM			= trim(Request("BMDD_Desc_BOM"))
BOM_Mask_Desc_BMD_Desc= trim(Request("BOM_Mask_Desc_BMD_Desc"))

BMDD_Desc_BOM = replace(BMDD_Desc_BOM,"'","''")
BOM_Mask_Desc_BMD_Desc = replace(BOM_Mask_Desc_BMD_Desc,"'","''")
set RS1 = Server.CreateObject("ADODB.RecordSet")
rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트

	strError_Temp = ""

	
	SQL = "select * from tblBOM_Mask_Desc where BMD_Desc='"&BOM_Mask_Desc_BMD_Desc&"' "
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		SQL = "insert into tblBOM_Mask_Desc values ('"&BOM_Mask_Desc_BMD_Desc&"',999)"
		sys_DBCon.execute(SQL)
	end if
	RS1.Close

	SQL = "select * from tblBOM_Mask_Desc_Detail where BMDD_Desc_BOM='"&BMDD_Desc_BOM&"' "
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		SQL = "insert into tblBOM_Mask_Desc_Detail (BMDD_Desc_BOM, BOM_Mask_Desc_BMD_Desc) "
		SQL = SQL & " values	('"&BMDD_Desc_BOM&"','"&BOM_Mask_Desc_BMD_Desc&"')"
		sys_DBCon.execute(SQL)
	end if
	RS1.Close

	strError = strError & strError_Temp

end if
set RS1 = nothing
%>

<%
dim Request_Fields
dim strRequestForm
dim strRequestQueryString
for each Request_Fields in Request.Form
	if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	end if
next
for each Request_Fields in Request.QueryString
	if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	end if
next
if strError = "" then
%>
<form name="frmRedirect" action="parts_desc_list.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
	'strError = strError & "* 일부의 수정이 취소되었습니다."
%>
<form name="frmRedirect" action="parts_desc_list.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>



<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->