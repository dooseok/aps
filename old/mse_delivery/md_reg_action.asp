<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim LM_Code
dim LM_Company
dim LM_Name
dim BOM_Sub_BS_D_No_1
dim BOM_Sub_BS_D_No_2
dim BOM_Sub_BS_D_No_3
dim BOM_Sub_BS_D_No_4

dim temp
dim strError
dim URL_Prev
dim URL_Next

LM_Code				= trim(Request("LM_Code"))
LM_Company			= trim(Request("LM_Company"))
LM_Name				= trim(Request("LM_Name"))
BOM_Sub_BS_D_No_1	= trim(Request("BOM_Sub_BS_D_No_1"))
BOM_Sub_BS_D_No_2	= trim(Request("BOM_Sub_BS_D_No_2"))
BOM_Sub_BS_D_No_3	= trim(Request("BOM_Sub_BS_D_No_3"))
BOM_Sub_BS_D_No_4	= trim(Request("BOM_Sub_BS_D_No_4"))

URL_Prev = Request("URL_Prev")
URL_Next = Request("URL_Next")

rem 에러메세지가 있을 경우 실행안됨

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select top 1 LM_Name from tbLGE_Model where LM_Name='"&LM_Name&"'"
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
	strError = "* 동일한 모델정보가 이미 등록되어있습니다.\n"
end if
RS1.Close

if strError = "" then

	SQL = "insert into tbLGE_Model (LM_Company,LM_Name,BOM_Sub_BS_D_No_1,BOM_Sub_BS_D_No_2,BOM_Sub_BS_D_No_3,BOM_Sub_BS_D_No_4) values "
	SQL = SQL & "('"&LM_Company&"','"&LM_Name&"','"&BOM_Sub_BS_D_No_1&"','"&BOM_Sub_BS_D_No_2&"','"&BOM_Sub_BS_D_No_3&"','"&BOM_Sub_BS_D_No_4&"')"
	
	sys_DBCon.execute(SQL)
end if
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
<form name="frmRedirect" action="lm_list.asp" method=post>
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="lm_list.asp" method=post>
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