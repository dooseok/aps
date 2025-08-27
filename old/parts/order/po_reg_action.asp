<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1
dim CNT1

dim PO_Date
dim PO_Due_Date
dim PO_State
dim Partner_P_Name

dim temp
dim strError
dim URL_Prev
dim URL_Next

PO_Date			= trim(Request("PO_Date"))
PO_Due_Date		= trim(Request("PO_Due_Date"))
PO_State		= trim(Request("PO_State"))
Partner_P_Name	= trim(Request("Partner_P_Name"))

dim PO_Code

URL_Prev = Request("URL_Prev")
URL_Next = Request("URL_Next")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	SQL = "insert into tbParts_Order (PO_Date,PO_Due_Date,PO_State,Partner_P_Name) values "
	SQL = SQL & "('"&PO_Date&"','"&PO_Due_Date&"','"&PO_State&"','"&Partner_P_Name&"')"
	sys_DBCon.execute(SQL)
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select max(PO_Code) from tbParts_Order where PO_Date='"&PO_Date&"' and Partner_P_Name='"&Partner_P_Name&"'"
	RS1.Open SQL,sys_DBCon
	PO_Code = RS1(0)
	RS1.Close
	set RS1 = Nothing

	'for CNT1 = 1 to 10
		'SQL = "insert into tbParts_Order_Detail (Parts_Order_PO_Code,Parts_P_P_No,POD_Price,POD_Due_Date,POD_Qty,POD_In_Qty,POD_Remark) values "
		'SQL = SQL & "("&PO_Code&",'',0,'"&PO_Due_Date&"',0,0,'')"
		'sys_DBCon.execute(SQL)
	'next
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
<form name="frmRedirect" action="PO_list.asp" method=post>
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
<form name="frmRedirect" action="PO_list.asp" method=post>
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