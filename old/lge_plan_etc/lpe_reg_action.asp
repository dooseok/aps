<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim BOM_Sub_BS_D_No
dim LPE_Type
dim LPE_Due_Date
dim LPE_Req_Qty
dim LPE_Buyer

dim temp
dim strError
dim URL_Prev
dim URL_Next

BOM_Sub_BS_D_No		= ucase(trim(Request("BOM_Sub_BS_D_No")))
LPE_Type			= trim(Request("LPE_Type"))
LPE_Due_Date		= trim(Request("LPE_Due_Date"))
LPE_Req_Qty			= trim(Request("LPE_Req_Qty"))
LPE_Buyer			= trim(Request("LPE_Buyer"))

URL_Prev			= Request("URL_Prev")
URL_Next			= Request("URL_Next")


rem 에러메세지가 있을 경우 실행안됨
if strError = "" then
	
	SQL = "insert into tbLGE_Plan_Etc (BOM_Sub_BS_D_No, LPE_Type, LPE_Due_Date, LPE_Req_Qty, LPE_Buyer,LPE_Complete_Qty) values "
	SQL = SQL & "('"&BOM_Sub_BS_D_No&"','"&LPE_Type&"','"&LPE_Due_Date&"',"&LPE_Req_Qty&",'"&LPE_Buyer&"',0)"
	
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
<form name="frmRedirect" action="lpe_list.asp" method=post>
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
<form name="frmRedirect" action="lpe_list.asp" method=post>
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