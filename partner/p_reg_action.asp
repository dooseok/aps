<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim P_Name
dim P_Name_Magic
dim P_Business_No
dim P_Owner
dim P_Pay_Method
dim P_Memo
dim P_Email

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")

P_Name			= trim(Request("P_Name"))
P_Name_Magic	= trim(Request("P_Name_Magic"))
P_Business_No	= trim(Request("P_Business_No"))
P_Owner			= trim(Request("P_Owner"))
P_Pay_Method	= trim(Request("P_Pay_Method"))
P_Memo			= trim(Request("P_Memo"))
P_Email			= trim(Request("P_Email"))

set RS1 = Server.CreateObject("ADODB.RecordSet")
rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	rem DB 업데이트
	RS1.Open "tbPartner",sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("P_Name")			= P_Name
		.Fields("P_Name_Magic")		= P_Name_Magic
		.Fields("P_Business_No")	= P_Business_No
		.Fields("P_Owner")			= P_Owner
		.Fields("P_Pay_Method")		= P_Pay_Method
		.Fields("P_Memo")			= P_Memo
		.Fields("P_Email")			= P_Email
		.Update
		.Close
	end with
end if

rem 객체 해제
Set RS1	= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="p_list.asp" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="p_list.asp" method=post>

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