<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem 변수선언
dim SQL
dim RS1

dim M_Code
dim M_P_No
dim M_Desc
dim M_Spec
dim M_Additional_Info
dim M_Qty
dim M_Process

dim temp
dim strError
dim URL_Prev
dim URL_Next

dim strDelete

rem 객체선언
Set RS1		= Server.CreateObject("ADODB.RecordSet")

URL_Prev	= Request("URL_Prev")
URL_Next	= Request("URL_Next")

M_Code				= Request("M_Code")
M_P_No				= Request("M_P_No")
M_Desc				= Request("M_Desc")
M_Spec				= Request("M_Spec")
M_Additional_Info	= Request("M_Additional_Info")
M_Qty				= Request("M_Qty")
M_Process			= Request("M_Process")

SQL = "select count(M_P_No) from tbMaterial where M_Code <> '"&M_Code&"' and M_P_No = '"&M_P_No&"'"
RS1.Open SQL,sys_DBCon
if RS1(0) > 0 then
	strError = "동일한 파트넘버가 있습니다."
end if
RS1.Close

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	rem DB 업데이트
	dim oldM_Qty
	SQL = "select * from tbMaterial where M_Code = '"&M_Code&"'"
	
	RS1.Open SQL,sys_DBCon
	oldM_Qty = RS1("M_Qty")
	RS1.Close	
	
	RS1.Open SQL,sys_DBconString,3,2,&H0001
	with RS1
		.Fields("M_P_No")			= M_P_No
		.Fields("M_Desc")			= M_Desc
		.Fields("M_Spec")			= M_Spec
		.Fields("M_Additional_Info")= M_Additional_Info
		.Fields("M_Qty")			= M_Qty
		.Fields("M_Process")		= M_Process
		.Update
		.Close
	end with
	
	SQL = "insert into tbMaterial_Stock_History (Material_M_P_No,MSH_Change_Stock,MSH_Applyed_Stock,MSH_Change_Type,MSH_Change_Date,MSH_Company) values ('"
	SQL = SQL &M_P_No&"',"
	SQL = SQL &M_Qty-oldM_Qty&","
	SQL = SQL &M_Qty&",'"
	SQL = SQL &"직접수정','"
	SQL = SQL &date()&"','"
	SQL = SQL &"엠에스이')"
	
	if M_Qty-oldM_Qty <> 0 then
		sys_DBCon.execute(SQL)
	end if
		
end if

rem 객체 해제
Set RS1	= nothing
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
<form name="frmRedirect" action="<%=URL_Next%>" method=post>
<input type="hidden" name="M_Code" value="<%=M_Code%>">

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
<form name="frmRedirect" action="<%=URL_Prev%>" method=post>
<input type="hidden" name="M_Code" value="<%=M_Code%>">

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