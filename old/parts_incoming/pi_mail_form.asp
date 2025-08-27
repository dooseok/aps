<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim SQL
dim RS1
dim RS2

dim Email
dim strPI_Code

strPI_Code = Request("strChecked_Value")

if right(strPI_Code,1) = "," then
	strPI_Code = left(strPI_Code,len(strPI_Code)-1)
end if

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")

SQL = 		"select "&vbcrlf
SQL = SQL & "	Partner_P_Name "&vbcrlf
SQL = SQL & "from "&vbcrlf
SQL = SQL & "	vwPI_List "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PI_Code in ("&strPI_Code&") "&vbcrlf
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	SQL = "select P_Email from tbPartner where P_Name = '"&RS1("Partner_P_Name")&"'"
	RS2.Open SQL,sys_DBCon
	if RS2.Eof or RS2.Bof then
		Email = ""
	else
		Email = RS2("P_Email")
	end if
	RS2.Close

	RS1.MoveNext
loop
RS1.Close
%>

<script language="javascript">
function frmCheck()
{
	if(!frmMail.Email.value)
	{
		alert("메일주소를 입력해주세요.");
	}
	else
	{
		frmMail.submit();
	}
}
</script>

<table width=100% cellpadding=0 cellspacing=0 border=0>
<form name="frmMail" action="pi_mail.asp" method="post">
<input type="hidden" name="Title" value="[엠에스이] 발주서입니다.">
<input type="hidden" name="URL" value="<%=DefaultURLAdmin%>/parts_incoming/pi_print.asp?strChecked_Value=<%=strPI_Code%>,">
<tr height=100px>
	<td width=20px></td>
	<td>
		<input type="text" name="Email" value="<%=Email%>" style="width:95%">
	</td>
	<td width=50>
		<%=Make_S_BTN("발송","javascript:frmCheck()","")%>
	</td>
	<td width=20px></td>
</tr>
</form>
</table>
<%
set RS1 = nothing
set RS2 = nothing
%>
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->