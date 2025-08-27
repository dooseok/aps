<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
Dim RS1
Dim SQL

Dim strEdit_Header
dim arrEdit_Form(5,1)
dim B_Code
dim Title
dim URL_Action
dim URL_Prev
dim URL_Next
dim URL_List
dim Form_Type
dim Column_Width
dim Value_Width

dim arrInputSelectG
dim arrInputSelect

dim M_ID
dim M_Code
dim M_Channel
dim M_Password
dim M_Email_1
dim M_Email_2
dim M_HP

M_Code = Request.Cookies("Admin")("M_Code")

Set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbMember where M_Code='"&M_Code&"'"
RS1.Open SQL,sys_DBCon
M_ID		= RS1("M_ID")
M_Channel	= RS1("M_Channel")
M_Password	= RS1("M_Password")
M_Email_1	= RS1("M_Email_1")
M_Email_2	= RS1("M_Email_2")
M_HP		= RS1("M_HP")
RS1.Close
Set RS1 = Nothing

strEdit_Header = "<input type='hidden' name='M_Code' value='"&M_Code&"'>" &vbcrlf

arrEdit_Form(0,0) = "아이디"
arrEdit_Form(0,1) = M_ID

arrEdit_Form(1,0) = "* 비밀번호"
arrEdit_Form(1,1) = "<input type='text' name='M_Password' value="&M_Password&" style='width:250px'>"

arrEdit_Form(2,0) = "* 소속회사"
arrEdit_Form(2,1) = M_Channel

arrEdit_Form(3,0) = "이메일 1"
arrEdit_Form(3,1) = "<input type='text' name='M_Email_1' value='"&M_Email_1&"' style='width:250px'>"

arrEdit_Form(4,0) = "이메일 2"
arrEdit_Form(4,1) = "<input type='text' name='M_Email_2' value='"&M_Email_2&"' style='width:250px'>"

arrEdit_Form(5,0) = "핸드폰번호"
arrEdit_Form(5,1) = "<input type='text' name='M_HP' value='"&M_HP&"' style='width:250px'>"


Title			= "개인정보수정"
URL_Action		= "m_edit_action.asp"
URL_Prev		= "/member/m_logout_action.asp"
URL_Next		= "/member/m_logout_action.asp"
URL_List		= "/member/m_logout_action.asp"
Form_Type		= ""
Column_Width	= 180
Value_Width		= 400
%>
<script language="javascript">
function Form_Check(form)
{
	var strError = '';
	if(!form.M_Password.value)
	{
		strError += "*비밀번호를 입력해주세요.\n"
	}
	if(strError == '')
	{
		form.submit();
	}
	else
	{
		alert(strError);
	}
}
</script>
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

call Common_Edit_Form(Title, URL_Action, URL_Next, URL_List, Form_Type, Column_Width, Value_Width, strEdit_Header, arrEdit_Form, strRequestForm)
%>
<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
