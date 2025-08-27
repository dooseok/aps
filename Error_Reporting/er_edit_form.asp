<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
Dim RS1
Dim SQL

Dim strEdit_Header
dim arrEdit_Form(6,1)
dim B_Code
dim Title
dim URL_Action
dim URL_Prev
dim URL_Next
dim URL_List
dim Form_Type
dim Column_Width
dim Value_Width

dim ER_Code
dim ER_Title
dim ER_Content
dim ER_Reg_Date
dim ER_Edit_Date
dim ER_File_1
dim ER_File_2
dim ER_File_3

ER_Code = Request("ER_Code")

Set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tberror_reporting where ER_Code='"&ER_Code&"'"
RS1.Open SQL,sys_DBCon
ER_Title			= RS1("ER_Title")
ER_Content		= RS1("ER_Content")
ER_Reg_Date		= RS1("ER_Reg_Date")
ER_Edit_Date		= RS1("ER_Edit_Date")
ER_File_1		= RS1("ER_File_1")
ER_File_2		= RS1("ER_File_2")
ER_File_3		= RS1("ER_File_3")
RS1.Close
Set RS1 = Nothing

strEdit_Header = "<input type='hidden' name='ER_Code' value='"&ER_Code&"'>" &vbcrlf

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

arrEdit_Form(0,0) = "*제목"
arrEdit_Form(0,1) = "<input type='text' name='ER_Title' value='"&ER_Title&"' style='width:300px'>"

arrEdit_Form(1,0) = "*내용"
arrEdit_Form(1,1) = "<textarea name='ER_Content' style='width:90%' rows=20 style='border:1px solid #999999'>"&ER_Content&"</textarea>"

arrEdit_Form(2,0) = "첨부파일1"
arrEdit_Form(2,1) = "<input type='hidden' name='oldER_File_1' value='"&ER_File_1&"'>"
If ER_File_1 <> "" then
	arrEdit_Form(2,1) = arrEdit_Form(2,1) & "<a href='/function/ifrm_download.asp?filepath="&DefaultPath_Error_Reporting&ER_File_1&"' target='ifrm_download'>"&ER_File_1&"</a>"
	arrEdit_Form(2,1) = arrEdit_Form(2,1) & "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='strDelete' value='ER_File_1'>삭제"
End if
arrEdit_Form(2,1) = arrEdit_Form(2,1) & "<br><input type='file' name='ER_File_1' style='width:90%'>"

arrEdit_Form(3,0) = "첨부파일2"
arrEdit_Form(3,1) = "<input type='hidden' name='oldER_File_2' value='"&ER_File_2&"'>"
If ER_File_2 <> "" then
	arrEdit_Form(3,1) = arrEdit_Form(3,1) & "<a href='/function/ifrm_download.asp?filepath="&DefaultPath_Error_Reporting&ER_File_2&"' target='ifrm_download'>"&ER_File_2&"</a>"
	arrEdit_Form(3,1) = arrEdit_Form(3,1) & "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='strDelete' value='ER_File_2'>삭제"
End if
arrEdit_Form(3,1) = arrEdit_Form(3,1) & "<br><input type='file' name='ER_File_2' style='width:90%'>"

arrEdit_Form(4,0) = "첨부파일3"
arrEdit_Form(4,1) = "<input type='hidden' name='oldER_File_3' value='"&ER_File_3&"'>"
If ER_File_3 <> "" then
	arrEdit_Form(4,1) = arrEdit_Form(4,1) & "<a href='/function/ifrm_download.asp?filepath="&DefaultPath_Error_Reporting&ER_File_3&"' target='ifrm_download'>"&ER_File_3&"</a>"
	arrEdit_Form(4,1) = arrEdit_Form(4,1) & "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='strDelete' value='ER_File_3'>삭제"
End if
arrEdit_Form(4,1) = arrEdit_Form(4,1) & "<br><input type='file' name='ER_File_3' style='width:90%'>"

arrEdit_Form(5,0) = "최초등록일"
arrEdit_Form(5,1) = ER_Reg_Date

arrEdit_Form(6,0) = "최종수정일"
arrEdit_Form(6,1) = ER_Edit_Date

Title			= "요청사항수정"
URL_Action		= "ER_edit_action.asp"
URL_Prev		= "ER_edit_form.asp"
URL_Next		= "ER_edit_form.asp"
URL_List		= "ER_list.asp"
Form_Type		= "enctype='MULTIPART/FORM-DATA'"
Column_Width	= 180
Value_Width		= 700
%>
<script language="javascript">
function Form_Check(form)
{
	var strError = '';
	if(!form.ER_Title.value)
	{
		strError += "*제목을 입력해주세요.\n"
	}
	if(!form.ER_Content.value)
	{
		strError += "*내용을 입력해주세요.\n"
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
call Common_Edit_Form(Title, URL_Action, URL_Next, URL_List, Form_Type, Column_Width, Value_Width, strEdit_Header, arrEdit_Form, strRequestForm)
%>

<script language="javascript">
function frmDelete_Check()
{
	if(confirm("정말 삭제하시겠습니까?"))
	{
		frmDelete.submit();
	}
}
</script>
<img src="/img/blank.gif" height="5px" width="1px"><br>
<table width="100px" border=0 cellspacing=0 cellpadding=0>
<form name="frmDelete" action="ER_delete_action.asp" method="post">
<%
response.write strRequestForm
%>
<input type="hidden" name="ER_Code" value="<%=ER_Code%>">
<tr>
	<td align=center>
		<table width=100px cellpadding=0 cellspacing=0 border=0>
		<tr>
			<td width=100><%=Make_L_BTN("삭제하기","javascript:frmDelete_Check()","")%></td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
