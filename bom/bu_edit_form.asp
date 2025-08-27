<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<%
Dim RS1
Dim SQL

Dim strEdit_Header
dim arrEdit_Form(9,1)
dim B_Code
dim Title
dim URL_Action
dim URL_Prev
dim URL_Next
dim URL_List
dim Form_Type
dim ColumBU_Width
dim Value_Width

dim BU_Code
dim BOM_B_D_No
dim BU_Content
dim BU_Receive_Date
dim BU_Apply_Date
dim BU_Reply_Date
dim BU_Request_Reply_Date
dim BU_File_1
dim BU_File_2
dim BU_File_3
dim BU_Type


BU_Code = Request("BU_Code")

Set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbBOM_Update where BU_Code='"&BU_Code&"'"
RS1.Open SQL,sys_DBCon
BOM_B_D_No		= RS1("BOM_B_D_No")
BU_Content		= RS1("BU_Content")
BU_Receive_Date	= RS1("BU_Receive_Date")
BU_Apply_Date	= RS1("BU_Apply_Date")
BU_Reply_Date	= RS1("BU_Reply_Date")
BU_Request_Reply_Date	= RS1("BU_Request_Reply_Date")
BU_File_1		= RS1("BU_File_1")
BU_File_2		= RS1("BU_File_2")
BU_File_3		= RS1("BU_File_3")
BU_Type			= RS1("BU_Type")

RS1.Close
Set RS1 = Nothing

strEdit_Header = "<input type='hidden' name='BU_Code' value='"&BU_Code&"'>" &vbcrlf

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

call BOM_Guide()
arrEdit_Form(0,0) = "*파트넘버"
arrEdit_Form(0,1) = "<input type='text' name='BOM_B_D_No' value='"&BOM_B_D_No&"' style='width:300px' onDblclick=""javascript:show_BOM_Guide(this,'frmCommonReg',0);"">"

arrEdit_Form(1,0) = "*구분"
arrEdit_Form(1,1) = "<input type=checkbox name='BU_Type_New' value='Y'"

if instr(BU_Type,"신규") > 0 then
	arrEdit_Form(1,1) = arrEdit_Form(1,1) & " checked"
end if
arrEdit_Form(1,1) = arrEdit_Form(1,1) & ">신규개발&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name='BU_Type_Add' value='Y'"
if instr(BU_Type,"추가") > 0 then
	arrEdit_Form(1,1) = arrEdit_Form(1,1) & " checked"
end if
arrEdit_Form(1,1) = arrEdit_Form(1,1) & ">작업추가&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name='BU_Type_Update' value='Y'"
if instr(BU_Type,"시방") > 0 then
	arrEdit_Form(1,1) = arrEdit_Form(1,1) & " checked"
end if
arrEdit_Form(1,1) = arrEdit_Form(1,1) & ">도면시방"

arrEdit_Form(2,0) = "*내용"
arrEdit_Form(2,1) = "<textarea name='BU_Content' style='width:90%' rows=40 style='border:1px solid #999999'>"&BU_Content&"</textarea>"

arrEdit_Form(3,0) = "*접수일"
arrEdit_Form(3,1) = "<input type='text' name='BU_Receive_Date' value='"&BU_Receive_Date&"' style='width:150px' readonly onclick='Calendar_D(document.frmEditForm.BU_Receive_Date);'>"

arrEdit_Form(4,0) = "*적용일"
arrEdit_Form(4,1) = "<input type='text' name='BU_Apply_Date' value='"&BU_Apply_Date&"' style='width:150px' readonly onclick='Calendar_D(document.frmEditForm.BU_Apply_Date);'>"

arrEdit_Form(5,0) = "회신일"
arrEdit_Form(5,1) = "<input type='text' name='BU_Reply_Date' value='"&BU_Reply_Date&"' style='width:150px' readonly onclick='Calendar_D(document.frmEditForm.BU_Reply_Date);'>"

arrEdit_Form(6,0) = "회신요구일"
arrEdit_Form(6,1) = "<input type='text' name='BU_Request_Reply_Date' value='"&BU_Request_Reply_Date&"' style='width:150px' readonly onclick='Calendar_D(document.frmEditForm.BU_Request_Reply_Date);'>"

arrEdit_Form(7,0) = "첨부파일1"
arrEdit_Form(7,1) = "<input type='hidden' name='oldBU_File_1' value='"&BU_File_1&"'>"
If BU_File_1 <> "" then
	arrEdit_Form(7,1) = arrEdit_Form(7,1) & "<a href='/function/ifrm_download.asp?filepath="&DefaultPath_BOM&BU_File_1&"' target='ifrm_download'>"&BU_File_1&"</a>"
	arrEdit_Form(7,1) = arrEdit_Form(7,1) & "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='strDelete' value='BU_File_1'>삭제"
End if
arrEdit_Form(7,1) = arrEdit_Form(7,1) & "<br><input type='file' name='BU_File_1' style='width:90%'>"

arrEdit_Form(8,0) = "첨부파일2"
arrEdit_Form(8,1) = "<input type='hidden' name='oldBU_File_2' value='"&BU_File_2&"'>"
If BU_File_2 <> "" then
	arrEdit_Form(8,1) = arrEdit_Form(8,1) & "<a href='/function/ifrm_download.asp?filepath="&DefaultPath_BOM&BU_File_2&"' target='ifrm_download'>"&BU_File_2&"</a>"
	arrEdit_Form(8,1) = arrEdit_Form(8,1) & "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='strDelete' value='BU_File_2'>삭제"
End if
arrEdit_Form(8,1) = arrEdit_Form(8,1) & "<br><input type='file' name='BU_File_2' style='width:90%'>"

arrEdit_Form(9,0) = "첨부파일3"
arrEdit_Form(9,1) = "<input type='hidden' name='oldBU_File_3' value='"&BU_File_3&"'>"
If BU_File_3 <> "" then
	arrEdit_Form(9,1) = arrEdit_Form(9,1) & "<a href='/function/ifrm_download.asp?filepath="&DefaultPath_BOM&BU_File_3&"' target='ifrm_download'>"&BU_File_3&"</a>"
	arrEdit_Form(9,1) = arrEdit_Form(9,1) & "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='strDelete' value='BU_File_3'>삭제"
End if
arrEdit_Form(9,1) = arrEdit_Form(9,1) & "<br><input type='file' name='BU_File_3' style='width:90%'>"

Title			= "시방사항수정"


if instr(admin_bu_list,"-"&gM_ID&"-") > 0 then
	URL_Action		= "BU_edit_action.asp"
else
	URL_Action		= "/function/admin_denide.asp?URL_Back=/bom/bu_edit_form.asp&strPK=BU_Code&strPK_Value="&BU_Code
end if

URL_Prev		= "BU_edit_form.asp"
URL_Next		= "BU_edit_form.asp"
URL_List		= "BU_list.asp"
Form_Type		= "enctype='MULTIPART/FORM-DATA'"
ColumBU_Width	= 180
Value_Width		= 700
%>
<script language="javascript">
function Form_Check(form)
{
	var strError = '';
	if(!form.BOM_B_D_No.value)
	{
		strError += "*파트넘버를 입력해주세요.\n"
	}
	if(!form.BU_Content.value)
	{
		strError += "*내용을 입력해주세요.\n"
	}
	if(!form.BU_Receive_Date.value)
	{
		strError += "*접수일을 입력해주세요.\n"
	}
	if(!form.BU_Apply_Date.value)
	{
		strError += "*적용일을 입력해주세요.\n"
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
call Common_Edit_Form(Title, URL_Action, URL_Next, URL_List, Form_Type, ColumBU_Width, Value_Width, strEdit_Header, arrEdit_Form, strRequestForm)
%>

<script language="javascript">
function printForm()
{
	alert("확인을 누르신 후 잠시 기다리시면\n인쇄창이 뜹니다.");
	window.open("bu_print.asp?bu_code=<%=BU_Code%>","PartsOrderSheet","height="+screen.height+",width="+screen.width+",status=yes,toolbar=yes,location=yes,directories=yes,location=yes,menubar=yes,resizable=yes,scrollbars=yes,titlebar=yes");
}	
</script>

<img src="/img/blank.gif" width=10px height=10px><br>
<table width=150px cellpadding=0 cellspacing=0 border=0>
<tr>
	<td width=150><%=Make_L_BTN("확인서인쇄","javascript:printForm();","")%></td>
</tr>
</table>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
