<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim Part_filePrefix
dim Part_Title
dim Part_Title_Eng
Part_filePrefix = "bpmq"
Part_Title		= "ǰ��"
Part_Title_Eng	= "QA"

Dim RS1
Dim SQL

Dim strEdit_Header
dim arrEdit_Form(3,1)
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

dim BPM_Code
dim BPM_PartNo
dim BPM_StartDate
dim BPM_EndDate
dim BPM_Memo

BPM_Code = Request("BPM_Code")

Set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbBOM_PartsOutSheet_Memo_"&Part_Title_Eng&" where BPM_Code='"&BPM_Code&"'"
RS1.Open SQL,sys_DBCon
BPM_PartNo		= RS1("BPM_PartNo")
BPM_StartDate	= RS1("BPM_StartDate")
BPM_EndDate		= RS1("BPM_EndDate")
BPM_Memo		= RS1("BPM_Memo")
RS1.Close
Set RS1 = Nothing

strEdit_Header = "<input type='hidden' name='BPM_Code' value='"&BPM_Code&"'>" &vbcrlf

call BOM_Guide()
arrEdit_Form(0,0) = "*��Ʈ�ѹ�"
arrEdit_Form(0,1) = "<input type='text' name='BPM_PartNo' style='width:150px' value='"&BPM_PartNo&"' onClick=""javascript:show_BOM_Guide(this,'frmCommonEdit',0);"">"

arrEdit_Form(1,0) = "*����Ⱓ(������)"
arrEdit_Form(1,1) = "<input type='text' name='BPM_StartDate' style='width:150px' readonly value='"&BPM_StartDate&"' onclick='Calendar_D(document.frmEditForm.BPM_StartDate);'>"

arrEdit_Form(2,0) = "*����Ⱓ(������)"
arrEdit_Form(2,1) = "<input type='text' name='BPM_EndDate' style='width:150px' readonly value='"&BPM_EndDate&"' onclick='Calendar_D(document.frmEditForm.BPM_EndDate);'>&nbsp;<input type='button' value='����' onclick=""javascript:document.frmEditForm.BPM_EndDate.value='2099-12-31'"">"

arrEdit_Form(3,0) = "*����"
arrEdit_Form(3,1) = "<textarea name='BPM_Memo' style='width:90%' rows=20 style='border:1px solid #999999'>"&BPM_Memo&"</textarea>"


Title			= "������Ʈ�ֱ�-"& Part_Title
URL_Action		= Part_filePrefix & "_edit_action.asp"
URL_Prev		= Part_filePrefix & "_list.asp"
URL_Next		= Part_filePrefix & "_edit_form.asp"
URL_List		= Part_filePrefix & "_list.asp"
Form_Type		= ""
Column_Width	= 180
Value_Width		= 400
%>
<script language="javascript">
function Form_Check(form)
{
	var strError = '';
	if(!form.BPM_PartNo.value)
	{
		strError += "*��Ʈ�ѹ��� �Է����ּ���.\n"
	}
	if(!form.BPM_StartDate.value)
	{
		strError += "*����Ⱓ(������)�� �Է����ּ���.\n"
	}
	if(!form.BPM_EndDate.value)
	{
		strError += "*����Ⱓ(������)�� �Է����ּ���.\n"
	}
	if(!form.BPM_Memo.value)
	{
		strError += "*������ �Է����ּ���.\n"
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
