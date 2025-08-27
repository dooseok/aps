<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" --> 
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim Part_filePrefix
dim Part_Title
dim Part_Title_Eng
Part_filePrefix = "bpmd"
Part_Title		= "개발"
Part_Title_Eng	= "Dev"

Dim RS1
Dim SQL

dim arrReg_Form(3,1)

dim Title
dim URL_Action
dim URL_Prev
dim URL_Next
dim URL_List
dim Form_Type
dim ColumBU_Width
dim Value_Width

call BOM_Guide()
arrReg_Form(0,0) = "*파트넘버"
arrReg_Form(0,1) = "<input type='text' name='BPM_PartNo' style='width:150px' onClick=""javascript:show_BOM_Guide(this,'frmCommonReg',0);"">"

arrReg_Form(1,0) = "*적용기간(시작일)"
arrReg_Form(1,1) = "<input type='text' name='BPM_StartDate' style='width:150px' readonly onclick='Calendar_D(document.frmRegForm.BPM_StartDate);'>"

arrReg_Form(2,0) = "*적용기간(종료일)"
arrReg_Form(2,1) = "<input type='text' name='BPM_EndDate' style='width:150px' readonly onclick='Calendar_D(document.frmRegForm.BPM_EndDate);'>&nbsp;<input type='button' value='미정' onclick=""javascript:document.frmRegForm.BPM_EndDate.value='2099-12-31'"">"

arrReg_Form(3,0) = "*내용"
arrReg_Form(3,1) = "<textarea name='BPM_Memo' style='width:90%' rows=20 style='border:1px solid #999999'></textarea>"


Title			= "제원시트메모-"&Part_Title
URL_Action		= Part_filePrefix & "_reg_action.asp"
URL_Prev		= Part_filePrefix & "_reg_form.asp"
URL_Next		= Part_filePrefix & "_list.asp"
URL_List		= Part_filePrefix & "_list.asp"
Form_Type		= ""
ColumBU_Width	= 180
Value_Width		= 700
%>
<script language="javascript">
function Form_Check(form)
{
	var strError = '';
	if(!form.BPM_PartNo.value)
	{
		strError += "*파트넘버를 입력해주세요.\n"
	}
	if(!form.BPM_StartDate.value)
	{
		strError += "*적용기간(시작일)을 입력해주세요.\n"
	}
	if(!form.BPM_EndDate.value)
	{
		strError += "*적용기간(종료일)을 입력해주세요.\n"
	}
	if(!form.BPM_Memo.value)
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
call Common_Reg_Form(Title, URL_Action, URL_Next, URL_List, Form_Type, ColumBU_Width, Value_Width, arrReg_Form)
%>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
