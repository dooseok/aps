<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
Dim RS1
Dim SQL

dim arrReg_Form(4,1)

dim Title
dim URL_Action
dim URL_Prev
dim URL_Next
dim URL_List
dim Form_Type
dim Column_Width
dim Value_Width

arrReg_Form(0,0) = "*����"
arrReg_Form(0,1) = "<input type='text' name='N_Title' style='width:300px'>"

arrReg_Form(1,0) = "*����"
arrReg_Form(1,1) = "<textarea name='N_Content' style='width:90%' rows=20 style='border:1px solid #999999'></textarea>"

arrReg_Form(2,0) = "÷������1"
arrReg_Form(2,1) = "<input type='file' name='N_File_1' style='width:300px'>"

arrReg_Form(3,0) = "÷������2"
arrReg_Form(3,1) = "<input type='file' name='N_File_2' style='width:300px'>"

arrReg_Form(4,0) = "÷������3"
arrReg_Form(4,1) = "<input type='file' name='N_File_3' style='width:300px'>"


Title			= "�ű԰������"
URL_Action		= "n_reg_action.asp"
URL_Prev		= "n_reg_form.asp"
URL_Next		= "n_list.asp"
URL_List		= "n_list.asp"
Form_Type		= "enctype='MULTIPART/FORM-DATA'"
Column_Width	= 180
Value_Width		= 700
%>
<script language="javascript">
function Form_Check(form)
{
	var strError = '';
	if(!form.N_Title.value)
	{
		strError += "*������ �Է����ּ���.\n"
	}
	if(!form.N_Content.value)
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
call Common_Reg_Form(Title, URL_Action, URL_Next, URL_List, Form_Type, Column_Width, Value_Width, arrReg_Form)
%>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
