<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" --> 
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
Dim RS1
Dim SQL

dim arrReg_Form(9,1)

dim Title
dim URL_Action
dim URL_Prev
dim URL_Next
dim URL_List
dim Form_Type
dim ColumBU_Width
dim Value_Width

call BOM_Guide()
arrReg_Form(0,0) = "*��Ʈ�ѹ�"
arrReg_Form(0,1) = "<input type='text' name='BOM_B_D_No' style='width:300px' onDblclick=""javascript:show_BOM_Guide(this,'frmCommonReg',0);"">"

arrReg_Form(1,0) = "*����"
arrReg_Form(1,1) = "<input type=checkbox name='BU_Type_New' value='Y'>�ű԰���&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name='BU_Type_Add' value='Y'>�۾��߰�&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name='BU_Type_Update' value='Y'>����ù�"

arrReg_Form(2,0) = "*����"
arrReg_Form(2,1) = "<textarea name='BU_Content' style='width:90%' rows=40 style='border:1px solid #999999'></textarea>"

arrReg_Form(3,0) = "*������"
arrReg_Form(3,1) = "<input type='text' name='BU_Receive_Date' style='width:150px' readonly onclick='Calendar_D(document.frmRegForm.BU_Receive_Date);'>"

arrReg_Form(4,0) = "*������"
arrReg_Form(4,1) = "<input type='text' name='BU_Apply_Date' style='width:150px' readonly onclick='Calendar_D(document.frmRegForm.BU_Apply_Date);'>"

arrReg_Form(5,0) = "ȸ����"
arrReg_Form(5,1) = "<input type='text' name='BU_Reply_Date' style='width:150px' readonly onclick='Calendar_D(document.frmRegForm.BU_Reply_Date);'>"

arrReg_Form(6,0) = "ȸ�ſ䱸��"
arrReg_Form(6,1) = "<input type='text' name='BU_Request_Reply_Date' style='width:150px' readonly onclick='Calendar_D(document.frmRegForm.BU_Request_Reply_Date);'>"

arrReg_Form(7,0) = "÷������1"
arrReg_Form(7,1) = "<input type='file' name='BU_File_1' style='width:300px'>"

arrReg_Form(8,0) = "÷������2"
arrReg_Form(8,1) = "<input type='file' name='BU_File_2' style='width:300px'>"

arrReg_Form(9,0) = "÷������3"
arrReg_Form(9,1) = "<input type='file' name='BU_File_3' style='width:300px'>"


Title			= "�űԽù���"
URL_Action		= "BU_reg_action.asp"
URL_Prev		= "BU_reg_form.asp"
URL_Next		= "BU_list.asp"
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
		strError += "*��Ʈ�ѹ��� �Է����ּ���.\n"
	}
	if(!form.BU_Content.value)
	{
		strError += "*������ �Է����ּ���.\n"
	}
	if(!form.BU_Receive_Date.value)
	{
		strError += "*�������� �Է����ּ���.\n"
	}
	if(!form.BU_Apply_Date.value)
	{
		strError += "*�������� �Է����ּ���.\n"
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
