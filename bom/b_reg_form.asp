<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<%
Dim RS1
Dim SQL

dim arrReg_Form(12,1)

dim Title
dim URL_Action
dim URL_Prev
dim URL_Next
dim URL_List
dim Form_Type
dim Column_Width
dim Value_Width

Dim NEW_YN
NEW_YN = Request("NEW_YN")

If (NEW_YN="Y") or (NEW_YN="") Then
	arrReg_Form(0,0) = "*파트넘버"
	arrReg_Form(0,1) = "<input type='hidden' name='NEW_YN' value='"&NEW_YN&"'><input type='text' name='B_D_No' style='width:150px'>"
Else 
	arrReg_Form(0,0) = "*파트넘버"
	arrReg_Form(0,1) = arrReg_Form(0,1) & "<select name='B_D_No' style='width:150px'>"
	arrReg_Form(0,1) = arrReg_Form(0,1) & "<option value=''>---도번선택---</option>"
	Set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select distinct B_D_No from tbBOM order by B_D_No"
	RS1.Open SQL,sys_DBCon
	Do Until RS1.Eof
		arrReg_Form(0,1) = arrReg_Form(0,1) & "<option value='"&RS1("B_D_NO")&"'>"&RS1("B_D_NO")&"</option>"
		RS1.MoveNext
	Loop
	RS1.Close
	Set RS1 = Nothing
	arrReg_Form(0,1) = arrReg_Form(0,1) & "</select>"
End If

arrReg_Form(1,0) = "*시방번호"
arrReg_Form(1,1) = "<input type='text' name='B_Version_Code' value='' style='width:150px'>"

arrReg_Form(2,0) = "*현재적용중"
arrReg_Form(2,1) = arrReg_Form(2,1) & "<select name='B_Version_Current_YN'  style='width:50px'>"
arrReg_Form(2,1) = arrReg_Form(2,1) & "<option value='Y'"
arrReg_Form(2,1) = arrReg_Form(2,1) & ">Y</option>"
arrReg_Form(2,1) = arrReg_Form(2,1) & "<option value='N' selected"
arrReg_Form(2,1) = arrReg_Form(2,1) & ">N</option>"
arrReg_Form(2,1) = arrReg_Form(2,1) & "</select>"

arrReg_Form(3,0) = "*시방적용일"
arrReg_Form(3,1) = "<input type='text' name='B_Version_Date' style='width:150px'><img src='/img/ico_calender.jpg' onclick='Calendar_D(document.frmRegForm.B_Version_Date);' style='cursor:pointer'>"

arrReg_Form(4,0) = "*등록일"
arrReg_Form(4,1) = "<input type='text' name='B_Issue_Date' style='width:150px'><img src='/img/ico_calender.jpg' onclick='Calendar_D(document.frmRegForm.B_Issue_Date);' style='cursor:pointer'>"

arrReg_Form(5,0) = "메모"
arrReg_Form(5,1) = "<textarea name='B_Memo' style='width:90%;border:1px solid #999999' rows=10></textarea>"

arrReg_Form(6,0) = "모델"
arrReg_Form(6,1) = "<input type='text' name='B_Tool' style='width:100px'>"

arrReg_Form(7,0) = "구분"
arrReg_Form(7,1) = "<input type='text' name='B_Desc' style='width:100px'>"

arrReg_Form(8,0) = "스펙"
arrReg_Form(8,1) = "<input type='text' name='B_Spec' style='width:90%'>"

arrReg_Form(9,0) = "첨부파일1"
arrReg_Form(9,1) = "<input type='file' name='B_File_1' style='width:90%'>"

arrReg_Form(10,0) = "첨부파일2"
arrReg_Form(10,1) = "<input type='file' name='B_File_2' style='width:90%'>"

arrReg_Form(11,0) = "첨부파일3"
arrReg_Form(11,1) = "<input type='file' name='B_File_3' style='width:90%'>"

arrReg_Form(12,0) = "첨부파일4"
arrReg_Form(12,1) = "<input type='file' name='B_File_4' style='width:90%'>"

If (NEW_YN="Y") Then
	Title			= "BOM신규등록"
Else
	Title			= "BOM변경등록"
End If
URL_Action		= "b_reg_action.asp"
URL_Prev		= "b_reg_form.asp"
URL_Next		= "b_list.asp"
URL_List		= "b_list.asp"
Form_Type		= "enctype='MULTIPART/FORM-DATA'"
Column_Width	= 180
Value_Width		= 400
%>
<script language="javascript">
var date_pattern = /^(19|20)\d{2}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[0-1])$/; 

function Form_Check(form)
{
	var strError = '';
	if(!form.B_D_No.value)
	{
		strError += "*도번을 입력해주세요.\n"
	}
	if(!form.B_Issue_Date.value)
	{
		strError += "*등록일을 입력해주세요.\n"
	}
	if(!date_pattern.test(form.B_Version_Date.value))
	{
		strError += "*시방적용일을 날짜형식(YYYY-MM-DD)으로 입력해주세요.\n"
	}
	if(!date_pattern.test(form.B_Issue_Date.value))
	{
		strError += "*등록일을 날짜형식(YYYY-MM-DD)으로 입력해주세요.\n"
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
<!-- #include virtual = "/header/session_check_tail.asp" -->