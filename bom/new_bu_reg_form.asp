<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" --> 
<!-- include virtual = "/header/html_header_ex.asp" --> 
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
Dim RS1
Dim SQL

dim arrReg_Form(18,1)

dim Title
dim URL_Action
dim URL_Prev
dim URL_Next
dim URL_List
dim Form_Type
dim ColumBU_Width
dim Value_Width

call BOM_Guide()


'select 
'	BU_Code,

'	BU_LG_Part,
'	BU_LG_Staff,
'	BU_Eco_No,
'	BOM_B_D_No,
'	BU_Parts_PNO,
'	BU_Content= left(BU_Content,150),
	
'	BU_Apply_Date= convert(char(10),BU_Apply_Date,121),	
'	BU_Receive_Date= convert(char(10),BU_Receive_Date,121),
'	BU_Last_Use_Date= convert(char(10),BU_Last_Use_Date,121),

'BU_File_1,
'BU_File_2,
'BU_File_3,
'Member_M_ID,




'BU_Type = case right(BU_Type,1) when '-' then left(BU_Type,len(BU_Type)-1) else BU_Type end 



arrReg_Form(0,0) = "사업부"
arrReg_Form(0,1) = "<input type='text' name='BU_LG_Part' style='width:300px' >"

arrReg_Form(1,0) = "담당연구원"
arrReg_Form(1,1) = "<input type='text' name='BU_LG_Staff' style='width:300px' >"

arrReg_Form(2,0) = "구분"
arrReg_Form(2,1) = arrReg_Form(2,1) & "<input type=checkbox name='BU_Type_SW' value='Y'>S/W&nbsp;&nbsp;&nbsp;&nbsp;"
arrReg_Form(2,1) = arrReg_Form(2,1) & "<input type=checkbox name='BU_Type_HW' value='Y'>H/W&nbsp;&nbsp;&nbsp;&nbsp;"
arrReg_Form(2,1) = arrReg_Form(2,1) & "<input type=checkbox name='BU_Type_REAL' value='Y'>현실화&nbsp;&nbsp;&nbsp;&nbsp;"
arrReg_Form(2,1) = arrReg_Form(2,1) & "<input type=checkbox name='BU_Type_SAMPLE' value='Y'>샘플 후 폐기&nbsp;&nbsp;&nbsp;&nbsp;"

arrReg_Form(3,0) = "ECO No"
arrReg_Form(3,1) = "<input type='text' name='BU_Eco_No' style='width:300px'>"

arrReg_Form(4,0) = "시방번호"
arrReg_Form(4,1) = "<input type='text' name='BU_Sibang_No' style='width:300px'>"

arrReg_Form(5,0) = "도면 파트넘버"
arrReg_Form(5,1) = "<input type='text' name='BOM_B_D_No' style='width:300px' onDblclick=""javascript:show_BOM_Guide(this,'frmCommonReg',0);"">"

arrReg_Form(6,0) = "부품 파트넘버"
arrReg_Form(6,1) = "<textarea name='BU_Parts_PNO' style='width:90%' rows=5 style='border:1px solid #999999'></textarea>"

arrReg_Form(7,0) = "*시방내용"
arrReg_Form(7,1) = "<textarea name='BU_Content' style='width:90%' rows=15 style='border:1px solid #999999'></textarea>"

arrReg_Form(8,0) = "*접수일"
arrReg_Form(8,1) = "<input type='text' name='BU_Receive_Date' style='width:150px' readonly onclick='Calendar_D(document.frmRegForm.BU_Receive_Date);'>"
'arrReg_Form(8,1) = "<input type='text' name='BU_Receive_Date' style='width:150px' id='idBU_Receive_Date'>"

arrReg_Form(9,0) = "*적용일"
arrReg_Form(9,1) = "<input type='text' name='BU_Apply_Date' style='width:150px' readonly onclick='Calendar_D(document.frmRegForm.BU_Apply_Date);'>"
'arrReg_Form(9,1) = "<input type='text' name='BU_Apply_Date' style='width:150px' id='idBU_Apply_Date'>"

arrReg_Form(10,0) = "소진일정"
arrReg_Form(10,1) = "<textarea name='BU_Last_Use_Date' style='width:90%' rows=2 style='border:1px solid #999999'></textarea>"

arrReg_Form(11,0) = "기준라인"
arrReg_Form(11,1) = "<select name='BU_MSE_LG' style='width:70px'>"
arrReg_Form(11,1) = arrReg_Form(11,1) & "<option value='MSE'"
arrReg_Form(11,1) = arrReg_Form(11,1) & " style=''>MSE</option>"
arrReg_Form(11,1) = arrReg_Form(11,1) & "<option value='LG'"
arrReg_Form(11,1) = arrReg_Form(11,1) & " style=''>LG</option>"
arrReg_Form(11,1) = arrReg_Form(11,1) & "</select>"

arrReg_Form(12,0) = "전송여부"
arrReg_Form(12,1) = "<select name='BU_Link_YN' style='width:70px'>"
arrReg_Form(12,1) = arrReg_Form(12,1) & "<option value='전송안함'"
arrReg_Form(12,1) = arrReg_Form(12,1) & " style=''>전송안함</option>"
arrReg_Form(12,1) = arrReg_Form(12,1) & "<option value='전송함'"
arrReg_Form(12,1) = arrReg_Form(12,1) & " style=''>전송함</option>"
arrReg_Form(12,1) = arrReg_Form(12,1) & "</select>"

arrReg_Form(13,0) = "첨부파일1 ( 품 번 )"
arrReg_Form(13,1) = "<input type='File' name='BU_File_PartNo' style='width:300px'>"

arrReg_Form(14,0) = "첨부파일2 (시방서)"
arrReg_Form(14,1) = "<input type='File' name='BU_File_1' style='width:300px'>"

arrReg_Form(15,0) = "첨부파일3 ( 도 면 )"
arrReg_Form(15,1) = "<input type='File' name='BU_File_2' style='width:300px'>"

arrReg_Form(16,0) = "첨부파일4 ( 기타1 )"
arrReg_Form(16,1) = "<input type='File' name='BU_File_3' style='width:300px'>"

arrReg_Form(17,0) = "첨부파일5 ( 기타2 )"
arrReg_Form(17,1) = "<input type='File' name='BU_File_4' style='width:300px'>"

arrReg_Form(18,0) = "첨부파일6 ( 기타3 )"
arrReg_Form(18,1) = "<input type='File' name='BU_File_5' style='width:300px'>"

Title			= "신규시방등록"
URL_Action		= "new_BU_reg_action.asp"
URL_Prev		= "new_BU_reg_form.asp"
URL_Next		= "new_BU_list.asp"
URL_List		= "new_BU_list.asp"
Form_Type		= "enctype='MULTIPART/FORM-DATA'"
ColumBU_Width	= 180
Value_Width		= 700
%>



<script language="javascript">

function Form_Check(form)
{
	var strError = '';
	if(!form.BU_Content.value)
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

console.log(3);
</script>
<%
call Common_Reg_Form(Title, URL_Action, URL_Next, URL_List, Form_Type, ColumBU_Width, Value_Width, arrReg_Form)
%>

<script language="javascript">
<!--$(document).ready(function(){	
	$('#idBU_Receive_Date').flatpickr({locale: 'ko', dateFormat: 'Y-m-d', enableTime:false,<%'=strFlatPickrArg%>});
	$('#idBU_Apply_Date').flatpickr({locale: 'ko', dateFormat: 'Y-m-d', enableTime:false,<%'=strFlatPickrArg%>});
});-->
</script>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
