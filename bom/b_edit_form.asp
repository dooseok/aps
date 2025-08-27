<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<%
Dim RS1
Dim SQL

Dim strEdit_Header
dim arrEdit_Form(15,1)
dim B_Code
dim Title
dim URL_Action
dim URL_Prev
dim URL_Next
dim URL_List
dim Form_Type
dim Column_Width
dim Value_Width

dim B_D_No
dim B_Version_Code
dim B_Version_Date
dim B_Version_Current_YN
dim B_Issue_Date
dim B_Tool
dim B_Desc
dim B_Spec
Dim B_File_1
Dim B_File_2
Dim B_File_3
Dim B_File_4
Dim B_File_5
Dim B_File_6
Dim B_State
Dim B_Memo
Dim B_Reg_Date
Dim B_Edit_Date
dim Bom_Sub_Cnt

Dim NEW_YN

B_Code = Request("B_Code")
NEW_YN = Request("NEW_YN")

Set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from vwB_List where B_Code='"&B_Code&"'"
RS1.Open SQL,sys_DBCon
B_D_No			= RS1("B_D_No")
B_Version_Code			= RS1("B_Version_Code")
B_Version_Date			= RS1("B_Version_Date")
B_Version_Current_YN	= RS1("B_Version_Current_YN")
B_Issue_Date	= RS1("B_Issue_Date")
B_Tool			= RS1("B_Tool")
B_Desc			= RS1("B_Desc")
B_Spec			= RS1("B_Spec")
B_File_1		= RS1("B_File_1")
B_File_2		= RS1("B_File_2")
B_File_3		= RS1("B_File_3")
B_File_4		= RS1("B_File_4")
B_Memo			= RS1("B_Memo")
B_Reg_Date		= RS1("B_Reg_Date")
B_Edit_Date		= RS1("B_Edit_Date")
Bom_Sub_Cnt		 = RS1("Bom_Sub_Cnt")
RS1.Close
set RS1 = nothing

strEdit_Header = "<input type='hidden' name='B_Code' value='"&B_Code&"'>" &vbcrlf
strEdit_Header = strEdit_Header & "<input type='hidden' name='B_Version_Current_YN_old' value='"&B_Version_Current_YN&"'>" &vbcrlf
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
%>

<script language="javascript">
function delBOM(B_Code)
{
	if(confirm("[경고!]\n[확인]을 누르시면 [<%=B_D_No%>]와 관련된 모든 정보가 삭제됩니다.\n문의사항이 있으시면 IT담당자에게 문의해주세요."))
	{
		location.href="b_edit_form_del_action.asp?B_Code=<%=B_Code%>&B_D_No=<%=B_D_No%>"
	}
}

function delBOMSub(B_Code)
{
	var strBS_D_No = prompt("삭제할 품번을 입력해주세요. 예) EBR35935601","");
	if(strBS_D_No != null)
	{
		location.href="b_edit_form_del_sub_action.asp?B_Code=<%=B_Code%>&B_D_No=<%=B_D_No%>&BS_D_No="+strBS_D_No;
	}
}

function copyBOM(B_Code)
{
	if(confirm("!주의! 반드시 완료메세지가 보일 때까지 기다려주세요.\n(예상 소요시간 1분)\n[확인]을 누르시면 [<%=B_D_No%>]를 복사합니다."))
	{
		location.href="b_edit_form_copy_action.asp?B_Code=<%=B_Code%>&B_D_No=<%=B_D_No%>&B_Version_Code=<%=server.URLEncode(B_Version_Code)%>"
	}
}


function XLS_Download()
{
	confirm('다운로드 창이 뜰 때까지, 잠시만 기다려주세요.\n예상 소요시간 1분')
	{
		ifrmXLSDown.location.href="/bom/xls_download_action.asp?b_d_no=<%=B_D_No%>&b_code=<%=B_Code%>";
	}
}

function DiffView()
{
	if(confirm('[<%=B_D_No%>]의 핸재시방과 이전시방을 비교하는 뷰어입니다.\n일반 뷰어보다 로딩속도가 느릴 수 있습니다.'))
	{	
		window.open('db_load_action.asp?Diff_YN=Y&B_Code=<%=B_Code%>');
	}
}

function BOMPrint()
{
	<%if Bom_Sub_Cnt > 30 then%>
	alert("옵션이 30개를 초과하는 경우에는 인쇄를 지원하지 않습니다.\n도면을 엑셀로 다운로드하여 출력해 주십시오.");
	<%else%>
	if(confirm('도면을 인쇄합니다.\n인쇄 창이 뜰 때까지 잠시만 기다려주세요.'))
	{	
		window.open("/bom/db_load_action.asp?b_code=<%=B_Code%>&mode=print","BOMPrint","height=100px,width=100px,top=2000px,lef=2000px,status=yes,toolbar=yes,location=yes,directories=yes,location=yes,menubar=yes,resizable=yes,scrollbars=yes,titlebar=yes")
	}
	<%end if%>
}
</script>


<%
arrEdit_Form(0,0) = "*파트넘버"
arrEdit_Form(0,1) = B_D_No
arrEdit_Form(0,1) = ""
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "<table width=428px cellpadding=0 cellspacing=0 border=0>"
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "<tr>"
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "	<td width=90px align=left><input type='text' name='B_D_No' value='"&B_D_No&"' readonly style='width:150px'></td><td width=10px>&nbsp;</td>"
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "	<td width=77px align=left>"&Make_BTN("BOM복사","javascript:copyBOM('"&B_Code&"');","")&"</td>"
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "	<td width=10px align=left>&nbsp;</td>" 
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "	<td width=77px align=left>"&Make_BTN("BOM삭제","javascript:delBOM('"&B_Code&"');","")&"</td>"
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "	<td width=10px align=left>&nbsp;</td>"
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "	<td width=77px align=left>"&Make_BTN("품번삭제","javascript:delBOMSub('"&B_Code&"');","")&"</td>"
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "	<td width=10px align=left>&nbsp;</td>"
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "	<td align=left>&nbsp;</td>"  
arrEdit_Form(0,1) = arrEdit_Form(0,1) & "</table>"


arrEdit_Form(1,0) = "*시방번호"
arrEdit_Form(1,1) = ""
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "<table width=428px cellpadding=0 cellspacing=0 border=0>"
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "<tr>"
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "	<td width=90px align=left><input type='text' name='B_Version_Code' value='"&B_Version_Code&"' style='width:150px'></td><td width=10px>&nbsp;</td>"
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "	<td width=77px align=left>"&Make_BTN("BOM뷰어","javascript:window.open('db_load_action.asp?B_Code="&B_Code&"');","")&"</td>"
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "	<td width=10px align=left>&nbsp;</td>" 
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "	<td width=77px align=left>"&Make_BTN("DIFF뷰어","javascript:DiffView();","")&"</td>"
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "	<td width=10px align=left>&nbsp;<iframe id='ifrmXLSDown' src='about:blank' frameborder=0 width=0px height=0px></iframe></td>"
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "	<td width=77px align=left>"&Make_BTN("PRINT","javascript:BOMPrint();","")&"</td>"
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "	<td width=10px align=left>&nbsp;<iframe id='ifrmBOMPrint' src='about:blank' frameborder=0 width=0px height=0px></iframe></td>"
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "	<td align=left>&nbsp;</td>" 
arrEdit_Form(1,1) = arrEdit_Form(1,1) & "</table>"


arrEdit_Form(2,0) = "*현재적용중"
arrEdit_Form(2,1) = arrEdit_Form(2,1) & "<select name='B_Version_Current_YN'  style='width:50px'>"
arrEdit_Form(2,1) = arrEdit_Form(2,1) & "<option value='Y'"
if B_Version_Current_YN = "Y" then
	arrEdit_Form(2,1) = arrEdit_Form(2,1) & " selected"
end if
arrEdit_Form(2,1) = arrEdit_Form(2,1) & ">Y</option>"
arrEdit_Form(2,1) = arrEdit_Form(2,1) & "<option value='N'"
if B_Version_Current_YN = "N" then
	arrEdit_Form(2,1) = arrEdit_Form(2,1) & " selected"
end if
arrEdit_Form(2,1) = arrEdit_Form(2,1) & ">N</option>"
arrEdit_Form(2,1) = arrEdit_Form(2,1) & "</select>"

arrEdit_Form(3,0) = "*시방적용일"
arrEdit_Form(3,1) = "<input type='text' name='B_Version_Date' value="&B_Version_Date&" style='width:150px'><img src='/img/ico_calender.jpg' onclick='Calendar_D(document.frmEditForm.B_Version_Date);' style='cursor:pointer'>"

arrEdit_Form(4,0) = "*등록일"
arrEdit_Form(4,1) = "<input type='text' name='B_Issue_Date' value="&B_Issue_Date&" style='width:150px'><img src='/img/ico_calender.jpg' onclick='Calendar_D(document.frmEditForm.B_Issue_Date);' style='cursor:pointer'>"

arrEdit_Form(5,0) = "메모"
arrEdit_Form(5,1) = "<textarea name='B_Memo' style='width:90%;border:1px solid #999999' rows=3>"&B_Memo&"</textarea>"

arrEdit_Form(6,0) = "모델"
arrEdit_Form(6,1) = "<input type='text' name='B_Tool' value='"&B_Tool&"' style='width:100px'>"

arrEdit_Form(7,0) = "구분"
arrEdit_Form(7,1) = "<input type='text' name='B_Desc' value='"&B_Desc&"' style='width:100px'>"

arrEdit_Form(8,0) = "스펙"
arrEdit_Form(8,1) = "<input type='text' name='B_Spec' value='"&B_Spec&"' style='width:90%'>"

arrEdit_Form(9,0) = "첨부파일 1"
arrEdit_Form(9,1) = "<input type='hidden' name='oldB_File_1' value='"&B_File_1&"'>"
If B_File_1 <> "" then
	arrEdit_Form(9,1) = arrEdit_Form(9,1) & "<a href='/function/ifrm_download.asp?filepath="&DefaultPath_BOM&B_File_1&"' target='ifrm_download'>"&B_File_1&"</a>"
	arrEdit_Form(9,1) = arrEdit_Form(9,1) & "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='strDelete' value='B_File_1'>delete"
End if
arrEdit_Form(9,1) = arrEdit_Form(9,1) & "<br><input type='file' name='B_File_1' style='width:90%'>"

arrEdit_Form(10,0) = "첨부파일 2"
arrEdit_Form(10,1) = "<input type='hidden' name='oldB_File_2' value='"&B_File_2&"'>"
If B_File_2 <> "" then
	arrEdit_Form(10,1) = arrEdit_Form(10,1) & "<a href='/function/ifrm_download.asp?filepath="&DefaultPath_BOM&B_File_2&"' target='ifrm_download'>"&B_File_2&"</a>"
	arrEdit_Form(10,1) = arrEdit_Form(10,1) & "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='strDelete' value='B_File_2'>delete"
End if
arrEdit_Form(10,1) = arrEdit_Form(10,1) & "<br><input type='file' name='B_File_2' style='width:90%'>"

arrEdit_Form(11,0) = "첨부파일 3"
arrEdit_Form(11,1) = "<input type='hidden' name='oldB_File_3' value='"&B_File_3&"'>"
If B_File_3 <> "" then
	arrEdit_Form(11,1) = arrEdit_Form(11,1) & "<a href='/function/ifrm_download.asp?filepath="&DefaultPath_BOM&B_File_3&"' target='ifrm_download'>"&B_File_3&"</a>"
	arrEdit_Form(11,1) = arrEdit_Form(11,1) & "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='strDelete' value='B_File_3'>delete"
End if
arrEdit_Form(11,1) = arrEdit_Form(11,1) & "<br><input type='file' name='B_File_3' style='width:90%'>"

arrEdit_Form(12,0) = "첨부파일 4"
arrEdit_Form(12,1) = "<input type='hidden' name='oldB_File_4' value='"&B_File_4&"'>"
If B_File_4 <> "" Then
	arrEdit_Form(12,1) = arrEdit_Form(12,1) & "<a href='/function/ifrm_download.asp?filepath="&DefaultPath_BOM&B_File_4&"' target='ifrm_download'>"&B_File_4&"</a>"
	arrEdit_Form(12,1) = arrEdit_Form(12,1) & "&nbsp;&nbsp;&nbsp;<input type='checkbox' name='strDelete' value='B_File_4'>delete"
End if
arrEdit_Form(12,1) = arrEdit_Form(12,1) & "<br><input type='file' name='B_File_4' style='width:90%'>"

arrEdit_Form(13,0) = "최초등록일"
arrEdit_Form(13,1) = B_Reg_Date

arrEdit_Form(14,0) = "최종수정일"
arrEdit_Form(14,1) = B_Edit_Date

Title				= "Edit BOM"
URL_Action	= "b_edit_action.asp"
URL_Prev		= "b_edit_form.asp"
URL_Next		= "b_edit_form.asp"
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
		if(confirm("!주의! 반드시 완료메세지가 보일 때까지 기다려주세요.\n(예상 소요시간 1분)"))
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

<br>
<br>
<table width="580px" border=0 cellspacing=1 cellpadding=0 bgcolor="#999999" class="Common_List">
<tr height=33px bgcolor="#e0e0e0">

	
	<td width="110px"><b style="color:navy">파트넘버</b></td>
	<td width="170px"><b style="color:navy">시방번호</b></td>
	<td width="100px"><b style="color:navy">현재적용</b></td>
	<td width="100px"><b style="color:navy">시방적용일</b></td>
	<td width="100px"><b style="color:navy">작업</b></td>
</tr>
<%
set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbBOM where B_D_No = '"&B_D_No&"' order by B_Code desc"
RS1.Open SQL, sys_DBCon
do until RS1.Eof
%>
<tr height=28px bgcolor="#ffffff" valign=top <%if int(B_Code) = int(RS1("B_Code")) then%>style="background-Color='skyblue';"<%end if%>>
	<td align="Center" valign="Center"><a href="db_load_action.asp?b_code=<%=RS1("B_Code")%>" style='color:blue' target="_blank"><%=RS1("B_D_No")%></a></td>
	<td align="Center" valign="Center"><%=RS1("B_Version_Code")%></a></td>
	<td align="Center" valign="Center"><%=RS1("B_Version_Current_YN")%></a></td>
	<td align="Center" valign="Center"><%=RS1("B_Version_Date")%></a></td>
	<td valign=middle>
		<span style="cursor:hand;color:navy" onclick="javascript:location.href='/bom/b_edit_form.asp?b_code=<%=RS1("B_Code")%>'"><u>보기</u></span>
	</td>
</tr>

<%
	RS1.MoveNext
loop
RS1.Close
set RS1 = nothing
%>
</table>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->