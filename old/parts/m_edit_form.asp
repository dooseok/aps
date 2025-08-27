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

dim M_Code
dim M_P_No
dim M_Desc
dim M_Spec
dim M_Additional_Info
dim M_Qty
dim M_Process

M_Code = Request("M_Code")
M_P_No = Request("M_P_No")

Set RS1 = Server.CreateObject("ADODB.RecordSet")
if M_P_No <> "" then
	SQL = "select * from tbMaterial where M_P_No='"&M_P_No&"'"
else
	SQL = "select * from tbMaterial where M_Code='"&M_Code&"'"
end if
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
%>
<form name="frmRedirect" action="m_edit_form.asp" method="post">
<input type="hidden" name="M_Code" value="<%=M_Code%>">
</form>
<script language="javascript">
alert("해당 파트넘버의 자재가 없습니다.");
frmRedirect.submit();
</script>
<%
else
	M_P_No				= RS1("M_P_No")
	M_Desc				= RS1("M_Desc")
	M_Spec				= RS1("M_Spec")
	M_Additional_Info	= RS1("M_Additional_Info")
	M_Qty				= RS1("M_Qty")
	M_Process			= RS1("M_Process")
end if
RS1.Close
Set RS1 = Nothing

strEdit_Header = "<input type='hidden' name='M_Code' value='"&M_Code&"'>" &vbcrlf

call Material_Guide()
%>
<table width=580px cellpadding=0 cellspacing=0 border=0>
<form name="frmSelectM_P_No" action="m_edit_form.asp" method="post">
<tr>
	<td align=right>
		파트넘버:<input type="hidden" name="M_Code" value="<%=M_Code%>">
		<input type="text" name="M_P_No" onclick="javascript:show_Material_Guide(this);"><input type="submit" value="이동">
	<td>
</tr>
</form>
</table>
<br>
<%
arrEdit_Form(0,0) = "* 파트넘버"
arrEdit_Form(0,1) = "<input type='text' name=""M_P_No"" value="""&M_P_No&""" style='width:200px'>"

arrEdit_Form(1,0) = "  종류"
arrEdit_Form(1,1) = "<input type='text' name=""M_Desc"" value="""&M_Desc&""" style='width:300px'>"

arrEdit_Form(2,0) = "  스펙"
arrEdit_Form(2,1) = "<input type='text' name=""M_Spec"" value="""&M_Spec&""" style='width:300px'>"

arrEdit_Form(3,0) = "  설명"
arrEdit_Form(3,1) = "<input type='text' name=""M_Additional_Info"" value="""&M_Additional_Info&""" style='width:300px'>"

arrEdit_Form(4,0) = "* 수량"
arrEdit_Form(4,1) = "<input type='text' name=""M_Qty"" value="""&M_Qty&""" style='width:40px'>"

dim strM_Process
dim arrBasicDataMaterialProcess
dim arrTemp
dim CNT1
strM_Process = "<select name='M_Process' style='width:100px'>"
arrBasicDataMaterialProcess = split(replace(BasicDataMaterialProcess,"slt>",""),";")
for CNT1 = 0 to ubound(arrBasicDataMaterialProcess)
	arrTemp = split(arrBasicDataMaterialProcess(CNT1),":")
	strM_Process = strM_Process & "<option value='"&arrTemp(0)&"'>"&arrTemp(1)&"</option>"
next
strM_Process = strM_Process & "</select>"

arrEdit_Form(5,0) = "공정"
arrEdit_Form(5,1) = strM_Process

Title			= "자재상세정보"
URL_Action		= "m_edit_action.asp"
URL_Prev		= "m_list.asp"
URL_Next		= "m_edit_form.asp"
URL_List		= "m_list.asp"
Form_Type		= ""
Column_Width	= 180
Value_Width		= 400
%>
<script language="javascript">
function Form_Check(form)
{
	var strError = '';
	if(!form.M_P_No.value)
	{
		strError += "*파트넘버를 입력해주세요.\n"
	}
	if(!form.M_Qty.value)
	{
		strError += "*수량을 입력해주세요.\n"
	}
	if(!IsNum(form.M_Qty.value))
	{
		strError += "*수량은 숫자만 입력해주세요.\n"
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
<br><br>
<center>
<iframe width=1200px height=400px src="/material/price/mp_list.asp?s_Material_M_P_No=<%=M_P_No%>" frameborder=0 style="border:1px solid #cccccc"></iframe>
</center>
<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
