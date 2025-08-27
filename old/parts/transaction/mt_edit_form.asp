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
dim arrEdit_Form(4,1)
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

dim MT_Code
dim MT_Date
dim MT_Remark
dim MT_Company
dim MT_State

MT_Code = Request("MT_Code")

Set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbMaterial_Transaction where MT_Code='"&MT_Code&"'"
RS1.Open SQL,sys_DBCon
MT_Date			= RS1("MT_Date")
MT_Remark		= RS1("MT_Remark")
MT_State		= RS1("MT_State")
MT_Company		= RS1("MT_Company")

RS1.Close
Set RS1 = Nothing

strEdit_Header = "<input type='hidden' name='MT_Code' value='"&MT_Code&"'>" &vbcrlf

arrEdit_Form(0,0) = "번호"
arrEdit_Form(0,1) = MT_Code

arrEdit_Form(1,0) = "* 날짜"
arrEdit_Form(1,1) = "<input type='text' name='MT_Date' value="""&MT_Date&""" onclick=""Calendar_D(this)"" style='width:100px'>"

arrEdit_Form(2,0) = "구분"
arrEdit_Form(2,1) = MT_State

arrEdit_Form(3,0) = "거래처"

dim CNT3
arrInputSelectG	= split(replace(BasicDataMaterialTransactionCompany,"slt>",""),";")
arrEdit_Form(3,1) = "<select name='MT_Company'>"
for CNT3 = 0 to ubound(arrInputSelectG)
	arrInputSelect = split(arrInputSelectG(CNT3),":")
	if arrInputSelect(0) = "-1" then
		arrInputSelect(0) = ""
	elseif isnull(arrInputSelect(0)) then
		arrInputSelect(0) = ""
	end if
	arrEdit_Form(3,1) = arrEdit_Form(3,1) & "<option value='" & arrInputSelect(0) & "'"
	if cstr(MT_Company) = cstr(arrInputSelect(0)) then
		arrEdit_Form(3,1) = arrEdit_Form(3,1) & " selected"
	end if
	arrEdit_Form(3,1) = arrEdit_Form(3,1) & ">"&arrInputSelect(1)&"</option>"
next
arrEdit_Form(3,1) = arrEdit_Form(3,1) & "</select>"

arrEdit_Form(4,0) = "* 비고"
arrEdit_Form(4,1) = "<input type='text' name='MT_Remark' value="""&MT_Remark&""" style='width:300px'>"

Title			= "입출고상세정보"
URL_Action		= "mt_edit_action.asp"
URL_Prev		= "mt_list.asp"
URL_Next		= "mt_edit_form.asp"
URL_List		= "mt_list.asp"
Form_Type		= ""
Column_Width	= 180
Value_Width		= 400
%>
<script language="javascript">
function Form_Check(form)
{
	var strError = '';
	if(!form.MT_Date.value)
	{
		strError += "*날짜을 입력해주세요.\n"
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
<iframe width=1200px height=1024px src="detail/mtd_list.asp?s_Material_Transaction_MT_Code=<%=MT_Code%>&s_IpgoOrChulgo=<%if MT_State = "입고" then%>Ipgo<%else%>Chulgo<%end if%>" frameborder=0 style="border-top:1px solid #cccccc"></iframe>
</center>
<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
