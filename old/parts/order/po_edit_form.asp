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

dim PO_Code
dim PO_Date
dim PO_Due_Date
dim PO_State
dim Partner_P_Name

PO_Code = Request("PO_Code")

Set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbParts_Order where PO_Code='"&PO_Code&"'"
RS1.Open SQL,sys_DBCon
PO_Date			= RS1("PO_Date")
PO_Due_Date		= RS1("PO_Due_Date")
PO_State		= RS1("PO_State")
Partner_P_Name	= RS1("Partner_P_Name")

RS1.Close
Set RS1 = Nothing

strEdit_Header = "<input type='hidden' name='PO_Code' value='"&PO_Code&"'>" &vbcrlf

arrEdit_Form(0,0) = "발주번호"
arrEdit_Form(0,1) = PO_Code

arrEdit_Form(1,0) = "거래처"
arrEdit_Form(1,1) = Partner_P_Name

arrEdit_Form(2,0) = "* 발주일"
arrEdit_Form(2,1) = "<input type='text' name='PO_Date' value="""&PO_Date&""" readonly onclick=""Calendar_D(this)"" style='width:100px'>"

arrEdit_Form(3,0) = "* 납기일"
arrEdit_Form(3,1) = "<input type='text' name='PO_Due_Date' value="""&PO_Due_Date&""" readonly onclick=""Calendar_D(this)"" style='width:100px'>"

dim strEdit_PO_State
dim arrBasicDataPartsOrderState
dim arrTemp
dim CNT1
strEdit_PO_State = "<select name='PO_State' style='width:100px'>"
arrBasicDataPartsOrderState = split(replace(BasicDataPartsOrderState,"slt>",""),";")
for CNT1 = 0 to ubound(arrBasicDataPartsOrderState)
	arrTemp = split(arrBasicDataPartsOrderState(CNT1),":")
	
	if arrTemp(0) = PO_State then
		strEdit_PO_State = strEdit_PO_State & "<option value='"&arrTemp(0)&"' selected>"&arrTemp(1)&"</option>"
	else
		strEdit_PO_State = strEdit_PO_State & "<option value='"&arrTemp(0)&"'>"&arrTemp(1)&"</option>"
	end if
next
strEdit_PO_State = strEdit_PO_State & "</select>"

arrEdit_Form(4,0) = "진행상태"
arrEdit_Form(4,1) = strEdit_PO_State

Title			= "발주상세정보"
URL_Action		= "PO_edit_action.asp"
URL_Prev		= "PO_list.asp"
URL_Next		= "PO_edit_form.asp"
URL_List		= "PO_list.asp"
Form_Type		= ""
Column_Width	= 180
Value_Width		= 400
%>
<script language="javascript">
function Form_Check(form)
{
	var strError = '';
	if(!form.PO_Date.value)
	{
		strError += "*발주일을 입력해주세요.\n"
	}
	if(!form.PO_Due_Date.value)
	{
		strError += "*납기일을 입력해주세요.\n"
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
<iframe width=1220px height=1024px src="detail/pod_list.asp?s_edit_mode_yn=TRUE&S_PageSize=50&s_Parts_Order_PO_Code=<%=PO_Code%>&s_Partner_P_Name=<%=server.urlencode(Partner_P_Name)%>&s_PO_Due_Date=<%=PO_Due_Date%>" frameborder=0 style="border-top:1px solid #cccccc"></iframe>
</center>
<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
