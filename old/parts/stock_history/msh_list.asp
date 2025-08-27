<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim CNT1
dim CNT2

dim URL_This
dim URL_View
dim URL_Action
dim URL_Reg

dim S_PageNo
dim S_PageSize

dim MTD_Code
dim M_Code
dim M_P_No

MTD_Code	= Request("MTD_Code")
M_Code		= Request("M_Code")
M_P_No		= Request("s_Material_M_P_No")

set RS1 = Server.CreateObject("ADODB.RecordSet")
if MTD_Code <> "" then
	SQL = "select Material_M_P_No from tbMaterial_Transaction_Detail where MTD_Code = "&MTD_Code
	RS1.Open SQL,sys_DBCon
	M_P_No = RS1("Material_M_P_No")
	RS1.Close
end if

if M_Code <> "" then
	SQL = "select M_P_No from tbMaterial where M_Code = "&M_Code
	RS1.Open SQL,sys_DBCon
	M_P_No = RS1("M_P_No")
	RS1.Close
end if
set RS1 = nothing

if Request("s_Material_M_P_No") = "" then
	response.redirect "msh_list.asp?s_Material_M_P_No="&M_P_No
end if

if request("S_PageSize") <> "" then
	S_PageSize = request("S_PageSize")
elseif Request.cookies("ETC")("S_PageSize") <> "" then
	S_PageSize = Request.cookies("ETC")("S_PageSize")
else
	S_PageSize = 20
end if
S_PageNo		= request("S_PageNo")
if S_PageSize <> Request.cookies("ETC")("S_PageSize") then
	S_PageNo = 1
end if
if trim(S_PageNo) = "" then
	S_PageNo = 1
end if

Response.cookies("ETC")("S_PageSize")	= S_PageSize
Response.cookies("ETC").Path			= "/"

dim strRequestQueryString
strRequestQueryString = getRequestQueryString()

dim strSelectName
dim arrSelectName

dim strWidth
dim strAlign

dim arrWidth
dim strWidth_Total

dim strID
dim strID_Pos

dim strTable
dim strPK
dim strSelect
dim strWhere
dim strTransactionBy
dim strGroupBy

dim strReg
dim strEdit
Dim strPopup
Dim strDown

dim arrRecordSet
dim TotalRecordCount
dim Colspan

dim Reg_Form_YN

dim S_Order_By_1
dim S_Order_By_2
dim S_Order_By_3
dim S_Order_By_4

S_Order_By_1 = Request("S_Order_By_1")
S_Order_By_2 = Request("S_Order_By_2")
S_Order_By_3 = Request("S_Order_By_3")
S_Order_By_4 = Request("S_Order_By_4")

'1/9
'----------------------------------------------------------------------------------
strSelectName		= ""
strSelectName		= strSelectName & "번호,파트넘버,변동일,거래처,구분,재고량,변동량"
strWidth			= ""
strWidth			= strWidth		& "70,120,70,120,200,90,90"
'----------------------------------------------------------------------------------

arrWidth = split(strWidth,",")
for CNT1 = 0 to ubound(arrWidth)
	strWidth_Total = strWidth_Total + int(arrWidth(CNT1))
next
%>
<div style="width:<%=strWidth_Total%>px">
<%
'2/9
'----------------------------------------------------------------------------------
URL_This			= "/material/stock_history/msh_list.asp"
URL_View			= "/material/stock_history/msh_list.asp"
URL_Action			= "/material/stock_history/msh_list.asp"
URL_Reg				= "/material/stock_history/msh_list.asp"

strTable			= "tbMaterial_Stock_History"
strPK				= "MSH_Code"
strSelect			= ""
strSelect			= strSelect		& "MSH_Code,Material_M_P_No,MSH_Change_Date,MSH_Company,MSH_Change_Type,MSH_Applyed_Stock,MSH_Change_Stock"

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,"
else
	strEdit	= ",,,,,,"
end if
strPopup			= ",,,,,,"
strDown				= ",,,,,,"
strAlign			= "Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if Request("s_Material_M_P_No") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "Material_M_P_No = ''"&Request("s_Material_M_P_No")&"''"
end If

if Request("s_MSH_Company") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MSH_Company = ''"&Request("s_MSH_Company")&"''"
end If

if Len(Request("s_MSH_Change_Date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MSH_Change_Date between ''"&Left(Request("s_MSH_Change_Date"),10)&"'' and ''"&Right(Request("s_MSH_Change_Date"),10)&"''"
end if

if Request("s_MSH_Change_Type") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MSH_Change_Type = ''"&Request("s_MSH_Change_Type")&"''"
end If

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "msh_code"
	S_Order_By_2 	= "desc"
end if

strID				= "msh_Code"
strID_Pos			= "0"
'----------------------------------------------------------------------------------

arrSelectName		= split(strSelectName,",")

if S_Order_By_3 = "" then
	strTransactionBy			= S_Order_By_1&" "&S_Order_By_2
else
	strTransactionBy			= S_Order_By_1&" "&S_Order_By_2&", "&S_Order_By_3&" "&S_Order_By_4
end if

strGroupBy			= ""

dim strName
dim strColumn
dim strType

call Material_Guide()

dim RS1
dim SQL
dim strEdit_Partner

set RS1 = server.CreateObject("ADODB.RecordSet")
SQL ="select P_Name from tbPartner order by P_Sort asc, P_Name asc"
RS1.Open SQL,sys_DBCon
strEdit_Partner = ""
if RS1.Eof or RS1.Bof then
	strEdit_Partner = "slt>"
else
	strEdit_Partner = "slt>"
	do until RS1.Eof
		strEdit_Partner = strEdit_Partner & RS1("P_Name")
		strEdit_Partner = strEdit_Partner & ":"
		strEdit_Partner = strEdit_Partner & RS1("P_Name")
		strEdit_Partner = strEdit_Partner & ";"
		RS1.MoveNext
	loop
	strEdit_Partner = left(strEdit_Partner,len(strEdit_Partner)-1)
end if
RS1.Close
set RS1 = nothing

'4/9
'----------------------------------------------------------------------------------
strColumn		= "s_Material_M_P_No,s_MSH_Change_Date|/|s_MSH_Company,s_MSH_Change_Type"
strName			= "파트넘버,조회기간|/|거래처,구분"
strType			= "mtr,dt2|/|"&BasicDataMaterialTransactionCompany&replace(strEdit_Partner,"slt>",";")&","&BasicDataMaterialStockHistoryType
'----------------------------------------------------------------------------------

call Make_Search_Bar(strColumn, strName, strType, URL_This, strRequestQueryString)

Colspan	= ubound(arrSelectName) + 1
if left(strSelectName,2) = "체크" then
	Colspan	= Colspan + 1
end if
if right(strSelectName,2) = "작업" then
	Colspan	= Colspan + 1
end if

'5/9
'----------------------------------------------------------------------------------
Reg_Form_YN = "N"
call inc_tool_bar(Reg_Form_YN)
'----------------------------------------------------------------------------------

arrRecordSet		= getRecordSet(URL_This, S_PageNo, S_PageSize, strTable, strPK, strSelect, strWhere, strTransactionBy, strGroupBy)

TotalRecordCount	= arrRecordSet(0,ubound(arrRecordSet,2))
%>
<img src="/img/blank.gif" width=1px height=20px><br>
<%
if Reg_Form_YN = "Y" then
'6/9
'----------------------------------------------------------------------------------
	strReg	= ",,,,,"
'----------------------------------------------------------------------------------
	call inc_Common_List_Reg_Form(URL_Reg, Colspan, strRequestQueryString, strSelect, arrRecordSet, strWidth, strReg, strAlign, strWidth_Total, 1)
end if
%>
<img src="/img/blank.gif" width=1px height=10px><br>
<div id="idList" style="display:block;">
<%
call inc_Common_List(strID, strID_Pos, S_PageNo, URL_This, URL_View, URL_Action, arrSelectName, strSelect, arrRecordSet, TotalRecordCount, Colspan, strRequestQueryString, S_Order_By_1, S_Order_By_2, strPopup, strDown, strWidth, strEdit, strAlign, strWidth_Total)
%>
<img src="/img/blank.gif" width=1px height=5px><br>
<%
call inc_Common_Paging(URL_This, TotalRecordCount, S_PageSize, S_PageNo, strRequestQueryString)
%>
<img src="/img/blank.gif" width=1px height=50px><br>
</div>
</div>

<script>
function List_Reg()
{
<%
'7/9
'----------------------------------------------------------------------------------
%>
	var strError = List_Reg_Validater('MT_Date,MT_Company','날짜,거래처','txt,txt');
<%
'----------------------------------------------------------------------------------
%>
	if(!strError)
	{
		Show_Progress();
		frmCommonListReg.submit();
	}
	else
	{
		alert(strError);
		return false;
	}
}
</script>

<script language="javascript">
function List_Update()
{
<%
'8/9
'----------------------------------------------------------------------------------
%>
	var strError = List_Validater('MT_Date,MT_Company','날짜,거래처','txt,txt');
<%
'----------------------------------------------------------------------------------
%>
	if(!strError)
	{
		Show_Progress();
		frmCommonList.submit();
	}
	else
	{
		alert(strError);
		return false;
	}
}
</script>

<%
sub inc_tool_bar(Reg_Form_YN)
'9/9
'----------------------------------------------------------------------------------
%>
<script language="javascript">
function XLS_UP()
{
	var strChecked_Value = GetChecked_Value();

	if (strChecked_Value == "")
	{
		alert("한개 이상의 아이템을 선택해주십시오.")
	}
	else
	{
		//작업내용
		var arrChecked_Value = strChecked_Value.split(",");
		for (var cnt1=0; cnt1<arrChecked_Value.length-1; cnt1++)
		{

		}
	}
}

var RegForm_Toggle_YN = "N"
function RegForm_Toggle()
{
	if(RegForm_Toggle_YN == "N")
	{
		idRegForm.style.display = "block";
		idList.style.display = "block";

		idBtnRegForm.style.display = "none";
		idBtnList.style.display = "block";

		RegForm_Toggle_YN = "Y";
		return false;
	}
	else(RegForm_Toggle_YN == "Y")
	{
		idRegForm.style.display = "none";
		idList.style.display = "block";

		idBtnRegForm.style.display = "block";
		idBtnList.style.display = "none";

		RegForm_Toggle_YN = "N";
		return false;
	}
}

function List2Excel()
{
	frmList2Excel.submit();
}
</script>

<table width=100% cellpadding=0 cellspacing=0 border=0>
<tr>
	<td align=center>
		<table cellpadding=0 cellspacing=0 border=0>
		<tr>
<%
if Request("s_edit_mode_yn") <> "" then
%>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("수정완료","javascript:List_Update()","")%></td>
<%
end if
%>
<%
if Reg_Form_YN = "Y" then
%>
			<td width=5px></td>
			<td width=77px>
				<div id="idBtnRegForm"><%=Make_BTN("신규등록","javascript:RegForm_Toggle()","")%></div>
				<div id="idBtnList" style="display:none;"><%=Make_BTN("목록보기","javascript:RegForm_Toggle()","")%></div>
			</td>
<%
end if
%>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("EXCEL보기","List2Excel()","")%></td>
			<td width=5px></td>
		</tr>
		<iframe name="ifrmXLSDown" src="about:blank" frameborder=0 width=0px height=0px></iframe><form name="frmList2Excel" action="/function/inc_List2Excel.asp" method="post" target="ifrmXLSDown">
		<input type="hidden"	name="strSelectName"	value="<%=strSelectName%>">
		<input type="hidden"	name="strSelect"		value="<%=strSelect%>">
		<input type="hidden"	name="strTable"			value="<%=strTable%>">
		<input type="hidden"	name="strWhere"			value="<%=strWhere%>">
		<input type="hidden"	name="strTransactionBy"		value="<%=strTransactionBy%>">
		<input type="hidden"	name="strFileName"		value="<%=URL_This%>">
		</form>
		</table>
	</td>
</tr>
</table>

<%
'----------------------------------------------------------------------------------
end sub
%>
<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->