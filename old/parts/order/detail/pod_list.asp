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

if request("S_PageSize") <> "" then
	S_PageSize = request("S_PageSize")
elseif Request.cookies("ETC")("S_PageSize") <> "" then
	S_PageSize = Request.cookies("ETC")("S_PageSize")
else
	S_PageSize = 50
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
dim strOrderBy
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
strSelectName		= strSelectName & "번호,파트넘버,공정,스펙,설명,재고량,주문량,단가,주문가,납기일,비고,입고일,입고량,삭제"
strWidth			= ""
strWidth			= strWidth		& "50,120,70,170,170,60,60,60,70,70,90,70,70,50"
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
URL_This			= "/Parts/order/detail/POD_list.asp"
URL_View			= "/Parts/order/detail/POD_delete_action.asp"
URL_Action			= "/Parts/order/detail/POD_list_action.asp"
URL_Reg				= "/Parts/order/detail/POD_reg_action.asp"

strTable			= "vwParts_Order_Detail"
strPK				= "POD_Code"
strSelect			= ""
strSelect			= strSelect		& "POD_Code,Parts_P_P_No,Parts_M_Process,Parts_M_Spec,Parts_M_Additional_Info,Parts_M_Qty,POD_Qty,POD_Price,POD_Qty_Price,POD_Due_Date,POD_Remark,POD_In_Date,POD_In_Qty"

call Parts_Guide()

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,,"
else
	strEdit	= ",mtr,"&BasicDataPartsProcess&",,,,mny,mny,,dt1,txt,dt1,mny,"
end if
strPopup			= ",,,,,,,,,,,,,"
strDown				= ",,,,,,,,,,,,,"
strAlign			= "Center,Center,Center,Center,Center,Right,Right,Right,Right,Center,Center,Center,Right,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------

If Trim(strWhere) <> "" Then
	strWhere = strWhere & " and "
End If
strWhere = strWhere & "Parts_Order_PO_Code = ''"&Request("s_Parts_Order_PO_Code")&"''"

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "POD_code"
	S_Order_By_2 	= "asc"
end if

strID				= "POD_Code"
strID_Pos			= "0"
'----------------------------------------------------------------------------------

arrSelectName		= split(strSelectName,",")

if S_Order_By_3 = "" then
	strOrderBy			= S_Order_By_1&" "&S_Order_By_2
else
	strOrderBy			= S_Order_By_1&" "&S_Order_By_2&", "&S_Order_By_3&" "&S_Order_By_4
end if

strGroupBy			= ""

dim strName
dim strColumn
dim strType

'4/9
'----------------------------------------------------------------------------------
strColumn		= "s_edit_mode_yn"
strName			= "수정모드"
strType			= "chk"
'----------------------------------------------------------------------------------
%>
<form name="frmSearch_Bar" action="POD_list.asp" method="post">
<input type="hidden" name="s_Parts_Order_PO_Code" value="<%=Request("s_Parts_Order_PO_Code")%>">
<input type="hidden" name="s_Partner_P_Name" value="<%=Request("s_Partner_P_Name")%>">
수정모드 <input type="checkbox" name="s_edit_mode_yn" value="checked"<%if Request("s_edit_mode_yn") <> "" then%> checked<%end if%> onclick="javascript:Search_Bar_Submit_Check()">
</form>
<script language="javascript">
function Search_Bar_Submit_Check()
{
	Show_Progress();
	frmSearch_Bar.submit();
}
</script>
<%
'call Make_Search_Bar(strColumn, strName, strType, URL_This, strRequestQueryString)

Colspan	= ubound(arrSelectName) + 1
if left(strSelectName,2) = "체크" then
	Colspan	= Colspan + 1
end if
if right(strSelectName,2) = "작업" then
	Colspan	= Colspan + 1
end if

'5/9
'----------------------------------------------------------------------------------
Reg_Form_YN = "Y"
call inc_tool_bar(Reg_Form_YN)
'----------------------------------------------------------------------------------

arrRecordSet		= getRecordSet(URL_This, S_PageNo, S_PageSize, strTable, strPK, strSelect, strWhere, strOrderBy, strGroupBy)

TotalRecordCount	= arrRecordSet(0,ubound(arrRecordSet,2))
%>
<img src="/img/blank.gif" width=1px height=20px><br>
<%
if Reg_Form_YN = "Y" then
'6/9

'----------------------------------------------------------------------------------
	strReg	= ",mtr,,,,,txt,txt,,dt1"&Request("s_PO_Due_Date")&",txt,dt1"&date()&",txt,"
	'strReg	= ",mtr,,,,,txt,txt,,dt1"&Request("s_PO_Due_Date")&",txt,,,"
'----------------------------------------------------------------------------------
	call inc_Common_List_Reg_Form(URL_Reg, Colspan, strRequestQueryString, strSelect, arrRecordSet, strWidth, strReg, strAlign, strWidth_Total, 10)
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
<%
dim SQL
SQL = "select sum(convert(float,replace(POD_Qty_Price,'','',''''))) from "&strTable&" "
if trim(strWhere) <> "" then
	SQL = SQL & " where " & strWhere
end if
call Common_Display_Summary("총 발주금액 : |Sum_String|원", SQL, "N", strWidth_Total)

SQL = "select sum(convert(float,replace(POD_In_Qty_Price,'','',''''))) from "&strTable&" "
if trim(strWhere) <> "" then
	SQL = SQL & " where " & strWhere
end if

call Common_Display_Summary("총 입고금액 : |Sum_String|원", SQL, "N", strWidth_Total)
%>
<img src="/img/blank.gif" width=1px height=50px><br>
</div>
</div>

<script>
function List_Reg_Multi()
{
<%
'7/9
'----------------------------------------------------------------------------------
%>
	var strError = List_Reg_Multi_Validater('POD_Qty','주문량','num','Parts_P_P_No');
<%
'----------------------------------------------------------------------------------
%>
	if (strError=='CANCEL')
	{
		alert('값을 입력해주십시오.');
		return false;
	}	
	else if(!strError)
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
	//var strError = false;//List_Validater('','','');
<%
'----------------------------------------------------------------------------------
%>
	//if(!strError)
	//{
		Show_Progress();
		frmCommonList.submit();
	//}
	//else
	//{
		//alert(strError);
		//return false;
	//}
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

function Popup_POD_Print()
{
	if(confirm("확인을 클릭하신 후\n인쇄창이 뜰 때까지\n기다려주십시오."))
		window.open("POD_print.asp?PO_Code=<%=Request("s_Parts_Order_PO_Code")%>&Partner_P_Name=<%=Server.URLEncode(Request("s_Partner_P_Name"))%>","PartsOrderSheet","height="+screen.height+",width="+screen.width+",status=yes,toolbar=yes,location=yes,directories=yes,location=yes,menubar=yes,resizable=yes,scrollbars=yes,titlebar=yes");
}

function pop_Templete_Load()
{
	window.open("/Parts/templete/t_frame.asp?s_Opener_Type=Order&s_Opener_SubType=&s_Opener_Code=<%=Request("s_Parts_Order_PO_Code")%>","PartsTransactionSheet","height=800,width=1100,status=yes,toolbar=yes,location=yes,directories=yes,location=yes,menubar=yes,resizable=yes,scrollbars=yes,titlebar=yes");
}

function pop_Templete_Save()
{
	window.open("/Parts/templete/t_reg_form.asp?s_Opener_Type=Order&s_Opener_SubType=<%=Request("s_IpgoOrChulgo")%>&s_Opener_Code=<%=Request("s_Parts_Order_PO_Code")%>","PartsOrderSheet","height=800,width=1100,status=yes,toolbar=yes,location=yes,directories=yes,location=yes,menubar=yes,resizable=yes,scrollbars=yes,titlebar=yes");
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
			<td width=77px><table width=77 height=22 cellpadding=0 cellspacing=0 background='/img/btn_050204.gif' style='cursor:hand;' onclick="javascript:Popup_POD_Print()">
			<tr>
				<td height=22 align=center style='font-family:돋움;font-size:9pt;color:#1B3D8A'><img src='/img/blank.gif' width=1 height=1><br>발주서출력</td>
			</tr>
			</table>
			</td>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("EXCEL보기","List2Excel()","")%></td>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("항목저장","pop_Templete_Save()","")%></td>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("항목불러오기","pop_Templete_Load()","")%></td>
			<td width=5px></td>
<%
if Request("s_edit_mode_yn") <> "" then
%>
			<td width=77px><%=Make_BTN("입고일복사","Row_Copy_In_Date()","")%></td>
			<td width=5px></td>
<%
end if
%>
		</tr>
		<iframe name="ifrmXLSDown" src="about:blank" frameborder=0 width=0px height=0px></iframe><form name="frmList2Excel" action="/function/inc_List2Excel.asp" method="post" target="ifrmXLSDown">
		<input type="hidden"	name="strSelectName"	value="<%=strSelectName%>">
		<input type="hidden"	name="strSelect"		value="<%=strSelect%>">
		<input type="hidden"	name="strTable"			value="<%=strTable%>">
		<input type="hidden"	name="strWhere"			value="<%=strWhere%>">
		<input type="hidden"	name="strOrderBy"		value="<%=strOrderBy%>">
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