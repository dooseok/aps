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
strSelectName		= strSelectName & "체크,번호,파트넘버,구분,스펙,거래처,단가,확정여부,발주일,발주자,납기일,총미입고,발주량,입고량,입고일,입고자,비고"

strWidth			= ""
strWidth			= strWidth		& "40,40,100,120,200,100,60,70,70,60,70,70,60,60,80,70,200"
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
call Material_Guide()
call Partner_Guide()
call Material_Price_Log_Popup_List()
call Material_Qty_Log_Popup_List()

URL_This			= "/material/mo_list_in.asp"
URL_View			= "/material/mo_edit_form.asp"
URL_Action			= "/material/mo_list_in_action.asp"
URL_Reg				= "/material/mo_list_in_reg_action.asp"

strTable			= "vwMO_List"
strPK				= "mo_code"
strSelect			= ""
strSelect			= strSelect		& "MO_Code,Material_M_P_No,Material_M_Desc,Material_M_Spec,Partner_P_Name,MO_Price,MO_Price_Temp_YN,MO_Order_Date,MO_Reg_ID,MO_Due_Date,Material_M_Qty_Include_coming,MO_Qty,MO_Qty_In,MO_Qty_In_Date,MO_Qty_In_ID,MO_Qty_In_Desc"

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,,,,,,,,,,"
else
	strEdit	= ",,,,,,,,,,,,num,dt1,,txt"
end if
strPopup			= ",s_Material_M_P_No3,,,,s_Material_M_P_No2,,,,,,,,,,,,,,,,"
strDown				= ",,,,,,,,,,,,,,,,,,,,,"
strAlign			= ""
strAlign			= strAlign		& "Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if Request("s_Total") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "(Material_M_P_No like ''%"&Request("s_Total")&"%'' or Partner_P_Name like ''%"&Request("s_Total")&"%'' or Material_M_Desc like ''%"&Request("s_Total")&"%'' or Material_M_Spec like ''%"&Request("s_Total")&"%'')"
end If

if Request("s_M_P_No") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "Material_M_P_No like ''%"&Request("s_M_P_No")&"%''"
end If

if Request("s_Partner_P_Name") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "Partner_P_Name like ''%"&Request("s_Partner_P_Name")&"%''"
end If

if Request("s_MO_Reg_ID") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MO_Reg_ID = ''"&Request("s_MO_Reg_ID")&"''"
end If

if Request("s_MO_Qty_In_ID") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MO_Qty_In_ID = ''"&Request("s_MO_Qty_In_ID")&"''"
end If

if Len(Request("s_MO_Due_Date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MO_Due_Date between ''"&Left(Request("s_MO_Due_Date"),10)&"'' and ''"&Right(Request("s_MO_Due_Date"),10)&"''"
end If

if Len(Request("s_MO_Qty_In_Date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MO_Qty_In_Date between ''"&Left(Request("s_MO_Qty_In_Date"),10)&"'' and ''"&Right(Request("s_MO_Qty_In_Date"),10)&"''"
end If

if Request("s_MO_Ipgo_YN") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MO_Qty_In = 0"
end If

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "mo_code"
	S_Order_By_2 	= "desc"
end if

strID				= "mo_code"
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
if instr("-구매-자재-기획-총무-경영진-","-"&Request.Cookies("Admin")("M_Part")&"-") > 0  or instr("-shindk-kimdh-leehs-","-"&gM_ID&"-") > 0 then
	strColumn		= "s_Total,s_M_P_No,s_Partner_P_Name,s_MO_Ipgo_YN|/|s_MO_Reg_ID,s_MO_Due_Date,s_MO_Qty_In_ID,s_MO_Qty_In_Date,s_edit_mode_yn"
	strName			= "통합검색,파트넘버,거래처,미입고|/|발주자,발주기간,입고자,입고기간,수정모드"
	strType			= "txt,txt,ptn,slt>:전체;미입고:미입고|/|mnm-구매-자재-기획-,dt2,mnm-구매-자재-기획-,dt2,chk"
else
	strColumn		= "s_Total,s_M_P_No,s_Partner_P_Name,s_MO_Ipgo_YN|/|s_MO_Reg_ID,s_MO_Due_Date,s_MO_Qty_In_ID,s_MO_Qty_In_Date"
	strName			= "통합검색,파트넘버,거래처,미입고|/|발주자,발주기간,입고자,입고기간"
	strType			= "txt,txt,ptn,slt>:전체;미입고:미입고|/|mnm-구매-자재-기획-,dt2,mnm-구매-자재-기획-,dt2"
end if
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
if instr("-자재-기획-총무-경영진-","-"&Request.Cookies("Admin")("M_Part")&"-") > 0  or instr("-shindk-kimdh-leehs-","-"&gM_ID&"-") > 0 then
	Reg_Form_YN = "Y"
else
	Reg_Form_YN = "N"
end if
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
	strReg	= ",mtr,,,,,,,,,,num,num,dt1,,txt"
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
<%
dim SQL
SQL = "select sum(MO_Price * MO_Qty_In) from "&strTable&" "


	if trim(strWhere) <> "" then
		SQL = SQL & " where " & strWhere
	end if

	call Common_Display_Summary("금액 총계 : |Sum_String|", SQL, "Y", strWidth_Total)
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
	var strError_1	= "Material_M_P_No,MO_Qty"
	var strError_2	= "파트넘버,발주량"
	var strError_3	= "txt,num";
	var strError	= List_Reg_Validater(strError_1,strError_2,strError_3);
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
	var strError_1	= "";
	var strError_2	= "";
	var strError_3	= "";
	var strError	= List_Validater(strError_1,strError_2,strError_3);
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
		<form name=frmDateToCheckedItem>
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
			<td width=30px></td>
<%
if Len(Request("s_MO_Qty_In_Date")) = "22" Then
%>
			<td width=40px><%=Make_s_BTN("이전날","MovePrev()","")%></td>
			<td width=5px></td>
<%
else
%>
			<td width=40px>&nbsp;</td>
			<td width=5px></td>
<%
end if
%>
			<td width=40px><%=Make_s_BTN("오늘","MoveToday()","")%></td>
			<td width=5px></td>
<%
if Len(Request("s_MO_Qty_In_Date")) = "22" Then
%>
			<td width=40px><%=Make_s_BTN("다음날","MoveNext()","")%></td>
			<td width=5px></td>
<%
else
%>
			<td width=40px>&nbsp;</td>
			<td width=5px></td>
<%
end if
%>
			<td width=30px></td>
			<td width=180px>입고일 일괄입력 <input type=text name="strMultiDate" maxlength=10 size=10 onclick="Calendar_D(this)"></td>
			<td width=40px><%=Make_s_BTN("입력","DateToCheckedItem()","")%></td>
			<td width=5px></td>
		</tr>
		</form>
		<form name="frmList2Excel" action="/function/inc_List2Excel.asp" method="post" target="_blank">
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
<br>
<script language="javascript">
function DateToCheckedItem()
{
	var strChecked_Value = "";

	if(frmCommonList.strID.length)
	{
		for (cnt1 = 0; cnt1 < frmCommonList.strID.length; cnt1++)
		{
			if(frmCommonList.strID[cnt1].checked == true && frmCommonList.MO_Qty_In_Date[cnt1].value == '')
			{
				frmCommonList.MO_Qty_In_Date[cnt1].value = frmDateToCheckedItem.strMultiDate.value;
			}
		}
	}
	else
	{
		if(frmCommonList.strID.checked == true && frmCommonList.MO_Qty_In_Date.value == '')
		{
			frmCommonList.MO_Qty_In_Date.value = frmDateToCheckedItem.strMultiDate.value;
		}
	}
	return strChecked_Value;
}

<%
if Len(Request("s_MO_Qty_In_Date")) = "22" Then
%>
	function MovePrev()
	{
		frmSearch_Bar.s_MO_Qty_In_Date[0].value = "<%=dateadd("d",-1,left(request("s_MO_Qty_In_Date"),10))%>";
		frmSearch_Bar.s_MO_Qty_In_Date[1].value = "<%=dateadd("d",-1,left(request("s_MO_Qty_In_Date"),10))%>";
		frmSearch_Bar.submit()
	}
	function MoveNext()
	{
		frmSearch_Bar.s_MO_Qty_In_Date[0].value = "<%=dateadd("d",1,left(request("s_MO_Qty_In_Date"),10))%>";
		frmSearch_Bar.s_MO_Qty_In_Date[1].value = "<%=dateadd("d",1,left(request("s_MO_Qty_In_Date"),10))%>";
		frmSearch_Bar.submit()
	}
<%
end if
%>
	function MoveToday()
	{
		frmSearch_Bar.s_MO_Qty_In_Date[0].value = "<%=date()%>";
		frmSearch_Bar.s_MO_Qty_In_Date[1].value = "<%=date()%>";
		frmSearch_Bar.submit()
	}
</script>
	

<%
'----------------------------------------------------------------------------------
end sub
%>
<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->