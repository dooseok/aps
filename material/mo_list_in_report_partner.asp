
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
strSelectName		= strSelectName & "거래처,공급가액,세액,매입금액,결제방법"

strWidth			= ""
strWidth			= strWidth		& "100,100,100,100,100"
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


URL_This			= "/material/mo_list_in_report_partner.asp"
URL_View			= "/material/mo_list_in_report_partner.asp"
URL_Action			= "/material/mo_list_in_report_partner.asp"
URL_Reg				= "/material/mo_list_in_report_partner.asp"

strTable			= "tbMaterial_Order"
strPK				= "Partner_P_Name"
strSelect			= ""



strSelect			= strSelect		& "Partner_P_Name,sumPrice_SRC = SUM(MO_Price * MO_Qty_In),sumPrice_VAT = SUM(MO_Price * MO_Qty_In) * 0.1,sumPrice = SUM(MO_Price * MO_Qty_In) * 1.1,Pay_Method = (select top 1 P_Pay_Method from tbPartner where P_Name = Partner_P_Name)"

strGroupBy			= "Partner_P_Name"
if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,"
else
	strEdit	= ",,,,,,,,"
end if
strPopup			= ",,,,,,,,,,,"
strDown				= ",,,,,,,,,,,"
strAlign			= ""
strAlign			= strAlign		& "Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
strWhere = "Partner_P_Name <> ''''"
if Len(Request("s_MO_Qty_In_Date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MO_Qty_In_Date between ''"&Left(Request("s_MO_Qty_In_Date"),10)&"'' and ''"&Right(Request("s_MO_Qty_In_Date"),10)&"''"
end If

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "Partner_P_Name"
	S_Order_By_2 	= "asc"
end if

strID				= "Partner_P_Name"
strID_Pos			= "0"
'----------------------------------------------------------------------------------

arrSelectName		= split(strSelectName,",")

if S_Order_By_3 = "" then
	strOrderBy			= S_Order_By_1&" "&S_Order_By_2
else
	strOrderBy			= S_Order_By_1&" "&S_Order_By_2&", "&S_Order_By_3&" "&S_Order_By_4
end if


dim strName
dim strColumn
dim strType

'4/9
'----------------------------------------------------------------------------------
strColumn		= "s_MO_Qty_In_Date"
strName			= "기간"
strType			= "dt2"

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

arrRecordSet		= getRecordSet(URL_This, S_PageNo, S_PageSize, strTable, strPK, strSelect, strWhere, strOrderBy, strGroupBy)
TotalRecordCount	= arrRecordSet(0,ubound(arrRecordSet,2))
%>
<img src="/img/blank.gif" width=1px height=20px><br>
<%
if Reg_Form_YN = "Y" then
'6/9
'----------------------------------------------------------------------------------
	strReg	= ",,,,,,"
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
	var strError_1	= "Material_M_P_No,MT_Qty_Out"
	var strError_2	= "파트넘버,수량"
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

var bFrameChange = "true";
function FrameSizeChange()
{
	if(bFrameChange == "true")
	{
		bFrameChange = "false";
		parent.frmMoFrame.cols='1230px,*';
	}
	else
	{
		bFrameChange = "true";
		parent.frmMoFrame.cols='100px,*';
	}
}
</script>

<table width=100% cellpadding=0 cellspacing=0 border=0>
<tr>
<%
if Request("s_callby") = "mo_frame" then
%>
			<td width=70px>
				<table width=100% cellpadding=0 cellspacing=0 border=0>
				<tr>
					<td width=5px></td>
					<td width=55px><%=Make_S_BTN("<>","javascript:FrameSizeChange();","")%></td>
					<td width=5px></td>
				</tr>
				</table>
			</td>
<%
else
%>
			<td width=7px>&nbsp;</td>
<%
end if
%>	
	<td align=center>
		<table width=100% cellpadding=0 cellspacing=0 border=0>
		<tr>
			<td align=center><%=Make_BTN("EXCEL보기","List2Excel()","")%></td>
		</tr>
		<form name="frmList2Excel" action="/function/inc_List2Excel.asp" method="post" target="_blank">
		<input type="hidden"	name="strSelectName"	value="<%=strSelectName%>">
		<input type="hidden"	name="strSelect"		value="<%=strSelect%>">
		<input type="hidden"	name="strTable"			value="<%=strTable%>">
		<input type="hidden"	name="strWhere"			value="<%=strWhere%>">
		<input type="hidden"	name="strOrderBy"		value="<%=strOrderBy%>">
		<input type="hidden"	name="strGroupBy"		value="<%=strGroupBy%>">
		<input type="hidden"	name="strHaving"		value="SUM(MO_Price * MO_Qty_In) > 0">
		<input type="hidden"	name="strFileName"		value="<%=URL_This%>">
		</form>
		</table>
	</td>
</tr>
</table>
<script language="javascript">
<%
if Len(Request("s_MT_Date")) = "22" Then
%>
	function MovePrev()
	{
		frmSearch_Bar.s_MT_Date[0].value = "<%=dateadd("d",-1,left(request("s_MT_Date"),10))%>";
		frmSearch_Bar.s_MT_Date[1].value = "<%=dateadd("d",-1,left(request("s_MT_Date"),10))%>";
		frmSearch_Bar.submit()
	}
	function MoveNext()
	{
		frmSearch_Bar.s_MT_Date[0].value = "<%=dateadd("d",1,left(request("s_MT_Date"),10))%>";
		frmSearch_Bar.s_MT_Date[1].value = "<%=dateadd("d",1,left(request("s_MT_Date"),10))%>";
		frmSearch_Bar.submit()
	}
<%
end if
%>
	function MoveToday()
	{
		frmSearch_Bar.s_MT_Date[0].value = "<%=date()%>";
		frmSearch_Bar.s_MT_Date[1].value = "<%=date()%>";
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