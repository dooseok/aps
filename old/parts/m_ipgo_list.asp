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
strSelectName		= strSelectName & "번호,발주번호,거래처,파트넘버,스펙,납기일,주문량,입고일,입고량,잔량,비고"
strWidth			= ""
strWidth			= strWidth		& "70,70,150,150,200,80,70,80,70,70,100"
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
URL_This			= "/material/m_ipgo_list.asp"
URL_View			= "/material/m_ipgo_edit_form.asp"
URL_Action			= "/material/m_ipgo_list_action.asp"
URL_Reg				= "/material/m_ipgo_reg_action.asp"

strTable			= "vwMaterial_Order_Detail_Ipgo"
strPK				= "MOD_Code"
strSelect			= ""
strSelect			= strSelect		& "MOD_Code,Material_Order_MO_Code,Partner_P_Name,Material_M_P_No,Material_M_Spec,MOD_Due_Date,MOD_Qty,MOD_In_Date,MOD_In_Qty,MOD_Diff,MOD_Remark"

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,"
else
	strEdit = ",,,,,,,dt1,mny,,"
end if
strPopup			= ",,,,,,,,,,"
strDown				= ",,,,,,,,,,"
strAlign			= "Center,Center,Center,Center,Center,Center,Right,Center,Right,Right,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if Request("s_Material_Order_MO_Code") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "Material_Order_MO_Code = ''"&Request("s_Material_Order_MO_Code")&"''"
end If
if Request("s_Partner_P_Name") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "Partner_P_Name like ''%"&Request("s_Partner_P_Name")&"%''"
end If
if Request("s_Material_M_P_No") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "Material_M_P_No like ''%"&Request("s_Material_M_P_No")&"%''"
end If
if Request("s_Material_M_Spec") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "Material_M_Spec like ''%"&Request("s_Material_M_Spec")&"%''"
end If
if Len(Request("s_MOD_Due_Date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MOD_Due_Date between ''"&Left(Request("s_MOD_Due_Date"),10)&"'' and ''"&Right(Request("s_MOD_Due_Date"),10)&"''"
end if
if Len(Request("s_MOD_In_Date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MOD_In_Date between ''"&Left(Request("s_MOD_In_Date"),10)&"'' and ''"&Right(Request("s_MOD_In_Date"),10)&"''"
end if
if Request("s_MOD_Diff") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MOD_Diff <> 0"
end if

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "MOD_Code"
	S_Order_By_2 	= "desc"
end if

strID				= "MOD_Code"
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
call Material_Guide()
'----------------------------------------------------------------------------------
strColumn		= "s_Material_Order_MO_Code,s_Partner_P_Name,s_Material_M_P_No,s_Material_M_Spec|/|s_MOD_Due_Date,s_MOD_In_Date,s_MOD_Diff,s_edit_mode_yn"
strName			= "발주번호,거래처,파트넘버,스펙|/|납기일,입고일,미입고,수정모드"
strType			= "txt,txt,mtr,txt|/|dt2,dt2,chk,chk"
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
Reg_Form_YN = "Y"
call inc_tool_bar(Reg_Form_YN)
'----------------------------------------------------------------------------------

arrRecordSet		= getRecordSet(URL_This, S_PageNo, S_PageSize, strTable, strPK, strSelect, strWhere, strOrderBy, strGroupBy)

TotalRecordCount	= arrRecordSet(0,ubound(arrRecordSet,2))
%>
<img src="/img/blank.gif" width=1px height=20px><br>
<%
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

if Reg_Form_YN = "Y" then
'6/9
'----------------------------------------------------------------------------------
	strReg	= ",,"&strEdit_Partner&",mtr,,,,dt1"&date()&",txt,,"
'----------------------------------------------------------------------------------
	'call inc_Common_List_Reg_Form(URL_Reg, Colspan, strRequestQueryString, strSelect, arrRecordSet, strWidth, strReg, strAlign, strWidth_Total, 1)
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
	var strError = List_Reg_Multi_Validater('MOD_In_Qty','입고량','num','Material_M_P_No');
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
	var strError = List_Validater('MOD_In_Qty','입고량','num');
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