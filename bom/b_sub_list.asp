<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
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
elseif Request.Cookies("ETC")("S_PageSize") <> "" then
	S_PageSize = Request.Cookies("ETC")("S_PageSize")
else
	S_PageSize = 20
end if
S_PageNo		= request("S_PageNo")
if S_PageSize <> Request.Cookies("ETC")("S_PageSize") then
	S_PageNo = 1
end if
if trim(S_PageNo) = "" then
	S_PageNo = 1
end if

Response.Cookies("ETC")("S_PageSize")	= S_PageSize
Response.Cookies("ETC").Path			= "/"

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
'strSelectName		= "번호,파트넘버,IMD재고,SMD재고,MAN재고,ASM재고,IMD_Axial_PT,IMD_Radial_PT"
'strWidth			= "50,100,80,80,80,80,100,100"
strSelectName		= "번호,파트넘버,완성품재고"
strWidth			= "120,150,150"
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
URL_This			= "/bom/b_sub_list.asp"
URL_View			= "/bom/b_sub_edit_form.asp"
URL_Action			= "/bom/b_sub_list_action.asp"
URL_Reg				= "/bom/b_sub_reg_form.asp"

'strTable			= "vwBS_List"
'strPK				= "Bs_Code"
'strSelect			= "bs_code,bs_d_no,b_imd_qty,bs_smd_qty,bs_man_qty,bs_asm_qty,BS_IMD_Axial_Point,BS_IMD_Radial_Point"
strTable			= "tbBOM_Sub"
strPK				= "Bs_Code"
strSelect			= "bs_code,bs_d_no,bs_man_qty"

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,"
else
	strEdit	= ",,,,,,num,num"
end if
strPopup			= ",/bom/bs_qty_chart.asp,,,,,,,,,,,"
strDown				= ",,,,,,,,,,,,"
strAlign			= "Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if Request("s_bom_b_d_no") = "" then
	If Trim(strWhere) <> "" Then
		'strWhere = strWhere & " and "
	End If
	'strWhere = strWhere & "	bom_b_code in (select b_code from tbBOM where B_D_No like ''4781AG2149'')"
else
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "	bom_b_code in (select b_code from tbBOM where B_D_No like ''%"&Request("s_bom_b_d_no")&"%'')"
end If

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "bs_d_no"
	S_Order_By_2 	= "asc"
end if

strID				= "bs_code"
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
call BOM_Guide()

'strColumn		= "s_bom_b_d_no,s_edit_mode_yn"
'strName			= "모델파트넘버,수정모드"
'strType			= "dn1,chk"
strColumn		= "s_bom_b_d_no"
strName			= "모델파트넘버"
strType			= "dn1"
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
	strReg	= ",,,,,,,,,,,,,,"
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
	var strError = List_Reg_Validater('M_Channel,M_ID,M_Password,M_Part,M_Position,M_Name,M_Enter_Date','소속회사,아이디,암호,부서,직급,이름,입사일','txt,txt,txt,txt,txt,txt,txt,txt');
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
	var strError = List_Validater('BS_IMD_Axial_Point,BS_IMD_Radial_Point','IMD_Axial_Point,BS_IMD_Radial_Point','num,num');
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
			<!--<td width=77px>
				<table width=205px cellpadding=0 cellspacing=0 border=0>
				<tr>
					<td width=100px><%=Make_L_BTN("BOM신규등록","","b_reg_form.asp?NEW_YN=Y")%></td>
					<td width=5px><img src="/img/blank.gif" width=5px height=1px></td>
					<td width=100px><%=Make_L_BTN("BOM변경등록","","b_reg_form.asp?NEW_YN=N")%></td>
				</tr>
				</table>
			</td>-->
			
			<td width=100px><%=Make_L_BTN("BOM신규등록","","b_reg_form.asp?NEW_YN=Y")%></td>
<%
end if
%>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("EXCEL보기","List2Excel()","")%></td>
			<td width=5px></td>
		</tr>
		<iframe name="ifrmXLSDown" src="about:blank" frameborder=0 width=0px height=0px></iframe>
<form name="frmList2Excel" action="/function/inc_List2Excel.asp" method="post" target="ifrmXLSDown">
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