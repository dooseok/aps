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
dim URL_Del

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
strSelectName	= "문서번호,시방 PNO,적용일,개발,자재,"
strSelect			= "BU_Code,BOM_B_D_No,BU_Apply_Date,BU_RnD_Check,BU_JaJe_Check,"
strWidth			= "100,100,100,40,40,"
strPopup			= ",db_load_action.asp,,Bom_Update_Y_DECO,Bom_Update_Y_DECO,"

if instr(admin_bu_list,"-"&gM_ID&"-") > 0 then
	strSelectName	= strSelectName&"IMT,SMT,제조<br><img src='/img/blank.gif' width=1px height=5px><br>2,제조<br><img src='/img/blank.gif' width=1px height=5px><br>3,기술,IQC,PCB<br><img src='/img/blank.gif' width=1px height=5px><br>검사,C/B<br><img src='/img/blank.gif' width=1px height=5px><br>검사,슈퍼<br><img src='/img/blank.gif' width=1px height=5px><br>마켓,영업,단가,OTP,SM<br><img src='/img/blank.gif' width=1px height=5px><br>Tec,준비<br><img src='/img/blank.gif' width=1px height=5px><br>작업,"
	strSelect			= strSelect&"BU_IMT_Check,BU_SMT_Check,BU_JeJo2_Check,BU_JeJo3_Check,BU_Eng_Check,BU_IQC_Check,BU_PCBA_QC_Check,BU_CBOX_QC_Check,BU_SPMK_Check,BU_DLV_Check,BU_Price_Check,BU_OTP_Check,BU_SMTech_Check,BU_DSTech_Check"
	strWidth			= strWidth&"40,40,40,40,40,40,40,40,40,40,40,40,40,40,"
	strPopup			= strPopup&"Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO"
else
	Select Case lcase(gM_ID)
		case "jaje"
			'strSelectName	= strSelectName&"자재,"
			'strSelect			= strSelect&"BU_JaJe_Check"
			'strWidth			= strWidth&"40,"
			'strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "imt"
			strSelectName	= strSelectName&"IMT,SMT,"
			strSelect			= strSelect&"BU_IMT_Check,BU_SMT_Check"
			strWidth			= strWidth&"40,40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO,Bom_Update_Y_DECO"
		case "smt"
			strSelectName	= strSelectName&"IMT,SMT,"
			strSelect			= strSelect&"BU_IMT_Check,BU_SMT_Check"
			strWidth			= strWidth&"40,40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO,Bom_Update_Y_DECO"
		case "jejo2"
			strSelectName	= strSelectName&"제조<br><img src='/img/blank.gif' width=1px height=5px><br>2,"
			strSelect			= strSelect&"BU_JeJo2_Check"
			strWidth			= strWidth&"40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "jejo3"
			strSelectName	= strSelectName&"제조<br><img src='/img/blank.gif' width=1px height=5px><br>3,"
			strSelect			= strSelect&"BU_JeJo3_Check"
			strWidth			= strWidth&"40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "eng"
			strSelectName	= strSelectName&"기술,"
			strSelect			= strSelect&"BU_Eng_Check"
			strWidth			= strWidth&"40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "iqc"
			strSelectName	= strSelectName&"IQC,"
			strSelect			= strSelect&"BU_IQC_Check"
			strWidth			= strWidth&"40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "pcbqc"
			strSelectName	= strSelectName&"PCB<br><img src='/img/blank.gif' width=1px height=5px><br>검사,"
			strSelect			= strSelect&"BU_PCBA_QC_Check"
			strWidth			= strWidth&"40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "cboxqc"
			strSelectName	= strSelectName&"C/B<br><img src='/img/blank.gif' width=1px height=5px><br>검사,"
			strSelect			= strSelect&"BU_CBOX_QC_Check"
			strWidth			= strWidth&"40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "qa"
			strSelectName	= strSelectName&"IQC,PCB<br><img src='/img/blank.gif' width=1px height=5px><br>검사,C/B<br><img src='/img/blank.gif' width=1px height=5px><br>검사,슈퍼<br><img src='/img/blank.gif' width=1px height=5px><br>마켓,준비<br><img src='/img/blank.gif' width=1px height=5px><br>작업,"
			strSelect			= strSelect&"BU_IQC_Check,BU_PCBA_QC_Check,BU_CBOX_QC_Check,BU_SPMK_Check,BU_DSTech_Check"
			strWidth			= strWidth&"40,40,40,40,40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO,Bom_Update_Y_DECO"
		case "sales"
			strSelectName	= strSelectName&"영업,"
			strSelect			= strSelect&"BU_DLV_Check"
			strWidth			= strWidth&"40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "price"
			strSelectName	= strSelectName&"단가,"
			strSelect			= strSelect&"BU_Price_Check"
			strWidth			= strWidth&"40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "otp"
			strSelectName	= strSelectName&"OTP,"
			strSelect			= strSelect&"BU_OTP_Check"
			strWidth			= strWidth&"40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "smtech"
			strSelectName	= strSelectName&"SM<br><img src='/img/blank.gif' width=1px height=5px><br>Tec,"
			strSelect			= strSelect&"BU_SMTech_Check"
			strWidth			= strWidth&"40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO"
		case "dstech"
			strSelectName	= strSelectName&"제조<br><img src='/img/blank.gif' width=1px height=5px><br>3,준비<br><img src='/img/blank.gif' width=1px height=5px><br>작업,"
			strSelect			= strSelect&"BU_JeJo3_Check,BU_DSTech_Check,"
			strWidth			= strWidth&"40,40,"
			strPopup			= strPopup&"Bom_Update_Y_DECO,Bom_Update_Y_DECO,"
	end select
end if
if right(strSelect,1) = "," then
	strSelect = left(strSelect,len(strSelect)-1)
end if

strSelectName	= strSelectName&"작업_프레임"
strSelect		= strSelect&""
strWidth		= strWidth&"60"
strPopup		= strPopup&""

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
URL_This			= "/bom/new_BU_list_main.asp"
URL_View			= "/bom/new_BU_edit_form.asp"
URL_Action			= "/bom/new_BU_list_action.asp"
URL_Reg				= "/bom/new_BU_reg_action.asp"

strTable			= "vwBU_List_new"
strPK				= "BU_Code"

strDown				= ",,,,,,,,,,,,,,,,,,,,,,,,,"
strAlign			= "Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center"

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
else
	strEdit	= ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
end if
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if instr(admin_bu_list,"-"&gM_ID&"-") > 0 then
else
	strWhere = strWhere & "right(BU_RnD_Check,4) in (''해당없음'',''적용완료'')"
end if
If Trim(strWhere) <> "" Then
	strWhere = strWhere & " and "
End If
strWhere = strWhere & "BU_RnD_Check+BU_JaJe_Check+BU_IMT_Check+BU_SMT_Check+BU_JeJo2_Check+BU_JeJo3_Check+BU_IQC_Check+BU_PCBA_QC_Check+BU_CBOX_QC_Check+BU_SPMK_Check+BU_DLV_Check+BU_Price_Check+BU_OTP_Check+BU_Eng_Check+BU_SMTech_Check+BU_DSTech_Check like ''%확인%''"

if S_Order_By_1 & S_Order_By_2 = "" then
	'S_Order_By_1 	= "BU_Code"
	'S_Order_By_2 	= "desc"
	S_Order_By_1 	= "BU_Apply_Date"
	S_Order_By_2 	= "asc"
end if

strID				= "BU_Code"
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
call BOM_Guide
strColumn		= ""
strName			= ""
strType			= ""
'----------------------------------------------------------------------------------

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
	'strReg	= ",,,,,,,"&DefaultPath_Notice&","&DefaultPath_Notice&","&DefaultPath_Notice&",,,,,,,,"
	strReg = ",,,,,,,,,,,,,,,,,,,,,,,,,,"
'----------------------------------------------------------------------------------	
	call inc_Common_List_Reg_Form(URL_Reg, Colspan, strRequestQueryString, strSelect, arrRecordSet, strWidth, strReg, strAlign, strWidth_Total, 1)
end if
%>

<img src="/img/bu_check.jpg"><br>
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
	var strError = List_Reg_Validater('BOM_B_D_No','파트넘버','txt');
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
	var strError = List_Validater('BOM_B_D_No','파트넘버','txt');
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
</script>
<Br>
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
			<td width=150px><%=Make_L_BTN("시방리스트 보기","javascript:parent.location.href='/bom/new_bu_list.asp'","")%></td>
		</tr>
		</table>
	</td>
</tr>
</table>

<script language="javascript">
function printPeriod()
{
	window.open("new_frame_bu_period_print.asp","PartsOrderSheet","height="+screen.height+",width="+screen.width+",status=yes,toolbar=yes,location=yes,directories=yes,location=yes,menubar=yes,resizable=yes,scrollbars=yes,titlebar=yes");
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