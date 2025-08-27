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
strSelectName		= strSelectName & "üũ,��ȣ,��Ʈ�ѹ�,�ŷ�ó,�ŷ�ó��,�ܰ�Ȯ��,�ܰ�,������,������,���YN,����,����,BOM����,����,����Ŀ,�뵵,�������,�������,�̷����,����PartNO"

strWidth			= ""
strWidth			= strWidth		& "40,40,150,120,70,70,70,60,70,60,100,150,250,70,70,70,70,70,70,100"
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

URL_This			= "/material/bak_m_list.asp"
URL_View			= "/material/m_edit_form.asp"
URL_Action			= "/material/m_list_action.asp"
URL_Reg				= "/material/m_reg_action.asp"

strTable			= "vwMaterial_M_List_Bak"
strPK				= "m_code"
strSelect			= ""
strSelect			= strSelect		& "M_Code,M_P_No,Partner_P_Name,cntMulti_Partner,M_Price_Temp_YN,M_Price,M_Price_LGE,M_Price_Apply_Date,M_OSP_YN,M_Desc,M_Spec,M_Spec_Bom,M_Process,M_Maker,M_Division,M_Qty,M_Qty_Safe,M_Qty_Include_coming,M_P_No_Sub"

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,,,,,,,,,,,,,"
else
		
	if instr(admin_material_handler,"-"&gM_ID&"-") > 0 then
		strEdit	= ",,,,slt>Ȯ��:Ȯ��;���ܰ�:���ܰ�,num,num,dt1,"&BasicDataMaterialOSP&",mem,mem,mem,"&BasicDataPartsType&",mem,"&BasicDataMaterialDivision&",,,,txt,,,,"
	else
		strEdit	= ",,,,,,,,"&BasicDataMaterialOSP&",mem,mem,mem,"&BasicDataPartsType&",mem,"&BasicDataMaterialDivision&",,,,txt,,,,"
	end if
end if

	strPopup			= ",s_Material_M_P_No3,,,,s_Material_M_P_No1,,,,,,,,,,,,,,,,,,,,,"

strDown				= ",,,,,,,,,,,,,,,,,,,,,,"
strAlign			= ""
strAlign			= strAlign		& "Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if Request("s_Total") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "(M_P_No like ''%"&Request("s_Total")&"%'' or Partner_P_Name like ''%"&Request("s_Total")&"%'' or M_Desc like ''%"&Request("s_Total")&"%'' or M_Spec like ''%"&Request("s_Total")&"%'' or M_Spec_Bom like ''%"&Request("s_Total")&"%'' or M_Maker like ''%"&Request("s_Total")&"%'' or M_P_No_Sub like ''%"&Request("s_Total")&"%'')"
end If

if Request("s_M_P_No") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "M_P_No like ''%"&Request("s_M_P_No")&"%''"
end If

if Request("s_Partner_P_Name") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "Partner_P_Name like ''%"&Request("s_Partner_P_Name")&"%''"
end If

if Request("s_M_OSP_YN") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "M_OSP_YN = ''"&Request("s_M_OSP_YN")&"''"
end If

if Request("s_M_Division") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "M_Division = ''"&Request("s_M_Division")&"''"
end If

if Request("s_Multi_Partner_YN") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "cntMulti_Partner > 1"
end If

if Request("s_Price_Temp_YN") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "M_Price_Temp_YN = ''���ܰ�''"
end If

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "m_code"
	S_Order_By_2 	= "desc"
end if

strID				= "m_code"
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
'if instr("-����-����-��ȹ-�ѹ�-�濵��-","-"&Request.Cookies("Admin")("M_Part")&"-") > 0  or instr("-shindk-kimdh-leehs-","-"&gM_ID&"-") > 0 then
'	strColumn		= "s_Total,s_M_P_No,s_Partner_P_Name,s_Multi_Partner_YN|/|s_Price_Temp_YN,s_M_OSP_YN,s_M_Division,s_edit_mode_yn"
'	strName			= "���հ˻�,��Ʈ�ѹ�,�ŷ�ó,�����ŷ�ó�� ����|/|���ܰ��� ����,���YN,�뵵,�������"
'	strType			= "txt,txt,ptn,chk|/|chk,"&BasicDataMaterialOSP&","&BasicDataMaterialDivision&",chk"
'else
	strColumn		= "s_Total,s_M_P_No,s_Partner_P_Name,s_Multi_Partner_YN|/|s_Price_Temp_YN,s_M_OSP_YN,s_M_Division"
	strName			= "���հ˻�,��Ʈ�ѹ�,�ŷ�ó,�����ŷ�ó�� ����|/|���ܰ��� ����,���YN,�뵵"
	strType			= "txt,txt,ptn,chk|/|chk,"&BasicDataMaterialOSP&","&BasicDataMaterialDivision
'end if
'----------------------------------------------------------------------------------

call Make_Search_Bar(strColumn, strName, strType, URL_This, strRequestQueryString)

Colspan	= ubound(arrSelectName) + 1
if left(strSelectName,2) = "üũ" then
	Colspan	= Colspan + 1
end if
if right(strSelectName,2) = "�۾�" then
	Colspan	= Colspan + 1
end if

'5/9
'----------------------------------------------------------------------------------
'if instr("-����-����-��ȹ-�ѹ�-�濵��-","-"&Request.Cookies("Admin")("M_Part")&"-") > 0  or instr("-shindk-kimdh-leehs-","-"&gM_ID&"-") > 0 then
'	Reg_Form_YN = "Y"
'else
	Reg_Form_YN = "N"
'end if
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
	strReg	= ",txt,ptn,,slt>Ȯ��:Ȯ��;���ܰ�:���ܰ�,num,num,,"&BasicDataMaterialOSP&",mem,mem,mem,"&BasicDataPartsType&",txt,"&BasicDataMaterialDivision&",,,,txt"
'----------------------------------------------------------------------------------
	call inc_Common_List_Reg_Form(URL_Reg, Colspan, strRequestQueryString, strSelect, arrRecordSet, strWidth, strReg, strAlign, strWidth_Total, 1)
end if

if request("s_edit_mode_yn") = "" then
%>
<img src="/img/blank.gif" width=1px height=10px><br>
<%
else
%>
<img src="/img/blank.gif" width=1px height=3px><br>
<font color=darkgreen>�����ŷ�ó�� ��ϵ� �������� ������, ���� ���� ��ϵ� [��Ʈ�ѹ�-�ŷ�ó]�� ������ �����ϸ� ������Ʈ�� �˴ϴ�.<font>
<img src="/img/blank.gif" width=1px height=3px><br>
<%
end if	
%>
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
	var strError_1	= "M_P_No,Partner_P_Name,M_Price_Temp_YN,M_Price,M_Price_LGE,M_OSP_YN,M_Process,M_Division"
	var strError_2	= "��Ʈ�ѹ�,�ŷ�ó,�ܰ�Ȯ��,�ܰ�,������,��޿���,����,�뵵"
	var strError_3	= "txt,txt,txt,num,num,txt,txt,txt";
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
	var strError_1	= "M_OSP_YN";
	var strError_2	= "���";
	var strError_3	= "txt";
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
		alert("�Ѱ� �̻��� �������� �������ֽʽÿ�.")
	}
	else
	{
		//�۾�����
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
function List_BalJu()
{
	Show_Progress();
	frmCommonList.action = "m_list2mo_list.asp?dummy=<%=strRequestQueryString%>";
	frmCommonList.submit();
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
		<table cellpadding=0 cellspacing=0 border=0>
		<tr>
<%
if instr("-����-����-��ȹ-�ѹ�-�濵��-","-"&Request.Cookies("Admin")("M_Part")&"-") > 0  or instr("-shindk-kimdh-leehs-","-"&gM_ID&"-") > 0 then
	if Request("s_callby") = "mo_frame" then
%>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("����ó��","javascript:List_BalJu()","")%></td>
<%
	end if
end if
%>
<%
if Request("s_edit_mode_yn") <> "" then
%>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("�����Ϸ�","javascript:List_Update()","")%></td>
<%
end if
%>
<%
if Reg_Form_YN = "Y" then
%>
			<td width=5px></td>
			<td width=77px>
<%
	if instr(admin_material_handler,"-"&gM_ID&"-") > 0 then
%>
				<div id="idBtnRegForm"><%=Make_BTN("�űԵ��","javascript:RegForm_Toggle()","")%></div>
<%
	end if
%>
				<div id="idBtnList" style="display:none;"><%=Make_BTN("��Ϻ���","javascript:RegForm_Toggle()","")%></div>
			</td>
<%
end if
%>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("EXCEL����","List2Excel()","")%></td>
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
	<td width=70px>&nbsp;</td>
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