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
if instr(admin_lm_list,"-"&Request("Admin")("M_ID")&"-") > 0 then
	strSelectName		= strSelectName & "��ȣ,��ü����,�𵨸�,��Ʈ�ѹ�1,��Ʈ�ѹ�2,��Ʈ�ѹ�3,��Ʈ�ѹ�4,����"
else
	strSelectName		= strSelectName & "��ȣ,��ü����,�𵨸�,��Ʈ�ѹ�1,��Ʈ�ѹ�2,��Ʈ�ѹ�3,��Ʈ�ѹ�4"
end if
strWidth			= ""
strWidth			= strWidth		& "60,70,160,160,160,160,160,60"
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
URL_This			= "/lge_model/lm_list.asp"
URL_View			= "/lge_model/lm_delete_action.asp"
URL_Action			= "/lge_model/lm_list_action.asp"
URL_Reg				= "/lge_model/lm_reg_action.asp"

strTable			= "vwLM_List"
strPK				= "LM_Code"
strSelect			= ""
strSelect			= strSelect		& "LM_Code,LM_Company,LM_Name,BOM_Sub_BS_D_No_1,BOM_Sub_BS_D_No_2,BOM_Sub_BS_D_No_3,BOM_Sub_BS_D_No_4"

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,"
else
	strEdit	= ","&BasicModelCompany&",mem,mem,mem,mem,mem,mem,"
end if
strPopup			= ",,,,,,,,"
strDown				= ",,,,,,,,"
strAlign			= "Center,Center,Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------

if Request("s_LM_Company") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "LM_Company = ''"&Request("s_LM_Company")&"''"
end If
if Request("s_LM_Name") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "LM_Name like ''%"&Request("s_LM_Name")&"%''"
end If
if Request("s_Parts_P_P_No") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "(BOM_Sub_BS_D_No_1 like ''%"&Request("s_Parts_P_P_No")&"%'' or BOM_Sub_BS_D_No_2 like ''%"&Request("s_Parts_P_P_No")&"%'' or BOM_Sub_BS_D_No_3 like ''%"&Request("s_Parts_P_P_No")&"%'' or BOM_Sub_BS_D_No_4 like ''%"&Request("s_Parts_P_P_No")&"%'')"
end If

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "lm_code"
	S_Order_By_2 	= "desc"
end if

strID				= "LM_Code"
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
strColumn		= "s_LM_Company,s_LM_Name,s_Parts_P_P_No,s_edit_mode_yn"
strName			= "��ü����,�𵨸�,��Ʈ�ѹ�,�������"
strType			= BasicModelCompany&",txt,txt,chk"
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
	strReg	= ","&BasicModelCompany&",mem,mem,mem,mem,mem,mem,"
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
	var strError = List_Reg_Validater('LM_Company,LM_Name','��ü����,�𵨸�','txt,txt');
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
	var strError = List_Validater('LM_Company,LM_Name','��ü����,�𵨸�','txt,txt');
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
</script>

<table width=100% cellpadding=0 cellspacing=0 border=0>
<tr>
	<td align=center>
		<table cellpadding=0 cellspacing=0 border=0>
		<tr>
<%
if Request("s_edit_mode_yn") <> "" and instr(admin_lm_list,"-"&Request("Admin")("M_ID")&"-") > 0 then
%>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("�����Ϸ�","javascript:List_Update()","")%></td>
<%
end if
%>
<script language="javascript">
function Change_To_Other()
{
	if(confirm("�̺з� ������ ���� ���� Ÿ��� ��ȯ�ϰڽ��ϱ�?"))
	{
		location.href='lm_change_to_others.asp';
	}
}
</script>

<%
if instr(admin_lm_list,"-"&Request("Admin")("M_ID")&"-") > 0 then
%>
			<!--<td width=5px></td>
			<td width=77px><%=Make_BTN("Ÿ����ȯ","javascript:Change_To_Other()","")%></td>-->
<%
end if
%>

<%
if Reg_Form_YN = "Y" and instr(admin_lm_list,"-"&Request("Admin")("M_ID")&"-") > 0 then
%>		
			<td width=5px></td>
			<td width=77px>
				<div id="idBtnRegForm"><%=Make_BTN("�űԵ��","javascript:RegForm_Toggle()","")%></div>
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