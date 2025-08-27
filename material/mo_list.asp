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
strSelectName		= strSelectName & "��ȣ,��Ʈ�ѹ�,����,����,BOM����,�ŷ�ó,�ܰ�,Ȯ������,���ַ�,�԰�,�ѹ��԰�,������,������,������,���ֿϷ�,�μ���,�ӿ�"

strWidth			= ""
strWidth			= strWidth		& "40,100,120,150,250,100,60,70,60,60,60,70,70,70,70,70,70"
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
call Material_Qty_Log_Popup_List()'

URL_This			= "/material/mo_list.asp"
URL_View			= "/material/mo_edit_form.asp"
URL_Action			= "/material/mo_list_action.asp"
URL_Reg				= "/material/mo_reg_action.asp"

strTable			= "vwMO_List"
strPK				= "mo_code"
strSelect			= ""
strSelect			= strSelect		& "MO_Code,Material_M_P_No,Material_M_Desc,Material_M_Spec,Material_M_Spec_BOM,Partner_P_Name,MO_Price,MO_Price_Temp_YN,MO_Qty,MO_Qty_In,Material_M_Qty_Include_coming,MO_Reg_ID,MO_Order_Date,MO_Due_Date,MO_Check_1_YN,MO_Check_2_YN,MO_Check_3_YN"

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,,,,,,,,,,,"
else
	strEdit	= ",,,,,,,,num,,,,,dt1,cid,cid,cid,,,,,"
end if
strPopup			= ",s_Material_M_P_No3,,,,,s_Material_M_P_No2,,,,,,,,,,,,,,,,"
strDown				= ",,,,,,,,,,,,,,,,,,,,,,,"
strAlign			= ""
strAlign			= strAlign		& "Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if Request("s_Total") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "(Material_M_P_No like ''%"&Request("s_Total")&"%'' or Partner_P_Name like ''%"&Request("s_Total")&"%'' or Material_M_Desc like ''%"&Request("s_Total")&"%'' or Material_M_Spec like ''%"&Request("s_Total")&"%'' or Material_M_Spec_Bom like ''%"&Request("s_Total")&"%'')"
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
	strWhere = strWhere & "MO_Reg_ID like ''%"&Request("s_MO_Reg_ID")&"%''"
end If

if Len(Request("s_MO_Due_Date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "MO_Due_Date between ''"&Left(Request("s_MO_Due_Date"),10)&"'' and ''"&Right(Request("s_MO_Due_Date"),10)&"''"
end If

if Request("s_MO_Check_1_YN") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	if Request("s_MO_Check_1_YN") = "�̰���" then
		strWhere = strWhere & "MO_Check_1_YN = ''''"
	elseif Request("s_MO_Check_1_YN") = "����" then
		strWhere = strWhere & "MO_Check_1_YN <> ''''"
	end if
end If

if Request("s_MO_Check_2_YN") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	if Request("s_MO_Check_2_YN") = "�̰���" then
		strWhere = strWhere & "MO_Check_2_YN = ''''"
	elseif Request("s_MO_Check_2_YN") = "����" then
		strWhere = strWhere & "MO_Check_2_YN <> ''''"
	end if
end If

if Request("s_MO_Check_3_YN") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	if Request("s_MO_Check_3_YN") = "�̰���" then
		strWhere = strWhere & "MO_Check_3_YN = ''''"
	elseif Request("s_MO_Check_3_YN") = "����" then
		strWhere = strWhere & "MO_Check_3_YN <> ''''"
	end if
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
if instr("-����-��ȹ-�ѹ�-�濵��-","-"&Request.Cookies("Admin")("M_Part")&"-") > 0  or instr("-shindk-kimdh-leehs-","-"&gM_ID&"-") > 0 then
	strColumn		= "s_Total,s_M_P_No,s_Partner_P_Name,s_MO_Reg_ID,s_MO_Due_Date|/|s_MO_Check_1_YN,s_MO_Check_2_YN,s_MO_Check_3_YN,s_edit_mode_yn"
	strName			= "���հ˻�,��Ʈ�ѹ�,�ŷ�ó,������,���ֱⰣ|/|Ȯ��,�μ���,�ӿ�,�������"
	strType			= "txt,txt,ptn,mnm-����-����-��ȹ-,dt2|/|slt>:��ü;����:����;�̰���:�̰���,slt>:��ü;����:����;�̰���:�̰���,slt>:��ü;����:����;�̰���:�̰���,chk"
else
	strColumn		= "s_Total,s_M_P_No,s_Partner_P_Name,s_MO_Reg_ID,s_MO_Due_Date|/|s_MO_Check_1_YN,s_MO_Check_2_YN,s_MO_Check_3_YN"
	strName			= "���հ˻�,��Ʈ�ѹ�,�ŷ�ó,������,���ֱⰣ|/|Ȯ��,�μ���,�ӿ�"
	strType			= "txt,txt,ptn,mnm-����-����-��ȹ-,dt2|/|slt>:��ü;����:����;�̰���:�̰���,slt>:��ü;����:����;�̰���:�̰���,slt>:��ü;����:����;�̰���:�̰���"
end if
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
	strReg	= ",mtr,,,,,,num,,,,,,,,,,,,,,"
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
	var strError_1	= "Material_M_P_No,MO_Qty"
	var strError_2	= "��Ʈ�ѹ�,����"
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
	var strError_1	= "MO_Qty";
	var strError_2	= "����";
	var strError_3	= "num";
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
			<td width=77px><%=Make_BTN("�����Ϸ�","javascript:List_Update()","")%></td>
<%
end if
%>
<%
if Reg_Form_YN = "Y" then
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
			<td width=30px></td>
<%
if Len(Request("s_MO_Due_Date")) = "22" Then
%>
			<td width=40px><%=Make_s_BTN("������","MovePrev()","")%></td>
			<td width=5px></td>
<%
else
%>
			<td width=40px>&nbsp;</td>
			<td width=5px></td>
<%
end if
%>
			<td width=40px><%=Make_s_BTN("����","MoveToday()","")%></td>
			<td width=5px></td>
<%
if Len(Request("s_MO_Due_Date")) = "22" Then
%>
			<td width=40px><%=Make_s_BTN("������","MoveNext()","")%></td>
			<td width=5px></td>
<%
else
%>
			<td width=40px>&nbsp;</td>
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

<script language="javascript">
<%
if Len(Request("s_MO_Due_Date")) = "22" Then
%>
	function MovePrev()
	{
		frmSearch_Bar.s_MO_Due_Date[0].value = "<%=dateadd("d",-1,left(request("s_MO_Due_Date"),10))%>";
		frmSearch_Bar.s_MO_Due_Date[1].value = "<%=dateadd("d",-1,left(request("s_MO_Due_Date"),10))%>";
		frmSearch_Bar.submit()
	}
	function MoveNext()
	{
		frmSearch_Bar.s_MO_Due_Date[0].value = "<%=dateadd("d",1,left(request("s_MO_Due_Date"),10))%>";
		frmSearch_Bar.s_MO_Due_Date[1].value = "<%=dateadd("d",1,left(request("s_MO_Due_Date"),10))%>";
		frmSearch_Bar.submit()
	}
<%
end if
%>
	function MoveToday()
	{
		frmSearch_Bar.s_MO_Due_Date[0].value = "<%=date()%>";
		frmSearch_Bar.s_MO_Due_Date[1].value = "<%=date()%>";
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