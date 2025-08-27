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
strSelectName		= strSelectName & "체크,번호,파트넘버,파트이름,스펙,메이커,계획수량,현재재고,1주재고,2주재고,3주재고,4주재고,단가1,거래처1,단가2,거래처2"

strWidth			= ""
strWidth			= strWidth		& "30,40,120,100,200,100,65,65,65,65,65,65,75,120,75,120,60"
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
URL_This			= "/parts/P_Qty_list.asp"
URL_View			= "/parts/P_Qty_edit_form.asp"
URL_Action			= "/parts/P_Qty_list_action.asp"
URL_Reg				= "/parts/P_Qty_reg_action.asp"

strTable			= "vwP_Qty_List"
strPK				= "p_code"
strSelect			= ""
strSelect			= strSelect		& "P_Code,P_P_No,P_Desc,P_Spec,P_Maker,SUM_BQ_Qty,P_Qty,P_Incoming_Qty_1,P_Incoming_Qty_2,P_Incoming_Qty_3,P_Incoming_Qty_4,Partner_P_Price_1,Partner_P_Name_1,Partner_P_Price_2,Partner_P_Name_2"

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

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,,,,,"
else
	strEdit	= ",,,,,,,,,,,,,,,,"
	
end if
strPopup			= ",,,,,,,,,,,,,,,,"
strDown				= ",,,,,,,,,,,,,,,,"
strAlign			= ""
strAlign			= strAlign		& "Center,Center,Center,Center,Center,Center,"
strAlign			= strAlign		& "Center,Center,Center,Center,Center,"
strAlign			= strAlign		& "Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if Request("s_bom_sub_bs_d_no") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "P_P_No in (select Parts_P_P_No from tbBOM_Qty where BOM_Sub_BS_D_No like ''%"&Request("s_bom_sub_bs_d_no")&"%'')"
end If
if Request("s_p_p_no") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "p_p_no like ''%"&Request("s_p_p_no")&"%''"
end If

if Request("s_p_desc") <> "" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "p_desc like ''%"&Request("s_p_desc")&"%''"
end if

if Request("s_P_Qty_up") <> "" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "P_Qty >= "&Request("s_P_Qty_up")
end if

if Request("s_P_Qty_down") <> "" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "P_Qty <= "&Request("s_P_Qty_down")
end if

if Request("s_safe_qty") <> "" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "P_Qty <= p_safe_qty"
end if

if Request("s_partner_p_name") <> "" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "P_P_No in (select Parts_P_P_No from tbParts_Price where Partner_P_Name = ''"&Request("s_partner_p_name")&"'')"
end if

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "p_code"
	S_Order_By_2 	= "desc"
end if

strID				= "p_code"
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
strColumn		= "s_bom_sub_bs_d_no,s_p_p_no,s_p_desc|/|s_P_Qty_up,s_P_Qty_down,s_partner_p_name,s_safe_qty"
strName			= "모델파트넘버,부품파트넘버,파트이름|/|재고량(이상),재고량(이하),거래처,안전재고이하"
strType			= "txt,txt,txt|/|num,num,"&strEdit_Partner&",chk"
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
if Reg_Form_YN = "N" then
'6/9
'----------------------------------------------------------------------------------	
	strReg	= ",,,,,,,,,,,,,,,,"
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
'6.5/9
'----------------------------------------------------------------------------------
SQL = "select sum(SUM_BQ_Qty*Partner_P_Price_1) from "&strTable&" "

if trim(strWhere) <> "" then
	SQL = SQL & " where " & strWhere
end if

call Common_Display_Summary("계획대비 부품금액 총계 : |Sum_String|", SQL, "Y", strWidth_Total)
'----------------------------------------------------------------------------------
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
	var strError_1	= "P_P_No"
	var strError_2	= "파트넘버"
	var strError_3	= "txt";
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
	var strError_1	= "P_P_No"
	var strError_2	= "파트넘버"
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
function Set_Order()
{
	var strChecked_Value = GetChecked_Value();
	
	if (strChecked_Value == "")
	{
		alert("한개 이상의 아이템을 선택해주십시오.")
	}
	else
	{
		window.open('/parts_incoming/pi_list.asp?strChecked_Value='+strChecked_Value);
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
if Request("s_edit_mode_yn") <> "" and instr(admin_P_Qty_list,"-"&Request("Admin")("M_ID")&"-") > 0 then
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
			<td width=77px><%=Make_BTN("선택발주","javascript:Set_Order()","")%></td>
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