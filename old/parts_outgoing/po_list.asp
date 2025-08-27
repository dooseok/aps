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
if instr(admin_po_list,"-"&Request("Admin")("M_ID")&"-") > 0 then
	strSelectName		= strSelectName & "체크,번호,등록일,파트넘버,출고일,거래처,단가,수량,총계,총계(정산),비고,상태,거래구분,삭제"
else
	strSelectName		= strSelectName & "체크,번호,등록일,파트넘버,출고일,거래처,단가,수량,총계,총계(정산),비고,상태,거래구분"
end if
strWidth			= ""
strWidth			= strWidth		& "30,60,100,100,140,100,100,80,60,100,100,100,100,60"
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
URL_This			= "/parts_Outgoing/po_list.asp"
URL_View			= "/parts_Outgoing/po_delete_action.asp"
URL_Action			= "/parts_Outgoing/po_list_action.asp"
URL_Reg				= "/parts_Outgoing/po_reg_action.asp"

strTable			= "vwPO_List"
strPK				= "PO_Code"
strSelect			= ""
strSelect			= strSelect		& "PO_Code,PO_Issued_Date,Parts_P_P_No,PO_Out_Date,Partner_P_Name,PO_Price,PO_Qty,Sum_Price_Qty,Sum_Price_Qty_CAL,PO_Remark,PO_State,PO_Payment_Type"

dim RS1
dim SQL

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,"
else
	strEdit	= ",,pno,dt1,"&BasicDataPartsOutgoingComp&",txt,txt,txt,,txt,"&BasicDataPartsOutgoingState&","&BasicDataPartnerPaymentType
end if
strPopup			= ",,,,,,,,,,,,,,"
strDown				= ",,,,,,,,,,,,,,"
strAlign			= "Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------

if Len(Request("s_PO_Issued_Date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PO_Issued_Date between ''"&Left(Request("s_PO_Issued_Date"),10)&"'' and ''"&Right(Request("s_PO_Issued_Date"),10)&"''"
end if
if Len(Request("s_PO_Out_Date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PO_Out_Date between ''"&Left(Request("s_PO_Out_Date"),10)&"'' and ''"&Right(Request("s_PO_Out_Date"),10)&"''"
end if
if Request("s_Parts_LGE_PL_P_No") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "Parts_P_P_No like ''%"&Request("s_Parts_P_P_No")&"%''"
end If
if Request("s_PO_Price_up") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PO_Price >= "&Request("s_PO_Price_up")
end If
if Request("s_PO_Price_down") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PO_Price <= "&Request("s_PO_Price_down")
end If
if Request("s_PO_Qty_up") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PO_Qty >= "&Request("s_PO_Qty_up")
end If
if Request("s_PO_Qty_down") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PO_Qty <= "&Request("s_PO_Qty_down")
end If

if Request("s_PO_State") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PO_State = ''"&Request("s_PO_State")&"''"
end If

if Request("s_Partner_P_Name") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "Partner_P_Name = ''"&Request("s_Partner_P_Name")&"''"
end If

if Request("s_PO_Payment_Type") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PO_Payment_Type = ''"&Request("s_PO_Payment_Type")&"''"
end If

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "PO_Code"
	S_Order_By_2 	= "desc"
end if

strID				= "PO_Code"
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
strColumn		= "s_PO_Issued_Date,s_PO_Out_Date,s_Parts_P_P_No|/|s_PO_Price_up,s_PO_Qty_up,s_Partner_P_Name,s_PO_State|/|s_PO_Price_down,s_PO_Qty_down,s_PO_Payment_Type,s_edit_mode_yn"
strName			= "등록일,출고일,파트넘버|/|단가(이상),수량(이상),거래처,상태|/|단가(이하),수량(이하),거래구분,수정모드"
strType			= "dt2,dt2,txt|/|num,num,"&BasicDataPartsOutgoingComp&","&BasicDataPartsOutgoingState&"|/|num,num,"&BasicDataPartnerPaymentType&",chk"
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
if Reg_Form_YN = "Y" then
'6/9
'----------------------------------------------------------------------------------	
	strReg	= ",dt1"&date()&",pno,dt1,"&BasicDataPartsOutgoingComp&",txt,txt,txt,,txt,"&BasicDataPartsOutgoingState&","&BasicDataPartnerPaymentType
'----------------------------------------------------------------------------------	
	'call inc_Common_List_Reg_Form(URL_Reg, Colspan, strRequestQueryString, strSelect, arrRecordSet, strWidth, strReg, strAlign, strWidth_Total, 1)
	call Parts_Guide()
	call inc_Common_List_Reg_Form_Parts_Outgoing_Multi(URL_Reg, Colspan, strRequestQueryString, strSelect, arrRecordSet, strWidth, strReg, strAlign, strWidth_Total)
end if
%>
<img src="/img/blank.gif" width=1px height=10px><br>

<div id="idList" style="display:block;">
<%
'call inc_Common_List(strID, strID_Pos, S_PageNo, URL_This, URL_View, URL_Action, arrSelectName, strSelect, arrRecordSet, TotalRecordCount, Colspan, strRequestQueryString, S_Order_By_1, S_Order_By_2, strPopup, strDown, strWidth, strEdit, strAlign, strWidth_Total)
call inc_Common_List_Parts_Outgoing(strID, strID_Pos, S_PageNo, URL_This, URL_View, URL_Action, arrSelectName, strSelect, arrRecordSet, TotalRecordCount, Colspan, strRequestQueryString, S_Order_By_1, S_Order_By_2, strPopup, strDown, strWidth, strEdit, strAlign, strWidth_Total)
%>

<script language="javascript">
function cal_Sum_Price_Qty(strFormname,strObjIdx)
{
	objPO_Price			= eval(strFormname+".PO_Price["+strObjIdx+"]");
	objPO_Qty			= eval(strFormname+".PO_Qty["+strObjIdx+"]");
	objSum_Price_Qty	= eval(strFormname+".Sum_Price_Qty["+strObjIdx+"]");
	
	if(IsFloat(objPO_Price.value) && IsFloat(objPO_Qty.value))
	{
		objSum_Price_Qty.value = Math.round(parseFloat(objPO_Price.value) * parseFloat(objPO_Qty.value) * 100)/100;
	}
	else
	{
		objSum_Price_Qty.value = "";
	}
}

function Fill_Form(strFormname,strColumn,strValue)
{		
	strValue		= decodeURL(strValue);
	strObj 			= eval(strFormname+"." + strColumn);
	strObj.value 	= strValue;	
}

function Reset_Form(strFormname,strObjIdx)
{
	objPO_Price			= eval(strFormname+".PO_Price["+strObjIdx+"]");
	objPO_Qty			= eval(strFormname+".PO_Qty["+strObjIdx+"]");
	objPO_Payment_Type	= eval(strFormname+".PO_Payment_Type["+strObjIdx+"]");
	objSum_Price_Qty	= eval(strFormname+".Sum_Price_Qty["+strObjIdx+"]");

	objPO_Price.value			= "";
	objPO_Qty.value				= "";
	objPO_Payment_Type.value	= "";
	objSum_Price_Qty.value		= "";
}

function MakeSelectBox(strFormname,strColumn,strValue)
{
	strValue			= decodeURL(strValue);
	var arrValue		= strValue.split("|/|")
	
	strObj 				= eval(strFormname+"." + strColumn);
	strObj.length 		= arrValue.length-1;
	
	for (var CNT1=0; CNT1 < arrValue.length-1; CNT1++)
	{
		if (arrValue[CNT1] == "-선택-")
			strObj.options[CNT1]	= new Option(arrValue[CNT1],'');
		else
			strObj.options[CNT1]	= new Option(arrValue[CNT1],arrValue[CNT1]);
	}
	strObj.style.width = "100%"
}

function getParts_Info(strFormname,strObjIdx)
{
	var objP_P_No = eval(strFormname+".Parts_P_P_No["+strObjIdx+"]");
	
	if (objP_P_No.value.length != "")
	{
		ifrmParts_Info.location.href="/parts_outgoing/inc_parts_info.asp?strFormname="+strFormname+"&s_P_P_No=" + objP_P_No.value + "&strObjIdx="+strObjIdx;		
	}
}

function getParts_Partner_Info(strFormname,strObjIdx)
{
	var objP_P_No			= eval(strFormname+".Parts_P_P_No["+strObjIdx+"]");
	var objPartner_P_Name	= eval(strFormname+".Partner_P_Name["+strObjIdx+"]");
	if (objP_P_No.value.length != "")
	{
		ifrmParts_Info.location.href="/parts_outgoing/inc_parts_info.asp?strFormname="+strFormname+"&s_P_P_No=" + objP_P_No.value + "&s_Partner_P_Name=" + encodeURL(objPartner_P_Name.value) + "&strObjIdx="+strObjIdx;		
	}
}

</script>

<img src="/img/blank.gif" width=1px height=5px><br>
<%
call inc_Common_Paging(URL_This, TotalRecordCount, S_PageSize, S_PageNo, strRequestQueryString)
%>
<%
'6.5/9
'----------------------------------------------------------------------------------
SQL = "select sum(Sum_Price_Qty_CAL) from "&strTable&" "

if trim(strWhere) <> "" then
	SQL = SQL & " where " & strWhere
end if

call Common_Display_Summary("출고금액 총계 (정상) : |Sum_String|", SQL, "Y", strWidth_Total)
'----------------------------------------------------------------------------------
%>
<img src="/img/blank.gif" width=1px height=50px><br>
</div>
</div>

<script language="javascript">
function List_Reg()
{
<%
'7/9
'----------------------------------------------------------------------------------
%>
	//var strError = List_Reg_Validater_Multi('PO_Issued_Date,Parts_P_P_No,PO_Price,PO_Qty','출고일,파트넘버,단가,주문수량','txt,txt,num,num');
	var strError = "";
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
	//var strError = List_Validater('PO_Issued_Date,Parts_P_P_No,PO_Price,PO_Qty','출고일,파트넘버,단가,주문수량','txt,txt,num,num');
	var strError = "";
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
if Request("s_edit_mode_yn") <> "" and instr(admin_po_list,"-"&Request("Admin")("M_ID")&"-") > 0 then
%>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("수정완료","javascript:List_Update()","")%></td>
<%
end if
%>
<%
if Reg_Form_YN = "Y" and instr(admin_po_list,"-"&Request("Admin")("M_ID")&"-") > 0 then
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


<%
sub inc_Common_List_Reg_Form_Parts_Outgoing_Multi(URL_Reg, Colspan, strRequestQueryString, strSelect, arrRecordSet, strWidth, strReg, strAlign, strWidth_Total)
	dim CNT0
	dim CNT1
	dim CNT2
	dim CNT3
	
	dim nPaddingTop
	dim strEdit
	strEdit = strReg
	
	dim strRequestQueryString_dummy
	
	dim arrWidth
	dim strWidth_Cal
	
	dim arrSelect
	dim arrEdit
	dim arrAlign
	
	dim arrInputSelectG
	dim arrInputSelect
	
	arrWidth	= split(strWidth,",")
	arrSelect	= split(strSelect,",")
	arrEdit		= split(strEdit,",")
	arrAlign	= split(strAlign,",")
	
	dim strHeight
	
	if instr(strEdit,"mem") > 0 then
		strHeight = "40"
		nPaddingTop = "14"
	else
		strHeight = "20"
		nPaddingTop = "4"
	end if
%>
<div id=idRegForm style="display:none;">
<table width="<%=strWidth_Total%>px" border=0 cellspacing=1 cellpadding=0 bgcolor="#999999" class="Common_List_Reg_Form">
<form name="frmCommonListReg" action="<%=URL_Reg%>?dummy=<%=strRequestQueryString%>" method="post">
<tr height=33px bgcolor="#e0e0e0">
<%
	for CNT1 = 0 to ubound(arrSelectName)
		if arrSelectName(CNT1) = "체크" then
%>
	<td<%if arrWidth(CNT1) <> "" then%> width="<%=arrWidth(CNT1)%>px"<%end if%>>&nbsp;</td>
<%
		elseif arrSelectName(CNT1) = "작업" then
%>
	<td<%if arrWidth(CNT1) <> "" then%> width="<%=arrWidth(CNT1)%>px"<%end if%>>&nbsp;</td>
<%		
		elseif arrSelectName(CNT1) = "삭제" then
%>
	<td<%if arrWidth(CNT1) <> "" then%> width="<%=arrWidth(CNT1)%>px"<%end if%>>&nbsp;</td>
<%		
		else
			if arrSelectName(0) = "체크" then
%>
	<td<%if arrWidth(CNT1) <> "" then%> width="<%=arrWidth(CNT1)%>px"<%end if%>><b style="color:navy"><%=arrSelectName(CNT1)%></b></td>
<%
			else
%>
	<td<%if arrWidth(CNT1) <> "" then%> width="<%=arrWidth(CNT1)%>px"<%end if%>><b style="color:navy"><%=arrSelectName(CNT1)%></b></td>
<%
			end if
		end if
	next
%>
</tr>
<%
for CNT0 = 0 to 19
%>
<tr height=<%=strHeight%>px bgcolor="<%if (CNT2 mod 2) = 0 then%>#ffffff<%else%>#ffffff<%end if%>" valign=top onmouseover="this.style.backgroundColor='pink';" onmouseout="this.style.backgroundColor='<%if (CNT2 mod 2) = 0 then%>#ffffff<%else%>#ffffff<%end if%>';">
<input type="hidden" name="strID_All" value="<%=CNT0%>">
<%
		if arrSelectName(0)="체크" then
%>
			<td<%if arrWidth(0) <> "" then%> width="<%=arrWidth(0)%>px"<%end if%>>&nbsp;</td>
<%
		end if
		
		if arrSelectName(0)="체크" then
			strWidth_Cal = arrWidth(1)
		else
			strWidth_Cal = arrWidth(0)
		end if
		
		for CNT1 = 0 to ubound(arrRecordSet,1)
			if arrEdit(CNT1) <> "" then				
				select case left(arrEdit(CNT1),3)
				case "txt"
%>
			<td><input type="text" size=10 name="<%=arrSelect(CNT1)%>"<%if instr("-Sum_Price_Qty-","-"&arrSelect(CNT1)&"-") > 0 then%> readonly<%end if%> <%if instr("-PO_Qty-PO_Price-","-"&arrSelect(CNT1)&"-") > 0 then%> onkeyup="javascript:cal_Sum_Price_Qty('frmCommonListReg',<%=CNT0%>);"<%end if%> value="" class=input style="width:100%;text-align:<%=arrAlign(CNT1)%>;height:<%=strHeight-1%>px;padding-top:<%=nPaddingTop%>px; "></td>
<%
				case "mem"
%>
			<td><textarea wrap="soft" name="<%=arrSelect(CNT1)%>" class=input rows=2 style="width:100%;height:<%=strHeight-1%>px;overflow:auto;text-align:<%=arrAlign(CNT1)%>;padding-top:<%=nPaddingTop%>px;" scrollbar=auto></textarea></td>
<%
				case "pno"
%>
			<td><input type="text" size=10 name="<%=arrSelect(CNT1)%>" value="" onclick="javascript:show_Parts_Guide(this,'frmCommonListReg',<%=CNT0%>);" class=input style="width:100%;text-align:<%=arrAlign(CNT1)%>;height:<%=strHeight-1%>px;padding-top:<%=nPaddingTop%>px;"></td>
<%
				case "num"
%>
			<td><input type="text" size=10 name="<%=arrSelect(CNT1)%>" value="" class=input style="width:100%;text-align:<%=arrAlign(CNT1)%>;height:<%=strHeight-1%>px;padding-top:<%=nPaddingTop%>px;"></td>
<%
				case "dt1"
					if len(arrEdit(CNT1)) > 3 then
%>
			<td><input type="text" readonly size=10 name="<%=arrSelect(CNT1)%>" value="<%=right(arrEdit(CNT1),len(arrEdit(CNT1))-3)%>" class=input style="width:100%;text-align:<%=arrAlign(CNT1)%>;height:<%=strHeight-1%>px;padding-top:<%=nPaddingTop%>px;" onclick="Calendar_D(this)"></td>
<%
					else
%>
			<td><input type="text" readonly size=10 name="<%=arrSelect(CNT1)%>" value="" class=input style="width:100%;text-align:<%=arrAlign(CNT1)%>;height:<%=strHeight-1%>px;padding-top:<%=nPaddingTop%>px;" onclick="Calendar_D(this)"></td>
<%
				end if
				case "slt"
					arrInputSelectG	= split(replace(arrEdit(CNT1),"slt>",""),";")			
%>
			<td valign=middle><select name="<%=arrSelect(CNT1)%>"<%if arrSelect(CNT1) = "Partner_P_Name" then%> onchange="javascript:getParts_Partner_Info('frmCommonListReg',<%=CNT0%>);"<%end if%>>
<%
				if arrSelect(CNT1) = "PO_State" then
%>
				<option value="출고준비">출고준비</option>
<%				
				else
%>
				<option value="">-선택-</option>
<%
					for CNT3 = 0 to ubound(arrInputSelectG)
						arrInputSelect = split(arrInputSelectG(CNT3),":")
						if arrInputSelect(0) = "-1" then
							arrInputSelect(0) = ""
						elseif isnull(arrInputSelect(0)) then
							arrInputSelect(0) = ""
						end if
%>
				<option value="<%=arrInputSelect(0)%>"><%=arrInputSelect(1)%></option>
<%
					next
%>
			</select>
<%
				end if
%>
		</td>
<%
				end select
			else
				if arrSelectName(0)="체크" then
%>
			<td align="<%=arrAlign(CNT1)%>"><textarea rows=1 readonly style="width:<%if arrSelectName(CNT1+1)="번호" then%><%=strWidth_Cal%><%else%><%=arrWidth(CNT1)%><%end if%>px;height:<%=strHeight-1%>px;overflow:auto;text-align:center;padding-top:<%=nPaddingTop%>px;background-color:transparent"><%if arrSelectName(CNT1+1)="번호" then%>NEW<%end if%></textarea></td>
<%		
				else
%>
			<td align="<%=arrAlign(CNT1)%>"><textarea rows=1 readonly style="width:<%if arrSelectName(CNT1+1)="번호" then%><%=strWidth_Cal%><%else%><%=arrWidth(CNT1)%><%end if%>px;height:<%=strHeight-1%>px;overflow:auto;text-align:center;padding-top:<%=nPaddingTop%>px;background-color:transparent"><%if arrSelectName(CNT1)="번호" then%>NEW<%end if%></textarea></td>
<%				
				end if
			end if
		next
		if arrSelectName(ubound(arrSelectName))="작업" then
%>
			<td>&nbsp;</td>
<%
		end if
		if arrSelectName(ubound(arrSelectName))="삭제" then
%>
			<td>&nbsp;</td>
<%
		end if
%>
</tr>
<%
next
%>
</form>

<iframe name="ifrmParts_Info" src="about:blank;" frameborder=0 width=0px height=0px></iframe>
</table>
<img src="/img/blank.gif" width=1px height=10px><br>
<table width=100% cellpadding=0 cellspacing=0 border=0>
<tr>
	<td align=center>
		<table width=159px cellpadding=0 cellspacing=0 border=0>
		<tr>
			<td width=77px><%=Make_BTN("신규등록","javascript:List_Reg();","")%></td>
			<td width=5px><img src="/img/blank.gif" width=5px height=1px></td>
			<td width=77px><%=Make_BTN("폼지우기","javascript:frmCommonListReg.reset();","")%></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<img src="/img/blank.gif" width=1px height=10px><br>
</div>
<%
end sub
%>

<%
sub inc_Common_List_Parts_Outgoing(strID, strID_Pos, S_PageNo, URL_This,URL_View, URL_Action, arrSelectName, strSelect, arrRecordSet, TotalRecordCount, Colspan, strRequestQueryString, S_Order_By_1, S_Order_By_2, strPopup, strDown, strWidth, strEdit, strAlign, strWidth_Total)
	dim CNT1
	dim CNT2
	dim CNT3
	
	dim strRequestQueryString_dummy
	
	dim nPaddingTop
	
	dim arrWidth
	dim strWidth_Cal
	
	dim arrSelect
	dim arrEdit
	dim arrPopup
	dim arrDown
	dim arrAlign
	
	
	dim arrInputSelectG
	dim arrInputSelect
	
	arrWidth	= split(strWidth,",")
	arrSelect	= split(strSelect,",")
	arrEdit		= split(strEdit,",")
	arrPopup	= split(strPopup,",")
	arrDown		= split(strDown,",")
	arrAlign	= split(strAlign,",")
	
	dim strHeight
	
	if instr(strEdit,"mem") > 0 then
		strHeight = "40"
		nPaddingTop = "14"
	else
		strHeight = "20"
		nPaddingTop = "4"
	end if
	
	if TotalRecordCount = "" then
		TotalRecordCount = 0
	end if
%>
<script language="javascript">

function Check_All(form)
{
	var cnt1;
	
	if(frmCommonList.strID.length)
	{
		if(frmCommonList.Check_All_YN.checked == true)
		{
			for (cnt1 = 0; cnt1 < frmCommonList.strID.length; cnt1++)
				frmCommonList.strID[cnt1].checked = true;
		}
		else if(form.Check_All_YN.checked == false)
		{
			for (cnt1 = 0; cnt1 < frmCommonList.strID.length; cnt1++)
				frmCommonList.strID[cnt1].checked = false;
		}
	}
	else
	{
		if(frmCommonList.Check_All_YN.checked == true)
		{
			frmCommonList.strID.checked = true;
		}
		else if(frmCommonList.Check_All_YN.checked == false)
		{
			frmCommonList.strID.checked = false;
		}
	}
}

function GetChecked_Value()
{
	var strChecked_Value = "";
	
	if(typeof(frmCommonList.strID) == [object])
	{
		if(frmCommonList.strID.length)
		{
			for (cnt1 = 0; cnt1 < frmCommonList.strID.length; cnt1++)
			{
				if(frmCommonList.strID[cnt1].checked == true)
				{
					strChecked_Value += frmCommonList.strID[cnt1].value + ",";
				}
			}
		}
		else
		{
			if(frmCommonList.strID.checked == true)
			{
				strChecked_Value += frmCommonList.strID.value + ",";
			}
		}
	}
	return strChecked_Value;
}

function setSorting(S_Order_By_1,S_Order_By_2)
{
<%
strRequestQueryString_dummy = strRequestQueryString
strRequestQueryString_dummy = replace(strRequestQueryString_dummy,"S_Order_By_1=","Dummy_Order_By_1=")
strRequestQueryString_dummy = replace(strRequestQueryString_dummy,"S_Order_By_2=","Dummy_Order_By_2=")
strRequestQueryString_dummy = replace(strRequestQueryString_dummy,"S_Order_By_3=","Dummy_Order_By_3=")
strRequestQueryString_dummy = replace(strRequestQueryString_dummy,"S_Order_By_4=","Dummy_Order_By_4=")
%>
	if ("<%=lcase(S_Order_By_1)%>"==S_Order_By_1.toLowerCase())
	{
		location.href="<%=URL_This%>?S_Order_By_1="+S_Order_By_1+"&S_Order_By_2="+S_Order_By_2+"<%=strRequestQueryString_dummy%>";
	}
	else if ("<%=S_Order_By_1%>"=="")
	{
		location.href="<%=URL_This%>?S_Order_By_1="+S_Order_By_1+"&S_Order_By_2="+S_Order_By_2+"<%=strRequestQueryString_dummy%>";
	}
	else
	{
		location.href="<%=URL_This%>?S_Order_By_1="+S_Order_By_1+"&S_Order_By_2="+S_Order_By_2+"&S_Order_By_3=<%=S_Order_By_1%>&S_Order_By_4=<%=S_Order_By_2%><%=strRequestQueryString_dummy%>";
	}
}

function Delete_Check(strID_Value)
{
	if(confirm("정말 삭제하시겠습니까?"))
	{
		location.href="<%=URL_View%>?<%=strID%>="+strID_Value+"<%=strRequestQueryString%>";
	}
}
</script>

<table width="<%=strWidth_Total%>px" cellpadding=0 cellspacing=0 border=0>
<tr>
	<td align=left>&nbsp;<span style="font-face:돋움:font-size:8pt"><span style='color:red'><%=TotalRecordCount%>건</span>이 검색되었습니다.
		<img src="/img/blank.gif" width=5px>
<%
	response.write "현재 정렬기준은"
	for CNT1 = 0 to ubound(arrSelect)
		if lcase(arrSelect(CNT1)) = lcase(S_Order_By_1) then
			response.write "<span style='color:red'>"
			if arrSelectName(0)="체크" then
				response.write "「" & arrSelectName(CNT1+1)
			else
				response.write "「" & arrSelectName(CNT1)
			end if
			
		end if
	next
	
	if S_Order_By_2 = "asc" then
		response.write " - 오름차순" & "」"
	else
		response.write " - 내림차순" & "」"
	end if
	response.write "</span>"
	
	if S_Order_By_3 <> "" then
		response.write ","
		for CNT1 = 0 to ubound(arrSelect)
			if lcase(arrSelect(CNT1)) = lcase(S_Order_By_3) then
				response.write "<span style='color:blue'>"
				if arrSelectName(0)="체크" then
					response.write "「" & arrSelectName(CNT1) 
				else
					response.write "「" & arrSelectName(CNT1)
				end if
			end if
		next
		
		if S_Order_By_4 = "asc" then
			response.write " - 오름차순" & "」"
		else
			response.write " - 내림차순" & "」"
		end if
		response.write "</span>"
	end if
	
	response.write "입니다.</span>"
%>
	</td>
</tr>
</table>
<img src="/img/blank.gif" width=1px height=3px><br>
<table width="<%=strWidth_Total%>px" border=0 cellspacing=1 cellpadding=0 bgcolor="#999999" class="Common_List">
<form name="frmCommonList" action="<%=URL_Action%>?dummy=<%=strRequestQueryString%>" method="post">
<tr height=33px bgcolor="#e0e0e0">
<%
	for CNT1 = 0 to ubound(arrSelectName)
		if arrSelectName(CNT1) = "체크" then
%>
	<td<%if arrWidth(CNT1) <> "" then%> width="<%=arrWidth(CNT1)%>px"<%end if%>><input type="checkbox" name="Check_All_YN" style="border:0px none #ffffff;background-color:#ffffff" onclick="javascript:Check_All(frmCommonList)"></td>
<%
		elseif arrSelectName(CNT1) = "작업" then
%>
	<td<%if arrWidth(CNT1) <> "" then%> width="<%=arrWidth(CNT1)%>px"<%end if%>><b style="color:navy"><%=arrSelectName(CNT1)%></b></td>
<%		
		elseif arrSelectName(CNT1) = "삭제" then
%>
	<td<%if arrWidth(CNT1) <> "" then%> width="<%=arrWidth(CNT1)%>px"<%end if%>><b style="color:navy"><%=arrSelectName(CNT1)%></b></td>
<%		
		else
			if arrSelectName(0) = "체크" then
%>
	<td<%if arrWidth(CNT1) <> "" then%> width="<%=arrWidth(CNT1)%>px"<%end if%> style="line-height:10px"><span style="cursor:hand;" onclick="javascript:setSorting('<%=arrSelect(CNT1-1)%>','ASC')"><img src="/img/ico_sorting_up<%if lcase(S_Order_By_1) = lcase(arrSelect(CNT1-1)) and lcase(S_Order_By_2) = "asc" then%>_Red<%elseif lcase(S_Order_By_3) = lcase(arrSelect(CNT1-1)) and lcase(S_Order_By_4) = "asc" then%>_Blue<%end if%>.gif"></span><br><img src="/img/blank.gif" width=1px height=3px><br><b style="color:navy"><%=arrSelectName(CNT1)%></b><br><span style="cursor:hand;" onclick="javascript:setSorting('<%=arrSelect(CNT1-1)%>','DESC')"><img src="/img/ico_sorting_down<%if lcase(S_Order_By_1) = lcase(arrSelect(CNT1-1)) and lcase(S_Order_By_2) = "desc" then%>_Red<%elseif lcase(S_Order_By_3) = lcase(arrSelect(CNT1-1)) and lcase(S_Order_By_4) = "desc" then%>_Blue<%end if%>.gif"></span></td>
<%
			else
%>
	<td<%if arrWidth(CNT1) <> "" then%> width="<%=arrWidth(CNT1)%>px"<%end if%> style="line-height:10px"><span style="cursor:hand;" onclick="javascript:setSorting('<%=arrSelect(CNT1)%>','ASC')"><img src="/img/ico_sorting_up<%if lcase(S_Order_By_1) = lcase(arrSelect(CNT1)) and lcase(S_Order_By_2) = "asc" then%>_Red<%elseif lcase(S_Order_By_3) = lcase(arrSelect(CNT1)) and lcase(S_Order_By_4) = "asc" then%>_Blue<%end if%>.gif"></span><br><img src="/img/blank.gif" width=1px height=3px><br><b style="color:navy"><%=arrSelectName(CNT1)%></b><br><span style="cursor:hand;" onclick="javascript:setSorting('<%=arrSelect(CNT1)%>','DESC')"><img src="/img/ico_sorting_down<%if lcase(S_Order_By_1) = lcase(arrSelect(CNT1)) and lcase(S_Order_By_2) = "desc" then%>_Red<%elseif lcase(S_Order_By_3) = lcase(arrSelect(CNT1)) and lcase(S_Order_By_4) = "desc" then%>_Blue<%end if%>.gif"></span></td>
<%
			end if
		end if
	next
%>
</tr>
<%
	for CNT2 = 0 to ubound(arrRecordSet,2)-1
%>
<tr height=<%=strHeight%>px bgcolor="<%if (CNT2 mod 2) = 0 then%>#ffffff<%else%>#ffffff<%end if%>" valign=top onmouseover="this.style.backgroundColor='pink';" onmouseout="this.style.backgroundColor='<%if (CNT2 mod 2) = 0 then%>#ffffff<%else%>#ffffff<%end if%>';">
<input type="hidden" name="strID_All" value="<%=arrRecordSet(strID_Pos, CNT2)%>">
<%
		if arrSelectName(0)="체크" then
%>
			<td<%if arrWidth(0) <> "" then%> width="<%=arrWidth(0)%>px"<%end if%> valign=middle><input type="checkbox" name="strID" value="<%=arrRecordSet(strID_Pos, CNT2)%>" style="border:0px none #ffffff;background-color:<%if (CNT2 mod 2) = 0 then%>#FEFFD6<%else%>#ffffff<%end if%>"></td>
<%
		end if
		for CNT1 = 0 to ubound(arrRecordSet,1)
			if arrSelectName(0)="체크" then
				strWidth_Cal = arrWidth(CNT1 + 1)
			else
				strWidth_Cal = arrWidth(CNT1)
			end if
			
			if arrEdit(CNT1) <> "" then				
				select case left(arrEdit(CNT1),3)
				case "txt"
%>
			<td title="<%=arrRecordSet(CNT1, CNT2)%>"><input type="text" size=10 name="<%=arrSelect(CNT1)%>"<%if instr("-Sum_Price_Qty-","-"&arrSelect(CNT1)&"-") > 0 then%> readonly<%end if%> <%if instr("-PO_Qty-PO_Price-","-"&arrSelect(CNT1)&"-") > 0 then%> onkeyup="javascript:cal_Sum_Price_Qty('frmCommonList',<%=CNT2%>);"<%end if%> value="<%=arrRecordSet(CNT1, CNT2)%>" class=input style="width:100%;text-align:<%=arrAlign(CNT1)%>;height:<%=strHeight-1%>px;padding-top:<%=nPaddingTop%>px; "></td>
<%
				case "mem"
%>
			<td title="<%=arrRecordSet(CNT1, CNT2)%>"><textarea wrap="soft" name="<%=arrSelect(CNT1)%>" class=input rows=2 style="width:<%=strWidth_Cal%>px;height:<%=strHeight-1%>px;overflow:auto;text-align:<%=arrAlign(CNT1)%>;padding-top:<%=nPaddingTop%>px;" scrollbar=auto><%=arrRecordSet(CNT1, CNT2)%></textarea></td>
<%
				case "pno"
%>
			<td title="<%=arrRecordSet(CNT1, CNT2)%>"><input type="text" size=10 name="<%=arrSelect(CNT1)%>" value="<%=arrRecordSet(CNT1, CNT2)%>" onclick="javascript:show_Parts_Guide(this,'frmCommonList',<%=CNT2%>);" class=input style="width:100%;text-align:<%=arrAlign(CNT1)%>;height:<%=strHeight-1%>px;padding-top:<%=nPaddingTop%>px;"></td>
<%
				case "num"
%>
			<td title="<%=arrRecordSet(CNT1, CNT2)%>"><input type="text" size=10 name="<%=arrSelect(CNT1)%>" value="<%=arrRecordSet(CNT1, CNT2)%>" class=input style="width:100%;text-align:<%=arrAlign(CNT1)%>;height:<%=strHeight-1%>px;padding-top:<%=nPaddingTop%>px;"></td>
<%
				case "mny"
%>
			<td title="<%=arrRecordSet(CNT1, CNT2)%>"><input type="text" size=10 name="<%=arrSelect(CNT1)%>" value="<%=replace(arrRecordSet(CNT1, CNT2),"원","")%>" class=input style="width:100%;text-align:<%=arrAlign(CNT1)%>;height:<%=strHeight-1%>px;padding-top:<%=nPaddingTop%>px;"></td>
<%
				case "dt1"
%>
			<td title="<%=arrRecordSet(CNT1, CNT2)%>"><input type="text" readonly size=10 name="<%=arrSelect(CNT1)%>" value="<%=arrRecordSet(CNT1, CNT2)%>" class=input style="width:100%;text-align:<%=arrAlign(CNT1)%>;height:<%=strHeight-1%>px;padding-top:<%=nPaddingTop%>px;" onclick="Calendar_D(document.frmCommonList.<%=arrSelect(CNT1)%><%if ubound(arrRecordSet,2) > 1 then%>[<%=CNT2%>]<%end if%>)"></td>
<%
				case "slt"
					arrInputSelectG	= split(replace(arrEdit(CNT1),"slt>",""),";")			
%>
			<td valign=middle title="<%=arrRecordSet(CNT1, CNT2)%>"><select name="<%=arrSelect(CNT1)%>"<%if arrSelect(CNT1) = "Partner_P_Name" then%> onchange="javascript:getParts_Partner_Info('frmCommonList',<%=CNT2%>);"<%end if%>>
<%
					for CNT3 = 0 to ubound(arrInputSelectG)
						arrInputSelect = split(arrInputSelectG(CNT3),":")
						if arrInputSelect(0) = "-1" then
							arrInputSelect(0) = ""
						elseif isnull(arrInputSelect(0)) then
							arrInputSelect(0) = ""
						end if
						
						if isnull(arrRecordSet(CNT1, CNT2)) then
							arrRecordSet(CNT1, CNT2) = ""
						end if
%>
				<option value="<%=arrInputSelect(0)%>"<%if cstr(arrRecordSet(CNT1, CNT2)) = cstr(arrInputSelect(0)) then%> selected<%end if%>><%=arrInputSelect(1)%></option>
<%
					next
%>
			</select></td>
<%
				end select
		
			elseif arrDown(CNT1) <> "" then
%>
			<td align="<%=arrAlign(CNT1)%>" title="<%=arrRecordSet(CNT1, CNT2)%>"><a href="/function/ifrm_download.asp?filepath=<%=arrDown(CNT1)%><%=arrRecordSet(CNT1, CNT2)%>" style='color:blue' target="ifrm_download"><%=arrRecordSet(CNT1, CNT2)%></a></td>
<%				
			elseif arrPopup(CNT1) <> "" then
%>
			<td align="<%=arrAlign(CNT1)%>" title="<%=arrRecordSet(CNT1, CNT2)%>"><a href="<%=arrPopup(CNT1)%>?<%=strID%>=<%=arrRecordSet(0, CNT2)%>" style='color:blue' target="_blank"><%=arrRecordSet(CNT1, CNT2)%></a></td>
<%
			else
%>
			<td align="<%=arrAlign(CNT1)%>" title="<%=arrRecordSet(CNT1, CNT2)%>"><%if arrRecordSet(CNT1, CNT2)="-1" then%>미입력<%else%><textarea rows=1 readonly style="width:<%=strWidth_Cal%>px;height:<%=strHeight-1%>px;overflow:auto;text-align:<%=arrAlign(CNT1)%>;<%if Request("S_Edit_Mode_YN") = "checked" then%>padding-top:<%=nPaddingTop%>px;<%else%>padding-top:3px;<%end if%>background-color:transparent"><%=arrRecordSet(CNT1, CNT2)%></textarea><%end if%></td>
<%		
			end if
		next
		if arrSelectName(ubound(arrSelectName))="작업" then
%>
			<td valign=middle>
				<span style="cursor:hand;color:navy" onclick="javascript:location.href='<%=URL_View&"?"&strID&"="&arrRecordSet(0, CNT2)&strRequestQueryString%>'"><u>보기</u></span>
			</td>
<%
		end if
		if arrSelectName(ubound(arrSelectName))="삭제" then
%>
			<td valign=middle>
				<span style="cursor:hand;color:navy" onclick="javascript:Delete_Check('<%=arrRecordSet(0, CNT2)%>')"><u>삭제</u></span>
			</td>
<%
		end if
%>
</tr>
<%
	next
%>
</form>
<iframe name="ifrmParts_Info" src="about:blank;" frameborder=0 width=0px height=0px></iframe>
<iframe name="ifrm_download" src="about:blank" width=0 height=0 frameborder=0></iframe>
</table>
<%
end sub
%>



<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->