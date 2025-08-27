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
strSelectName		= strSelectName & "번호,파트넘버,출발일,수량,단가,금액,LG제번,기타코드"

strWidth			= ""
strWidth			= strWidth		& "70,120,120,80,80,80,100,100"
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
URL_This			= "/product_send/ps_list.asp"
URL_View			= "/product_send/ps_edit_form.asp"
URL_Action			= "/product_send/ps_list_action.asp"
URL_Reg				= "/product_send/ps_reg_action.asp"

call BOMSub_Guide

strTable			= "vwPS_List"
strPK				= "ps_code"
strSelect			= ""
strSelect			= strSelect		& "PS_Code,BOM_Sub_BS_D_No,PS_Send_Date,PS_Qty,BOM_Price_BP_Price,Sum_Price,LGE_Plan_LP_Work_Order,LGE_Plan_ETC_LPE_Code"

dim RS1
dim SQL
dim strLP_Work_Order
dim strLPE_Code
set RS1 = Server.CreateObject("ADODB.RecordSet")

strLP_Work_Order	= "slt>"
SQL = "select distinct(LP_Work_Order) from tbLGE_Plan where exists (select LPD_Code from tbLGE_Plan_Date where LGE_Plan_LP_Work_Order = LP_Work_Order and LPD_Input_Date between '"&dateadd("d",-30,date())&"' and '"&dateadd("d",30,date())&"') order by LP_Work_Order asc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strLP_Work_Order = strLP_Work_Order & RS1("LP_Work_Order") & ":" & RS1("LP_Work_Order") & ";"
	RS1.MoveNext
loop
RS1.Close
strLP_Work_Order	= left(strLP_Work_Order,len(strLP_Work_Order)-1)

strLPE_Code			= "slt>"
SQL = "select distinct LPE_Code = LPE_Type+'_'+convert(varchar,LPE_Code) from tbLGE_Plan_ETC order by LPE_Code"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strLPE_Code = strLPE_Code & RS1("LPE_Code") & ":" & RS1("LPE_Code") & ";"
	RS1.MoveNext
loop
RS1.Close
if trim(strLPE_Code) = "slt>" then
	strLPE_Code = "txt"
else
	strLPE_Code	= left(strLPE_Code,len(strLPE_Code)-1) 
end if

set RS1 = nothing

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,,,,,,,"
else
	strEdit	= ",dn2,dt1,num,,,"&strLP_Work_Order&","&strLPE_Code
end if
strPopup			= ",,,,,,,,,,,,,,,,,,,,"
strDown				= ",,,,,,,,,,,,,,,,,,,,"
strAlign			= ""
strAlign			= strAlign		& "Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if Len(Request("s_PS_Send_Date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PS_Send_Date between ''"&Left(Request("s_PS_Send_Date"),10)&"'' and ''"&Right(Request("s_PS_Send_Date"),10)&"''"
end if

if Request("s_BOM_Sub_BS_D_No") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "BOM_Sub_BS_D_No like ''%"&Request("s_BOM_Sub_BS_D_No")&"%''"
end If

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "ps_code"
	S_Order_By_2 	= "desc"
end if

strID				= "ps_code"
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
strColumn		= "s_BOM_Sub_BS_D_No,s_PS_Send_Date,s_edit_mode_yn"
strName			= "파트넘버,출발일,수정모드"
strType			= "dn2,dt2,chk"
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
	strReg	= ",dn2,dt1,num,,,"&strLP_Work_Order&","&strLPE_Code
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
	var strError_1	= "BOM_Sub_BS_D_No,PS_Send_Date,PS_Qty"
	var strError_2	= "파트넘버,출발일,수량"
	var strError_3	= "txt,txt,num";
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
	var strError_1	= "BOM_Sub_BS_D_No,PS_Send_Date,PS_Qty"
	var strError_2	= "파트넘버,출발일,수량"
	var strError_3	= "txt,txt,num";
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
		<tr>
<%
if Request("s_edit_mode_yn") <> "" and instr(admin_ps_list,"-"&Request.Cookies("Admin")("M_ID")&"-") > 0 then
%>
			<td width=5px></td>
			<td width=77px><%=Make_BTN("수정완료","javascript:List_Update()","")%></td>
<%
end if
%>

<%
if instr(admin_ps_list,"-"&Request.Cookies("ADMIN")("M_ID")&"-") > 0 then
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