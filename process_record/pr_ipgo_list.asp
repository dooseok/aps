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
dim S_PR_Process
S_PR_Process = Request("S_PR_Process")
if S_PR_Process = "" then
	S_PR_Process = "MAN"
end if
if instr("-IMD-SMD-",S_PR_Process) > 0 then
	strSelectName		= strSelectName & "번호,작업날짜,제번,라인,직접원,간접원,작업구분,파트넘버,계획,양품,불량,<계획,계획>,<실적,실적>,LOSS,비고,소요,점수,점수계,삭제"
elseif instr("-MAN-",S_PR_Process) > 0 then
	strSelectName		= strSelectName & "번호,작업날짜,제번,라인,직접원,간접원,작업구분,파트넘버,계획,양품,불량,<계획,계획>,<실적,실적>,LOSS,비고,소요,S/T,삭제"
elseif instr("-ASM-DLV-",S_PR_Process) > 0 then
	strSelectName		= strSelectName & "번호,작업날짜,제번,라인,직접원,간접원,작업구분,파트넘버,계획,양품,불량,<계획,계획>,<실적,실적>,LOSS,비고,소요,삭제"
end if

strWidth			= ""

if instr("-IMD-SMD-",S_PR_Process) > 0 then
	strWidth			= strWidth		& "40,70,80,50,50,50,70,90,40,40,40,50,50,50,50,50,200,50,50,50,50"
elseif instr("-MAN-",S_PR_Process) > 0 then
	strWidth			= strWidth		& "40,70,80,50,50,50,70,90,40,40,40,50,50,50,50,50,200,50,50,50"
elseif instr("-ASM-DLV-",S_PR_Process) > 0 then
	strWidth			= strWidth		& "40,70,80,50,50,50,70,90,40,40,40,50,50,50,50,50,200,50,50"
end if
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
URL_This			= "/process_record/pr_list.asp"
URL_View			= "/process_record/pr_delete_action.asp"
URL_Action			= "/process_record/pr_list_action.asp"
URL_Reg				= "/process_record/pr_reg_action.asp"

strTable			= "vwPR_List"
strPK				= "PR_Code"
strSelect			= ""
if instr("-IMD-SMD-",S_PR_Process) > 0 then
	strSelect			= strSelect		& "PR_Code,PR_Work_Date,PR_Work_Order,PR_Line,PR_Worker_CNT,PR_Supporter_CNT,PR_WorkType,BOM_Sub_BS_D_No,PR_Plan_Amount,PR_Amount,PR_Amount_NG,PR_Plan_Start_Time,PR_Plan_End_Time,PR_Start_Time,PR_End_Time,PR_Loss_Time,PR_Memo,PR_Time_Diff,PR_Point,PR_Calc_Point"
elseif instr("-MAN-",S_PR_Process) > 0 then	
	strSelect			= strSelect		& "PR_Code,PR_Work_Date,PR_Work_Order,PR_Line,PR_Worker_CNT,PR_Supporter_CNT,PR_WorkType,BOM_Sub_BS_D_No,PR_Plan_Amount,PR_Amount,PR_Amount_NG,PR_Plan_Start_Time,PR_Plan_End_Time,PR_Start_Time,PR_End_Time,PR_Loss_Time,PR_Memo,PR_Time_Diff,PR_ST"
elseif instr("-ASM-DLV-",S_PR_Process) > 0 then	
	strSelect			= strSelect		& "PR_Code,PR_Work_Date,PR_Work_Order,PR_Line,PR_Worker_CNT,PR_Supporter_CNT,PR_WorkType,BOM_Sub_BS_D_No,PR_Plan_Amount,PR_Amount,PR_Amount_NG,PR_Plan_Start_Time,PR_Plan_End_Time,PR_Start_Time,PR_End_Time,PR_Loss_Time,PR_Memo,PR_Time_Diff"
end if

'Call WorkOrder_Guide()

select case S_PR_Process
case "IMD"
	Call BOM_Guide()
case "SMD"
	Call BOMSub_Guide()
case "MAN"
	Call BOMSub_Guide()
case "ASM"
	Call BOMSub_Guide()
case "DLV"
	Call BOMSub_Guide()
end select

dim RS1
dim SQL

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,,,,,,,,"
else
	select case S_PR_Process
	case "IMD"	
		strEdit	= ",dt1,w/o,"&BasicDataIMDLine&",txt,txt,"&BasicDataFAWorkType&",dn1,,txt,txt,,,txt,txt,txt0,txt,,,,,"	
	case "SMD"
		strEdit	= ",dt1,w/o,"&BasicDataSMDLine&",txt,txt,"&BasicDataFAWorkType&",dn2,,txt,txt,,,txt,txt,txt0,txt,,,,,"
	case "MAN"
		strEdit	= ",dt1,w/o,"&BasicDataMANLine&",txt,txt,"&BasicDataMANWorkType&",dn2,,txt,txt,,,txt,txt,txt0,txt,,,,"
	case "ASM"
		strEdit	= ",dt1,w/o,"&BasicDataASMLine&",txt,txt,"&BasicDataMANWorkType&",dn2,,txt,txt,,,txt,txt,txt0,txt,,,"
	case "DLV"
		strEdit	= ",dt1,w/o,"&BasicDataDLVLine&",txt,txt,"&BasicDataDLVWorkType&",dn2,,txt,txt,,,txt,txt,txt0,txt,,,"
	end select
end if
strPopup			= ",,,,,,,,,,,,,,,,,,,,"
strDown				= ",,,,,,,,,,,,,,,,,,,,"
strAlign			= "Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if S_PR_Process <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PR_Process = ''"&S_PR_Process&"''"
end If

if Request("s_pr_work_order") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "pr_work_order like ''%"&Request("s_pr_work_order")&"%''"
end If

if Request("s_BOM_Sub_BS_D_No") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "BOM_Sub_BS_D_No = ''"&Request("s_BOM_Sub_BS_D_No")&"''"
end If

if Request("s_PR_Work_Date") = "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PR_Work_Date = ''"&date()&"''"
end If

if Request("s_PR_Work_Date") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PR_Work_Date = ''"&Request("s_PR_Work_Date")&"''"
end If

if Request("s_pr_worktype") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "pr_worktype = ''"&Request("s_pr_worktype")&"''"
end If
if Request("s_pr_line") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "pr_line = ''"&Request("s_pr_line")&"''"
end If

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "pr_work_date"
	S_Order_By_2 	= "asc"
	S_Order_By_3 	= "pr_start_time"
	S_Order_By_4 	= "asc"
end if

strID				= "pr_code"
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
strColumn		= "s_pr_process,s_pr_work_order,s_BOM_Sub_BS_D_No,s_pr_work_date,s_pr_line,s_pr_worktype,s_edit_mode_yn"
strName			= "공정,제번,파트넘버,작업일,라인,작업구분,수정모드"
select case S_PR_Process
case "IMD"
	strType	= BasicDataProcess&",txt,dn1,dt1"&date()&","&BasicDataIMDLine&","&BasicDataFAWorkType&",chk"
case "SMD"
	strType	= BasicDataProcess&",txt,dn2,dt1"&date()&","&BasicDataSMDLine&","&BasicDataFAWorkType&",chk"
case "MAN"
	strType	= BasicDataProcess&",txt,dn2,dt1"&date()&","&BasicDataMANLine&","&BasicDataManWorkType&",chk"
case "ASM"
	strType	= BasicDataProcess&",txt,dn2,dt1"&date()&","&BasicDataASMLine&","&BasicDataManWorkType&",chk"
case "DLV"
	strType	= BasicDataProcess&",txt,dn2,dt1"&date()&","&BasicDataDLVLine&","&BasicDataDLVWorkType&",chk"
end select
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

dim DefaultDate
if Request("s_PR_Work_Date") <> "" then
	DefaultDate = Request("s_PR_Work_Date")
else
	DefaultDate = date()
end if

dim DefaultLine
if Request("s_pr_line") <> "" then
	DefaultLine = Request("s_pr_line")
else
	DefaultLine = ""
end if

select case S_PR_Process
case "IMD"
	strReg	= ",dt1"&DefaultDate&",w/o,"&BasicDataIMDLine&"/d"&DefaultLine&",txt1,txt0,"&BasicDataFAWorkType&"/d작업,dn1,,txt,txt0,,,txt,txt,txt0,txt,,,,"
case "SMD"
	strReg	= ",dt1"&DefaultDate&",w/o,"&BasicDataSMDLine&"/d"&DefaultLine&",txt1,txt0,"&BasicDataFAWorkType&"/d작업,dn2,,txt,txt0,,,txt,txt,txt0,txt,,,,,"
case "MAN"
	strReg	= ",dt1"&DefaultDate&",w/o,"&BasicDataMANLine&"/d"&DefaultLine&",txt5,txt,"&BasicDataMANWorkType&"/d작업,dn2,,txt,txt0,,,txt,txt,txt0,txt,,,,"
case "ASM"
	strReg	= ",dt1"&DefaultDate&",w/o,"&BasicDataASMLine&"/d"&DefaultLine&",txt5,txt,"&BasicDataMANWorkType&"/d작업,dn2,,txt,txt0,,,txt,txt,txt0,txt,,,"
case "DLV"
	strReg	= ",dt1"&DefaultDate&",w/o,"&BasicDataDLVLine&"/d"&DefaultLine&",txt1,txt0,"&BasicDataDLVWorkType&"/d작업,dn2,,txt,txt0,,,txt,txt,txt0,txt,,,"
end select
	
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
if instr("-IMD-SMD-",S_PR_Process) > 0 then
	SQL = "select sum(PR_Calc_Point) from "&strTable&" "
	
	if trim(strWhere) <> "" then
		SQL = SQL & " where " & strWhere
	end if
	
	call Common_Display_Summary("점수 총계 : |Sum_String|", SQL, "N", strWidth_Total)
elseif instr("-DLV-",S_PR_Process) > 0 then
	SQL = "select sum(PR_Price) from "&strTable&" "
	
	
	if trim(strWhere) <> "" then
		SQL = SQL & " where " & strWhere
	end if
	
	call Common_Display_Summary("납품금액 총계 : |Sum_String|", SQL, "Y", strWidth_Total)
end if

if instr("-DLV-",S_PR_Process) > 0 then
	SQL = "select sum(PR_Amount) from "&strTable&" "
	
	if trim(strWhere) <> "" then
		SQL = SQL & " where " & strWhere
	end if
	
	call Common_Display_Summary("납품수량 총계 : |Sum_String|", SQL, "N", strWidth_Total)
else
	SQL = "select sum(PR_Amount) from "&strTable&" "
	
	if trim(strWhere) <> "" then
		SQL = SQL & " where " & strWhere
	end if
	
	call Common_Display_Summary("생산수량 총계 : |Sum_String|", SQL, "N", strWidth_Total)
end if
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
	var strError = List_Reg_Validater('PR_Work_Date,PR_Line,PR_Worker_CNT,PR_WorkType,BOM_Sub_BS_D_No,PR_Amount,PR_Amount_NG,PR_Start_Time,PR_End_Time','작업날짜,라인,작업인원,작업구분,파트넘버,양품수,불량수,시작시각,종료시각','txt,txt,num,txt,txt,num,num,fit4,fit4');
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
	var strError = List_Validater('PR_Work_Date,PR_Line,PR_Worker_CNT,PR_WorkType,BOM_Sub_BS_D_No,PR_Amount,PR_Amount_NG,PR_Start_Time,PR_End_Time','작업날짜,라인,작업인원,작업구분,파트넘버,양품수,불량수,시작시각,종료시각','txt,txt,num,txt,txt,num,num,fit4,fit4');
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
<img src="/img/blank.gif" width=1px height=10px><br>
<img src="/img/bgTimeGuide.gif"><br>
<img src="/img/blank.gif" width=1px height=5px><br>
<%
'----------------------------------------------------------------------------------
end sub
%>
<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->