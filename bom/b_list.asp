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
if instr(admin_b_list,"-"&gM_ID&"-") > 0 then
	strSelectName		= "체크,번호,파트넘버,옵션수,시방번호,현재적용,시방적용일,등록일,비 매칭자재,작업"
else
	strSelectName		= "체크,번호,파트넘버,옵션수,시방번호,현재적용,시방적용일,등록일,비 매칭자재"
end if
strWidth			= "50,70,100,80,80,80,100,100,90,80"
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
URL_This			= "/bom/b_list.asp"
URL_View			= "/bom/b_edit_form.asp"
URL_Action			= "/bom/b_list_action.asp"
URL_Reg				= "/bom/b_reg_form.asp"

strTable			= "vwB_List"
strPK				= "B_Code"
strSelect			= "B_Code,B_D_No,Bom_Sub_Cnt,B_Version_Code,B_Version_Current_YN,B_Version_Date,B_Issue_Date,cntP_P_No2_Miss"

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,,,,"
else
	strEdit	= ",,,,,,dt1,txt,txt,txt,,,,,,"
end if

if gM_ID="shindk" then
	strPopup			= ",db_load_action.asp,,,,,,,,,,,,,,,,"
else
	strPopup			= ",db_load_action.asp,,,,,,,,,,,,,,,,"
end if
strDown				= ",,,,,,,,,,,,,,,"
strAlign			= "Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
'if Trim(strWhere) <> "" Then
'	strWhere = strWhere & " and "
'end if
'strWhere = strWhere & " (B_Version_Code <> ''devtemp'') "

if Request("s_show_old") = "" then
	'strWhere = strWhere & "B_Current_YN = ''Y''"
end if

if Request("s_bom_b_d_no") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "	(B_Code in "
	strWhere = strWhere & "		(select "
	strWhere = strWhere & "			BOM_B_Code "
	strWhere = strWhere & "		from "
	strWhere = strWhere & "			tbBOM "
	strWhere = strWhere & "			left outer join "
	strWhere = strWhere & "			tbBOM_Sub "
	strWhere = strWhere & "			on B_Code = BOM_B_Code "
	strWhere = strWhere & "		where "
	strWhere = strWhere & "			"
	strWhere = strWhere & "			(Bs_D_No) like ''%"&Request("s_bom_b_d_no")&"%'') or "
	strWhere = strWhere & " b_d_no like ''%"&Request("s_bom_b_d_no")&"%'')"
	 
end If
if Request("s_parts_p_no_current") <> "" and Request("s_b_version_current_yn") <> "N" Then '현재적용중인 BOM기준 부품만 조회하는 경우
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "B_Code in (select distinct BOM_B_Code from tbBOM_Sub where "
	strWhere = strWhere & "			BS_Code in (select distinct BOM_Sub_BS_Code from tbBOM_Qty where Parts_P_P_No = ''"&Request("s_parts_p_no_current")&"''))"
end if

if Request("s_parts_p_no_archive") <> "" and Request("s_b_version_current_yn") <> "Y" Then 'Archive BOM기준 부품만 조회하는 경우
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "B_Code in (select distinct BOM_B_Code from tbBOM_Sub where "
	strWhere = strWhere & "			BS_Code in (select distinct BOM_Sub_BS_Code from tbBOM_Qty_Archive where Parts_P_P_No = ''"&Request("s_parts_p_no_archive")&"''))"
end if

if Request("s_parts_p_no_all") <> "" Then '모든 BOM기준 부품만 조회하는 경우
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	if Request("s_b_version_current_yn") = "Y" then
		strWhere = strWhere & "B_Code in (select distinct BOM_B_Code from tbBOM_Sub where "
		strWhere = strWhere & "			BS_Code in (select distinct BOM_Sub_BS_Code from tbBOM_Qty where Parts_P_P_No = ''"&Request("s_parts_p_no_all")&"''))"
	elseif Request("s_b_version_current_yn") = "N" then
		strWhere = strWhere & "B_Code in (select distinct BOM_B_Code from tbBOM_Sub where "
		strWhere = strWhere & "			BS_Code in (select distinct BOM_Sub_BS_Code from tbBOM_Qty_Archive where Parts_P_P_No = ''"&Request("s_parts_p_no_all")&"''))"
	else
		strWhere = strWhere & "B_Code in (select distinct BOM_B_Code from tbBOM_Sub where "
		strWhere = strWhere & "			BS_Code in (select distinct BOM_Sub_BS_Code from tbBOM_Qty where Parts_P_P_No = ''"&Request("s_parts_p_no_all")&"'') "
		strWhere = strWhere & "			or "
		strWhere = strWhere & "			BS_Code in (select distinct BOM_Sub_BS_Code from tbBOM_Qty_Archive where Parts_P_P_No = ''"&Request("s_parts_p_no_all")&"'') "
		strWhere = strWhere & ") "
	end if
end if

if Request("s_b_version_code") <> "" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "b_version_code like ''%"&Request("s_b_version_code")&"%''"
end if

if Request("s_b_desc_spec") <> "" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "(b_tool like ''%"&Request("s_b_desc_spec")&"%'' or b_desc like ''%"&Request("s_b_desc_spec")&"%'' or b_spec like ''%"&Request("s_b_desc_spec")&"%'')"
end if

if Request("s_b_version_current_yn") <> "" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "b_version_current_yn = ''"&Request("s_b_version_current_yn")&"''"
end if

if Len(Request("s_b_issue_date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "b_issue_date between ''"&Left(Request("s_b_issue_date"),10)&"'' and ''"&Right(Request("s_b_issue_date"),10)&"''"
end if

if Len(Request("s_b_version_date")) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "b_version_date between ''"&Left(Request("s_b_version_date"),10)&"'' and ''"&Right(Request("s_b_version_date"),10)&"''"
end if

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "b_code"
	S_Order_By_2 	= "desc"
end if

strID				= "b_code"
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
call Parts_Guide()

strColumn		= "s_bom_b_d_no,s_b_issue_date,s_b_version_date|/|s_b_version_code,s_b_version_current_yn|/|s_parts_p_no_current,s_parts_p_no_archive,s_parts_p_no_all"
strName			= "도면품번,등록일,시방적용일|/|시방번호,현재적용|/|부품-현재적용,부품-미적용,부품-전체"
strType			= "dn1,dt2,dt2|/|txt,slt>Y:Y;N:N|/|pno,pno,pno"
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
if Reg_Form_YN = "N" then
'6/9
'----------------------------------------------------------------------------------	
	strReg	= ",,txt,txt,"&BasicDataPart&","&BasicDataPosition&",txt,txt,txt,txt,txt,txt,txt,dt1,,mem"
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
	var strError = List_Validater('B_Issue_Date','등록일','txt');
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
'if Request("s_edit_mode_yn") = "" then
%>
<%
'elseif instr(admin_b_list,"-"&Request.Cookies("Admin")("M_Authority")&"-") > 0 then
%>
			<!--<td width=5px></td>
			<td width=77px><%=Make_BTN("수정완료","javascript:List_Update()","")%></td>-->
<%
'else
%><!--
			<td width=5px></td>
			<td width=77px><%=Make_BTN("<strike>Submit</strike>","javascript:alert('"&Request.Cookies("Admin")("M_Name")&"님의 권한은 ["&Request.Cookies("Admin")("M_Authority")&"]입니다.\n작업가능한 권한은 ["&admin_b_list&"]입니다.\n문의사항이 있으시면 전산담당자에게 문의하여주십시오.');","")%></td>
--><%
'end if
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
			
<%
	'if instr(admin_b_list,"-"&Request.Cookies("ADMIN")("M_Authority")&"-") > 0 then
%>
			
<%
'else
%><!--
			<td width=100px><%=Make_L_BTN("<strike>Registration</strike>","javascript:alert('"&Request.Cookies("Admin")("M_Name")&"님의 권한은 ["&Request.Cookies("Admin")("M_Authority")&"]입니다.\n작업가능한 권한은 ["&admin_b_list&"]입니다.\n문의사항이 있으시면 전산담당자에게 문의하여주십시오.');","")%></td>
--><%
'end if
%>
			
<%
end if
%>		
			<td width=77px><%=Make_BTN("신규등록","","b_reg_form.asp?NEW_YN=Y")%></td>
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