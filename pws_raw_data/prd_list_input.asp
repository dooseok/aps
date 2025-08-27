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
strSelectName		= strSelectName & "번호,투입일,생산라인,파트넘버,수량"
strWidth			= ""
strWidth			= strWidth		& "60,100,100,120,60"
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
URL_This			= "/pws_raw_data/prd_list_Input.asp"
URL_View			= "/pws_raw_data/prd_edit_form.asp"
URL_Action			= "/pws_raw_data/prd_list_action.asp"
URL_Reg				= "/pws_raw_data/prd_reg_action.asp"

strTable			= "vwPRD_list_Input"
strPK				= "minPRD_Code"
strSelect			= ""
strSelect			= strSelect		& "minPRD_Code, PRD_Input_Date, PRD_Line, PRD_PartNo, cntPRD_Code"

dim RS1
dim SQL
dim strEdit_Partner

if Request("s_edit_mode_yn") = "" then
	strEdit = ",,,,,,,,,,,,,,,"
else
	strEdit	= ",,,,,,,,,,,,,,,"
	
end if
strPopup			= ",,,,,,,,,,,,,,,"
strDown				= ",,,,,,,,,,,,,,,"
strAlign			= ""
strAlign			= strAlign		& "Center,Center,Center,Center,Center,Center,right,Center,Center,Center,Center,Center,Center,Center,Center"
'----------------------------------------------------------------------------------

'3/9
'----------------------------------------------------------------------------------
if Request("s_PRD_PartNo") <> "" then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PRD_PartNo like ''%"&Request("s_PRD_PartNo")&"%''"
end If
dim s_PRD_Input_Date
s_PRD_Input_Date = Request("s_PRD_Input_Date")
if Len(s_PRD_Input_Date) <> "22"  then
	s_PRD_Input_Date = date()&", "&date()
end if

if Len(s_PRD_Input_Date) = "22" Then
	If Trim(strWhere) <> "" Then
		strWhere = strWhere & " and "
	End If
	strWhere = strWhere & "PRD_Input_Date between ''"&Left(s_PRD_Input_Date,10)&"'' and ''"&Right(s_PRD_Input_Date,10)&"''"
end If

if S_Order_By_1 & S_Order_By_2 = "" then
	S_Order_By_1 	= "PRD_Input_Date"
	S_Order_By_2 	= "desc"
end if

strID				= "minPRD_Code"
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

strColumn		= "s_PRD_PartNo,s_PRD_Input_Date"
strName			= "파트넘버,투입일"
strType			= "dn2,dt2"
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
	strReg	= ",,,,,,,,,,"
'----------------------------------------------------------------------------------	
	call inc_Common_List_Reg_Form(URL_Reg, Colspan, strRequestQueryString, strSelect, arrRecordSet, strWidth, strReg, strAlign, strWidth_Total, 1)
end if
%>

<%
'call Price_Memo
'sub Price_Memo
%>
<!--<img src="/img/blank.gif" width=1px height=10px><br>
<form name="frmMemo action="" method="post">
<textarea cols=60 rows=3></textarea>
<input type="submit" value="등록">
</form>-->
<%
'end sub
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
	var strError_1	= "BP_Type,BP_LGE_Staff,BP_MSE_Staff,BP_Desc";
	var strError_2	= "구분,LGE합의자,MSE합의자,비고";
	var strError_3	= "txt,txt,txt,txt";
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

	var strError_1	= "BP_LGE_Staff,BP_MSE_Staff,BP_Desc";
	var strError_2	= "LGE합의자,MSE합의자,비고";
	var strError_3	= "txt,txt,txt";

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

function frmFileUpload_Check()
{
	if(!frmFileUpload.strFile.value)
	{
		alert("파일을 선택해주세요.")
		return false;
	}
	else
	{
		Show_Progress();
		//if(confirm("업로드가 완료되면, 완료 메세지가 보입니다.\n완료 메세지가 나올때까지 작업을 중지하고 기다려주십시오."))
		frmFileUpload.submit();
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
		<form name="frmFileUpload" action="lr_xls_upload_action.asp" method="post" enctype="MULTIPART/FORM-DATA">
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
		</form>
		<form name="frmList2Excel" action="/function/inc_List2Excel.asp" method="post" target="_blank">
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

<Script language="javascript">
	if(frmSearch_Bar.s_PRD_Input_Date[0].value == "" || frmSearch_Bar.s_PRD_Input_Date[1].value == "")
	{
		frmSearch_Bar.s_PRD_Input_Date[0].value="<%=date()%>";
		frmSearch_Bar.s_PRD_Input_Date[1].value="<%=date()%>";
	}
</script>
<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->