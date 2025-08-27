<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim RS1
dim RS2
dim SQL
dim CNT1
dim CNT2

dim s_Min_LPD_Input_Date
dim s_Diff_LPD_Input_Date
dim s_LM_Company
dim s_Edit_Process

dim Max_LPD_Input_Date

dim LM_Company
dim LP_Code
dim LP_Line
dim LP_Work_Order
dim LP_Model
dim LP_Suffix
dim LP_Tool
dim LP_Tool_Type
dim LP_Input_Time
dim LP_Buyer
dim LP_LOT
dim LP_LOT_Remain

dim LP_LOT_Sum
dim LP_LOT_Remain_Sum
dim LPD_Input_Qty_Sum(100)

dim BOM_Sub_BS_D_No
dim old_BOM_Sub_BS_D_No

dim BOM_Sub_BS_D_No_CNT
dim BOM_Sub_BS_D_No_1
dim BOM_Sub_BS_D_No_2
dim BOM_Sub_BS_D_No_3
dim BOM_Sub_BS_D_No_4
dim BOM_Sub_BS_D_No_Str

dim BOM_B_Code_1
dim BOM_B_Code_2
dim BOM_B_Code_3
dim BOM_B_Code_4

dim LP_IMD_Complete_Qty
dim LP_SMD_Complete_Qty
dim LP_MAN_Complete_Qty
dim LP_ASM_Complete_Qty
dim LP_DLV_Complete_Qty

dim old_BOM_Sub_BS_D_No_Str

dim LPD_Input_Qty

dim arrPlan_Date
dim strDate_Offset
dim strLPD_Input_Date
dim strLPD_Input_Qty

dim Exist_YN

dim strInput_Cnt

dim S_Order_By_1
dim S_Order_By_2
dim S_Order_By_3
dim S_Order_By_4

dim strOrderBy

dim strBGColor
dim strFontColor_sum
dim strBGColor_sum

dim MPD_Qty
dim strQty

dim arrInputSelectG_1
dim arrInputSelect_1
dim arrInputSelectG_2
dim arrInputSelect_2

dim strDiff_MPD_Date
dim strMPD_Process
dim strMPD_Qty_Sum
dim arrDiff_MPD_Date
dim arrMPD_Process
dim arrMPD_Qty_Sum

dim bgMSE_Plan_Editor

dim strCall_MSE_Plan_Editor

S_Order_By_1 = Request("S_Order_By_1")
S_Order_By_2 = Request("S_Order_By_2")
S_Order_By_3 = Request("S_Order_By_3")
S_Order_By_4 = Request("S_Order_By_4")

s_Min_LPD_Input_Date	= Request("s_Min_LPD_Input_Date")
s_Diff_LPD_Input_Date	= Request("s_Diff_LPD_Input_Date")
s_LM_Company			= Request("s_LM_Company")
s_Edit_Process			= Request("s_Edit_Process")

if s_Edit_Process <> "" then
	if S_Order_By_1 & S_Order_By_2 = "" then
		S_Order_By_1 	= "LP_Tool_Type"
		S_Order_By_2 	= "asc"
		S_Order_By_3 	= "BOM_Sub_BS_D_No"
		S_Order_By_4 	= "asc"
	end if
else
	if S_Order_By_1 & S_Order_By_2 = "" then
		S_Order_By_1 	= "LP_Line"
		S_Order_By_2 	= "asc"
		S_Order_By_3 	= "Min_LPD_Input_Date"
		S_Order_By_4 	= "asc"
	end if
end if

if S_Order_By_3 = "" then
	strOrderBy			= S_Order_By_1&" "&S_Order_By_2
else
	strOrderBy			= S_Order_By_1&" "&S_Order_By_2&", "&S_Order_By_3&" "&S_Order_By_4
end if
	
strFontColor_sum	= "red"
strBGColor_Sum		= "white"

if s_Min_LPD_Input_Date = "" then
	s_Min_LPD_Input_Date = date()
end if

if s_Diff_LPD_Input_Date = "" then
	s_Diff_LPD_Input_Date = 13
end if

if s_LM_Company = "" then
	s_LM_Company = "MSE"
end if

if s_Edit_Process = "" then
	if Request.Cookies("Admin")("M_Part") = "제조1" then
		s_Edit_Process = "MAN"
	elseif Request.Cookies("Admin")("M_Part") = "제조2" then
		s_Edit_Process = "SMD"
	end if
end if

dim LPD_Input_Qty_TR_Sum

Max_LPD_Input_Date = dateadd("d",s_Diff_LPD_Input_Date,s_Min_LPD_Input_Date)

s_Min_LPD_Input_Date	= CDate(s_Min_LPD_Input_Date)
Max_LPD_Input_Date		= CDate(Max_LPD_Input_Date)
%>

<%
dim strRequestQueryString
dim Request_Fields

strRequestQueryString = ""
for each Request_Fields in Request.QueryString
	if lcase(left(Request_Fields,2))="s_" then
		strRequestQueryString = strRequestQueryString & "&"&Request_Fields&"="&server.URLEncode(Request(Request_Fields))
	end if
next
for each Request_Fields in Request.Form
	if lcase(left(Request_Fields,2))="s_" then
		strRequestQueryString = strRequestQueryString & "&"&Request_Fields&"="&server.URLEncode(Request(Request_Fields))
	end if
next

dim strRequestQueryString_dummy
strRequestQueryString_dummy = strRequestQueryString
strRequestQueryString_dummy = replace(strRequestQueryString_dummy,"S_Order_By_1=","Dummy_Order_By_1=")
strRequestQueryString_dummy = replace(strRequestQueryString_dummy,"S_Order_By_2=","Dummy_Order_By_2=")
strRequestQueryString_dummy = replace(strRequestQueryString_dummy,"S_Order_By_3=","Dummy_Order_By_3=")
strRequestQueryString_dummy = replace(strRequestQueryString_dummy,"S_Order_By_4=","Dummy_Order_By_4=")
%>

<%
if Request("menucall_YN") = "Y" then
	SQL = "select top 0 * from ("
else
	SQL = "select * from ("
end if
SQL = SQL & "select "&vbcrlf
SQL = SQL & "	LP_Code = LPE_Code, "&vbcrlf
SQL = SQL & "	LM_Company = 'MSE', "&vbcrlf
SQL = SQL & "	LP_Line = '', "&vbcrlf
SQL = SQL & "	LP_Work_Order = LPE_Type+'_'+convert(varchar,LPE_Code), "&vbcrlf
SQL = SQL & "	LP_Model = '', "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No, "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No_1	= BOM_Sub_BS_D_No, "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No_2	= '', "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No_3	= '', "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No_4	= '', "&vbcrlf
SQL = SQL & "	BOM_B_Code_1 = (select st.BOM_B_Code from tbBOM_Sub st where st.BS_D_No = BOM_Sub_BS_D_No and exists (select B_Code from tbBOM where B_Current_YN='Y' and B_Code = BOM_B_Code)), "&vbcrlf
SQL = SQL & "	BOM_B_Code_2 = '', "&vbcrlf
SQL = SQL & "	BOM_B_Code_3 = '', "&vbcrlf
SQL = SQL & "	BOM_B_Code_4 = '', "&vbcrlf
SQL = SQL & "	LP_IMD_Complete_Qty = LPE_IMD_Complete_Qty, "&vbcrlf
SQL = SQL & "	LP_SMD_Complete_Qty = LPE_SMD_Complete_Qty, "&vbcrlf
SQL = SQL & "	LP_MAN_Complete_Qty = LPE_MAN_Complete_Qty, "&vbcrlf
SQL = SQL & "	LP_ASM_Complete_Qty = LPE_ASM_Complete_Qty, "&vbcrlf
SQL = SQL & "	LP_DLV_Complete_Qty = LPE_DLV_Complete_Qty, "&vbcrlf
SQL = SQL & "	LP_Suffix = '', "&vbcrlf
SQL = SQL & "	LP_Buyer = LPE_Buyer, "&vbcrlf
SQL = SQL & "	LP_Tool = '', "&vbcrlf
SQL = SQL & "	LP_Tool_Type = '', "&vbcrlf
SQL = SQL & "	LP_Input_Time = '', "&vbcrlf
SQL = SQL & "	LP_Lot = LPE_Req_Qty, "&vbcrlf
SQL = SQL & "	LP_Lot_Remain = LPE_Req_Qty, "&vbcrlf
for CNT1 = 0 to s_Diff_LPD_Input_Date
	SQL = SQL & "	Date_Qty_"&CNT1&" = case LPE_Due_Date when '"&dateadd("d",CNT1,s_Min_LPD_Input_Date)&"' then LPE_Req_Qty end, "&vbcrlf
next
SQL = SQL & "	Min_LPD_Input_Date = LPE_Due_Date "&vbcrlf
SQL = SQL & "from tbLGE_Plan_ETC "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	LPE_Due_Date between '"&s_Min_LPD_Input_Date&"' and '"&Max_LPD_Input_Date&"' "&vbcrlf
SQL = SQL & " "&vbcrlf
SQL = SQL & "union "&vbcrlf
SQL = SQL & " "&vbcrlf
SQL = SQL & "select "&vbcrlf
SQL = SQL & "	LP_Code, "&vbcrlf
SQL = SQL & "	LM_Company, "&vbcrlf
SQL = SQL & "	LP_Line, "&vbcrlf 
SQL = SQL & "	LP_Work_Order, "&vbcrlf 
SQL = SQL & "	LP_Model, "&vbcrlf 
SQL = SQL & "	BOM_Sub_BS_D_No = BOM_Sub_BS_D_No_1 + '<br>' + BOM_Sub_BS_D_No_2 + '<br>' + BOM_Sub_BS_D_No_3 + '<br>' + BOM_Sub_BS_D_No_4, "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No_1, "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No_2, "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No_3, "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No_4, "&vbcrlf
SQL = SQL & "	BOM_B_Code_1, "&vbcrlf
SQL = SQL & "	BOM_B_Code_2, "&vbcrlf
SQL = SQL & "	BOM_B_Code_3, "&vbcrlf
SQL = SQL & "	BOM_B_Code_4, "&vbcrlf
SQL = SQL & "	LP_IMD_Complete_Qty, "&vbcrlf
SQL = SQL & "	LP_SMD_Complete_Qty, "&vbcrlf
SQL = SQL & "	LP_MAN_Complete_Qty, "&vbcrlf
SQL = SQL & "	LP_ASM_Complete_Qty, "&vbcrlf
SQL = SQL & "	LP_DLV_Complete_Qty, "&vbcrlf
SQL = SQL & "	LP_Suffix , "&vbcrlf
SQL = SQL & "	LP_Buyer, "&vbcrlf
SQL = SQL & "	LP_Tool, "&vbcrlf
SQL = SQL & "	LP_Tool_Type = TI_Type, "&vbcrlf
SQL = SQL & "	LP_Input_Time, "&vbcrlf
SQL = SQL & "	LP_Lot, "&vbcrlf
SQL = SQL & "	LP_Lot_Remain, "&vbcrlf
for CNT1 = 0 to s_Diff_LPD_Input_Date
	SQL = SQL & "	Date_Qty_"&CNT1&" = (select sum(LPD_Input_Qty) from tbLGE_Plan_Date where LGE_Plan_LP_Work_Order=LP_Work_Order and LPD_Input_Date = '"&dateadd("d",CNT1,s_Min_LPD_Input_Date)&"'), "&vbcrlf
next
SQL = SQL & "	Min_LPD_Input_Date = (select min(LPD_Input_Date) from tbLGE_Plan_Date where LGE_Plan_LP_Work_Order=LP_Work_Order and '2007-12-26' <= LPD_Input_Date) "&vbcrlf
SQL = SQL & "from "&vbcrlf
SQL = SQL & "	tbLGE_Plan "&vbcrlf
SQL = SQL & "	left outer join "&vbcrlf
SQL = SQL & "	vwLM_List "&vbcrlf
SQL = SQL & "	on LP_Model = LM_Name "&vbcrlf
SQL = SQL & "	left outer join "&vbcrlf
SQL = SQL & "	tbTool_Info "&vbcrlf
SQL = SQL & "	on TI_Name = LP_Tool "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	exists (select LGE_Plan_LP_Work_Order from tbLGE_Plan_Date where LGE_Plan_LP_Work_Order=LP_Work_Order and LPD_Input_Date between '"&s_Min_LPD_Input_Date&"' and '"&Max_LPD_Input_Date&"') "&vbcrlf
SQL = SQL & ") tb "&vbcrlf
if s_LM_Company = "ALL" then
elseif s_LM_Company = "타사제외" then
	SQL = SQL & "where LM_Company <> '타사' "&vbcrlf'
else
	SQL = SQL & "where LM_Company = '"&s_LM_Company&"' "&vbcrlf'
end if
SQL = SQL & "order by " & strOrderBy &vbcrlf

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")

RS1.Open SQL,sys_DBCon
%>

<script language="javascript">
function frmDate_Search_Check()
{
	alert("확인을 클릭하신 후 잠시 기다려주시기 바랍니다.");
	Show_Progress();
	frmDate_Search.submit();
}

function frmFile_Upload_Check()
{
	if (!frmFile_Upload.strFile.value)
	{
		alert("파일을 선택해주세요.");
	}
	else
	{
		Show_Progress();
		frmFile_Upload.submit();
	}
}

function setSorting2(S_Order_By_1,S_Order_By_2,S_Order_By_3,S_Order_By_4)
{
	location.href="lp_view.asp?S_Order_By_1="+S_Order_By_1+"&S_Order_By_2="+S_Order_By_2+"&S_Order_By_3="+S_Order_By_3+"&S_Order_By_4="+S_Order_By_4+"<%=strRequestQueryString_dummy%>";
}

function setSorting(S_Order_By_1,S_Order_By_2)
{
	Show_Progress();
	if ("<%=S_Order_By_1%>"==S_Order_By_1)
	{
		location.href="lp_view.asp?S_Order_By_1="+S_Order_By_1+"&S_Order_By_2="+S_Order_By_2+"<%=strRequestQueryString_dummy%>";
	}
	else if ("<%=S_Order_By_1%>"=="")
	{
		location.href="lp_view.asp?S_Order_By_1="+S_Order_By_1+"&S_Order_By_2="+S_Order_By_2+"<%=strRequestQueryString_dummy%>";
	}
	else if (S_Order_By_1=="" && S_Order_By_2=="")
	{
		location.href="lp_view.asp?dummy=<%=strRequestQueryString_dummy%>";
	}
	else
	{
		location.href="lp_view.asp?S_Order_By_1="+S_Order_By_1+"&S_Order_By_2="+S_Order_By_2+"&S_Order_By_3=<%=S_Order_By_1%>&S_Order_By_4=<%=S_Order_By_2%><%=strRequestQueryString_dummy%>";
	}
}

function List2Excel()
{
	frmList2Excel.submit();
}
</script>

<table border=0 cellspacing=1 cellpadding=0 width=1000px bgcolor="#999999" align=center>
<form name="frmDate_Search" action="lp_view.asp" method="post">
<tr height=25px>
	<td bgcolor=white>
		<table border=0 cellspacing=2 cellpadding=0 width=100% bgcolor="#ffffff">
		<tr>
			<td width=5px>&nbsp;</td>
			<td width=30px align=right>기간</td>
			<td width=180px align=center>
				<input type="text" name="s_Min_LPD_Input_Date" size=10 class="input" readonly value="<%=s_Min_LPD_Input_Date%>" onclick="Calendar_D(document.frmDate_Search.s_Min_LPD_Input_Date);">
				부터
				<select name="s_diff_LPD_Input_Date">
<%
for CNT1 = 1 to 30
%>
				<option value="<%=CNT1%>"<%if int(s_diff_LPD_Input_Date)=CNT1 then%> selected<%end if%>><%=CNT1+1%></option>
<%
next
%>
				</select>일간
			</td>
			<td width=5px></td>
			<td width=40px align=right>업체</td>
			<td width=60px align=left>
				<select name="s_LM_Company">
				<option value="ALL"<%if s_LM_Company="ALL" then%> selected<%end if%>>-전체-</option>
				<option value="MSE"<%if s_LM_Company="MSE" then%> selected<%end if%>>MSE</option>
				<option value="타사"<%if s_LM_Company="타사" then%> selected<%end if%>>타사</option>
				<option value="타사제외"<%if s_LM_Company="타사제외" then%> selected<%end if%>>타사제외</option>
				<option value="미분류"<%if s_LM_Company="미분류" then%> selected<%end if%>>미분류</option>
				</select>
			</td>
			<td width=70px align=right>공정</td>
			<td width=50px align=left>
				<select name="s_Edit_Process">
				<option value=""<%if s_Edit_Process="" then%> selected<%end if%>>-선택-</option>
				<option value="IMD"<%if s_Edit_Process="IMD" then%> selected<%end if%>>IMD</option>
				<option value="SMD"<%if s_Edit_Process="SMD" then%> selected<%end if%>>SMD</option>
				<option value="MAN"<%if s_Edit_Process="MAN" then%> selected<%end if%>>MAN</option>
				<option value="ASM"<%if s_Edit_Process="ASM" then%> selected<%end if%>>ASM</option>
				</select>
			</td>
			<td width=50px><%=Make_S_BTN("조회","javascript:frmDate_Search_Check();","")%></td>
			<td width=77px><%=Make_BTN("EXCEL보기","List2Excel()","")%></td>
			<td><td>
			<td width=60px>SCS파일</td>
		</form>
		<form name="frmFile_Upload" action="lp_upload_action.asp" method="post" enctype="MULTIPART/FORM-DATA">
			<td width=200px><input type="file" name="strFile" style="width:95%" class="input"></td>
			<td width=77px><%=Make_BTN("파일업로드","javascript:frmFile_Upload_Check();","")%></td>
			<td width=5px>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</form>
<form name="frmList2Excel" action="/function/inc_List2Excel_LP_View.asp" method="post" target="_blank">
<input type="hidden"	name="SQL"						value="<%=SQL%>">
<input type="hidden"	name="s_diff_LPD_Input_Date"	value="<%=s_diff_LPD_Input_Date%>">
<input type="hidden"	name="s_Min_LPD_Input_Date"		value="<%=s_Min_LPD_Input_Date%>">
</form>
<tr height=25px>
	<td bgcolor=white>
		<table width=100% cellpadding=0 cellspacing=0 border=0 bgcolor="#999999">
		<form name="frmLP_View" action="lp_view_change_action.asp?dummy=<%=strRequestQueryString%>" method="post">
		<input type="hidden" name="Min_LPD_Input_Date" value="<%=s_Min_LPD_Input_Date%>">
		<input type="hidden" name="Max_LPD_Input_Date" value="<%=Max_LPD_Input_Date%>">
		<script language="javascript">
		function frmLP_View_Submit_Check(strType)
		{
			if(strType == "Change")
			{
				if(!frmLP_View.Company.value)
				{
					alert("변경할 업체를 선택해주세요.");
					return false;
				}
				if(confirm("선택한 모델을 전부 ["+frmLP_View.Company.value+"]로 변경하시겠습니까?"))
				{
					frmLP_View.action = "lp_view_change_action.asp?dummy=<%=strRequestQueryString%>"
					frmLP_View.submit();
				}
			}
			else if(strType == "Check")
			{
				frmLP_View.action = "lp_view_check_action.asp?dummy=<%=strRequestQueryString%>"
				frmLP_View.submit();
			}
		}
		</script>
		<tr bgcolor="white">
			<td width=150px align=left>
				<img src="/img/blank.gif" width=11px height=1px>
				일괄변경
				<select name="Company" style="font-size:10px;font-family:arial,돋움;color:#333333;">
				<option value="" style="font-size:10px;font-family:arial,돋움;color:#333333;">-선택-</option>
				<option value="MSE" style="font-size:10px;font-family:arial,돋움;color:#333333;">MSE</option>
				<option value="미분류" style="font-size:10px;font-family:arial,돋움;color:#333333;">미분류</option>
				<option value="타사" style="font-size:10px;font-family:arial,돋움;color:#333333;">타사</option>
				</select>
			</td>
			<td width=50px><%=Make_S_BTN("변환","javascript:frmLP_View_Submit_Check('Change');","")%></td>
			<td width=47px></td>
			<td width=77px><%=Make_BTN("체크저장","javascript:frmLP_View_Submit_Check('Check');","")%></td>
			<td align=right>
				<table width=520px cellpadding=0 cellspacing=0 border=0>
				<tr>
					<%if S_Order_By_1="LP_Tool_Type" and S_Order_By_2="asc" and S_Order_By_3="BOM_Sub_BS_D_No" and S_Order_By_4="asc" then%>
					<td width=100px><%=Make_L_BTN("<span style='color:red;'>TYPE-P/NO</span>","javascript:setSorting2('LP_Tool_Type','asc','BOM_Sub_BS_D_No','asc')",		"")%></td>
					<%else%>
					<td width=100px><%=Make_L_BTN("<span>TYPE-P/NO</span>","javascript:setSorting2('LP_Tool_Type','asc','BOM_Sub_BS_D_No','asc')",		"")%></td>
					<%end if%>
					<td width=5px></td>
					<%if S_Order_By_1="LP_Line" and S_Order_By_2="asc" and S_Order_By_3="Min_LPD_Input_Date" and S_Order_By_4="asc" then%>
					<td width=100px><%=Make_L_BTN("<span style='color:red;'>LINE-DATE</span>","javascript:setSorting2('LP_Line','asc','Min_LPD_Input_Date','asc')",			"")%></td>
					<%else%>
					<td width=100px><%=Make_L_BTN("<span>LINE-DATE</span>","javascript:setSorting2('LP_Line','asc','Min_LPD_Input_Date','asc')",			"")%></td>
					<%end if%>
					<td width=5px></td>
					<%if S_Order_By_1="BOM_Sub_BS_D_No" and S_Order_By_2="asc" and S_Order_By_3="Min_LPD_Input_Date" and S_Order_By_4="asc" then%>
					<td width=100px><%=Make_L_BTN("<span style='color:red;'>P/NO-DATE</span>","javascript:setSorting2('BOM_Sub_BS_D_No','asc','Min_LPD_Input_Date','asc')",	"")%></td>
					<%else%>
					<td width=100px><%=Make_L_BTN("<span>P/NO-DATE</span>","javascript:setSorting2('BOM_Sub_BS_D_No','asc','Min_LPD_Input_Date','asc')",	"")%></td>
					<%end if%>
					<td width=5px></td>
					<%if S_Order_By_1="Min_LPD_Input_Date" and S_Order_By_2="asc" and S_Order_By_3="BOM_Sub_BS_D_No" and S_Order_By_4="asc" then%>
					<td width=100px><%=Make_L_BTN("<span style='color:red;'>DATE-P/NO</span>","javascript:setSorting2('Min_LPD_Input_Date','asc','BOM_Sub_BS_D_No','asc')",	"")%></td>
					<%else%>
					<td width=100px><%=Make_L_BTN("<span>DATE-P/NO</span>","javascript:setSorting2('Min_LPD_Input_Date','asc','BOM_Sub_BS_D_No','asc')",	"")%></td>
					<%end if%>
					<td width=5px></td>
					<%if S_Order_By_1="LP_LOT_Remain" and S_Order_By_2="desc" and S_Order_By_3="BOM_Sub_BS_D_No" and S_Order_By_4="asc" then%>
					<td width=100px><%=Make_L_BTN("<span style='color:red;'>PLAN-P/NO</span>","javascript:setSorting2('LP_LOT_Remain','desc','BOM_Sub_BS_D_No','asc')",		"")%></td>
					<%else%>
					<td width=100px><%=Make_L_BTN("<span>PLAN-P/NO</span>","javascript:setSorting2('LP_LOT_Remain','desc','BOM_Sub_BS_D_No','asc')",		"")%></td>
					<%end if%>
				</tr>
				</table>
			</td>
			<td width=17px></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<br>
<img src="/img/blank.gif" width=1px height=10px><br>

<img src="/img/blank.gif" width=1px height=10px><br>
<table width="<%=717+30*s_diff_LPD_Input_Date%>px" cellpadding=0 cellspacing=0 border=0 bgcolor="#999999" class="LGE_Plan">
<tr bgcolor="white" height=5>
	<td width=15px></td>
	<td width=40px style="cursor:hand;"<%if S_Order_By_1="LM_Company" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="LM_Company" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LM_Company','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=30px style="cursor:hand;"<%if S_Order_By_1="LP_Line" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Line" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Line','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=60px style="cursor:hand;"<%if S_Order_By_1="LP_Work_Order" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Work_Order" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Work_Order','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=100px style="cursor:hand;"<%if S_Order_By_1="LP_Model" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Model" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Model','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=90px style="cursor:hand;"<%if S_Order_By_1="BOM_Sub_BS_D_No" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="BOM_Sub_BS_D_No" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('BOM_Sub_BS_D_No','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=60px style="cursor:hand;"<%if S_Order_By_1="LP_Suffix" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Suffix" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Suffix','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=40px style="cursor:hand;"<%if S_Order_By_1="LP_Tool" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Tool" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Tool','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=60px style="cursor:hand;"<%if S_Order_By_1="LP_Tool_Type" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Tool_Type" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Tool_Type','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=40px style="cursor:hand;"<%if S_Order_By_1="LP_Input_Time" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Input_Time" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Input_Time','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=35px style="cursor:hand;"<%if S_Order_By_1="LP_LOT" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_LOT" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_LOT','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=35px style="cursor:hand;"<%if S_Order_By_1="LP_LOT_Remain" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_LOT_Remain" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_LOT_Remain','asc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=30px><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=30px><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=30px><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=30px><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=30px><img src="/img/blank.gif" width=1px height=5px></td>
	<td width=2px></td>
<%
for CNT1 = 0 to s_diff_LPD_Input_Date
%>
	<td width=30px style="cursor:hand;"<%if S_Order_By_1="Min_LPD_Input_Date" and S_Order_By_2="asc" then%> bgcolor="red"<%elseif S_Order_By_3="Min_LPD_Input_Date" and S_Order_By_4="asc" then%> bgcolor="orange"<%end if%> onclick="setSorting('Min_LPD_Input_Date','asc')"></td>
<%
next
%>
</tr>
<tr bgcolor="dimgray">
	<td style="color:white"></td>
	<td style="color:white"><b>COMP</b></td>
	<td style="color:white"><b>LINE</b></td>
	<td style="color:white"><b>W/O</b></td>
	<td style="color:white"><b>MODEL</b></td>
	<td style="color:white"><b>PART NO</b></td>
	<td style="color:white"><b>SUFFIX</b></td>
	<td style="color:white"><b>TOOL</b></td>
	<td style="color:white"><b>TYPE</b></td>
	<td style="color:white"><b>INPUT</b></td>
	<td style="color:white"><b>LOT</b></td>
	<td style="color:white"><b>PLAN</b></td>
	<td style="color:white" style="cursor:hand" onclick="javascript:frmLP_View_Submit_Check('Check');"><b>I</b></td>
	<td style="color:white" style="cursor:hand" onclick="javascript:frmLP_View_Submit_Check('Check');"><b>S</b></td>
	<td style="color:white" style="cursor:hand" onclick="javascript:frmLP_View_Submit_Check('Check');"><b>M</b></td>
	<td style="color:white" style="cursor:hand" onclick="javascript:frmLP_View_Submit_Check('Check');"><b>A</b></td>
	<td style="color:white" style="cursor:hand" onclick="javascript:frmLP_View_Submit_Check('Check');"><b>D</b></td>
	<td style="color:white">|</td>
<%
for CNT1 = 0 to s_diff_LPD_Input_Date
%>
	<td style="color:white"><b><%=Right(dateadd("d",CNT1,s_Min_LPD_Input_Date),2)%></td>
<%
next
%>
</tr>
<tr bgcolor="white" height=5>
	<td></td>	
	<td style="cursor:hand;"<%if S_Order_By_1="LM_Company" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="LM_Company" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LM_Company','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td style="cursor:hand;"<%if S_Order_By_1="LP_Line" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Line" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Line','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td style="cursor:hand;"<%if S_Order_By_1="LP_Work_Order" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Work_Order" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Work_Order','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td style="cursor:hand;"<%if S_Order_By_1="LP_Model" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Model" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Model','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td style="cursor:hand;"<%if S_Order_By_1="BOM_Sub_BS_D_No" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="BOM_Sub_BS_D_No" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('BOM_Sub_BS_D_No','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td style="cursor:hand;"<%if S_Order_By_1="LP_Suffix" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Suffix" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Suffix','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td style="cursor:hand;"<%if S_Order_By_1="LP_Tool" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Tool" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Tool','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td style="cursor:hand;"<%if S_Order_By_1="LP_Tool_Type" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Tool_Type" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Tool_Type','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td style="cursor:hand;"<%if S_Order_By_1="LP_Input_Time" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_Input_Time" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_Input_Time','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td style="cursor:hand;"<%if S_Order_By_1="LP_LOT" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_LOT" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_LOT','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td style="cursor:hand;"<%if S_Order_By_1="LP_LOT_Remain" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="LP_LOT_Remain" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('LP_LOT_Remain','desc')"><img src="/img/blank.gif" width=1px height=5px></td>
	<td><img src="/img/blank.gif" width=1px height=5px></td>
	<td><img src="/img/blank.gif" width=1px height=5px></td>
	<td><img src="/img/blank.gif" width=1px height=5px></td>
	<td><img src="/img/blank.gif" width=1px height=5px></td>
	<td><img src="/img/blank.gif" width=1px height=5px></td>
	<td></td>
<%
for CNT1 = 0 to s_diff_LPD_Input_Date
%>
	<td  style="cursor:hand;"<%if S_Order_By_1="Min_LPD_Input_Date" and S_Order_By_2="desc" then%> bgcolor="red"<%elseif S_Order_By_3="Min_LPD_Input_Date" and S_Order_By_4="desc" then%> bgcolor="orange"<%end if%> onclick="setSorting('Min_LPD_Input_Date','desc')"></td>
<%
next
%>
</tr>
<%
if not(RS1.Eof or RS1.Bof) then
	old_BOM_Sub_BS_D_No = RS1("BOM_Sub_BS_D_No")
end if

LP_LOT_Sum			= 0
LP_LOT_Remain_Sum	= 0

do until RS1.Eof

	LP_LOT_Sum				= LP_LOT_Sum		+ cint(LP_LOT)
	LP_LOT_Remain_Sum		= LP_LOT_Remain_Sum	+ cint(LP_LOT_Remain)
	
	LM_Company				= RS1("LM_Company")
	
	LP_Code					= RS1("LP_Code")
	LP_Line					= RS1("LP_Line")
	LP_Work_Order			= RS1("LP_Work_Order")
	LP_Model				= RS1("LP_Model")

	LP_Suffix				= RS1("LP_Suffix")
	LP_Tool					= RS1("LP_Tool")
	LP_Tool_Type			= RS1("LP_Tool_Type")
	LP_Input_time			= RS1("LP_Input_time")
	LP_Buyer				= RS1("LP_Buyer")
	LP_LOT					= RS1("LP_LOT")
	LP_LOT_Remain			= RS1("LP_LOT_Remain")
	
	BOM_Sub_BS_D_No			= RS1("BOM_Sub_BS_D_No")
	
	BOM_Sub_BS_D_No_1		= RS1("BOM_Sub_BS_D_No_1")
	BOM_Sub_BS_D_No_2		= RS1("BOM_Sub_BS_D_No_2")
	BOM_Sub_BS_D_No_3		= RS1("BOM_Sub_BS_D_No_3")
	BOM_Sub_BS_D_No_4		= RS1("BOM_Sub_BS_D_No_4")
	
	BOM_B_Code_1			= RS1("BOM_B_Code_1")
	BOM_B_Code_2			= RS1("BOM_B_Code_2")
	BOM_B_Code_3			= RS1("BOM_B_Code_3")
	BOM_B_Code_4			= RS1("BOM_B_Code_4")
	
	LP_IMD_Complete_Qty		= RS1("LP_IMD_Complete_Qty")
	LP_SMD_Complete_Qty		= RS1("LP_SMD_Complete_Qty")
	LP_MAN_Complete_Qty		= RS1("LP_MAN_Complete_Qty")
	LP_ASM_Complete_Qty		= RS1("LP_ASM_Complete_Qty")
	LP_DLV_Complete_Qty		= RS1("LP_DLV_Complete_Qty")
	
	if BOM_B_Code_1 = "" then
		BOM_Sub_BS_D_No_1 = "<span style='color:gray'>" & BOM_Sub_BS_D_No_1 & "</span>"
	else
		BOM_Sub_BS_D_No_1 = "<span style='cursor:hand' onclick=""javascript:window.open('/bom/db_load_action.asp?B_Code="&BOM_B_Code_1&"')"">" & BOM_Sub_BS_D_No_1 & "</span>"
	end if
	
	if BOM_B_Code_2 = "" then
		BOM_Sub_BS_D_No_2 = "<span style='color:gray'>" & BOM_Sub_BS_D_No_2 & "</span>"
	else
		BOM_Sub_BS_D_No_2 = "<span style='cursor:hand' onclick=""javascript:window.open('/bom/db_load_action.asp?B_Code="&BOM_B_Code_2&"')"">" & BOM_Sub_BS_D_No_2 & "</span>"
	end if
	
	if BOM_B_Code_3 = "" then
		BOM_Sub_BS_D_No_3 = "<span style='color:gray'>" & BOM_Sub_BS_D_No_3 & "</span>"
	else
		BOM_Sub_BS_D_No_3 = "<span style='cursor:hand' onclick=""javascript:window.open('/bom/db_load_action.asp?B_Code="&BOM_B_Code_3&"')"">" & BOM_Sub_BS_D_No_3 & "</span>"
	end if
	
	if BOM_B_Code_4 = "" then
		BOM_Sub_BS_D_No_4 = "<span style='color:gray'>" & BOM_Sub_BS_D_No_4 & "</span>"
	else
		BOM_Sub_BS_D_No_4 = "<span style='cursor:hand' onclick=""javascript:window.open('/bom/db_load_action.asp?B_Code="&BOM_B_Code_4&"')"">" & BOM_Sub_BS_D_No_4 & "</span>"
	end if
	
	
	BOM_Sub_BS_D_No_Str	 	= ""
	BOM_Sub_BS_D_No_Str = RS1("BOM_Sub_BS_D_No")
	if not(ISNULL(BOM_Sub_BS_D_No_Str)) then
		BOM_Sub_BS_D_No_Str = replace(BOM_Sub_BS_D_No_Str,"<br><br><br>","")
		BOM_Sub_BS_D_No_Str = replace(BOM_Sub_BS_D_No_Str,"<br><br>","")
	
		if right(BOM_Sub_BS_D_No_Str,4) = "<br>" then
			BOM_Sub_BS_D_No_Str = left(BOM_Sub_BS_D_No_Str,len(BOM_Sub_BS_D_No_Str)-4)
		end if
	end if
		
	if old_BOM_Sub_BS_D_No <> BOM_Sub_BS_D_No and (S_Order_By_1 = "BOM_Sub_BS_D_No" or S_Order_By_3 = "BOM_Sub_BS_D_No") then
		call TR_Sum(strBGColor_Sum,strFontColor_sum,old_BOM_Sub_BS_D_No_Str,LP_LOT_Sum,LP_LOT_Remain_Sum,s_diff_LPD_Input_Date,s_Min_LPD_Input_Date,LPD_Input_Qty_Sum)
		
		LP_LOT_Sum			= 0
		LP_LOT_Remain_Sum	= 0
		for CNT1 = 0 to s_diff_LPD_Input_Date
			LPD_Input_Qty_Sum(CNT1) = 0
		next
%>
<tr bgcolor=white height=30px>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<!--<td><%=LP_Buyer%></td>-->
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
<%
		for CNT1 = 0 to s_diff_LPD_Input_Date		
%>
	<td bgcolor="white"></td>
<%
		next

		old_BOM_Sub_BS_D_No		= BOM_Sub_BS_D_No
		old_BOM_Sub_BS_D_No_Str	= BOM_Sub_BS_D_No_Str
		LP_LOT_Sum				= 0
		LP_LOT_Remain_Sum		= 0
%>
</tr>
<tr bgcolor="dimgray">
	<td style="color:white"></td>
	<td style="color:white"><b>COMP</td>
	<td style="color:white"><b>LINE</td>
	<td style="color:white"><b>W/O</td>
	<td style="color:white"><b>MODEL</td>
	<td style="color:white"><b>PART NO</td>
	<td style="color:white"><b>SUFFIX</td>
	<td style="color:white"><b>TOOL</td>
	<td style="color:white"><b>TYPE</td>
	<td style="color:white"><b>INPUT</td>
	<td style="color:white"><b>LOT</td>
	<td style="color:white"><b>PLAN</td>
	<td style="color:white" style="cursor:hand" onclick="javascript:frmLP_View_Submit_Check('Check');"><b>I</td>
	<td style="color:white" style="cursor:hand" onclick="javascript:frmLP_View_Submit_Check('Check');"><b>S</td>
	<td style="color:white" style="cursor:hand" onclick="javascript:frmLP_View_Submit_Check('Check');"><b>M</td>
	<td style="color:white" style="cursor:hand" onclick="javascript:frmLP_View_Submit_Check('Check');"><b>A</td>
	<td style="color:white" style="cursor:hand" onclick="javascript:frmLP_View_Submit_Check('Check');"><b>D</td>
	<td style="color:white">|</td>
<%
		for CNT1 = 0 to s_diff_LPD_Input_Date
%>
	<td style="color:white"><b><%=Right(dateadd("d",CNT1,s_Min_LPD_Input_Date),2)%></td>
<%
		next
%>
</tr>
<%
	else
%>
<tr height=1px><td colspan=100><img src="/img/black.gif" width=100% height=1px></td></tr>
<%
		old_BOM_Sub_BS_D_No_Str = BOM_Sub_BS_D_No_Str
	end if
	
	if LP_DLV_Complete_Qty > 0 then
		strBGColor = "#CFCFCF"
	elseif LP_ASM_Complete_Qty > 0 then
		strBGColor = "#DBCEE6"
	elseif LP_MAN_Complete_Qty > 0 then
		strBGColor = "#C6EBFE"
	elseif LP_SMD_Complete_Qty > 0 then
		strBGColor = "#D2F6C9"
	elseif LP_IMD_Complete_Qty > 0 then
		strBGColor = "#F2D4D4"
	else
		strBGColor = "white"
	end if
%>
<tr bgcolor="<%=strBGColor%>">
	<td><input type="checkbox" name="strLP_Model" value="<%=LP_Model%>" style="border:0px none #ffffff;background-color:<%=strBGColor%>"></td>
	<td><%=LM_Company%></td>
	<td><%=LP_Line%></td>
	<td><%=LP_Work_Order%></td>
	<td><%=LP_Model%></td>
	<td><%=BOM_Sub_BS_D_No_Str%></td>
	<td><%=LP_Suffix%></td>
	<td><%=LP_Tool%></td>
	<td><%=LP_Tool_Type%></td>
	<td><%=LP_Input_Time%></td>
	<!--<td><%=LP_Buyer%></td>-->
	<td><%=LP_LOT%></td>
	<td><%=LP_LOT_Remain%></td>
	<td><input type="text" name="strLP_IMD_Complete_Qty" value="<%=LP_IMD_Complete_Qty%>" style="font-family:arial;font-size:10px;width:22px;text-align:center;border:0px solid none;background-color:<%=strBGColor%>;border-bottom:1px solid #999999" maxlength=3></td>
	<td><input type="text" name="strLP_SMD_Complete_Qty" value="<%=LP_SMD_Complete_Qty%>" style="font-family:arial;font-size:10px;width:22px;text-align:center;border:0px solid none;background-color:<%=strBGColor%>;border-bottom:1px solid #999999" maxlength=3></td>
	<td><input type="text" name="strLP_MAN_Complete_Qty" value="<%=LP_MAN_Complete_Qty%>" style="font-family:arial;font-size:10px;width:22px;text-align:center;border:0px solid none;background-color:<%=strBGColor%>;border-bottom:1px solid #999999" maxlength=3></td>
	<td><input type="text" name="strLP_ASM_Complete_Qty" value="<%=LP_ASM_Complete_Qty%>" style="font-family:arial;font-size:10px;width:22px;text-align:center;border:0px solid none;background-color:<%=strBGColor%>;border-bottom:1px solid #999999" maxlength=3></td>
	<td><input type="text" name="strLP_DLV_Complete_Qty" value="<%=LP_DLV_Complete_Qty%>" style="font-family:arial;font-size:10px;width:22px;text-align:center;border:0px solid none;background-color:<%=strBGColor%>;border-bottom:1px solid #999999" maxlength=3></td>
	<td>|</td>
<%	
	LPD_Input_Qty_TR_Sum = 0
	for CNT1 = 0 to s_diff_LPD_Input_Date
		LPD_Input_Qty			= ""		
		
		LPD_Input_Qty = RS1("DATE_QTY_"&CNT1)
		
		if isnumeric(LPD_Input_Qty) then
			LPD_Input_Qty_Sum(CNT1)	= LPD_Input_Qty_Sum(CNT1) + cint(LPD_Input_Qty)
			LPD_Input_Qty_TR_Sum	= LPD_Input_Qty_TR_Sum + cint(LPD_Input_Qty)
		end if
		
		if dateadd("d",CNT1,s_Min_LPD_Input_Date) = date() then
%>
	<td bgcolor="#7EFF00" align=center><%=LPD_Input_Qty%></td>
<%
		else

			if weekday(dateadd("d",CNT1,s_Min_LPD_Input_Date)) = 1 then
%>
	<td bgcolor="pink" align=center><%=LPD_Input_Qty%></td>
		
<%
			elseif weekday(dateadd("d",CNT1,s_Min_LPD_Input_Date)) = 7 then
%>
	<td bgcolor="skyblue" align=center><%=LPD_Input_Qty%></td>
<%			
			else
%>
	<td bgcolor="<%if CNT1 mod 2 = 1 then%>white<%else%>#e3e3e3<%end if%>" align=center><%=LPD_Input_Qty%></td>
<%			
			end if
		end if
	next
%>
	<input type="hidden" name="strAll_Work_Order" value="<%=LP_Work_Order%>//<%=LPD_Input_Qty_TR_Sum%>//<%=BOM_Sub_BS_D_No_Str%>">
	<input type="hidden" name="strOLD_LP_IMD_Complete_Qty" value="<%=LP_IMD_Complete_Qty%>">
	<input type="hidden" name="strOLD_LP_SMD_Complete_Qty" value="<%=LP_SMD_Complete_Qty%>">
	<input type="hidden" name="strOLD_LP_MAN_Complete_Qty" value="<%=LP_MAN_Complete_Qty%>">
	<input type="hidden" name="strOLD_LP_ASM_Complete_Qty" value="<%=LP_ASM_Complete_Qty%>">
	<input type="hidden" name="strOLD_LP_DLV_Complete_Qty" value="<%=LP_DLV_Complete_Qty%>">
</tr>
<%
	RS1.MoveNext
loop
RS1.Close

if S_Order_By_1 = "BOM_Sub_BS_D_No" or S_Order_By_3 = "BOM_Sub_BS_D_No" then
	LP_LOT_Sum				= LP_LOT_Sum		+ cint(LP_LOT)
	LP_LOT_Remain_Sum		= LP_LOT_Remain_Sum	+ cint(LP_LOT_Remain)

	call TR_Sum(strBGColor_Sum,strFontColor_sum,old_BOM_Sub_BS_D_No_Str,LP_LOT_Sum,LP_LOT_Remain_Sum,s_diff_LPD_Input_Date,s_Min_LPD_Input_Date,LPD_Input_Qty_Sum)
	
	LP_LOT_Sum			= 0
	LP_LOT_Remain_Sum	= 0
	for CNT1 = 0 to s_diff_LPD_Input_Date
		LPD_Input_Qty_Sum(CNT1) = 0
	next
end if
%>
</form>
</table>

<%
if s_Edit_Process <> "" then

	select case s_Edit_Process
		case "IMD"
			bgMSE_Plan_Editor = "#F2D4D4"
		case "SMD"
			bgMSE_Plan_Editor = "#D2F6C9"
		case "MAN"
			bgMSE_Plan_Editor = "#C6EBFE"
		case "ASM"
			bgMSE_Plan_Editor = "#DBCEE6"		
	end select
%>
<div id="divMSE_Plan_Editor" style="width:50px;height:50px;position:absolute;display:none;border:1px solid #999999;filter:alpha(opacity=90);">
<table width=100% cellpadding=0 cellspacing=0 border=0 bgcolor=white class="MSE_Plan_Editor">
<form name="frmMSE_Plan_Editor" action="inc_MSE_Plan_reg_action.asp" method="post" target="ifrmMSE_Plan_Action">
<input type="hidden" name="idDIV">
<input type="hidden" name="BOM_Sub_BS_D_No">
<input type="hidden" name="MPD_Process" value="<%=s_Edit_Process%>">
<input type="hidden" name="MPD_Date">
<input type="hidden" name="MPD_Qty_Total">
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td colspan=2 align=left valign=top><img src="/img/ico_MSE_Plan_<%=s_Edit_Process%>.gif"></td>
</tr>
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td width=40px>공정 :</td>
	<td align=left><%=s_Edit_Process%></td>
</tr>
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td width=40px>날짜 :</td>
	<td align=left id="idMSE_Plan_Date"></td>
</tr>
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td width=40px>총계 :</td>
	<td align=left id="idMPD_Qty_Total">
		&nbsp;
	</td>
</tr>
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td colspan=2>
		<table width=100% cellpadding=1 cellspacing=0 border=0>
		<tr>
			<td width=30px></td>
<%
	select case s_Edit_Process
	case "IMD"
		arrInputSelectG_2 = split(replace(BasicDataIMDLine,"slt>",""),";")
	case "SMD"
		arrInputSelectG_2 = split(replace(BasicDataSMDLine,"slt>",""),";")
	case "MAN"
		arrInputSelectG_2 = split(replace(BasicDataMANLine,"slt>",""),";")
	case "ASM"
		arrInputSelectG_2 = split(replace(BasicDataASMLine,"slt>",""),";")
	end select
	
	for CNT2 = 0 to ubound(arrInputSelectG_2)
		arrInputSelect_2 = split(arrInputSelectG_2(CNT2),":")
%>
			<td><%=arrInputSelect_2(0)%></td>
<%
	next
%>			
			<td>&nbsp;</td>
		</tr>
<%	
	if s_Edit_Process="IMD" or s_Edit_Process="SMD" then
		arrInputSelectG_1	= split(replace(BasicDataFullTime,"slt>",""),";")
	else
		arrInputSelectG_1	= split(replace(BasicDataHalfTime,"slt>",""),";")
	end if

	for CNT1 = 0 to ubound(arrInputSelectG_1)
		arrInputSelect_1 = split(arrInputSelectG_1(CNT1),":")
%>
		<tr>
			<td>&nbsp;<%=arrInputSelect_1(0)%>&nbsp;</td>
<%
		select case s_Edit_Process
		case "IMD"
			arrInputSelectG_2 = split(replace(BasicDataIMDLine,"slt>",""),";")
		case "SMD"
			arrInputSelectG_2 = split(replace(BasicDataSMDLine,"slt>",""),";")
		case "MAN"
			arrInputSelectG_2 = split(replace(BasicDataMANLine,"slt>",""),";")
		case "ASM"
			arrInputSelectG_2 = split(replace(BasicDataASMLine,"slt>",""),";")
		end select
		for CNT2 = 0 to ubound(arrInputSelectG_2)
			arrInputSelect_2 = split(arrInputSelectG_2(CNT2),":")
%>
			<td><input type="text" name="<%=arrInputSelect_2(0)%>_<%=arrInputSelect_1(0)%>" value="" style="width:35px;text-align:center" maxlength=4 onkeyup="javascript:cal_MPD_Qty_Total()"></td>
<%
		next
%>
			<td>&nbsp;</td>
		</tr>
<%
	next
%>
		</table>
	</td>
</tr>
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td colspan=2>
		<input type="button" value="확인" onclick="javascript:frmMSE_Plan_Editor_Check();">
		<input type="button" value="취소" onclick="javascript:hide_MSE_Plan_Editor();">
	</td>
</tr>
<tr height=3px bgcolor="<%=bgMSE_Plan_Editor%>">
	<td colspan=2><img src="/img/blank.gif" width=1px height=3px></td>
</tr>
</form>
</table>
<iframe name="ifrmMSE_Plan_Action" src="about:blank" width=0px height=0px frameborder=0></iframe>
</div>

<%
dim ifrmPR_Monitor_Width
dim ifrmPR_Monitor_Height
select case s_Edit_Process
case "IMD"
	ifrmPR_Monitor_Width	= "507"
	ifrmPR_Monitor_Height	= "532"
case "SMD"
	ifrmPR_Monitor_Width	= "398"
	ifrmPR_Monitor_Height	= "532"
case "MAN"
	ifrmPR_Monitor_Width	= "507"
	ifrmPR_Monitor_Height	= "277"
case "ASM"
	ifrmPR_Monitor_Width	= "507"
	ifrmPR_Monitor_Height	= "277"
end select
%>
<div id="divMSE_Plan_Monitor" style="width:<%=ifrmPR_Monitor_Width%>px;height:<%=ifrmPR_Monitor_Height%>px;position:absolute;display:none;border:1px solid #999999;filter:alpha(opacity=90);">
<iframe name="ifrmMSE_Plan_Monitor" src="about:blank" width=<%=ifrmPR_Monitor_Width%>px height=<%=ifrmPR_Monitor_Height%>px frameborder=0></iframe>
</div>

<script language="javascript">
function cal_MPD_Qty_Total()
{
	var nTotal = 0;
<%
	if s_Edit_Process="IMD" or s_Edit_Process="SMD" then
		arrInputSelectG_1	= split(replace(BasicDataFullTime,"slt>",""),";")
	else
		arrInputSelectG_1	= split(replace(BasicDataHalfTime,"slt>",""),";")
	end if
		
	for CNT1 = 0 to ubound(arrInputSelectG_1)
		arrInputSelect_1 = split(arrInputSelectG_1(CNT1),":")
		
		select case s_Edit_Process
		case "IMD"
			arrInputSelectG_2 = split(replace(BasicDataIMDLine,"slt>",""),";")
		case "SMD"
			arrInputSelectG_2 = split(replace(BasicDataSMDLine,"slt>",""),";")
		case "MAN"
			arrInputSelectG_2 = split(replace(BasicDataMANLine,"slt>",""),";")
		case "ASM"
			arrInputSelectG_2 = split(replace(BasicDataASMLine,"slt>",""),";")
		end select
		
		for CNT2 = 0 to ubound(arrInputSelectG_2)
			arrInputSelect_2 = split(arrInputSelectG_2(CNT2),":")
%>
	if (frmMSE_Plan_Editor.<%=arrInputSelect_2(0)%>_<%=arrInputSelect_1(0)%>.value)
		nTotal += parseInt(frmMSE_Plan_Editor.<%=arrInputSelect_2(0)%>_<%=arrInputSelect_1(0)%>.value);
<%
		next
	next
%>
	frmMSE_Plan_Editor.MPD_Qty_Total.value = nTotal;
	
	var objDIV = document.getElementById("idMPD_Qty_Total");
	objDIV.innerHTML = nTotal;
}

function show_MSE_Plan_Editor(idDIV,strBOM_Sub_BS_D_No,strMSE_Plan_Date)
{
	frmMSE_Plan_Editor.reset();
	ifrmMSE_Plan_Action.location.href	= "inc_mse_plan_load_action.asp?MPD_Process=<%=s_Edit_Process%>&BOM_Sub_BS_D_No="+strBOM_Sub_BS_D_No+"&MPD_Date="+strMSE_Plan_Date;
	
	divMSE_Plan_Editor.style.posLeft	= event.x + 0 + document.body.scrollLeft;
	divMSE_Plan_Editor.style.posTop		= event.y + 0 + document.body.scrollTop;
	divMSE_Plan_Editor.style.display	= "block";
	
	ifrmMSE_Plan_Monitor.location.href	= "/process_record/inc_pr_monitor.asp?s_Process=<%=s_Edit_Process%>&s_Date="+strMSE_Plan_Date;
	divMSE_Plan_Monitor.style.posLeft	= event.x + 0 + document.body.scrollLeft - 600;
	divMSE_Plan_Monitor.style.posTop	= event.y + 0 + document.body.scrollTop;
	divMSE_Plan_Monitor.style.display	= "block";

	frmMSE_Plan_Editor.BOM_Sub_BS_D_No.value				= strBOM_Sub_BS_D_No;
	frmMSE_Plan_Editor.MPD_Date.value						= strMSE_Plan_Date;
	
	document.getElementById("idMSE_Plan_Date").innerHTML	= strMSE_Plan_Date;
	frmMSE_Plan_Editor.idDIV.value							= idDIV;
}

function hide_MSE_Plan_Editor()
{
	frmMSE_Plan_Editor.reset();
	divMSE_Plan_Editor.style.display="none";
	divMSE_Plan_Monitor.style.display="none";
	
	if(typeof(parent.ifrmPR_Monitor) == "object")
		parent.ifrmPR_Monitor.location.reload();
}

function frmMSE_Plan_Editor_Check()
{
	var objDIV = document.getElementById(frmMSE_Plan_Editor.idDIV.value);
	if (parseInt(frmMSE_Plan_Editor.MPD_Qty_Total.value) > 0)
	{
		objDIV.innerHTML		= frmMSE_Plan_Editor.MPD_Qty_Total.value;
		objDIV.style.display	= "block";
	
<%
	select case s_Edit_Process
	case "IMD"
%>
		objDIV.style.backgroundColor = "red";
<%
	case "SMD"
%>
		objDIV.style.backgroundColor = "green";
<%
	case "MAN"
%>
		objDIV.style.backgroundColor = "blue";
<%
	case "ASM"
%>
		objDIV.style.backgroundColor = "#7306C6";
<%
	end select
%>
	}
	else
	{
		objDIV.style.display			= "none";
		objDIV.innerHTML				= "";
		objDIV.style.backgroundColor	= "transparent";
	}
	frmMSE_Plan_Editor.submit();
}

function Load_frmMSE_Plan_Editor(strMPD_Line,strMPD_Time,strMPD_Qty)
{
	objForm = eval("frmMSE_Plan_Editor."+strMPD_Line+"_"+strMPD_Time)
	objForm.value = strMPD_Qty;
}
</script>
<%
end if
%>

<%
set RS1 = nothing
set RS2 = nothing
%>


<%
sub TR_Sum(strBGColor_Sum,strFontColor_sum,old_BOM_Sub_BS_D_No_Str,LP_LOT_Sum,LP_LOT_Remain_Sum,s_diff_LPD_Input_Date,s_Min_LPD_Input_Date,LPD_Input_Qty_Sum)
%>
<tr height=1px>
	<td bgcolor="#333333" colspan=100><img src="/img/blank.gif" width=1px height=1px></td>
</tr>
<%
	dim RS1
	dim CNT1
	dim CNT2
	dim CNT3
	
	dim strHidden
	dim strBlank
	dim strEdit_Process
	dim strOther_Process
	dim Edit_Process_Exist_YN
	dim Other_Process_Exist_YN
	
	dim arrBOM_Sub_BS_D_No
	
	dim B_IMD_Qty
	dim BS_SMD_Qty
	dim BS_MAN_Qty
	dim BS_ASM_Qty
	
	if not(ISNULL(old_BOM_Sub_BS_D_No_Str)) then
		old_BOM_Sub_BS_D_No_Str = replace(old_BOM_Sub_BS_D_No_Str,"<br><br><br>","")
		old_BOM_Sub_BS_D_No_Str = replace(old_BOM_Sub_BS_D_No_Str,"<br><br>","")
		if right(old_BOM_Sub_BS_D_No_Str,4) = "<br>" then
			old_BOM_Sub_BS_D_No_Str = left(old_BOM_Sub_BS_D_No_Str,len(old_BOM_Sub_BS_D_No_Str)-4)
		end if
	else
		old_BOM_Sub_BS_D_No_Str = ""
	end if
	
	
	arrBOM_Sub_BS_D_No = split(old_BOM_Sub_BS_D_No_Str,"<br>")
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)
		SQL = "select BS_SMD_Qty,BS_MAN_Qty,BS_ASM_Qty from tbBOM_Sub where BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			BS_SMD_Qty = 0
			BS_MAN_Qty = 0
			BS_ASM_Qty = 0
		else
			BS_SMD_Qty = RS1("BS_SMD_Qty")
			BS_MAN_Qty = RS1("BS_MAN_Qty")
			BS_ASM_Qty = RS1("BS_ASM_Qty")
		end if
		RS1.Close
		
		SQL = "select B_IMD_Qty from tbBOM where B_Code in (select BOM_B_Code from tbBOM_Sub where BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"')"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			B_IMD_Qty = 0
		else
			B_IMD_Qty = RS1("B_IMD_Qty")
		end if
		RS1.Close
%>
<tr bgcolor="<%=strBGColor_Sum%>" height=19px>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td style="color:<%=strFontColor_sum%>"><%=arrBOM_Sub_BS_D_No(CNT1)%></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<!--<td><%=LP_Buyer%></td>-->
	<td style="color:<%=strFontColor_sum%>"><%=LP_LOT_Sum%></td>
	<td style="color:<%=strFontColor_sum%>"><%=LP_LOT_Remain_Sum%></td>
	<td style="color:<%=strFontColor_sum%>"><%=B_IMD_Qty%></td>
	<td style="color:<%=strFontColor_sum%>"><%=BS_SMD_Qty%></td>
	<td style="color:<%=strFontColor_sum%>"><%=BS_MAN_Qty%></td>
	<td style="color:<%=strFontColor_sum%>"><%=BS_ASM_Qty%></td>
	<td></td>
	<td style="color:<%=strFontColor_sum%>">|</td>
<%
		strDiff_MPD_Date		= ""
		strMPD_Process			= ""
		strMPD_Qty_Sum			= ""
		SQL =		"select distinct"&vbcrlf
		SQL = SQL & "	Code		= (left(convert(char,MPD_Date,121),10)+BOM_Sub_BS_D_No), "&vbcrlf
		SQL = SQL & "	Diff_MPD_Date = datediff(day,'"&s_Min_LPD_Input_Date&"',MPD_Date), "&vbcrlf
		SQL = SQL & "	MPD_Process, "&vbcrlf
		SQL = SQL & "	MPD_Qty_Sum	= sum(MPD_Qty) "&vbcrlf
		SQL = SQL & "from "&vbcrlf
		SQL = SQL & "	tbMSE_Plan_Date "&vbcrlf
		SQL = SQL & "where "&vbcrlf
		SQL = SQL & "	BOM_Sub_BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"' "&vbcrlf
		SQL = SQL & "	group by BOM_Sub_BS_D_No, MPD_Date, MPD_Process "&vbcrlf
		
		RS2.Open SQL,sys_DBCon
		do until RS2.Eof
			strDiff_MPD_Date	= strDiff_MPD_Date	& RS2("Diff_MPD_Date")	& "//"
			strMPD_Process		= strMPD_Process	& RS2("MPD_Process")	& "//"
			strMPD_Qty_Sum		= strMPD_Qty_Sum	& RS2("MPD_Qty_Sum")	& "//"
			RS2.MoveNext
		loop
		RS2.Close
		arrDiff_MPD_Date	= split(strDiff_MPD_Date,"//")
		arrMPD_Process		= split(strMPD_Process,"//")
		arrMPD_Qty_Sum		= split(strMPD_Qty_Sum,"//")
				
		for CNT2 = 0 to s_diff_LPD_Input_Date
			
			if s_Edit_Process <> "" then			'수정 모드라면 수정을 위한 공용코드를 만들어 둔다.
				strCall_MSE_Plan_Editor	= " style='cursor:hand' onclick=""show_MSE_Plan_Editor('"&arrBOM_Sub_BS_D_No(CNT1)&"_"&CNT2&"','"&arrBOM_Sub_BS_D_No(CNT1)&"','"&dateadd("d",CNT2,s_Min_LPD_Input_Date)&"')"" "
			end if
			strHidden	= "<span"&strCall_MSE_Plan_Editor&" id='"&arrBOM_Sub_BS_D_No(CNT1)&"_"&CNT2&"' class='"&s_Edit_Process&"' style='display:none;'>&nbsp;</span>"
			strBlank	= "<span"&strCall_MSE_Plan_Editor&" id='"&arrBOM_Sub_BS_D_No(CNT1)&"_"&CNT2&"' class='BLANK'>&nbsp;</span>"
			
			strQty = LPD_Input_Qty_Sum(CNT2)
			if strQty = 0 then
				strQty = ""
			end if
			strQty = "<span"&strCall_MSE_Plan_Editor&" class='LGE_DUE'>"&strQty&"</span>"
			
			Edit_Process_Exist_YN	= "N"
			Other_Process_Exist_YN	= "N"	
			strEdit_Process			= ""						'한 셀용 변수 초기화
			strOther_Process		= ""
			for CNT3 = 0 to ubound(arrDiff_MPD_Date) - 1	'조회된 데이터를 검색
				if arrDiff_MPD_Date(CNT3) = cstr(CNT2) then			'그 중 현재 셀에 표시할 것이 잇다면.
					if arrMPD_Process(CNT3) = s_Edit_Process then	'수정 대상 공정이면,
						Edit_Process_Exist_YN	= "Y"
						strEdit_Process = strEdit_Process & "<span"&strCall_MSE_Plan_Editor&" id='"&arrBOM_Sub_BS_D_No(CNT1)&"_"&CNT2&"' class='"&arrMPD_Process(CNT3)&"'>"&arrMPD_Qty_Sum(CNT3)&"</span>"
					else											'수정 대상 공정이 아니면,
						Other_Process_Exist_YN	= "Y"				
						strOther_Process = strOther_Process & "<span"&strCall_MSE_Plan_Editor&" class='"&arrMPD_Process(CNT3)&"'>"&arrMPD_Qty_Sum(CNT3)&"</span>"
					end if
				end if
			next		
			
			if LPD_Input_Qty_Sum(CNT2) = 0 then				'납기정보 없음
				if Edit_Process_Exist_YN = "Y" then				'수정대상 공정 있음
					strQty = strEdit_Process & strOther_Process
				elseif Other_Process_Exist_YN = "Y" then		'다른 공정만 있음
					strQty = strHidden & strEdit_Process & strOther_Process 
				else											'공정정보 없음
					strQty = strBlank
				end if
			elseif strQty <> "" then						'납기정보 있음
				if Edit_Process_Exist_YN = "Y" then				'수정대상 공정 있음
					strQty = strEdit_Process & strOther_Process & strQty
				elseif Other_Process_Exist_YN = "Y" then		'다른 공정만 있음
					strQty = strHidden & strEdit_Process & strOther_Process & strQty
				else											'공정정보 없음
					strQty = strHidden & strQty
				end if
			end if			
			
			if dateadd("d",CNT2,s_Min_LPD_Input_Date) = date() then
%>
	<td style="color:<%=strFontColor_sum%>" bgcolor="#7EFF00" align=center><%=strQty%></td>
<%
			else
				if weekday(dateadd("d",CNT2,s_Min_LPD_Input_Date)) = 1 then
%>
	<td style="color:<%=strFontColor_sum%>" bgcolor="pink"><%=strQty%></td>
<%
				elseif weekday(dateadd("d",CNT2,s_Min_LPD_Input_Date)) = 7 then
%>
	<td style="color:<%=strFontColor_sum%>" bgcolor="skyblue"><%=strQty%></td>
<%			
				else
%>
	<td style="color:<%=strFontColor_sum%>" bgcolor="<%if CNT2 mod 2 = 1 then%>white<%else%>#e3e3e3<%end if%>"><%=strQty%></td>
<%			
				end if
			end if
			
			if CNT1 = ubound(arrBOM_Sub_BS_D_No) then
				LPD_Input_Qty_Sum(CNT2) = 0
			end if
		next
%>
</tr>
<%
		if CNT1 <= ubound(arrBOM_Sub_BS_D_No) then
%>
<tr height=1px>
	<td bgcolor="#ffffff" colspan=4></td>
	<td bgcolor="#999999" colspan="<%=10+s_diff_LPD_Input_Date%>"><img src="/img/blank.gif" width=1px height=1px></td>
</tr>
<%
		end if
	next
%>
<!--
<tr height=1px>
	<td bgcolor="#333333" colspan=100><img src="/img/blank.gif" width=1px height=1px></td>
</tr>
-->
<%
	set RS1 = nothing
end sub
%>


<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->