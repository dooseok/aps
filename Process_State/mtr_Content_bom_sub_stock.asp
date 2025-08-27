<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_Work_Date
dim nRawCount
dim strLine
dim oldLine

dim strBG_Color

's_Work_Date = request("s_Work_Date")
s_Work_Date = date()

dim RS1
dim RS2
dim SQL

'반복문에 사용하기 위한 변수 선언
dim CNT1
dim CNT2
dim CNT3

dim arrBOM_Sub_Stock(300,4)
dim nRowSpanP1
dim nRowSpanP2
dim nRowSpanP3
dim nRowSpanP4
dim nRowSpanP5
dim nRowSpanC5

dim nCurrentP1
dim nCurrentP2
dim nCurrentP3
dim nCurrentP4
dim nCurrentP5
dim nCurrentC5

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")

SQL = "select count(distinct PRD_Partno) from tbPWS_Raw_Data where PRD_Box_Date = '"&s_Work_Date&"' and PRD_Line = 'pwsbox1'"
RS1.Open SQL,sys_DBCon,1
nRowSpanP1 = RS1(0)
RS1.Close
SQL = "select count(distinct PRD_Partno) from tbPWS_Raw_Data where PRD_Box_Date = '"&s_Work_Date&"' and PRD_Line = 'pwsbox2'"
RS1.Open SQL,sys_DBCon,1
nRowSpanP2 = RS1(0)
RS1.Close
SQL = "select count(distinct PRD_Partno) from tbPWS_Raw_Data where PRD_Box_Date = '"&s_Work_Date&"' and PRD_Line = 'pwsbox3'"
RS1.Open SQL,sys_DBCon,1
nRowSpanP3 = RS1(0)
RS1.Close
SQL = "select count(distinct PRD_Partno) from tbPWS_Raw_Data where PRD_Box_Date = '"&s_Work_Date&"' and PRD_Line = 'pwsbox4'"
RS1.Open SQL,sys_DBCon,1
nRowSpanP4 = RS1(0)
RS1.Close
SQL = "select count(distinct PRD_Partno) from tbPWS_Raw_Data where PRD_Box_Date = '"&s_Work_Date&"' and PRD_Line = 'pwsbox51'"
RS1.Open SQL,sys_DBCon,1
nRowSpanP5 = RS1(0)
RS1.Close
SQL = "select count(distinct PRD_Partno) from tbPWS_Raw_Data where PRD_Box_Date = '"&s_Work_Date&"' and PRD_Line = 'pwsbox52'"
RS1.Open SQL,sys_DBCon,1
nRowSpanC5 = RS1(0)
RS1.Close

SQL = "select top 1 PRD_PartNo from tbPWS_Raw_Data where PRD_Line = 'pwsbox1' and PRD_Box_Date = '"&s_Work_Date&"' order by PRD_Box_Time desc"
RS1.Open SQL,sys_DBCon,1
if not(RS1.Eof or RS1.Bof) then
	nCurrentP1 = RS1("PRD_PartNo")
end if
RS1.Close
SQL = "select top 1 PRD_PartNo from tbPWS_Raw_Data where PRD_Line = 'pwsbox2' and PRD_Box_Date = '"&s_Work_Date&"' order by PRD_Box_Time desc"
RS1.Open SQL,sys_DBCon,1
if not(RS1.Eof or RS1.Bof) then
	nCurrentP2 = RS1("PRD_PartNo")
end if
RS1.Close
SQL = "select top 1 PRD_PartNo from tbPWS_Raw_Data where PRD_Line = 'pwsbox3' and PRD_Box_Date = '"&s_Work_Date&"' order by PRD_Box_Time desc"
RS1.Open SQL,sys_DBCon,1
if not(RS1.Eof or RS1.Bof) then
	nCurrentP3 = RS1("PRD_PartNo")
end if
RS1.Close
SQL = "select top 1 PRD_PartNo from tbPWS_Raw_Data where PRD_Line = 'pwsbox4' and PRD_Box_Date = '"&s_Work_Date&"' order by PRD_Box_Time desc"
RS1.Open SQL,sys_DBCon,1
if not(RS1.Eof or RS1.Bof) then
	nCurrentP4 = RS1("PRD_PartNo")
end if
RS1.Close
SQL = "select top 1 PRD_PartNo from tbPWS_Raw_Data where PRD_Line = 'pwsbox51' and PRD_Box_Date = '"&s_Work_Date&"' order by PRD_Box_Time desc"
RS1.Open SQL,sys_DBCon,1
if not(RS1.Eof or RS1.Bof) then
	nCurrentP5 = RS1("PRD_PartNo")
end if
RS1.Close
SQL = "select top 1 PRD_PartNo from tbPWS_Raw_Data where PRD_Line = 'pwsbox52' and PRD_Box_Date = '"&s_Work_Date&"' order by PRD_Box_Time desc"
RS1.Open SQL,sys_DBCon,1
if not(RS1.Eof or RS1.Bof) then
	nCurrentC5 = RS1("PRD_PartNo")
end if
RS1.Close

SQL = "select PRD_Line, PRD_Partno, cntPartNo = count(PRD_PartNo) from tbPWS_Raw_Data where PRD_Box_Date = '"&s_Work_Date&"' group by PRD_Line, PRD_PartNo order by PRD_Line, PRD_PartNo"
RS1.Open SQL,sys_DBCon,1
CNT1 = 0
do until RS1.Eof
	
	select case lcase(RS1("PRD_Line"))
		case "pwsbox1"
			strLine = "P-1"
		case "pwsbox2"
			strLine = "P-2"
		case "pwsbox3"
			strLine = "P-3"
		case "pwsbox4"
			strLine = "P-4"
		case "pwsbox51"
			strLine = "P-5"
		case "pwsbox52"
			strLine = "C-5"	
		case else
			strLine = "ETC"	
	end select
	arrBOM_Sub_Stock(CNT1,0) = strLine
			
	arrBOM_Sub_Stock(CNT1,1) = RS1("PRD_Partno")
	
	SQL = "select BS_MAN_Qty = BS_MAN_Qty+BS_ASM_Qty from tbBOM_Sub where BS_D_No = '"&RS1("PRD_Partno")&"'"
	RS2.Open SQL,sys_DBCon
	if RS2.Eof or RS2.Bof then
		arrBOM_Sub_Stock(CNT1,2) = 0
	else
		arrBOM_Sub_Stock(CNT1,2) = RS2("BS_MAN_Qty")
	end if
	RS2.Close
	
	arrBOM_Sub_Stock(CNT1,3) = RS1("cntPartNo")
	
	
	SQL = "select sumPR_Amount = sum(PR_Amount) from tbProcess_Record where BOM_Sub_BS_D_No = '"&RS1("PRD_Partno")&"' and PR_Process = 'DLV' and PR_Work_Date = '"&s_Work_Date&"'"
	RS2.Open SQL,sys_DBCon
	if RS2.Eof or RS2.Bof then
		arrBOM_Sub_Stock(CNT1,4) = 0
	elseif isnull(RS2("sumPR_Amount")) then
		arrBOM_Sub_Stock(CNT1,4) = 0
	else
		arrBOM_Sub_Stock(CNT1,4) = RS2("sumPR_Amount")
	end if
	RS2.Close
	
	CNT1 = CNT1 + 1
	RS1.MoveNext
loop
nRawCount = RS1.RecordCount
RS1.Close

set RS2 = nothing
set RS1 = nothing
%>

<script language="javascript">
var strHTML = "";
strHTML += "<table width=100% cellpadding=0 cellspacing=1 bgcolor='white' style='color:white;font-size:37px;text-align:center;font-weight:bold'>";
strHTML += "<col width=200px></col>";
strHTML += "<col></col>";
strHTML += "<col width=200px></col>";
strHTML += "<col width=200px></col>";
strHTML += "<col width=200px></col>";

<%
for CNT1 = 0 to nRawCount
	if arrBOM_Sub_Stock(CNT1,0) <> "ETC" and arrBOM_Sub_Stock(CNT1,1) <> ""  then

		if arrBOM_Sub_Stock(CNT1,1) = nCurrentP1 and arrBOM_Sub_Stock(CNT1,0) = "P-1" then
			strBG_Color = "green"
		elseif arrBOM_Sub_Stock(CNT1,1) = nCurrentP2 and arrBOM_Sub_Stock(CNT1,0) = "P-2" then
			strBG_Color = "green"
		elseif arrBOM_Sub_Stock(CNT1,1) = nCurrentP3 and arrBOM_Sub_Stock(CNT1,0) = "P-3" then
			strBG_Color = "green"
		elseif arrBOM_Sub_Stock(CNT1,1) = nCurrentP4 and arrBOM_Sub_Stock(CNT1,0) = "P-4" then
			strBG_Color = "green"
		elseif arrBOM_Sub_Stock(CNT1,1) = nCurrentP5 and arrBOM_Sub_Stock(CNT1,0) = "P-5" then
			strBG_Color = "green"
		elseif arrBOM_Sub_Stock(CNT1,1) = nCurrentC5 and arrBOM_Sub_Stock(CNT1,0) = "C-5" then
			strBG_Color = "green"
		else
			strBG_Color = "black"
		end if
%>
strHTML += "<tr bgcolor=black>";
//라인
<%
		if oldLine <> arrBOM_Sub_Stock(CNT1,0) and arrBOM_Sub_Stock(CNT1,0) = "P-1" then
%>
strHTML += "	<td rowspan='<%=nRowSpanP1%>' align=center><%=arrBOM_Sub_Stock(CNT1,0)%></td>";
<%
		elseif oldLine <> arrBOM_Sub_Stock(CNT1,0) and arrBOM_Sub_Stock(CNT1,0) = "P-2" then
%>
strHTML += "	<td rowspan='<%=nRowSpanP2%>' align=center><%=arrBOM_Sub_Stock(CNT1,0)%></td>";
<%
		elseif oldLine <> arrBOM_Sub_Stock(CNT1,0) and arrBOM_Sub_Stock(CNT1,0) = "P-3" then
%>
strHTML += "	<td rowspan='<%=nRowSpanP3%>' align=center><%=arrBOM_Sub_Stock(CNT1,0)%></td>";
<%
		elseif oldLine <> arrBOM_Sub_Stock(CNT1,0) and arrBOM_Sub_Stock(CNT1,0) = "P-4" then
%>
strHTML += "	<td rowspan='<%=nRowSpanP4%>' align=center><%=arrBOM_Sub_Stock(CNT1,0)%></td>";
<%
		elseif oldLine <> arrBOM_Sub_Stock(CNT1,0) and arrBOM_Sub_Stock(CNT1,0) = "P-5" then
%>
strHTML += "	<td rowspan='<%=nRowSpanP5%>' align=center><%=arrBOM_Sub_Stock(CNT1,0)%></td>";
<%
		elseif oldLine <> arrBOM_Sub_Stock(CNT1,0) and arrBOM_Sub_Stock(CNT1,0) = "C-5" then
%>
strHTML += "	<td rowspan='<%=nRowSpanC5%>' align=center><%=arrBOM_Sub_Stock(CNT1,0)%></td>";
<%
		end if
		oldLine = arrBOM_Sub_Stock(CNT1,0)
%>
//파트넘버
//strHTML += "	<td bgcolor='<%=strBG_Color%>'><span style='cursor:hand' onclick=\"javascript:Pop_Print('<%=arrBOM_Sub_Stock(CNT1,1)%>');\">"
//strHTML += 			"<%=arrBOM_Sub_Stock(CNT1,1)%>";
//strHTML += 		"</span></td>";
strHTML += "	<td bgcolor='<%=strBG_Color%>'>"
strHTML += 			"<%=arrBOM_Sub_Stock(CNT1,1)%>";
strHTML += 		"</td>";

<%
		if arrBOM_Sub_Stock(CNT1,2) > arrBOM_Sub_Stock(CNT1,3) then
			arrBOM_Sub_Stock(CNT1,2) = arrBOM_Sub_Stock(CNT1,3)
		end if
		
		arrBOM_Sub_Stock(CNT1,2) = arrBOM_Sub_Stock(CNT1,3) - arrBOM_Sub_Stock(CNT1,4)
%>
//재고수량
strHTML += "	<td bgcolor='<%=strBG_Color%>' align=center><%=arrBOM_Sub_Stock(CNT1,2)%></td>";
//생산수량
strHTML += "	<td bgcolor='<%=strBG_Color%>' align=center><%=arrBOM_Sub_Stock(CNT1,3)%></td>";
//출하수량
strHTML += "	<td bgcolor='<%=strBG_Color%>' align=center><%=arrBOM_Sub_Stock(CNT1,4)%></td>";
strHTML += "</tr>";
<%
	end if
next
%>
</script>

<script language="javascript">
parent.idContent.innerHTML = strHTML;
var nScroll = <%=nRawCount%>*43-600;
//if(nScroll > 0)
	//parent.scrollTo(0,nScroll);
function fRun()
{
	if(document.readyState == "complete")
	{
		location.reload();
	}
	else
	{
		setTimeout("fRun()",30000);
	}
}
fRun();
</script>

<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	
	