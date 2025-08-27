<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->



<%
dim RS1
dim SQL

dim nRawCount

'반복문에 사용하기 위한 변수 선언
dim CNT1
dim CNT2
dim CNT3

dim s_Date
dim Array_Size_DK_Plan
dim Array_Size_PWS_Raw_Data
dim Array_Size_PWS_DLV_Data
dim Array_Size_PWS_STK_Data

dim apply_Flag_YN

s_Date = "2010-07-28"

Array_Size_DK_Plan		= 450
Array_Size_PWS_Raw_Data	= 100
Array_Size_PWS_DLV_Data	= 100
Array_Size_PWS_STK_Data	= 100

dim arrPWS_DK_Plan(1000,6)
dim arrPWS_Raw_Data(100,1)
dim arrPWS_DLV_Data(100,1)
dim arrPWS_STK_Data(100,1)

dim arrPWS_Raw_Data_Add(0,1)
arrPWS_Raw_Data_Add(0,0)	= "EBR71383302"
arrPWS_Raw_Data_Add(0,1)	= 120

dim arrPWS_DLV_Data_Add(0,1)
arrPWS_DLV_Data_Add(0,0)	= "EBR71383302"
arrPWS_DLV_Data_Add(0,1)	= 0


set RS1 = Server.CreateObject("ADODB.RecordSet")









'------------------------------------DK계획 가져오기----------------------------------------
SQL = "select * from tbDK100728 order by DKDate, DKTime"
RS1.Open SQL,sys_DBCon

CNT1 = 0
do until RS1.Eof
	arrPWS_DK_Plan(CNT1,0)	= RS1("PNO")	'파트넘버
	arrPWS_DK_Plan(CNT1,1)	= RS1("DKQty")	'DK계획수량
	arrPWS_DK_Plan(CNT1,2)	= RS1("DKDate")	'DK계획날짜
	arrPWS_DK_Plan(CNT1,3)	= RS1("DKTime")	'DK계획시각
	arrPWS_DK_Plan(CNT1,4)	= 0				'생산실적
	arrPWS_DK_Plan(CNT1,5)	= 0				'출하실적
	arrPWS_DK_Plan(CNT1,6)	= 0				'재고수량
	CNT1 = CNT1 + 1
	RS1.MoveNext
loop
RS1.Close

'------------------------------------생산실적(실제) 가져오기----------------------------------------
SQL = ""
SQL = SQL & "select "&vbcrlf
SQL = SQL & "	PRD_PartNo, "&vbcrlf
SQL = SQL & "	cntPRD_Barcode = count(PRD_Barcode) "&vbcrlf
SQL = SQL & "from "&vbcrlf
SQL = SQL & "	tbPWS_Raw_Data "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PRD_Input_Date = '"&s_Date&"' and "&vbcrlf
SQL = SQL & "	PRD_Line = 'pwsbox3' "&vbcrlf
SQL = SQL & "group by "&vbcrlf
SQL = SQL & "	PRD_PartNo "&vbcrlf
RS1.Open SQL,sys_DBCon

CNT1 = 0
do until RS1.Eof
	arrPWS_Raw_Data(CNT1,0)	= RS1("PRD_PartNo")		'파트넘버
	arrPWS_Raw_Data(CNT1,1)	= RS1("cntPRD_Barcode")	'실적수량
	CNT1 = CNT1 + 1
	RS1.MoveNext
loop
RS1.Close


'------------------------------------생산실적(증분) 추가하기----------------------------------------
CNT3 = 0
for CNT1 = 0 to ubound(arrPWS_Raw_Data_Add)
	
	apply_Flag_YN = "N"
	for CNT2 = 0 to ubound(arrPWS_Raw_Data)
		if arrPWS_Raw_Data_Add(CNT1,0) = arrPWS_Raw_Data(CNT2,0) then
			arrPWS_Raw_Data(CNT2,1) = int(arrPWS_Raw_Data(CNT2,1)) + int(arrPWS_Raw_Data_Add(CNT1,1))
			apply_Flag_YN = "Y"
		end if
	next
	
	if apply_Flag_YN = "N" then
		arrPWS_Raw_Data(Array_Size_PWS_Raw_Data - CNT3, 0)	= arrPWS_Raw_Data_Add(CNT1,0)
		arrPWS_Raw_Data(Array_Size_PWS_Raw_Data - CNT3, 1)	= arrPWS_Raw_Data_Add(CNT1,1)
		CNT3 = CNT3 + 1
	end if 
next

'------------------------------------출하실적(실제) 가져오기----------------------------------------
SQL = ""
SQL = SQL & "select "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No, "&vbcrlf
SQL = SQL & "	sumPR_Amount = sum(PR_Amount) "&vbcrlf
SQL = SQL & "from "&vbcrlf
SQL = SQL & "	tbProcess_Record "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PR_Work_Date = '"&s_Date&"' and "&vbcrlf
SQL = SQL & "	PR_Process = 'DLV' "&vbcrlf
SQL = SQL & "group by "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No "&vbcrlf
RS1.Open SQL,sys_DBCon

CNT1 = 0
do until RS1.Eof
	arrPWS_DLV_Data(CNT1,0)	= RS1("BOM_Sub_BS_D_No")	'파트넘버
	arrPWS_DLV_Data(CNT1,1)	= RS1("sumPR_Amount")		'출하수량
	CNT1 = CNT1 + 1
	RS1.MoveNext
loop
RS1.Close

'------------------------------------출하실적(증분) 추가하기----------------------------------------
CNT3 = 0
for CNT1 = 0 to ubound(arrPWS_DLV_Data_Add)
	
	apply_Flag_YN = "N"
	for CNT2 = 0 to ubound(arrPWS_DLV_Data)
		if arrPWS_DLV_Data_Add(CNT1,0) = arrPWS_DLV_Data(CNT2,0) then
			arrPWS_DLV_Data(CNT2,1) = int(arrPWS_DLV_Data(CNT2,1)) + int(arrPWS_DLV_Data_Add(CNT1,1))
			apply_Flag_YN = "Y"
		end if
	next
	
	if apply_Flag_YN = "N" then
		arrPWS_DLV_Data(Array_Size_PWS_DLV_Data - CNT3, 0)	= arrPWS_DLV_Data_Add(CNT1,0)
		arrPWS_DLV_Data(Array_Size_PWS_DLV_Data - CNT3, 1)	= arrPWS_DLV_Data_Add(CNT1,1)
		CNT3 = CNT3 + 1
	end if 
next
'---------------------------------------------------------------------------------------------------



'------------------------------------------재고배열 가져오기-----------------------------------------------
SQL = "select distinct PNO from tbDK100727"
RS1.Open SQL,sys_DBCon

CNT1 = 0

do until RS1.Eof
	arrPWS_STK_Data(CNT1,0) = RS1("PNO")
	arrPWS_STK_Data(CNT1,1) = 0
	
	for CNT2 = 0 to ubound(arrPWS_Raw_Data)
		if arrPWS_STK_Data(CNT1,0) = arrPWS_Raw_Data(CNT2,0) then
			arrPWS_STK_Data(CNT1,1) = arrPWS_Raw_Data(CNT2,1)
		end if
	next
	
	for CNT2 = 0 to ubound(arrPWS_DLV_Data)
		if arrPWS_STK_Data(CNT1,0) = arrPWS_DLV_Data(CNT2,0) then
			arrPWS_STK_Data(CNT1,1) = arrPWS_STK_Data(CNT1,1) - arrPWS_DLV_Data(CNT2,1)
		end if
	next
	
	CNT1 = CNT1 + 1
	RS1.MoveNext
loop
RS1.Close

'-------------------------------------실적 데이터 반영------------------------------------------
for CNT1=0 to ubound(arrPWS_DK_Plan)
	if arrPWS_DK_Plan(CNT1,0) <> "" then
		for CNT2=0 to ubound(arrPWS_Raw_Data)
			if arrPWS_Raw_Data(CNT2,0) <> "" and arrPWS_DK_Plan(CNT1,0) = arrPWS_Raw_Data(CNT2,0) then
				if int(arrPWS_DK_Plan(CNT1,1)) > int(arrPWS_Raw_Data(CNT2,1)) then
					arrPWS_DK_Plan(CNT1,4) = arrPWS_Raw_Data(CNT2,1)
					arrPWS_Raw_Data(CNT2,1) = 0
				else
					arrPWS_DK_Plan(CNT1,4) = arrPWS_DK_Plan(CNT1,1)
					arrPWS_Raw_Data(CNT2,1) = arrPWS_Raw_Data(CNT2,1) - arrPWS_DK_Plan(CNT1,1)
				end if
			end if
		next
'-------------------------------------출하 데이터 반영------------------------------------------	
		for CNT2=0 to ubound(arrPWS_DLV_Data)
			if arrPWS_DLV_Data(CNT2,0) <> "" and arrPWS_DK_Plan(CNT1,0) = arrPWS_DLV_Data(CNT2,0) then
				if int(arrPWS_DK_Plan(CNT1,1)) > int(arrPWS_DLV_Data(CNT2,1)) then
					arrPWS_DK_Plan(CNT1,5) = arrPWS_DLV_Data(CNT2,1)
					arrPWS_DLV_Data(CNT2,1) = 0
				else
					arrPWS_DK_Plan(CNT1,5) = arrPWS_DK_Plan(CNT1,1)
					arrPWS_DLV_Data(CNT2,1) = arrPWS_DLV_Data(CNT2,1) - arrPWS_DK_Plan(CNT1,1)
				end if
			end if
		next
'-------------------------------------재고 데이터 반영------------------------------------------			
		for CNT2=0 to ubound(arrPWS_STK_Data)
			if arrPWS_STK_Data(CNT2,0) <> "" and arrPWS_DK_Plan(CNT1,0) = arrPWS_STK_Data(CNT2,0) then
				arrPWS_DK_Plan(CNT1,6) = arrPWS_STK_Data(CNT2,1)
			end if
		next
	end if
next
%>

<script language="javascript">
var strHTML = "";
strHTML += "<table width=100% cellpadding=0 cellspacing=1 bgcolor='white' style='color:white;font-size:37px;text-align:center;font-weight:bold'>";
strHTML += "<col></col>";
strHTML += "<col width=190px></col>";
strHTML += "<col width=240px></col>";
strHTML += "<col width=190px></col>";
strHTML += "<col width=190px></col>";
strHTML += "<col width=150px></col>";

<%	
nRawCount = 0
for CNT1=0 to ubound(arrPWS_DK_Plan)
	if len(arrPWS_DK_Plan(CNT1,0)) = "11" then
		if arrPWS_DK_Plan(CNT1,5) > 0 or arrPWS_DK_Plan(CNT1,6) > 0 then
			nRawCount = nRawCount + 1
		end if
%>
strHTML += "<tr bgcolor=black>";
strHTML += "	<td align=center><%=arrPWS_DK_Plan(CNT1,0)%></td>";
strHTML += "	<td align=center><%=arrPWS_DK_Plan(CNT1,1)%></td>";
strHTML += "	<td align=center><%=int(mid(arrPWS_DK_Plan(CNT1,2),6,2))%>.<%=right(arrPWS_DK_Plan(CNT1,2),2)%>&nbsp;<%=left(arrPWS_DK_Plan(CNT1,3),2)%>:<%=mid(arrPWS_DK_Plan(CNT1,3),3,2)%></td>";
strHTML += "	<td align=center><%=arrPWS_DK_Plan(CNT1,4)%></td>";
strHTML += "	<td align=center><%=arrPWS_DK_Plan(CNT1,5)%></td>";
strHTML += "	<td align=center><%=arrPWS_DK_Plan(CNT1,6)%></td>";
strHTML += "</tr>";
<%
	end if
next
%>
strHTML += "</table>";

parent.idContent.innerHTML = strHTML;
//var nScroll = <%=nRawCount%>*20-600;
//if(nScroll > 0)
	//parent.scrollTo(0,nScroll);
	parent.scrollTo(0,0);
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

<%
set RS1 = nothing
%>

<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	
	