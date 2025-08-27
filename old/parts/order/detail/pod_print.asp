<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
call usePrinter()
%>

<%
dim CNT1
dim CNT2

dim SQL
dim RS1

dim PO_Code
dim Partner_P_Name

dim strPOD_Code
dim strParts_P_P_No
dim strParts_M_Additional_Info
dim strParts_M_Spec
dim strPOD_Due_Date
dim strPOD_Price
dim strPOD_Qty
dim strCalc_Price_Qty
dim strPOD_Remark

dim arrPOD_Code
dim arrParts_P_P_No
dim arrParts_M_Additional_Info
dim arrParts_M_Spec
dim arrPOD_Due_Date
dim arrPOD_Price
dim arrPOD_Qty
dim arrCalc_Price_Qty
dim arrPOD_Remark

dim Total_Calc_Price_Qty
dim Number
dim oldPartner_P_Name

set RS1 = Server.CreateObject("ADODB.RecordSet")

PO_Code			= Request("PO_Code")
Partner_P_Name	= Request("Partner_P_Name")

SQL = 		"select * "&vbcrlf
SQL = SQL & "from "&vbcrlf
SQL = SQL & "	vwParts_Order_Detail "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	Parts_Order_PO_Code = "&PO_Code&" "&vbcrlf
SQL = SQL & "order by "&vbcrlf
SQL = SQL & "	POD_Code asc "&vbcrlf

RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strPOD_Code						= strPOD_Code					& RS1("POD_Code")					& "|/|"
	strParts_P_P_No				= strParts_P_P_No			& RS1("Parts_P_P_No")			& "|/|"
	strParts_M_Additional_Info	= strParts_M_Additional_Info	& RS1("Parts_M_Additional_Info")	& "|/|"
	strParts_M_Spec				= strParts_M_Spec			& RS1("Parts_M_Spec")			& "|/|"
	strPOD_Due_Date					= strPOD_Due_Date				& RS1("POD_Due_Date")				& "|/|"
	strPOD_Price					= strPOD_Price					& RS1("POD_Price")					& "|/|"
	strPOD_Qty						= strPOD_Qty					& RS1("POD_Qty")					& "|/|"
	strCalc_Price_Qty				= strCalc_Price_Qty				& RS1("Calc_Price_Qty")				& "|/|"
	strPOD_Remark					= strPOD_Remark					& RS1("POD_Remark")					& "|/|"
	RS1.MoveNext
loop
RS1.Close
set RS1 = nothing

arrPOD_Code						= split(strPOD_Code,					"|/|")
arrParts_P_P_No				= split(strParts_P_P_No,				"|/|")
arrParts_M_Additional_Info	= split(strParts_M_Additional_Info,	"|/|")
arrParts_M_Spec				= split(strParts_M_Spec,				"|/|")
arrPOD_Due_Date					= split(strPOD_Due_Date,				"|/|")
arrPOD_Price					= split(strPOD_Price,					"|/|")
arrPOD_Qty						= split(strPOD_Qty,						"|/|")
arrCalc_Price_Qty				= split(strCalc_Price_Qty,				"|/|")
arrPOD_Remark					= split(strPOD_Remark,					"|/|")
%>

<%
call Header(Partner_P_Name)

Number = 1
Total_Calc_Price_Qty = 0
for CNT1 = 0 to ubound(arrPOD_Code)-1
	if arrParts_P_P_No(CNT1) <> "" then
%>
<tr bgcolor=white>
	<td><%=Number%>&nbsp;</td>
	<td><%=arrParts_P_P_No(CNT1)%>&nbsp;</td>
	<td><%=arrParts_M_Spec(CNT1)%>&nbsp;</td>
	<!--<td><%=arrParts_M_Additional_Info(CNT1)%>&nbsp;</td>-->
	<td align=right><%=replace(arrPOD_Qty(CNT1),"원","")%>&nbsp;</td>
	<td>EA</td>
	<td align=right><%=formatnumber(arrPOD_Price(CNT1),2)%>&nbsp;</td>
	<td align=right><%=formatnumber(arrCalc_Price_Qty(CNT1),2)%>&nbsp;</td>
	<td><%=arrPOD_Remark(CNT1)%>&nbsp;</td>
	<td><%=arrPOD_Due_Date(CNT1)%>&nbsp;</td>
</tr>
<%
	Number = Number + 1
	Total_Calc_Price_Qty = Total_Calc_Price_Qty + arrCalc_Price_Qty(CNT1)
	end if
next
for CNT2 = 1 to 25-(Number mod 25)
%>
<tr bgcolor=white>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<%
next
'call Tail(oldPartner_P_Name,Total_Calc_Price_Qty)
call Tail(Partner_P_Name,Total_Calc_Price_Qty)
%>

<%
sub Header(strP_Name)
	dim SQL
	dim RS1

	dim P_Name
	dim P_Owner
	dim P_Zipcode
	dim P_Address
	dim P_Business_No

	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select * from tbPartner where P_Name = '"&strP_Name&"'"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		P_Name			= strP_Name

		P_Owner			= ""
		P_Address		= ""
		P_Zipcode		= ""
		P_Business_No	= ""

	else
		P_Name			= strP_Name

		P_Owner			= RS1("P_Owner")
		P_Address		= RS1("P_Address")
		P_Zipcode		= RS1("P_Zipcode")
		P_Business_No	= RS1("P_Business_No")

	end if
	RS1.Close
%>
<table class="pi_print_1" width=1040px cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed" style="font-face:굴림;font-size:12px;font-weight:bold;">
<tr bgcolor=white>
	<td>
		<table width=100% cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed">
		<tr height=50 bgcolor=white>
			<td style="font-size:40px"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;발 주 서</b></td>
		</tr>
		<tr bgcolor=white align=Left>
			<td><%=left(date(),4)%>년 <%=mid(date(),6,2)%>월 <%=right(date(),2)%>일</td>
		</tr>
		</table>
	</td>
	<td width=250px align=right valign=bottom>
		<img src="/img/blank.gif" width=10px height=2px><br>
		<table class="pi_print_2" width=240px cellpadding=0 cellspacing=0 border=1 bgcolor="#333333" style="table-layout:fixed" style="border-collapse:collapse">
		<tr bgcolor=white>
			<td width=30px rowspan=2>결<br>제</td>
			<td>담 당</td>
			<td>검 토</td>
			<td>검 토</td>
			<td>승 인</td>
		</tr>
		<tr bgcolor=white height=50px>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<img src="/img/blank.gif" width=10px height=3px><br>
<table class="pi_print_1" width=1040px cellpadding=0 cellspacing=0 border=1 bgcolor="#333333" style="table-layout:fixed" style="border-collapse:collapse">
<col width=35px></col>
<col width=80px></col>
<col width=402px></col>
<col width=35px></col>
<col width=80px></col>
<col width=402px></col>
<tr bgcolor=white>
	<td rowspan=4>공<br>급<br>자</td>
	<td>등록번호</td>
	<td align=left>&nbsp;<%=P_Business_No%></td>
	<td rowspan=4>공<br>급<br>받<br>는<br>자</td>
	<td>등록번호</td>
	<td align=left>&nbsp;<%=DefaultBusinessNo%></td>
</tr>
<tr bgcolor=white>
	<td>상<img src="/img/blank.gif" width=24px height=1px>호</td>
	<td align=left>&nbsp;<%=P_Name%></td>
	<td>상<img src="/img/blank.gif" width=24px height=1px>호</td>
	<td align=left>&nbsp;엠에스이(주)</td>
</tr>
<tr bgcolor=white>
	<td>대<img src="/img/blank.gif" width=6px height=1px>표<img src="/img/blank.gif" width=6px height=1px>자</td>
	<td align=left>&nbsp;<%=P_Owner%></td>
	<td>대<img src="/img/blank.gif" width=6px height=1px>표<img src="/img/blank.gif" width=6px height=1px>자</td>
	<td align=left>&nbsp;김유숙, 양재순</td>
</tr>
<tr bgcolor=white>
	<td>주<img src="/img/blank.gif" width=24px height=1px>소</td>
	<td align=left>&nbsp;<%=P_Address%><%if P_Zipcode <> "" then%>&nbsp;&nbsp;&nbsp;우)<%=P_Zipcode%><%end if%></td>
	<td>주<img src="/img/blank.gif" width=24px height=1px>소</td>
	<td align=left>&nbsp;경남 마산시 양덕동 973-1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;tel)055-298-9448</td>
</tr>
</table>
<img src="/img/blank.gif" width=10px height=3px><br>
<table class="pi_print_1" width=1040px cellpadding=0 cellspacing=0 border=1 bgcolor="#333333" style="table-layout:fixed" style="border-collapse:collapse">
<col width=35px></col>
<col width=120px></col>
<col></col>
<col width=70px></col>
<col width=35px></col>
<col width=90px></col>
<col width=120px></col>
<col width=120px></col>
<col width=100px></col>
<tr bgcolor=white>
	<td>번호</td>
	<td>품번</td>
	<td>규격</td>
	<td>수량</td>
	<td>단위</td>
	<td>단가</td>
	<td>금액</td>
	<td>비고</td>
	<td>납기일</td>
</tr>
<%
	set RS1 = nothing
end sub
%>

<%
sub Tail(strP_Name, Total_Calc_Price_Qty)
	dim SQL
	dim RS1

	dim P_Name
	dim P_Tel
	dim P_Fax
	dim P_EMail

	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select * from tbPartner where P_Name = '"&strP_Name&"'"
	
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		P_Name			= strP_Name

		P_Tel			= ""
		P_Fax			= ""
		P_EMail			= ""

	else
		P_Name			= strP_Name

		P_Tel			= RS1("P_Tel")
		P_Fax			= RS1("P_Fax")
		P_EMail			= RS1("P_EMail")

	end if
	RS1.Close

	dim Tax

	Total_Calc_Price_Qty	= formatnumber(Total_Calc_Price_Qty, 4)
	Tax					= formatnumber(Total_Calc_Price_Qty * 0.1, 4)
%>
</table>
<img src="/img/blank.gif" width=10px height=3px><br>
<table class="pi_print_1" width=1040px cellpadding=0 cellspacing=0 border=1 bgcolor="#333333" style="table-layout:fixed" style="border-collapse:collapse">
<tr bgcolor=white>
	<td rowspan=3 align=left valign=top>비고</td>
	<td rowspan=3 align=left valign=top>담당자확인</td>
	<td width=80px>전화번호</td>
	<td><%=P_Tel%>&nbsp;</td>
	<td width=80px>공급가액</td>
	<td align=right><%=formatnumber(Total_Calc_Price_Qty, 2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor=white>
	<td width=80px>팩스번호</td>
	<td><%=P_Fax%>&nbsp;</td>
	<td width=80px>부<img src="/img/blank.gif" width=6px height=1px>가<img src="/img/blank.gif" width=6px height=1px>세</td>
	<td align=right><%=formatnumber(Tax, 2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor=white>
	<td width=80px>MAIL주소</td>
<%
if trim(P_EMail) <> "" then
	if instr(P_EMail,";") > 0 then
%>
	<td><%=left(P_EMail,instr(P_EMail,";")-1)%>&nbsp;</td>
<%
	else
%>
	<td><%=P_EMail%>&nbsp;</td>
<%		
	end if
else
%>
	<td>&nbsp;</td>
<%
end if
%>
	<td width=80px>총<img src="/img/blank.gif" width=6px height=1px>금<img src="/img/blank.gif" width=6px height=1px>액</td>
	<td align=right><%=formatnumber(Total_Calc_Price_Qty + (Total_Calc_Price_Qty * 0.1), 2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
<img src="/img/blank.gif" width=10px height=3px><br>
<%
end sub
%>

<script language="javascript">
factory.printing.header				= "";
factory.printing.footer				= "";
factory.printing.portrait			= false;
factory.printing.leftMargin			= 0.5;
factory.printing.rightMargin		= 0.5;
factory.printing.topMargin			= 0.5;
factory.printing.bottomMargin		= 0.5;
factory.printing.print(true, window);
</script>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->