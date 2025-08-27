<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim CNT1
dim CNT2

dim SQL
dim RS1

dim strPI_Code
dim strParts_P_P_No
dim strParts_Desc
dim strParts_Spec
dim strPI_To_Date
dim strPI_In_Date
dim strPI_Issued_Date
dim strPI_Price
dim strPI_Qty
dim strSum_Price_Qty
dim strPartner_P_Name
dim strPI_State
dim strPI_Payment_Type
dim strPI_Remark

dim arrPI_Code
dim arrParts_P_P_No
dim arrParts_Desc
dim arrParts_Spec
dim arrPI_To_Date
dim arrPI_In_Date
dim arrPI_Issued_Date
dim arrPI_Price
dim arrPI_Qty
dim arrSum_Price_Qty
dim arrPartner_P_Name
dim arrPI_State
dim arrPI_Payment_Type
dim arrPI_Remark

dim Total_Sum_Price_Qty
dim Number
dim oldPartner_P_Name

set RS1 = Server.CreateObject("ADODB.RecordSet")

strPI_Code = Request("strChecked_Value")

if right(strPI_Code,1) = "," then
	strPI_Code = left(strPI_Code,len(strPI_Code)-1)
end if

SQL = 		"select "&vbcrlf
SQL = SQL & "	PI_Code, "&vbcrlf
SQL = SQL & "	Parts_P_P_No, "&vbcrlf
SQL = SQL & "	Parts_Desc, "&vbcrlf
SQL = SQL & "	Parts_Spec, "&vbcrlf
SQL = SQL & "	PI_To_Date, "&vbcrlf
SQL = SQL & "	PI_In_Date, "&vbcrlf
SQL = SQL & "	PI_Issued_Date, "&vbcrlf
SQL = SQL & "	PI_Price, "&vbcrlf
SQL = SQL & "	PI_Qty, "&vbcrlf
SQL = SQL & "	Sum_Price_Qty, "&vbcrlf
SQL = SQL & "	Partner_P_Name, "&vbcrlf
SQL = SQL & "	PI_State, "&vbcrlf
SQL = SQL & "	PI_Payment_Type, "&vbcrlf
SQL = SQL & "	PI_Remark "&vbcrlf
SQL = SQL & "from "&vbcrlf
SQL = SQL & "	vwPI_List "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PI_Code in ("&strPI_Code&") "&vbcrlf
SQL = SQL & "order by "&vbcrlf
SQL = SQL & "	Partner_P_Name "&vbcrlf

RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strPI_Code			= strPI_Code			& RS1("PI_Code")			& "|/|"
	strParts_P_P_No		= strParts_P_P_No		& RS1("Parts_P_P_No")		& "|/|"
	strParts_Desc		= strParts_Desc			& RS1("Parts_Desc")			& "|/|"
	strParts_Spec		= strParts_Spec			& RS1("Parts_Spec")			& "|/|"
	strPI_To_Date		= strPI_To_Date			& RS1("PI_To_Date")			& "|/|"
	strPI_In_Date		= strPI_In_Date			& RS1("PI_In_Date")			& "|/|"
	strPI_Issued_Date	= strPI_Issued_Date		& RS1("PI_Issued_Date")		& "|/|"
	strPI_Price			= strPI_Price			& RS1("PI_Price")			& "|/|"
	strPI_Qty			= strPI_Qty				& RS1("PI_Qty")				& "|/|"
	strSum_Price_Qty	= strSum_Price_Qty		& RS1("Sum_Price_Qty")		& "|/|"
	strPartner_P_Name	= strPartner_P_Name		& RS1("Partner_P_Name")		& "|/|"
	strPI_State			= strPI_State			& RS1("PI_State")			& "|/|"
	strPI_Payment_Type	= strPI_Payment_Type	& RS1("PI_Payment_Type")	& "|/|"
	strPI_Remark		= strPI_Remark			& RS1("PI_Remark")			& "|/|"
	RS1.MoveNext
loop
RS1.Close
set RS1 = nothing

arrPI_Code			= split(strPI_Code,			"|/|")
arrParts_P_P_No		= split(strParts_P_P_No,	"|/|")
arrParts_Desc		= split(strParts_Desc,		"|/|")
arrParts_Spec		= split(strParts_Spec,		"|/|")
arrPI_To_Date		= split(strPI_To_Date,		"|/|")
arrPI_In_Date		= split(strPI_In_Date,		"|/|")
arrPI_Issued_Date	= split(strPI_Issued_Date,	"|/|")
arrPI_Price			= split(strPI_Price,		"|/|")
arrPI_Qty			= split(strPI_Qty,			"|/|")
arrSum_Price_Qty	= split(strSum_Price_Qty,	"|/|")
arrPartner_P_Name	= split(strPartner_P_Name,	"|/|")
arrPI_State			= split(strPI_State,		"|/|")
arrPI_Payment_Type	= split(strPI_Payment_Type,	"|/|")
arrPI_Remark		= split(strPI_Remark,	"|/|")
%>

<%
oldPartner_P_Name = ""
Total_Sum_Price_Qty = 0
for CNT1 = 0 to ubound(arrPI_Code)-1
	
	if oldPartner_P_Name <> arrPartner_P_Name(CNT1) then
		if oldPartner_P_Name <> "" then
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
	<td>&nbsp;</td>
</tr>
			<%
			next
			call Tail(oldPartner_P_Name,Total_Sum_Price_Qty)
			Total_Sum_Price_Qty = 0
		end if
		call Header(arrPartner_P_Name(CNT1))
		
		Number = 1
	end if
%>
<tr bgcolor=white>
	<td><%=Number%>&nbsp;</td>
	<td><%=arrParts_P_P_No(CNT1)%>&nbsp;</td>
	<td><%=arrParts_Desc(CNT1)%>&nbsp;</td>
	<td><%=arrParts_Spec(CNT1)%>&nbsp;</td>
	<td align=right><%=arrPI_Qty(CNT1)%>&nbsp;</td>
	<td>EA</td>
	<td align=right><%=formatnumber(arrPI_Price(CNT1),2)%>&nbsp;</td>
	<td align=right><%=formatnumber(arrSum_Price_Qty(CNT1),2)%>&nbsp;</td>
	<td><%=arrPI_Remark(CNT1)%>&nbsp;</td>
	<td><%=arrPI_To_Date(CNT1)%>&nbsp;</td>
</tr>
<%
	Number = Number + 1
	Total_Sum_Price_Qty = Total_Sum_Price_Qty + arrSum_Price_Qty(CNT1)
	oldPartner_P_Name = arrPartner_P_Name(CNT1)
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
	<td>&nbsp;</td>
</tr>
<%
next
call Tail(oldPartner_P_Name,Total_Sum_Price_Qty)
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
		<tr bgcolor=white align=left>
			<td><%=left(date(),4)%>년 <%=mid(date(),6,2)%>월 <%=right(date(),2)%>일</td>
		</tr>
		</table>
	</td>
	<td width=240px align=right valign=bottom>
		<img src="/img/blank.gif" width=10px height=2px><br>
		<table class="pi_print_2" width=200px cellpadding=0 cellspacing=0 border=1 bgcolor="#333333" style="table-layout:fixed" style="border-collapse:collapse">
		<tr bgcolor=white>
			<td width=30px rowspan=2>결<br>제</td>
			<td>담 당</td>
			<td>검 토</td>
			<td>검 토</td>
			<td>승 인</td>
		</tr>
		<tr bgcolor=white height=40px>
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
	<td align=left>&nbsp;김유숙</td>
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
<col width=110px></col>
<col width=160px></col>
<col></col>
<col width=50px></col>
<col width=35px></col>
<col width=90px></col>
<col width=100px></col>
<col width=120px></col>
<col width=100px></col>
<tr bgcolor=white>
	<td>번호</td>
	<td>품번</td>
	<td>품명</td>
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
sub Tail(strP_Name, Total_Sum_Price_Qty)
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
	
	Total_Sum_Price_Qty	= formatnumber(Total_Sum_Price_Qty, 4)
	Tax					= formatnumber(Total_Sum_Price_Qty * 0.1, 4)
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
	<td align=right><%=Total_Sum_Price_Qty%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor=white>
	<td width=80px>팩스번호</td>
	<td><%=P_Fax%>&nbsp;</td>
	<td width=80px>부<img src="/img/blank.gif" width=6px height=1px>가<img src="/img/blank.gif" width=6px height=1px>세</td>
	<td align=right><%=Tax%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor=white>
	<td width=80px>MAIL주소</td>
	<td><%=P_EMail%>&nbsp;</td>
	<td width=80px>총<img src="/img/blank.gif" width=6px height=1px>금<img src="/img/blank.gif" width=6px height=1px>액</td>
	<td align=right><%=formatnumber(Total_Sum_Price_Qty + (Total_Sum_Price_Qty * 0.1), 4)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
<img src="/img/blank.gif" width=10px height=3px><br>
<%
end sub
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->