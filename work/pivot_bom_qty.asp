<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->

<%
dim RS1
dim RS2
dim RS3
dim SQL

dim BOM_Sub_BS_D_No
dim Parts_P_P_No
dim BQ_Qty

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")
set RS3 = Server.CreateObject("ADODB.RecordSet")

SQL = "select distinct BOM_Sub_BS_D_No from tbBOM_Qty order by BOM_Sub_BS_D_No asc"
RS1.Open SQL,sys_DBCon
%>
<table border=1>
<tr>
	<td>ÀÚÀç</td>
	
<%
do until RS1.Eof
%>
	<td><%=RS1("BOM_Sub_BS_D_No")%></td>
<%
	RS1.MoveNext
loop
RS1.Close
%>
</tr>
<%
SQL = "select distinct Parts_P_P_No from tbBOM_Qty order by Parts_P_P_No asc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	Parts_P_P_No = RS1("Parts_P_P_No")
%>
<tr>
	<td><%=Parts_P_P_No%></td>
<%
	SQL = "select distinct BOM_Sub_BS_D_No from tbBOM_Qty order by BOM_Sub_BS_D_No asc"
	RS2.Open SQL,sys_DBCon
	do until RS2.Eof
		BOM_Sub_BS_D_No = RS2("BOM_Sub_BS_D_No")
		SQL = "select top 1 BQ_Qty from tbBOM_Qty where Parts_P_P_No = '"&Parts_P_P_No&"' and BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"'"
		RS3.Open SQL,sys_DBCon
		if RS3.Eof or RS3.Bof then
			BQ_Qty = 0
		else
			BQ_Qty = RS3("BQ_Qty")
		end if
%>
	<td><%=BQ_Qty%></td>
<%
		RS3.Close
		RS2.MoveNext
	loop
	RS2.Close
%>
</tr>
<%
	RS1.MoveNext
loop
RS1.Close
%>
<tr>
</tr>
</table>
<%
set RS1 = nothing
%>
<!-- #include virtual = "/header/db_tail.asp" -->