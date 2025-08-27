<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->

<%
dim s_Material_M_P_No
dim RS1
dim SQL

s_Material_M_P_No = Request("s_Material_M_P_No")

dim MT_Date
dim MT_Qty_Now
dim MT_Desc

SQL = "select MT_Date, MT_Qty_Now, MT_Desc from tbMaterial_Transaction where Material_M_P_No = '"&s_Material_M_P_No&"' order by MT_Code desc"
set RS1 = Server.CreateObject("ADODB.RecordSet")
%>

<table width=450px cellpadding=0 cellspacing=1 border=0 bgcolor=dimgray>
<tr bgcolor=#eeeeee>
	<td colspan=3><B><%=s_Material_M_P_No%></td>
</tr>
<tr bgcolor=#eeeeee onclick="javascript:parent.divMaterial_Qty_Log_Popup_List.style.display='none';">
	<td width=180px><b>날짜</b></td>
	<td width=90px><b>재고</b></td>
	<td width=180px><b>비고</b>&nbsp;&nbsp;&nbsp;<span style="cursor:hand;" onclick="javascript:parent.divMaterial_Qty_Log_Popup_List.style.display='none';">▼</span>&nbsp;&nbsp;</td>
</tr>

<%
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	MT_Date		= RS1("MT_Date")
	MT_Qty_Now	= RS1("MT_Qty_Now")
	MT_Desc		= RS1("MT_Desc")
%>
<tr bgcolor=white>
	<td><%=MT_Date%></td>
	<td><%=MT_Qty_Now%></td>
	<td><%=MT_Desc%></td>
<%
	RS1.MoveNext
loop
RS1.Close
%>
	</td>
</tr>
</table>
<%
set RS1 = nothing
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->