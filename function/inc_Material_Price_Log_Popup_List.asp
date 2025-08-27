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

dim Partner_P_Name
dim Material_M_P_No
dim MPL_Price
dim MPL_Price_LGE
dim MPL_Reg_Date

SQL = "select Partner_P_Name, Material_M_P_No, MPL_Reg_Date, MPL_Price, MPL_Price_LGE from tbMaterial_Price_Log where Material_M_P_No = '"&s_Material_M_P_No&"' order by MPL_Code desc"
set RS1 = Server.CreateObject("ADODB.RecordSet")
%>

<table width=430px cellpadding=0 cellspacing=1 border=0 bgcolor=dimgray>
<tr bgcolor=#eeeeee onclick="javascript:parent.divMaterial_Price_Log_Popup_List.style.display='none';">
	<td width=160px><b>거래처</b></td>
	<td width=70px><b>단가</b></td>
	<td width=70px><b>인증가</b></td>
	<td align=right><b>날짜</b>&nbsp;&nbsp;&nbsp;<span style="cursor:hand;" onclick="javascript:parent.divMaterial_Price_Log_Popup_List.style.display='none';">▼</span>&nbsp;&nbsp;</td>
</tr>

<%
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	Partner_P_Name	= RS1("Partner_P_Name")
	if Partner_P_Name = "" then
		Partner_P_Name = "[거래처미상]"
	end if
	MPL_Price		= RS1("MPL_Price")
	MPL_Price_LGE	= RS1("MPL_Price_LGE")
	Material_M_P_No	= RS1("Material_M_P_No")
	MPL_Reg_Date	= RS1("MPL_Reg_Date")
%>
<tr bgcolor=white>
	<td><%=Partner_P_Name%></td>
	<td><%=MPL_Price%></td>
	<td><%=MPL_Price_LGE%></td>
	<td><%=MPL_Reg_Date%></td>
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