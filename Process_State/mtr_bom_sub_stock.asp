<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_Work_Date

s_Work_Date = Request("s_Work_Date")
if s_Work_Date = "" then
	s_Work_Date = date()
end if
%>
<html>
<head>
	
</head>
<body topmargin=0 leftmargin=0 bgcolor=black>

<table width=100% height=1000px cellpadding=0 cellspacing=1 bgcolor="white" style="color:white;font-size:42px;text-align:center;font-weight:bold">
<col width=200px></col>
<col></col>
<col width=200px></col>
<col width=200px></col>
<col width=200px></col>
<col width=17px></col>
<tr height=85px bgcolor=skyblue style="color:navy">
	<td>LINE</td>
	<td>PartNO</td>
	<td>재고</td>
	<td>생산</td>
	<td>출하</td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td colspan=7><iframe src="mtr_ifrm_bom_sub_stock.asp?s_Work_Date=<%=s_Work_Date%>" width=100% height=100% frameborder=0 scrolling=yes></iframe></td>
</tr>
</table>
</body>
</html>



<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	