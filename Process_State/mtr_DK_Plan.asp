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
<col></col>
<col width=190px></col>
<col width=240px></col>
<col width=190px></col>
<col width=190px></col>
<col width=150px></col>
<col width=17px></col>
<tr height=50px bgcolor=skyblue>
	<td style="color:navy">PNO</td>
	<td style="color:navy">DK<br>계획수량</td>
	<td style="color:navy">DK<br>계획시간</td>
	<td style="color:navy">MSE<br>생산실적</td>
	<td style="color:navy">MSE<br>출하실적</td>
	<td style="color:navy">MSE<br>재고</td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td colspan=7><iframe src="mtr_ifrm_dk_plan.asp?s_Work_Date=<%=s_Work_Date%>" width=100% height=100% frameborder=0 scrolling=yes></iframe></td>
</tr>
</table>
</body>
</html>



<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	