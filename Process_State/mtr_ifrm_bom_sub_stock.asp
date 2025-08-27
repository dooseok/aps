<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_Work_Date
s_Work_Date = request("s_Work_Date")
%>
<html>
<head>
	
<script language="javascript">
function Pop_Print(strBOM_Sub_BS_D_No)
{
	window.open("/bom/bs_qty_chart.asp?s_BS_D_No="+strBOM_Sub_BS_D_No);
}
</script>

</head>
<body topmargin=0 leftmargin=0 bgcolor=black>
<div id="idContent" style="width:100%;height:100%;background-color:black;"></div>
</body>
</html>
<%
if gM_ID = "shindk" then
%>
<iframe src="mtr_Content_bom_sub_stock.asp?s_Work_Date=<%=s_Work_Date%>" width=1000px height=50px frameborder=0></iframe>
<%
else
%>
<iframe src="mtr_Content_bom_sub_stock.asp?s_Work_Date=<%=s_Work_Date%>" width=0px height=0px frameborder=0></iframe>
<%
end if
%>
<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	