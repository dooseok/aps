<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_Work_Date
dim s_Process

s_Work_Date = request("s_Work_Date")
s_Process = request("s_Process")
%>
<html>
<head>
<script language=javascript>
function ifrmContent_Reload()
{
	ifrmContent.location.href="mtr_Content_process_state_all_record.asp?s_Multi_YN=<%=Request("s_Multi_YN")%>&s_Process=<%=s_Process%>&s_Work_Date=<%=s_Work_Date%>";
}
</script>
</head>
<body topmargin=0 leftmargin=0 bgcolor=black>
<div id="idContent" style="width:100%;height:100%;background-color:black;"></div>
</body>
<form name="frmTemp" method="post" action="#">
	<input type="hidden" name="sum1" value="0">
	<input type="hidden" name="sum2" value="0">
	<input type="hidden" name="sum3" value="0">
	<input type="hidden" name="sum4" value="0">
	<input type="hidden" name="sum5" value="0">
	<input type="hidden" name="sumTotal" value="0">
</form>
</html>
<%
if gM_ID = "shindk" then
%>
<iframe src="mtr_Content_process_state_all_record_reloader.asp" width=1000px height=50px frameborder=1></iframe>
<iframe name="ifrmContent" src="mtr_Content_process_state_all_record.asp?s_Multi_YN=<%=Request("s_Multi_YN")%>&s_Process=<%=s_Process%>&s_Work_Date=<%=s_Work_Date%>" width=1000px height=800px frameborder=1></iframe>
<%
else
%>
<iframe src="mtr_Content_process_state_all_record_reloader.asp" width=0px height=0px frameborder=0></iframe>
<iframe name="ifrmContent" src="mtr_Content_process_state_all_record.asp?s_Multi_YN=<%=Request("s_Multi_YN")%>&s_Process=<%=s_Process%>&s_Work_Date=<%=s_Work_Date%>" width=0px height=0px frameborder=0></iframe>
<%
end if
%>
<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	