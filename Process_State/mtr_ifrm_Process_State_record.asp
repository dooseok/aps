<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_Work_Date
dim s_Line
dim s_Jaje_YN

s_Line = request("s_Line")
s_Jaje_YN = request("s_Jaje_YN")
%>
<html>
<head>
	
<script language="javascript">
function Pop_Print(strBOM_Sub_BS_D_No)
{
	window.open("/bom/b_parts_out_sheet.asp?part=ChulGo&s_BOM_Sub_BS_D_No="+strBOM_Sub_BS_D_No+"&s_P_Work_Type=MAN");
}

function ifrmContent_Reload()
{
	ifrmContent.location.href="mtr_Content_process_state_record.asp?s_Line=<%=s_Line%>&s_Jaje_YN=<%=s_Jaje_YN%>";
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
<iframe src="mtr_Content_process_state_record_reloader.asp" width=600px height=600px frameborder=1></iframe>
<iframe name="ifrmContent" src="mtr_Content_process_state_record.asp?s_Line=<%=s_Line%>&s_Jaje_YN=<%=s_Jaje_YN%>" width=600px height=600px frameborder=1></iframe>
<%
else
%>
<iframe src="mtr_Content_process_state_record_reloader.asp" width=0px height=0px frameborder=0></iframe>
<iframe name="ifrmContent" src="mtr_Content_process_state_record.asp?s_Line=<%=s_Line%>&s_Jaje_YN=<%=s_Jaje_YN%>" width=0px height=0px frameborder=0></iframe>
<%
end if
%>

<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	