<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_Work_Date
dim s_Line
dim s_Process

s_Process	= Request("s_Process")
if s_Work_Date = "" then
	s_Work_Date = date()
end if
%>
<html>
<head>

</head>
<body topmargin=0 leftmargin=0 bgcolor=white onload="startTime()">
<table width=100% height=60px cellpadding="0" cellspacing="0" border=0 bgcolor=white>
<col width=50%></col>
<col width=50%></col>
<tr height=60px>
	<td align=left style="font-size:40pt"><b><div id=idCurrentDate></div></b></td>
	<td align=right style="font-size:40pt"><b><div id=idCurrentTime></div></b></td>
</tr>
</table>
<table width=100% cellpadding="0" cellspacing="5" border=0 bgcolor=black>
<col width=16%></col>
<col width=14%></col>
<col width=14%></col>
<col width=14%></col>
<col width=14%></col>
<col width=14%></col>
<col width=14%></col>
<tr bgcolor=skyblue>
	<td align=center style="color:navy;font-size:50pt"><b>라인</b></td>
	<td align=center style="color:navy;font-size:50pt"><b>계획</b></td>
	<td align=center style="color:navy;font-size:50pt"><b>목표</b></td>
	<td align=center style="color:navy;font-size:50pt"><b>실적</b></td>
	<td align=center style="color:navy;font-size:50pt"><b>달성률</b></td>
	<td align=center style="color:navy;font-size:50pt"><b>무작업</b></td>
	<td align=center style="color:navy;font-size:50pt"><b>비고</b></td>
</tr>
<tr bgcolor=white>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr bgcolor=white>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr bgcolor=white>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr bgcolor=white>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr bgcolor=white>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
</table>
<br>

<%
if gM_ID = "shindk" then
%>
<iframe src="new_mtr_Content_process_state_all_record_reloader.asp" width=1000px height=50px frameborder=0></iframe>
<iframe name="ifrmContent" src="new_mtr_Content_process_state_all_record.asp?s_Multi_YN=<%=Request("s_Multi_YN")%>&s_Process=<%=s_Process%>&s_Work_Date=<%=s_Work_Date%>" width=1000px height=100px frameborder=1></iframe>
<%
else
%>
<iframe src="new_mtr_Content_process_state_all_record_reloader.asp" width=0px height=0px frameborder=0></iframe>
<iframe name="ifrmContent" src="new_mtr_Content_process_state_all_record.asp?s_Multi_YN=<%=Request("s_Multi_YN")%>&s_Process=<%=s_Process%>&s_Work_Date=<%=s_Work_Date%>" width=0px height=0px frameborder=0></iframe>
<%
end if
%>
</body>


<script language=javascript>
function startTime() {
    var today=new Date();
    var yy=today.getFullYear();
    var mm=today.getMonth();
    var dd=today.getDate();
    var h=today.getHours();
    var m=today.getMinutes();
    mm = make2Digit(mm);
    dd = make2Digit(dd);
    m = make2Digit(m);
    //alert("<style='color:red'>"+yy+"</div>" + "년 " + "<div style='color:red'>"+mm+"</div>" + "월 " + "<div style='color:red'>"+dd+"</div>" + "일");
    document.getElementById('idCurrentDate').innerHTML = "&nbsp;&nbsp;<span style='color:red'>"+yy+"</span>" + "년 " + "<span style='color:red'>"+mm+"</span>" + "월 " + "<span style='color:red'>"+dd+"</span>" + "일";
    document.getElementById('idCurrentTime').innerHTML = "현재시간 " + "<span style='color:red'>"+h+"</span>" + ":" + "<span style='color:red'>"+m+"</span>&nbsp;&nbsp;";
    var t = setTimeout(function(){startTime()},3000);
}

function make2Digit(i) {
    if (i<10) {i = "0" + i};  // add zero in front of numbers < 10
    return i;
}
</script>



<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	