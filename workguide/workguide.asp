<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_tb_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<div>
	<center>
	<div class="page-header">
	<h2>라인별 작업지도서 런쳐 선택</h2>
	</div>
	<table>
<%
dim CNT1
dim SQL
dim RS1
dim RS2
dim strLine
set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")
SQL = "select LI_Line from tblLine_Info where LI_Line not like '%cbox%' order by LI_Line"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	
	strLine = ucase(RS1("LI_Line"))
%>
	<tr height=40px>
		<td width=100px><%=strLine%>라인</td>
		<td width=100px><input type="button" class="btn btn-sm btn-primary" value="전공정" onclick="javascript:window.open('workguide_launcher.asp?s_PRD_Line=<%=strLine%>&s_Base_WG_Pos=1&s_Checked=1,7');"></td>
		<td width=100px><input type="button" class="btn btn-sm btn-primary" value="후공정" onclick="javascript:window.open('workguide_launcher.asp?s_PRD_Line=<%=strLine%>&s_Base_WG_Pos=1&s_Checked=8,15');"></td>
	</tr>
<%
	
	SQL = "select top 1 PRD_Line from tbWorkGuide where PRD_Line = '"&strLine&"'"
	RS2.Open SQL,sys_DBCon
	if RS2.Eof or RS2.Bof then
		for CNT1 = 1 to 15
			SQL = ""
			SQL = SQL	& "insert into tbWorkGuide ("
			SQL = SQL & "PRD_Line, WG_Pos, WG_ResX, WG_ResY, WG_MCDelay, WG_SlideDelay, WG_SlideDelay_Main, WG_Auto_YN "
			SQL = SQL & ") values ("
			SQL = SQL & "'"&strLine&"',"
			SQL = SQL & CNT1&","
			SQL = SQL & "1920,"
			SQL = SQL & "1080,"
			SQL = SQL & "0,"
			SQL = SQL & "10,"
			SQL = SQL & "10,"
			SQL = SQL & "'Y' "
			SQL = SQL & ") "
			sys_DBcon.execute(SQL)
		next
	end if
	RS2.Close

	RS1.MoveNext
loop
RS1.Close
set RS2 = nothing
set RS1 = nothing
%>
	</table>
	</center>
</div>


<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->