<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<%
dim s_PRD_Line
dim arrWG_Res
dim arrWG_MCDelay
dim arrWG_SlideDelay
dim arrWG_SlideDelay_Main
dim arrWG_Auto_YN
dim WG_Res
dim WG_ResX
dim WG_ResY
dim WG_MCDelay
dim WG_SlideDelay
dim WG_SlideDelay_Main
dim WG_Auto_YN
dim CNT1
dim SQL

s_PRD_Line				= request("s_PRD_Line")
arrWG_Res				= split(request("WG_Res"),", ")
arrWG_MCDelay			= split(request("WG_MCDelay"),", ")
arrWG_SlideDelay		= split(request("WG_SlideDelay"),", ")
arrWG_SlideDelay_Main	= split(request("WG_SlideDelay_Main"),", ")
arrWG_Auto_YN			= split(request("WG_Auto_YN"),", ")

for CNT1 = 0 to ubound(arrWG_Res)
	WG_Res				= trim(arrWG_Res(CNT1))
	WG_ResX				= left(WG_Res, instr(WG_Res,"x")-1)
	WG_ResY				= right(WG_Res, len(WG_Res)-instr(WG_Res,"x"))
	WG_MCDelay			= trim(arrWG_MCDelay(CNT1))
	WG_SlideDelay		= trim(arrWG_SlideDelay(CNT1))
	WG_SlideDelay_Main	= trim(arrWG_SlideDelay_Main(CNT1))
	WG_Auto_YN			= trim(arrWG_Auto_YN(CNT1))
	SQL = "update tbWorkGuide set WG_ResX = "&WG_ResX&", WG_ResY = "&WG_ResY&", WG_MCDelay = "&WG_MCDelay&", WG_SlideDelay = "&WG_SlideDelay&", WG_SlideDelay_Main = "&WG_SlideDelay_Main&", WG_Auto_YN = '"&WG_Auto_YN&"' where WG_Pos = "&CNT1+1&" and PRD_Line='"&s_PRD_Line&"'"
	sys_DBCon.execute(SQL)
next


%>

<form name="frmRedirect" action="workguide_launcher.asp?s_PRD_Line=<%=s_PRD_Line%>" method=post>
</form>
<script language="javascript">
frmRedirect.submit();
</script>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->