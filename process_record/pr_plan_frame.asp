<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_Process
dim s_Frame_Large_YN

dim frameHeight_1
dim frameHeight_2

s_Process			= Request("s_Process")
s_Frame_Large_YN	= Request("s_Frame_Large_YN")

if s_Frame_Large_YN = "Y" then
	frameHeight_1 = "300"
	frameHeight_2 = "450"	
else
	frameHeight_1 = "200"
	frameHeight_2 = "280"
end if
%>

<table width=100% height=100% cellpadding=0 cellspacing=0>
<tr height=<%=frameHeight_1%>px>
	<td height=<%=frameHeight_1%>px><iframe name="ifrmPR_Monitor" src="/process_record/pr_monitor.asp?s_Process=<%=s_Process%>" frameborder=0 style="border:1px solid #999999" scrolling=auto width="100%" height="100%"></iframe></td>
</tr>
<tr height=1px><td></td></tr>
<tr height=<%=frameHeight_2%>px>
	<td height=<%=frameHeight_2%>px><iframe src="/lge_plan/lp_view.asp?s_Edit_Process=<%=s_Process%>" frameborder=0 style="border:1px solid #999999" scrolling=auto width="100%" height="100%"></iframe></td>
</tr>
</table>



<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->