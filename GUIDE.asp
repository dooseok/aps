<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- include virtual = "/header/layout_full_header.asp" -->

<%
dim strURL

strURL = Request("strURL")
if strURL = "" then
	if gM_ID = "smtech" then
		strURL = "/bom/b_parts_out_sheet.asp"
	elseif gM_ID = "dstech" then
		strURL = "/bom/b_parts_out_sheet.asp"
	else
		strURL = "main.asp"
end if	
end if
%>

<frameset rows="88,*" border=0>
<frame src="/menu/inc_menu_top.asp?strURL=<%=Server.URLEncode(strURL)%>" name="frameTop" scrolling="no" noresize style="border-bottom:1px solid #999999">
<frame src="<%=strURL%>" name="frameMain" scrolling="auto" noresize>
</frameset>

<!-- include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->