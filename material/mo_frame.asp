<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- include virtual = "/header/layout_full_header.asp" -->

<frameset name="frmMoFrame" cols="100px,*" border=10>
<frame src="/material/m_list.asp?s_callby=mo_frame" name="frameLeft" scrolling="auto">
<frame src="/material/mo_list.asp?postback_yn=Y&s_edit_mode_yn=checked&s_MO_Reg_ID=<%=Request.cookies("ADMIN")("M_Name")%>" name="frameMain" scrolling="auto">
</frameset>

<!-- include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->