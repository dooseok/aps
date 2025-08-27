<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim CNT1

dim arrPWS_Data

dim PRD_Date
dim PRD_Line

PRD_Date = trim(Request("PRD_Date"))
PRD_Line = trim(Request("PRD_Line"))

if PRD_Date = "" then
	PRD_Date = dateadd("d",-1,date())
end if
if PRD_Line = "" then
	PRD_Line = "1"
end if

arrPWS_data = getPWS_Data(PRD_Date,PRD_Line)

for CNT1 = 0 to ubound(arrPWS_Data)
	response.write arrPWS_Data(CNT1,0) & "_________"
	response.write arrPWS_Data(CNT1,1) & "_________"
	response.write arrPWS_Data(CNT1,2) & "_________"
	response.write arrPWS_Data(CNT1,3) & "_________"
	response.write arrPWS_Data(CNT1,4) & "<br>"
next
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->