<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim CR_Code
dim byWho
dim toState

CR_Code = Request("CR_Code")
byWho = Request("byWho")
toState = Request("toState")

SQL = "update tbCost_Report set "&byWho&" = '"&toState&"' where CR_Code = " & CR_Code
'response.write SQL
sys_DBCon.execute(SQL)
%>

<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->