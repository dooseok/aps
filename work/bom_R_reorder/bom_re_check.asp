<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
dim SQL
dim RS1

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(code) from tbBOM_table where bRun='0'"
RS1.Open SQL,sys_DBCon
response.write "<h1>"&3912-RS1(0)&"/3912 ("&round((3912-RS1(0))/3912*100)&"%)</h1>"
RS1.Close
SQL = "select top 1 idx from tbBOM_Qty_Sort order by idx desc"
RS1.Open SQL,sys_DBCon
response.write "<h1>"&RS1(0)&"</h1>"
RS1.Close
set RS1 = nothing

response.write "<h1>"&now()&"</h1>"
%>
<!-- #include virtual = "/header/db_tail.asp" -->
