<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<% 
dim RS1
dim CNT1

dim SQL
dim value1
dim value2
dim value3
dim Column
SQL			= trim(Request("SQL"))
value1		= trim(Request("value1"))
value2		= trim(Request("value2"))
value3		= trim(Request("value3"))

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = replace(SQL,"$value1$",value1)
SQL = replace(SQL,"$value2$",value2)
SQL = replace(SQL,"$value3$",value3)
'response.write sQL
RS1.Open SQL,sys_DBCon

if RS1.Eof or RS1.Bof then
%>
<script language="javascript">
alert("조회결과가 없습니다.")
window.close();
</script>
<%
else
	
	Response.Buffer = false
	Response.Expires = 0
	Response.ContentType = "application/vnd.ms-excel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition","attachment;filename=CustomQuery.xls"
	for each Column in RS1.Fields
		response.write Column.Name
		response.write vbtab
	next
	response.write vbcrlf
	
	do until RS1.Eof
		for each Column in RS1.Fields
		    response.write Column.value
			response.write vbtab
		next
		response.write vbcrlf
		
		RS1.MoveNext
	loop
	RS1.Close
end if
set RS1 = nothing
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->