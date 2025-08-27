<!-- #include virtual = "/header/asp_header.asp" -->
<%
Response.Buffer = true
Response.ContentType = "application/vnd.ms-excel"
Response.CacheControl = "public"	
%>
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<%
dim FileName

dim strFileName
dim arrFileName

strFileName			= Request("strFileName")

if instr(strFileName,"/") > 0 then
	arrFileName = split(strFileName,"/")
	FileName = arrFileName(ubound(arrFileName))
else
	FileName = strFileName
end if
FileName = replace(FileName,".asp","")
FileName = replace(replace(replace(replace(now(),"-",""),"오후","PM"),"오전","AM")," ","") & "_" & FileName

Response.AddHeader "Content-Disposition","attachment;filename="&FileName&".xls"

dim RS1
dim CNT1

dim SQL
dim strSelectName
dim arrSelectName
dim strSelect
dim arrSelect
dim strTable
dim strWhere
dim strOrderBy
dim strGroupBy
dim strHaving

SQL					= Request("SQL")
strSelectName		= Request("strSelectName")
strSelect			= Request("strSelect")
strTable			= Request("strTable")
strWhere			= Request("strWhere")
strOrderBy			= Request("strOrderBy")
strGroupBy			= Request("strGroupBy")
strHaving			= Request("strHaving")
arrSelectName		= split(strSelectName,",")
arrSelect			= split(strSelect,",")

dim strColumnName
if SQL = "" then
	SQL = " select " & strSelect & " from " & strTable
	if trim(strWhere) <> "" then
		SQL = SQL & " where " & strWhere
	end if
	if trim(strGroupBy) <> "" then
		SQL = SQL & " group by " & strGroupBy
	end if
	if trim(strHaving) <> "" then
		SQL = SQL & " Having " & strHaving
	end if
	if trim(strOrderBy) <> "" then
		SQL = SQL & " order by " & strOrderBy
	end if
	
	SQL = replace(SQL,"''","'")
end if
set RS1 = Server.CreateObject("ADODB.RecordSet") 
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
%>
<script language="javascript">
alert("조회결과가 없습니다.")
window.close();
</script>
<%
else
	for CNT1 = 0 to ubound(arrSelectName)
		strColumnName=arrSelectName(CNT1)
		if strColumnName <> "체크" and strColumnName <> "작업" and strColumnName <> "삭제" then
			strColumnName = replace(strColumnName,"<br>"," ")
			strColumnName = replace(strColumnName,"<img src='/img/black' width=1px height=5px>","")
			strColumnName = replace(strColumnName,"  "," ")
			response.write strColumnName
			response.write vbtab
		end if
	next
	response.write vbcrlf	
	do until RS1.Eof
		for CNT1 = 0 to ubound(arrSelect)
			if instr(arrSelect(CNT1),"=") > 0 then
				response.write RS1(left(arrSelect(CNT1),instr(arrSelect(CNT1)," =")-1))
				response.write vbtab 
			elseif instr(arrSelect(CNT1),"BQ_Qty=BQ_Qty*") > 0 then
				response.write RS1("BQ_Qty") * int( right(arrSelect(CNT1),len(arrSelect(CNT1)) - instr(arrSelect(CNT1),"*")))
				response.write vbtab
			elseif RS1(trim(arrSelect(CNT1))) = "" or isnull(RS1(trim(arrSelect(CNT1)))) then
				response.write vbtab
			else
				response.write replace(RS1(trim(arrSelect(CNT1))),vbcrlf,"")
				'response.write RS1(trim(arrSelect(CNT1)))
				response.write vbtab 
			end if
		next
		response.write vbcrlf	
		RS1.MoveNext
	loop
end if
RS1.close
set RS1 = nothing 
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->