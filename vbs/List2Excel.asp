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
strHaving			= Request("strHaving")c
arrSelectName		= split(strSelectName,",")
arrSelect			= split(strSelect,",")

SQL= ""

SQL = SQL & "select "
SQL = SQL & "BOM_Sub_BS_D_No, "
SQL = SQL & "Parts_P_P_No, "
SQL = SQL & "P_Desc = (select top 1 P_Desc from tbParts where P_P_No = Parts_P_P_No), "
SQL = SQL & "P_Spec = (select top 1 P_Spec from tbParts where P_P_No = Parts_P_P_No), "
SQL = SQL & "P_Maker = (select top 1 P_Maker from tbParts where P_P_No = Parts_P_P_No), "
SQL = SQL & "BQ_Qty, "
SQL = SQL & "BQ_Order, "
SQL = SQL & "BQ_Remark, "
SQL = SQL & "BQ_CheckSum "

SQL = SQL & " from tbBOM_Qty where  "
SQL = SQL & "LEFT(BOM_Sub_BS_D_No,3) not in ('AKB', 'EAV') and "
SQL = SQL & "LEFT(BOM_Sub_BS_D_No,4) not in ('6711', '3911') and "
SQL = SQL & "BQ_Qty > 0  "


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
	do until RS1.Eof
		for CNT1 = 0 to 8	
				response.write RS1(CNT1)
				response.write vbtab
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