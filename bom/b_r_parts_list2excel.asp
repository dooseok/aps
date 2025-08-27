<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
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
FileName = right(replace(date(),"-",""),6) & "_" & FileName

dim RS1
dim CNT1

dim strAppend
dim arrAppend
dim arrAppend2

dim s_Parts_P_P_No

strAppend		= Request("strAppend")
s_Parts_P_P_No	= Request("s_Parts_P_P_No")

if strAppend = "" then
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
	Response.AddHeader "Content-Disposition","attachment;filename="&FileName&".xls"


	response.write "Assy PartNo"
	response.write vbtab
	response.write "Material PartNo"
	response.write vbtab
	response.write "Material R-PartNo"
	response.write vbcrlf	
	
	arrAppend = split(strAppend,"|%|")
	for CNT1 = 0 to ubound(arrAppend)-1
		arrAppend2 = split(arrAppend(CNT1),"|/|")
		response.write arrAppend2(0)
		response.write vbtab
		response.write s_Parts_P_P_No
		response.write vbtab
		response.write arrAppend2(1)
		response.write vbcrlf
	next
end if
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->