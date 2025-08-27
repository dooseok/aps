<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<%
dim SQL

On Error Resume Next
Response.Clear
Dim objError
Set objError = Server.GetLastError()

dim E_COM
E_COM = ""
if len(CStr(objError.Number)) > 0 then
	E_COM = objError.Number&" (0x" & Hex(objError.Number) &")"
end if

dim E_FileName
E_FileName = ""
if len(CStr(objError.File)) > 0 then
	E_FileName = objError.File
end if

dim E_Line
E_Line = ""
if len(CStr(objError.Line)) > 0 then
	E_Line = objError.Line
end if

dim E_Desc
E_Desc = ""
if len(CStr(objError.Description)) > 0 then
	E_Desc = "//Desc:"&objError.Description
end if
if len(CStr(objError.ASPDescription)) > 0 then
	E_Desc = E_Desc & "<Br>//ASPdesc:"&objError.ASPDescription
end if
if len(CStr(objError.ASPCode)) > 0 then
	E_Desc = E_Desc & "<Br>//ASPCode:"&objError.ASPCode
end if
if len(CStr(objError.Category)) > 0 then
	E_Desc = E_Desc & "<Br>//Category:"&objError.Category
end if
if len(CStr(objError.Column)) > 0 then
	E_Desc = E_Desc & "<Br>//Column:"&objError.Column
end if
if len(CStr(objError.Source)) > 0 then
	E_Desc = E_Desc & "<Br>//Source:"&objError.Source
end if
E_Desc = replace(E_Desc,"'","''")
'!!!!!!!!!! get 파라미터 to query 처리 시작

dim strRequestQueryString
dim Request_Fields
strRequestQueryString = ""
for each Request_Fields in Request.QueryString
	strRequestQueryString = strRequestQueryString & "&"&Request_Fields&"="&server.URLEncode(Request(Request_Fields))
next
strRequestQueryString = strRequestQueryString & "//"
for each Request_Fields in Request.Form
	strRequestQueryString = strRequestQueryString & "&"&Request_Fields&"="&server.URLEncode(Request(Request_Fields))
next

dim E_Server
E_Server = gHOST
SQL = "insert into tblError (E_Server,E_Query,E_COM,E_FileName,E_Line,E_Desc,E_AppVar,E_Memo) values ("
SQL = SQL & "'"&E_Server&"',"
SQL = SQL & "'"&strRequestQueryString&"',"
SQL = SQL & "'"&E_COM&"',"
SQL = SQL & "'"&E_FileName&"',"
SQL = SQL & "'"&E_Line&"',"
SQL = SQL & "'"&E_Desc&"',"
SQL = SQL & "'"&replace(Application("Error"),"'","''")&"',"
SQL = SQL & "'')"
Application("Error") = ""
if gM_ID = "shindk" then
	response.write SQL&"<br>"
end if
sys_DBCon.execute(SQL)
SQL = "update tblError_Flag set EF_Flag = 1"
'response.write SQL&"<br>"
sys_DBCon.execute(SQL)


'''''''''''''''''''''''''''''''''''''''''
%>
<center>
	<br><br><br><br><br>

		<h2>에러가 발생하였습니다.</h2>
		<br>
		<h4>IT 담당자에게 에러사항이 보고되었습니다.
		<br>
		신속한 처리를 원하시면 IT 담당자에게 문의바랍니다.</h4>
		<input type="button" onclick="javascript:location.href='/index.asp'" value="첫 페이지로 이동">
	
</center>
<!-- #include Virtual = "/header/db_tail.asp" -->