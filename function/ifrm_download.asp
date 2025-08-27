<%
'response.redirect 
'Response.Buffer = False
filepath = Request.QueryString("filepath")


if instr(filepath,".xls") > 0 then
	filepath = replace(filepath, "d:\my_website\msekorea\admin", "")
	filepath = replace(filepath,"\","/")
	response.redirect filepath
	response.end
end if



filename = Mid(filepath, InStrRev(filepath, "\")+1)

'If InStr(Request.ServerVariables("HTTP_USER_AGENT"), "Chrome") > 0 Then
'Else
	Response.AddHeader "Content-Disposition","attachment;filename=""" & Server.URLEncode(filename) & """"
'end if

set objFS	= Server.CreateObject("Scripting.FileSystemObject")
set objF	= objFS.GetFile(filepath)
'Response.AddHeader "Content-Length", objF.Size
set objF	= nothing
set objFS	= nothing

Response.ContentType	= "application/unknown" 
Response.CacheControl	= "public" 
	
Set objDownload	= Server.CreateObject("DEXT.FileDownload")
objDownload.Download filepath
Set objDownload	= Nothing
%>
<script language="javascript">
history.back();
</script>