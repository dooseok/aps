<%
dim objConn
dim intCampaignRecipientID

dim strFilePath
dim strFileName

strFilePath = Request.QueryString("filepath")
strFileName = Mid(strFilePath, InStrRev(strFilePath, "\")+1)

if left(strFilePath,1) = "/" then
    strFilePath = server.Mappath(strFilePath)&"\"
end if
    
if strFileName <> "" then
    Response.Buffer = False
    Dim objStream
    set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Type = 1 'adTypeBinary
    objStream.Open
    application("ERROR2") = strFilePath
    objStream.LoadFromFile(strFilePath)
    Response.ContentType = "application/x-unknown"
    Response.Addheader "Content-Disposition", "attachment; filename=""" & strFileName & """"
    Response.BinaryWrite objStream.Read
    objStream.Close
    set objStream = Nothing

end if
%>

