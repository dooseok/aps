<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = false%>
<!-- #include file="db_connection.inc" --> 
{"data": [
<%

Set oConn=Server.CreateObject("ADODB.Connection") 
oConn.Open strConnect
strSQL1 = "SELECT top 50 dbo.Artists.artistname, dbo.Recordings.RecordingTitle, dbo.Tracks.TrackTitle, dbo.Tracks.TrackFileName FROM  dbo.Artists INNER JOIN dbo.Recordings ON dbo.Artists.artistid = dbo.Recordings.ArtistID INNER JOIN dbo.Tracks ON dbo.Recordings.RecordingID = dbo.Tracks.RecordingID "
Set oRs1=oConn.Execute(strSQL1,lngRecs,1)


' send data to an array'
if not oRs1.eof then
	myArray=oRs1.getrows()
end if

oRs1.Close
Set oRs1 = Nothing
oConn.Close
Set oConn = Nothing



for r = LBound (myArray,2) to UBound(myArray,2)
	artistname 		= myArray(0, r)
	RecordingTitle 	= myArray(1, r)
	TrackTitle 		= myArray(2, r)
	TrackFileName 	= myArray(3, r)

	TrackFileUrl = replace(TrackFileName,"M:\Music\","\mp3\")

	'M:\Music\MP3MusicAlbums\!!!\Louden Up Now\01-Louden Up Now-When The Going Gets Tough The Tough Gets Krazee.mp3

	' convert to'
	'\mp3\MP3MusicAlbums\!!!\Louden Up Now\01-Louden Up Now-When The Going Gets Tough The Tough Gets Krazee.mp3

	TrackFileLink = "<a href="&"[x]"&TrackFileUrl&"[x]"&">"&TrackTitle&"</a>"

    theURL = replace("["&chr(34)&artistname&chr(34)&","&chr(34)&RecordingTitle&chr(34)&","&chr(34)&TrackTitle&chr(34)&","&chr(34)&TrackFileLink&chr(34)&"]","\","\\")
    theURL = replace(theURL,"[x]","\"&chr(34))
    response.write theURL

        if r < UBound(myArray,2) then
    	response.write ","
    end if


next


%>
]}