<%
call Mail_Sender(Request("Email"),Request("Title"),Request("Text"),Request("URL"))
%>
<script language="javascript">
alert("발송이 완료되었습니다.");
self.close();
</script>

<%
sub Mail_Sender(Email,Title,Text,URL)
	dim objMail

	dim MailBody
	MailBody	= getMailBody(Text,URL)

	Set objMail = Server.CreateObject("CDO.Message")

	objMail.From		= "엠에스이<parkhk@msekorea.com>"
	objMail.To			= Email&";엠에스이<parkhk@msekorea.com>"
	objMail.Subject		= Title
	objMail.HTMLBody	= MailBody

	objMail.Send
	Set objMail = Nothing
end sub
%>

<%
function getMailBody(Text,URL)
	dim MailBody_Top
	dim MailBody_Content
	dim MailBody_Bottom
	dim MailBody
	
	dim objMessage
	
	dim MailBody_Start
	dim MailBody_End
	
	MailBody_Top		= ""
	MailBody_Bottom		= ""
	
	set objMessage		= Server.CreateObject("CDO.Message")
	objMessage.CreateMHTMLBody URL, 31
	MailBody_Content	= objMessage.HTMLBody
	
	'MailBody_Content = replace(MailBody_Content,"src=""/","src=""http://www.hanuljasu.com/")

	'MailBody_Start	= instr(MailBody_Content,"<!--MailBody_Start-->")
	'MailBody_End	= instr(MailBody_Content,"<!--MailBody_End-->")

	'if int(MailBody_End) > int(MailBody_Start) then
	'	MailBody_Content = mid(MailBody_Content, MailBody_Start, MailBody_End - MailBody_Start)
	'end if

	MailBody_Content = replace(MailBody_Content,"<td","<td align='center' ")
	MailBody = MailBody_Top & MailBody_Content & MailBody_Bottom
	set objMessage = nothing
	getMailBody = MailBody
end function
%>