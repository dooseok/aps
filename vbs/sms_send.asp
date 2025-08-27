<%
dim strPhone
dim arrPhone

dim strMSG
dim send_YN
dim CNT1


strPhone	= Request("strPhone")

strMSG		= Request("strMSG")
send_YN		= Request("send_YN")

if strPhone = "" then
	strPhone = "01087305740"
end if
arrPhone = split(strPhone,";")

if strMSG = "" then
	strMSG = "SMSÅ×½ºÆ®_"&now()
end if
%>

<html>
<head>
</head>
<body>	
	
<form name="frmSMS_Send" method="post" action="http://sms.nanuminet.com/utf8.php">
<input type="hidden" name="sms_id" value="moonsunge">
<input type="hidden" name="sms_pw" value="ms7750">
<%
for CNT1 = 0 to ubound(arrPhone)
%>
<input type="hidden" name="phone[]" value="<%=arrPhone(CNT1)%>">
<%
next
%>
<input type="hidden" name="callback" value="01047195740">
<input type="hidden" name="senddate" value="">
<input type="hidden" name="return_url" value="http://kr.msekorea.com:1080/vbs/sms_send.asp">
<input type="hidden" name="return_data" value="">
<input type="hidden" name="msg[]" value="<%=strMSG%>">
</form>

</body>
</html>

<script language="javascript">
<%
if request("get_code") <> "" then
	response.write "/"&request("get_code")&"/"
else
%>
frmSMS_Send.submit();
//window.open('about:blank','_self').close();
<%
end if
%>
</script>