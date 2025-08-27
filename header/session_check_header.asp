<%
sub Redirect_M_Login()
	dim login_URL
	dim URL
	dim Query_String
	dim Request_Fields

	URL				= Request.ServerVariables("URL")
	Query_String	= "?login_URL=Yes"
		
	for each Request_Fields in Request.QueryString
		Query_String = Query_String & "&"&Request_Fields&"="&server.URLEncode(Request(Request_Fields))
	next
	for each Request_Fields in Request.Form
		Query_String = Query_String & "&"&Request_Fields&"="&server.URLEncode(Request(Request_Fields))
	next
	
	login_URL	= URL & Query_String
%>

<%
if Request("autologin")="yes" then
%>
<form name="redirect_form" action="/member/m_login.asp?autologin=yes" method="post">
<input type="hidden" name="autologinID" value="<%=request("autologinID")%>">
<input type="hidden" name="autologinPWD" value="<%=request("autologinPWD")%>">
<%
else
%>
<form name="redirect_form" action="/member/m_login.asp" method="post">
<%
end if
%>
<input type="hidden" name="login_URL" value="<%=login_URL%>">
</form>
<script language="javascript">
redirect_form.submit();
</script>
<%
end sub
%>

<%
function Check_Authority_YN()
	dim CNT1
	
	dim URL	
	URL = lcase(Request.ServerVariables("URL"))
	
	dim strCheck_Authority_YN
	strCheck_Authority_YN = "Y"

	dim arrBasicDataAuthoriy	
	arrBasicDataAuthoriy = split(BasicDataAuthoriy,";")
	
	for CNT1 = 0 to ubound(arrBasicDataAuthoriy)
		if instr(arrBasicDataAuthoriy(CNT1),Request.cookies("ADMIN")("M_Authority")) and instr(lcase(arrBasicDataAuthoriy(CNT1)),URL) > 0 then
			strCheck_Authority_YN = "N"
		end if
	next

	Check_Authority_YN = strCheck_Authority_YN
end function
%>

<%
dim arrREMOTE_ADDR 
dim strREMOTE_ADDR_CLASS
arrREMOTE_ADDR = split(Request.ServerVariables("REMOTE_ADDR"),".")
strREMOTE_ADDR_CLASS = arrREMOTE_ADDR(0)&"."&arrREMOTE_ADDR(1)&"."

if gM_ID = "" or (instr("-shindk-rnd-no7008-shindh-simjy-eng-leejw-smtech-dstech-leehg-sales-qa-ydw-smt-parksg-iqc-","-"&gM_ID&"-") = 0 and instr("-192.168.-112.222.-","-"&strREMOTE_ADDR_CLASS&"-") = 0) then
	if gM_ID <> "" then
%>
<script language="javascript">
	alert("전산담당자에게 원격지 접속권한을 요청하시기 바랍니다.");
</script>
<%		
	end if
	call Redirect_M_Login()
else
%>
