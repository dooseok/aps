<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<% 
rem ��������
dim RS1
dim SQL
dim CNT1

dim strError
dim strError_Temp

dim arrID_All
dim arrBP_PPH

arrID_All	= split(Request("strID_All")&" ",", ")
arrBP_PPH	= split(Request("BP_PPH")&" "	,", ")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)			= trim(arrID_All(CNT1))
	arrBP_PPH(CNT1)	= trim(arrBP_PPH(CNT1))
next

set RS1 = Server.CreateObject("ADODB.RecordSet")
rem �����޼����� ���� ��� ����ȵ�
if strError = "" then	
	
	rem DB ������Ʈ
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""

		if strError_Temp = "" then
			SQL = "update tbBOM_PPH set "
			if isnumeric(arrBP_PPH(CNT1)) then
			else
				arrBP_PPH(CNT1) = 0
			end if
			SQL = SQL & "	BP_PPH	= "&arrBP_PPH(CNT1)&" "
			SQL = SQL & "where BP_Code='"&arrID_All(CNT1)&"' "
			sys_DBCon.execute(SQL)
		end if
		
		strError = strError & strError_Temp
	next
end if
%>

<%
dim Request_Fields
dim strRequestForm
dim strRequestQueryString
for each Request_Fields in Request.Form
	if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	end if
next
for each Request_Fields in Request.QueryString
	if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	end if
next
if strError = "" then
%>
<form name="frmRedirect" action="bom_pph_list.asp" method=post>
	
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
	'strError = strError & "* �Ϻ��� ������ ��ҵǾ����ϴ�."
%>
<form name="frmRedirect" action="bom_pph_list.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>



<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->