<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
dim SQL
dim RS1

dim objName
dim P_P_No
dim s_Partner_P_Name

dim strPrice

objName 			= replace(Request("objName"),"Parts_P_P_No","POD_Price")
P_P_No				= Request("P_P_No")
s_Partner_P_Name	= Request("s_Partner_P_Name")

set RS1 = Server.CreateObject("ADODB.RecordSet")

'동일한 거래처, 파트넘버의 가격데이타 조회
SQL = "select MP_Price from tbParts_Price where Parts_P_P_No ='"&P_P_No&"' and Partner_P_Name='"&Request("s_Partner_P_Name")&"'"
response.write SQL
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	strPrice = 0
else
	strPrice = RS1(0)
end if
RS1.Close
set RS1 = nothing
%>
<script language="javascript">

if (typeof(parent.frmCommonList.<%=objName%>)=="object")
	parent.frmCommonList.<%=objName%>.value = "<%=strPrice%>";
	
if (typeof(parent.frmCommonListReg.<%=objName%>)=="object")
	parent.frmCommonListReg.<%=objName%>.value = "<%=strPrice%>";
	
</script>



<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->