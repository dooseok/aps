<!-- #include virtual = "/header/asp_header.asp" -->

<%
Response.cookies("Admin")("M_Code")				= ""
Response.cookies("Admin")("M_Channel")			= ""
Response.cookies("Admin")("M_ID")				= ""
Response.cookies("Admin")("M_Password")			= ""
Response.cookies("Admin")("M_Part")				= ""
Response.cookies("Admin")("M_Position")			= ""
Response.cookies("Admin")("M_Name")				= ""
Response.cookies("Admin")("M_Email_1")			= ""
Response.cookies("Admin")("M_Email_2")			= ""
Response.cookies("Admin")("M_HP")				= ""
Response.cookies("Admin")("M_Enter_Date")		= ""
Response.cookies("Admin")("M_Retire_Date")		= ""
Response.cookies("Admin")("M_Authority")		= ""

Response.cookies("Admin").Path					= "/"
Response.Cookies("Admin").Expires 				= Date - 1
%>
<script language="javascript">
top.location.href='/index.asp';
</script>
