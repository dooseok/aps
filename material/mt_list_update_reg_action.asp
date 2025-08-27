<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim Material_M_P_No
dim MT_Date
dim MT_Qty_In
dim MT_Qty_Out
dim MT_Qty_Update
dim MT_Desc
dim MT_Qty_Last
dim MT_Qty_Now
dim MT_Price

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")

Material_M_P_No	= trim(Request("Material_M_P_No"))
MT_Date			= trim(Request("MT_Date"))
MT_Qty_Update	= trim(Request("MT_Qty_Update"))
MT_Desc			= trim(Request("MT_Desc"))

set RS1 = Server.CreateObject("ADODB.RecordSet")
rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	rem DB 업데이트
	
	SQL = "select Partner_P_Name,M_Qty,M_Price from tbMaterial where M_P_No = '"&Material_M_P_No&"'"
	RS1.Open SQL,sys_DBCon
	MT_Qty_Last		= RS1("M_Qty") 
	
	if cdbl(MT_Qty_Last) > cdbl(MT_Qty_Update) then
		MT_Qty_In	= 0
		MT_Qty_Out	= cdbl(MT_Qty_Last) - cdbl(MT_Qty_Update)
	elseif cdbl(MT_Qty_Last) < cdbl(MT_Qty_Update) then
		MT_Qty_In	=  cdbl(MT_Qty_Update) - cdbl(MT_Qty_Last)
		MT_Qty_Out	= 0
	end if
		
	MT_Qty_Now		= cdbl(MT_Qty_Update)
	MT_Price		= RS1("M_Price")
	RS1.Close
	
	RS1.Open "tbMaterial_Transaction",sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("Material_M_P_No")		= Material_M_P_No
		.Fields("Partner_P_Name")		= ""
		.Fields("MT_Out_byWho")			= ""
		.Fields("MT_Date")				= MT_Date
		.Fields("MT_Price")				= MT_Price
		.Fields("MT_Qty_In")			= MT_Qty_In
		.Fields("MT_Qty_Out")			= MT_Qty_Out
		.Fields("MT_Qty_Update")		= MT_Qty_Update
		.Fields("MT_Qty_Last")			= MT_Qty_Last
		.Fields("MT_Qty_Now")			= MT_Qty_Now
		.Fields("MT_Desc")				= MT_Desc
		.Fields("MT_Reg_Date")			= now()
		.Fields("MT_Reg_ID")			= gM_ID
		.Update	
		.Close
	end with
	
	SQL = "update tbMaterial set M_Qty = "&MT_Qty_Now&" where M_P_No = '"&Material_M_P_No&"'"
	sys_DBCon.execute(SQL)
end if

rem 객체 해제
Set RS1	= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="mt_list_update.asp" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="mt_list_update.asp" method=post>

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