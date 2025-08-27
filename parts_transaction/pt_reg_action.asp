<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim PT_Date
dim Parts_P_P_No
dim PT_Qty
dim PT_Type
dim PT_Description
dim Partner_P_Name
dim PT_Price
dim Member_M_ID

dim oldP_Qty
dim oldPartner_P_Name
dim oldP_CO_Price
dim oldP_LGE_Price
dim oldP_MSE_Price

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")

PT_Date			= Request("PT_Date")
Parts_P_P_No	= Request("Parts_P_P_No")
PT_Qty			= Request("PT_Qty")
PT_Type			= Request("PT_Type")
PT_Description	= Request("PT_Description")
Partner_P_Name	= Request("Partner_P_Name")
PT_Price		= Request("PT_Price")
Member_M_ID		= gM_ID

set RS1 = Server.CreateObject("ADODB.RecordSet")
rem 에러메세지가 있을 경우 실행안됨

SQL = "select * from tbParts where P_P_No = '"&Parts_P_P_No&"'"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
else
	oldP_Qty			= RS1("P_Qty")
	oldPartner_P_Name	= RS1("Partner_P_Name")
	oldP_CO_Price		= RS1("P_CO_Price")
	oldP_LGE_Price		= RS1("P_LGE_Price")
	oldP_MSE_Price		= RS1("P_MSE_Price")
end if
RS1.Close

if PT_Description <> "" then
	PT_Description = PT_Description & " - "
end if
if PT_Qty <> 0 then
	PT_Description = PT_Description & "재고변경/"
end if
if oldPartner_P_Name <> Partner_P_Name and Partner_P_Name <> "" then
	PT_Description = PT_Description & "거래처변경/"
end if
if oldP_MSE_Price <> PT_Price and PT_Price <> "" then
	PT_Description = PT_Description & "가격변경/"
end if
if PT_Description <> "" then
	if right(PT_Description,1)="/" then
		PT_Description = left(PT_Description,len(PT_Description)-1)
	end if
end if

if strError = "" then	
	rem DB 업데이트
	RS1.Open "tbParts_Transaction",sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("PT_Date")			= PT_Date
		.Fields("Parts_P_P_No")		= Parts_P_P_No
		.Fields("PT_Qty")			= PT_Qty
		.Fields("PT_Type")			= "직접수정"
		.Fields("PT_Description")	= PT_Description
		.Fields("Partner_P_Name")	= Partner_P_Name
		if isNumeric(PT_Price) then
			.Fields("PT_Price")			= PT_Price
		end if
		.Fields("Member_M_ID")		= Member_M_ID

		.Update
		.Close
	end with
end if

SQL = "update tbParts set "
if isnumeric(PT_Qty) then
	SQL = SQL & "P_Qty = P_Qty + "&PT_Qty&","
end if
if isnumeric(PT_Price) then
	SQL = SQL & "P_MSE_Price = "&PT_Price&","
end if
if oldPartner_P_Name <> Partner_P_Name and Partner_P_Name <> "" then
	SQL = SQL & "Partner_P_Name = '"&Partner_P_Name&"',"
end if
if right(SQL,1)="," then
	SQL = left(SQL,len(SQL)-1)
end if
SQL = SQL & " where P_P_No='"&Parts_P_P_No&"'"
sys_DBCon.execute(SQL)
rem 객체 해제
Set RS1	= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="pt_list.asp" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="pt_list.asp" method=post>

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