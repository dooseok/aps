<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim	P_P_No
dim	P_Work_Type
dim	P_Desc
dim	P_Spec_Short
dim	P_Spec
dim	P_Maker
dim	P_Safe_Qty
dim	P_LGE_Price
dim P_MSE_Price
dim Partner_P_Name
dim P_Same_Code
dim P_Real_Demand_P_P_No

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev			= Request("URL_Prev")
URL_Next			= Request("URL_Next")

P_P_No				= trim(Request("P_P_No"))
P_Work_Type			= trim(Request("P_Work_Type"))
P_Desc				= trim(Request("P_Desc"))
P_Spec_Short		= replace(trim(Request("P_Spec_Short")),"'","''")
P_Spec				= replace(trim(Request("P_Spec")),"'","''")
P_Maker				= trim(Request("P_Maker"))
P_Safe_Qty			= trim(Request("P_Safe_Qty"))
P_LGE_Price			= trim(Request("P_LGE_Price"))
P_MSE_Price			= trim(Request("P_MSE_Price"))
Partner_P_Name		= trim(Request("Partner_P_Name"))
P_Same_Code			= trim(Request("P_Same_Code"))
P_Real_Demand_P_P_No= trim(Request("P_Real_Demand_P_P_No"))

set RS1 = Server.CreateObject("ADODB.RecordSet")

SQL = "select P_P_No from tbParts where P_P_No='"&P_P_No&"'"
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
	strError = strError & "* 같은 파트넘버의 아이템이 이미 등록되어있습니다.\n"
end if
RS1.Close


rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	rem DB 업데이트
	
	RS1.Open "tbParts",sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("P_P_No")				= P_P_No
		.Fields("P_Work_Type")			= P_Work_Type
		.Fields("P_Desc")				= P_Desc
		.Fields("P_Spec_Short")			= P_Spec_Short
		.Fields("P_Spec")				= P_Spec
		.Fields("P_Maker")				= P_Maker
		if isNumeric(P_Safe_Qty) then
		else
			P_Safe_Qty = 0
		end if
		.Fields("P_Safe_Qty")		= P_Safe_Qty
		if isNumeric(P_LGE_Price) then
		else
			P_LGE_Price = 0
		end if
		.Fields("P_LGE_Price")		= P_LGE_Price
		if isNumeric(P_MSE_Price) then
		else
			P_MSE_Price = 0
		end if
		.Fields("P_MSE_Price")		= P_MSE_Price
		.Fields("Partner_P_Name")			= Partner_P_Name
		if isNumeric(P_Same_Code) then
			.Fields("P_Same_Code")			= P_Same_Code
		end if
		.Fields("P_Real_Demand_P_P_No")	= P_Real_Demand_P_P_No
		.Update
		.Close
	end with
	
	SQL = "insert tbParts_Transaction (Parts_P_P_No,PT_Qty,PT_Price,PT_Type,PT_Description,Partner_P_Name,Member_M_ID,PT_Date) values ('"&P_P_No&"',0,"&P_MSE_Price&",'부품정보','신규등록','"&Partner_P_Name&"','"&gM_ID&"','"&date()&"')"
	sys_DBCon.execute(SQL)
end if

rem 객체 해제
Set RS1	= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="p_list.asp" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="p_list.asp" method=post>

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