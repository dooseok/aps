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
dim	P_Qty
dim	P_Safe_Qty
dim	P_LGE_Price
dim	Partner_P_Price_1
dim	Partner_P_Name_1
dim	Partner_P_Price_2
dim	Partner_P_Name_2

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev			= Request("URL_Prev")
URL_Next			= Request("URL_Next")

P_P_No				= trim(Request("P_P_No"))
P_Work_Type			= trim(Request("P_Work_Type"))
P_Desc				= trim(Request("P_Desc"))
P_Qty				= trim(Request("P_Qty"))
P_Safe_Qty			= trim(Request("P_Safe_Qty"))
P_LGE_Price			= trim(Request("P_LGE_Price"))
Partner_P_Price_1	= trim(Request("Partner_P_Price_1"))
Partner_P_Name_1	= trim(Request("Partner_P_Name_1"))
Partner_P_Price_2	= trim(Request("Partner_P_Price_2"))
Partner_P_Name_2	= trim(Request("Partner_P_Name_2"))

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
		if isNumeric(P_Qty) then
			.Fields("P_Qty")			= P_Qty
		end if
		if isNumeric(P_Safe_Qty) then
			.Fields("P_Safe_Qty")		= P_Safe_Qty
		end if
		if isNumeric(P_LGE_Price) then
			.Fields("P_LGE_Price")		= P_LGE_Price
		end if
		.Update
		.Close
	end with
	
	if isnumeric(Partner_P_Price_1) then
		if Partner_P_Price_1 > 0 and Partner_P_Name_1 <> "" then
			RS1.Open "tbParts_Price",sys_DBConString,3,2,2
			with RS1
				.AddNew
				.Fields("Parts_P_P_No")		= P_P_No
				.Fields("PP_Price")			= Partner_P_Price_1
				.Fields("Partner_P_Name")	= Partner_P_Name_1
				.Fields("PP_Las_YN")		= "Y"
				.Update
			.Close
		end with
		end if
	end if
	
	if isnumeric(Partner_P_Price_2) then
		if Partner_P_Price_2 > 0 and Partner_P_Name_2 <> "" then
			RS1.Open "tbParts_Price",sys_DBConString,3,2,2
			with RS1
				.AddNew
				.Fields("Parts_P_P_No")		= P_P_No
				.Fields("PP_Price")			= Partner_P_Price_2
				.Fields("Partner_P_Name")	= Partner_P_Name_2
				.Fields("PP_Las_YN")		= ""
				.Update
			.Close
		end with
		end if
	end if
end if

rem 객체 해제
Set RS1	= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="p_data_list.asp" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="p_data_list.asp" method=post>

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