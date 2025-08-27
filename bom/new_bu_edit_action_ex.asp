<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<% 
dim CNT1
dim SQL
dim RS1

dim BU_Code
dim strPart
dim strCheck
dim strMemo
dim strAppliedPartNo

dim arrAppliedPartNo
dim strSQL_Done

dim temp
dim strError
dim URL_Prev
dim URL_Next


rem 객체선언
Set RS1		= Server.CreateObject("ADODB.RecordSet")

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

URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")

BU_Code			= Request("BU_Code")
strPart			= trim(request("strPart"))
strCheck		= trim(request("strCheck"))
strMemo			= trim(request("strMemo"))
strAppliedPartNo = trim(ucase(request("strAppliedPartNo")))

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then
	
	rem DB 업데이트
	SQL = "select * from tbBOM_Update_New where BU_Code = '"&BU_Code&"'"
	RS1.Open SQL,sys_DBconString,3,2,&H0001
	with RS1
		if strCheck="미확인" then
			.Fields("BU_"&strPart&"_Date") = null 
		else
			.Fields("BU_"&strPart&"_Date") = date()
		end if
		
		.Fields("BU_"&strPart&"_Check")	= strCheck
		.Fields("BU_"&strPart&"_Memo")	= strMemo
		
		if instr("-IMT-SMT-JEJO1-JEJO3-","-"&ucase(strPart)&"-") > 0 then
			.Fields("BU_"&strPart&"_PartNo") = strAppliedPartNo
		end if
		
		.Update
		.Close
	end with
end if

if strAppliedPartNo <> "" then
	strSQL_Done = ""
	if UCASE(strPart) = "IMT" then
		strSQL_Done = "set BUP_DONE_IMT = 'Y'"
	elseif UCASE(strPart) = "SMT" then 
		strSQL_Done = "set BUP_DONE_SMT = 'Y'"
	elseif UCASE(strPart) = "JEJO2" then
		strSQL_Done = "set BUP_DONE_JeJo2 = 'Y'"
	elseif UCASE(strPart) = "JEJO3" then
		strSQL_Done = "set BUP_DONE_JeJo3 = 'Y'"
	end if

	if strSQL_Done <> "" then  
		arrAppliedPartNo = split(strAppliedPartNo,vbCrlf)
		for CNT1 = 0 to ubound(arrAppliedPartNo)
			if arrAppliedPartNo(CNT1) <> "" then
				SQL = "update tbBOM_Update_PartNo "
				SQL = SQL & strSQL_Done
				SQL = SQL & "where "
				SQL = SQL & "	BOM_Update_BU_Code = '"&BU_Code&"' and "
				SQL = SQL & "	BUP_PartNo = '"&arrAppliedPartNo(CNT1)&"' "
				sys_DBCon.execute(SQL)
			end if
		next
	end if
end if

Set RS1		= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="new_bu_edit_form.asp" method=post>
<input type="hidden" name="BU_Code" value="<%=BU_Code%>">
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
alert("수정이 완료되었습니다.");
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="new_bu_edit_form.asp" method=post>
<input type="hidden" name="BU_Code" value="<%=BU_Code%>">
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