<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem 변수선언
dim RS1
dim SQL
dim CNT1

dim strError
dim strError_Temp

dim arrID_All
dim arrBMDD_Desc_BOM
dim arrBOM_Mask_Desc_BMD_Desc


arrID_All					= split(Request("strID_All")&" "				,", ")
arrBMDD_Desc_BOM				= split(Request("BMDD_Desc_BOM")&" "				,", ")
arrBOM_Mask_Desc_BMD_Desc	= split(Request("BOM_Mask_Desc_BMD_Desc")&" "	,", ")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)						= trim(arrID_All(CNT1))
	arrBMDD_Desc_BOM(CNT1)				= trim(arrBMDD_Desc_BOM(CNT1))
	arrBOM_Mask_Desc_BMD_Desc(CNT1)	= trim(arrBOM_Mask_Desc_BMD_Desc(CNT1))
next
set RS1 = Server.CreateObject("ADODB.RecordSet")
rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
			arrBMDD_Desc_BOM(CNT1) = replace(arrBMDD_Desc_BOM(CNT1),"'","''")
			arrBOM_Mask_Desc_BMD_Desc(CNT1) = replace(arrBOM_Mask_Desc_BMD_Desc(CNT1),"'","''")
			
		
			SQL = "select * from tblBOM_Mask_Desc where BMD_Desc='"&arrBOM_Mask_Desc_BMD_Desc(CNT1)&"' "
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				SQL = "insert into tblBOM_Mask_Desc values ('"&arrBOM_Mask_Desc_BMD_Desc(CNT1)&"',999)"
				sys_DBCon.execute(SQL)
			end if
			RS1.Close
			
			
			SQL = "update tblBOM_Mask_Desc_Detail set "
			SQL = SQL & "	BMDD_Desc_BOM='"&arrBMDD_Desc_BOM(CNT1)&"', "
			SQL = SQL & "	BOM_Mask_Desc_BMD_Desc='"&arrBOM_Mask_Desc_BMD_Desc(CNT1)&"' "
			SQL = SQL & "where BMDD_Code='"&arrID_All(CNT1)&"'"
			sys_DBCon.execute(SQL)
		

		strError = strError & strError_Temp
	next
end if
set RS1 = nothing
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
<form name="frmRedirect" action="parts_desc_list.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
	'strError = strError & "* 일부의 수정이 취소되었습니다."
%>
<form name="frmRedirect" action="parts_desc_list.asp" method=post>

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