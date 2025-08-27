<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem 변수선언
dim SQL
dim strError
dim CNT1
dim CNT2

dim BOM_Sub_BS_D_No
dim MPD_Process
dim MPD_Date
dim MPD_Time
dim MPD_Line
dim MPD_Qty

dim arrInputSelectG_1
dim arrInputSelect_1
dim arrInputSelectG_2
dim arrInputSelect_2

BOM_Sub_BS_D_No		= Request("BOM_Sub_BS_D_No")
MPD_Process				= Request("MPD_Process")
MPD_Date				= Request("MPD_Date")


rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	rem DB 업데이트
	SQL = "delete from tbMSE_Plan_Date where BOM_Sub_BS_D_No='"&BOM_Sub_BS_D_No&"' and MPD_Process='"&MPD_Process&"' and MPD_Date='"&MPD_Date&"'"
	sys_DBCon.execute(SQL)
	
	select case MPD_Process
	case "IMD"
		arrInputSelectG_1 = split(replace(BasicDataIMDLine,"slt>",""),";")
	case "SMD"
		arrInputSelectG_1 = split(replace(BasicDataSMDLine,"slt>",""),";")
	case "MAN"
		arrInputSelectG_1 = split(replace(BasicDataMANLine,"slt>",""),";")
	case "ASM"
		arrInputSelectG_1 = split(replace(BasicDataASMLine,"slt>",""),";")
	end select
	
	for CNT1 = 0 to ubound(arrInputSelectG_1)
		arrInputSelect_1 = split(arrInputSelectG_1(CNT1),":")
		
		if MPD_Process="IMD" or MPD_Process="SMD" then
			arrInputSelectG_2	= split(replace(BasicDataFullTime,"slt>",""),";")
		else
			arrInputSelectG_2	= split(replace(BasicDataHalfTime,"slt>",""),";")
		end if
		
		for CNT2 = 0 to ubound(arrInputSelectG_2)
			arrInputSelect_2 = split(arrInputSelectG_2(CNT2),":")
			MPD_Line	= arrInputSelect_1(0)
			MPD_Time	= arrInputSelect_2(0)
			MPD_Qty		= Request(arrInputSelect_1(0)&"_"&arrInputSelect_2(0))
			if isNumeric(MPD_Qty) then
				if MPD_Qty > 0 then
					SQL = "insert into tbMSE_Plan_Date (BOM_Sub_BS_D_No,MPD_Process,MPD_Date,MPD_Time,MPD_Line,MPD_Qty) values"
					SQL = SQL & "('"&BOM_Sub_BS_D_No&"','"&MPD_Process&"','"&MPD_Date&"','"&MPD_Time&"','"&MPD_Line&"',"&MPD_Qty&")"
					sys_DBCon.execute(SQL)
				end if
			end if
		next
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
<script language="javascript">
parent.hide_MSE_Plan_Editor();
</script>
<%
else
%>
<script language="javascript">
alert("<%=strError%>");
parent.hide_MSE_Plan_Editor();
</script>
<%
end if
%>



<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->