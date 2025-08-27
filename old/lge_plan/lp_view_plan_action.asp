<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
dim SQL
dim RS1
dim CNT1
dim CNT2

dim PR_Work_Date
dim PR_Process
dim strPlan
dim arrPlan
dim arrTemp

dim PR_Work_Order
dim PR_Plan_Amount
dim strBOM_Sub_BS_D_No
dim arrBOM_Sub_BS_D_No

dim PR_ST
dim PR_Point

dim Pass_YN

PR_Work_Date	= Request("PR_Work_Date")
PR_Process		= Request("PR_Process")
strPlan			= Request("strPlan")
arrPlan			= split(strPlan,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrPlan)
	arrTemp	= split(arrPlan(CNT1),"|/|")
	PR_Work_Order		= arrTemp(0)
	PR_Plan_Amount		= arrTemp(1)
	strBOM_Sub_BS_D_No	= arrTemp(2)
	arrBOM_Sub_BS_D_No	= split(strBOM_Sub_BS_D_No,"<br>")
	
	for CNT2 = 0 to ubound(arrBOM_Sub_BS_D_No)
		
		Pass_YN = "Y"
		if left(arrBOM_Sub_BS_D_No(CNT2),4) = "6871" or left(arrBOM_Sub_BS_D_No(CNT2),3) = "EBR" then
			if PR_Process = "ASM" then
				Pass_YN = "N"
			end if
		else
			if instr("-IMD-SMD-MAN-",PR_Process) > 0 then
				Pass_YN = "N"
			end if
		end if
		
		if Pass_YN = "Y" then
			SQL = "select * from tbBOM where B_Code in (select BOM_B_Code from tbBOM_Sub where BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT2)&"')"
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				PR_ST		= 10
				PR_Point	= 200
			else
				PR_ST		= int(RS1("B_ST"))
				PR_Point	= int(RS1("B_Point"))
			end if
			RS1.Close
			
			if not(isnumeric(PR_Plan_Amount)) then
				PR_Plan_Amount = 0
			end if
		
			'실적 데이터 등록
			SQL = "insert into tbProcess_Record (PR_Work_Order, PR_WorkType,BOM_Sub_BS_D_No,PR_Process,PR_Worker_CNT,PR_Supporter_CNT,PR_Line,PR_Plan_Amount,PR_Work_Date,PR_Plan_Start_Time,PR_Plan_End_Time,PR_Memo,PR_ST,PR_Point) values "
			SQL = SQL & "('"&PR_Work_Order&"','작업','"&arrBOM_Sub_BS_D_No(CNT2)&"','"&PR_Process&"',0,0,'',"&PR_Plan_Amount&",'"&PR_Work_Date&"','','','',"&PR_ST&","&PR_Point&")"
			sys_DBCon.execute(SQL)
		end if
	next	
next

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

%>
<form name="frmRedirect" action="lp_view.asp" method=post>
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>


<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->