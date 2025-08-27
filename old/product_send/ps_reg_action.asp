<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim PS_Code
dim BOM_Sub_BS_D_No
dim PS_Send_Date
dim PS_Qty
dim LGE_Plan_LP_Work_Order
dim LGE_Plan_ETC_LPE_Code

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")

PS_Code					= Request("PS_Code")
BOM_Sub_BS_D_No			= Request("BOM_Sub_BS_D_No")
PS_Send_Date			= Request("PS_Send_Date")
PS_Qty					= Request("PS_Qty")
LGE_Plan_LP_Work_Order	= Request("LGE_Plan_LP_Work_Order")
LGE_Plan_ETC_LPE_Code	= Request("LGE_Plan_ETC_LPE_Code")

set RS1 = Server.CreateObject("ADODB.RecordSet")
rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	rem DB 업데이트
	RS1.Open "tbProduct_Send",sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("BOM_Sub_BS_D_No")			= BOM_Sub_BS_D_No
		.Fields("PS_Send_Date")				= PS_Send_Date
		.Fields("PS_Qty")					= PS_Qty
		.Fields("LGE_Plan_LP_Work_Order")	= LGE_Plan_LP_Work_Order
		if isnumeric(LGE_Plan_ETC_LPE_Code) then
			.Fields("LGE_Plan_ETC_LPE_Code")	= LGE_Plan_ETC_LPE_Code
		end if
		.Update
		.Close
	end with
end if

rem 객체 해제
Set RS1	= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="ps_list.asp" method=post>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="ps_list.asp" method=post>
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