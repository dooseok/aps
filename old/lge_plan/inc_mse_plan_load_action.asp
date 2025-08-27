<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
dim SQL
dim RS1

dim BOM_Sub_BS_D_No
dim MPD_Process
dim MPD_Date

BOM_Sub_BS_D_No	= Request("BOM_Sub_BS_D_No")
MPD_Process			= Request("MPD_Process")
MPD_Date			= Request("MPD_Date")
%>
<script language="javascript">
<%
set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select MPD_Line, MPD_Time, MPD_Qty from tbMSE_Plan_Date where BOM_Sub_BS_D_No='"&BOM_Sub_BS_D_No&"' and MPD_Process='"&MPD_Process&"' and MPD_Date='"&MPD_Date&"'"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
else
	do until RS1.Eof
%>
parent.Load_frmMSE_Plan_Editor('<%=RS1("MPD_Line")%>','<%=RS1("MPD_Time")%>','<%=RS1("MPD_Qty")%>');
<%
		RS1.MoveNext
	loop
end if
RS1.Close
set RS1 = nothing
%>
parent.cal_MPD_Qty_Total();
</script>

<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->