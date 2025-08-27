<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->

<%
dim txtWork_Order
dim RS1
dim SQL

txtWork_Order = Request("txtWork_Order")

dim Work_Order

dim BOM_Sub_BS_D_No

BOM_Sub_BS_D_No = Request("BOM_Sub_BS_D_No")

if len(BOM_Sub_BS_D_No) = 11 then
	SQL =		"select LP_Work_Order from ( "
	SQL = SQL &	"	select distinct(LP_Work_Order) from tbLGE_Plan where exists (select LPD_Code from tbLGE_Plan_Date where LGE_Plan_LP_Work_Order = LP_Work_Order and LPD_Input_Date between '"&dateadd("d",-30,date())&"' and '"&dateadd("d",30,date())&"') and  LP_Model in (select LM_Name from tbLGE_Model where (BOM_Sub_BS_D_No_1 = '"&BOM_Sub_BS_D_No&"' or BOM_Sub_BS_D_No_2 = '"&BOM_Sub_BS_D_No&"' or BOM_Sub_BS_D_No_3 = '"&BOM_Sub_BS_D_No&"' or BOM_Sub_BS_D_No_4 = '"&BOM_Sub_BS_D_No&"') and LM_Company = 'MSE')"
	SQL = SQL & "	union "
	SQL = SQL & "	select distinct LP_Work_Order = LPE_Type+'_'+convert(varchar,LPE_Code) from tbLGE_Plan_ETC where LPE_Due_Date between '"&dateadd("d",-30,date())&"' and '"&dateadd("d",30,date())&"' and BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"'"
	SQL = SQL & ") t1 "
else
	SQL =		"select LP_Work_Order from ( "
	SQL = SQL &	"	select distinct(LP_Work_Order) from tbLGE_Plan where exists (select LPD_Code from tbLGE_Plan_Date where LGE_Plan_LP_Work_Order = LP_Work_Order and LPD_Input_Date between '"&dateadd("d",-30,date())&"' and '"&dateadd("d",30,date())&"') and  LP_Model in (select LM_Name from tbLGE_Model where LM_Company = 'MSE')"
	SQL = SQL & "	union "
	SQL = SQL & "	select distinct LP_Work_Order = LPE_Type+'_'+convert(varchar,LPE_Code) from tbLGE_Plan_ETC where LPE_Due_Date between '"&dateadd("d",-30,date())&"' and '"&dateadd("d",30,date())&"' "
	SQL = SQL & ") t1 "
end if

if txtWork_Order <> "" then
	SQL = SQL & "where LP_Work_Order like '%"&txtWork_Order&"%'"
end if

set RS1 = Server.CreateObject("ADODB.RecordSet")
%>

<script language="javascript">
function press_enter(strName)
{ 
	if(event.keyCode == 13) 
	{ 
		frmWorkOrderGuide.submit();
	}
}
</script>

<table width=250px cellpadding=0 cellspacing=0 border=0>
<form name="frmWorkOrderGuide" action="inc_workorder_guide.asp" method="post">
<input type="hidden" name="BOM_Sub_BS_D_No" value="<%=BOM_Sub_BS_D_No%>">
<tr>
	<td align=left><img src="/img/blank.gif" width=1px height=1px><input type="text" name="txtWork_Order" value="<%=txtWork_Order%>" style="width:92%" onDblClick="javascript:parent.divWorkOrder_Guide.style.display='none';" onkeydown="javascript:press_enter('txtWork_Order')">&nbsp;<span style="cursor:hand;" onclick="javascript:parent.divWorkOrder_Guide.style.display='none';">¡å</span></td></td>
</tr>
<tr>
	<td>
		<select name="sltWork_Order" size=17 onDblClick="javascript:parent.OnDoubleClickWorkOrder(this.value)" style="width:100%;height:261px">
		<option value=""></option>
<%
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	Work_Order = RS1("LP_Work_Order")
%>
		<option value="<%=Work_Order%>"><%=Work_Order%></option>
<%
	RS1.MoveNext
loop
RS1.Close
%>
		</select>	
	</td>
</tr>
<script language="javascript">
if(parent.divWorkOrder_Guide.style.display == "block")
{
	frmWorkOrderGuide.txtWork_Order.focus();
}
</script>
</form>
</table>
<%
set RS1 = nothing
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->