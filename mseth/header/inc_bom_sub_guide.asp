<!-- #include virtual = "/header/asp_header.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->

<%
dim txtBS_D_No
dim RS1
dim SQL

txtBS_D_No = Request("txtBS_D_No")

dim BS_D_No

if txtBS_D_No <> "" then
	SQL = "select BS_D_No from tbBOM_Sub where BS_D_No like '%"&txtBS_D_No&"%' order by BS_D_No"
else
	SQL = "select BS_D_No from tbBOM_Sub order by BS_D_No"
end if
set RS1 = Server.CreateObject("ADODB.RecordSet")
%>

<script language="javascript">

function press_enter(strName)
{ 
	if(event.keyCode == 13) 
	{ 
		frmBOMSubGuide.submit();
	}
}
</script>

<table width=250px cellpadding=0 cellspacing=0 border=0>
<form name="frmBOMSubGuide" action="inc_bom_sub_guide.asp" method="post">
<tr>
	<td align=left><img src="/img/blank.gif" width=1px height=1px><input type="text" name="txtBS_D_No" value="<%=txtBS_D_No%>" style="width:92%" onDblClick="javascript:parent.divBOMSub_Guide.style.display='none';" onkeydown="javascript:press_enter('txtBS_D_No')">&nbsp;<span style="cursor:hand;" onclick="javascript:parent.divBOMSub_Guide.style.display='none';">¡å</span></td></td>
</tr>
<tr>
	<td>
		<select name="sltBS_D_No" size=20 onDblClick="javascript:parent.OnDoubleClickBOMSub(this.value)" style="width:100%">
<%
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	BS_D_No			= RS1("BS_D_No")
%>
		<option value="<%=BS_D_No%>"><%=BS_D_No%></option>
<%
	RS1.MoveNext
loop
RS1.Close
%>
		</select>	
	</td>
</tr>
<script language="javascript">
if(parent.divBOMSub_Guide.style.display == "block")
{
	frmBOMSubGuide.txtBS_D_No.focus();
}
</script>
</form>
</table>
<%
set RS1 = nothing
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->