<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->

<%
dim txtP_Name
dim RS1
dim SQL

txtP_Name = Request("txtP_Name")

dim P_Name

if txtP_Name <> "" then
	SQL = "select P_Name from tbPartner where P_Name like '%"&txtP_Name&"%' order by P_Name"
else
	SQL = "select P_Name from tbPartner order by P_Name"
end if
set RS1 = Server.CreateObject("ADODB.RecordSet")
%>

<script language="javascript">

function press_enter(strName)
{
	if(event.keyCode == 13)
	{
		frmPartnerGuide.submit();
	}
}
</script>

<table width=450px cellpadding=0 cellspacing=0 border=0>
<form name="frmPartnerGuide" action="inc_partner_guide.asp" method="post">
<tr>
	<td align=left><img src="/img/blank.gif" width=1px height=1px><input type="text" name="txtP_Name" value="<%=txtP_Name%>" style="width:92%" onDblClick="javascript:parent.divPartner_Guide.style.display='none';" onkeydown="javascript:press_enter('txtP_Name')">&nbsp;<span style="cursor:hand;" onclick="javascript:parent.divPartner_Guide.style.display='none';">¡å</span></td></td>
</tr>
<tr>
	<td>
		<select name="sltP_Name" size=26 onDblClick="javascript:parent.OnDoubleClickPartner(this.value)" style="width:100%;height:392px">
<%
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	P_Name			= RS1("P_Name")
%>
		<option value="<%=P_Name%>"><%=P_Name%></option>
<%
	RS1.MoveNext
loop
RS1.Close
%>
		</select>
	</td>
</tr>
<script language="javascript">
if(parent.frmPartnerGuide.style.display == "block")
{
	frmPartnerGuide.txtP_Name.focus();
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