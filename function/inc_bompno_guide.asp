<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->

<%
dim txtParts_P_P_No
dim RS1
dim SQL

dim Parts_P_P_No

txtParts_P_P_No = Request("txtParts_P_P_No")

set RS1 = Server.CreateObject("ADODB.RecordSet")

dim strTable
SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&Request("strValue1")
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
else
	if RS1("B_Version_Current_YN") = "Y" then
		strTable = "tbBOM_Qty"
	else
		strTable = "tbBOM_Qty_Archive"
	end if
end if
RS1.Close
	
if txtParts_P_P_No <> "" then
	SQL = "select distinct Parts_P_P_No from "&strTable&" where BOM_B_Code = "&Request("strValue1")&" and Parts_P_P_No like '%"&txtParts_P_P_No&"%' order by Parts_P_P_No"
else
	SQL = "select distinct Parts_P_P_No from "&strTable&" where BOM_B_Code = "&Request("strValue1")&" order by Parts_P_P_No"
end if
%>

<script language="javascript">

function press_enter(strName)
{
	if(event.keyCode == 13)
	{
		frmBOMPNOGuide.submit();
	}
}
</script>

<table width=450px cellpadding=0 cellspacing=0 border=0>
<form name="frmBOMPNOGuide" action="inc_bompno_guide.asp" method="post">
<input type="hidden" name="strValue1" value="<%=request("strValue1")%>">
<tr>
	<td align=left><img src="/img/blank.gif" width=1px height=1px><input type="text" name="txtParts_P_P_No" value="<%=txtParts_P_P_No%>" style="width:92%" onDblClick="javascript:parent.divBOMPNO_Guide.style.display='none';" onkeydown="javascript:press_enter('txtParts_P_P_No')">&nbsp;<span style="cursor:hand;" onclick="javascript:parent.divBOMPNO_Guide.style.display='none';">¡å</span></td></td>
</tr>
<tr>
	<td>
		<select name="sltParts_P_P_No" size=26 onDblClick="javascript:parent.OnDoubleClickBOMPNO(this.value)" style="width:100%;height:392px">
<%
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	Parts_P_P_No			= RS1("Parts_P_P_No")
%>
		<option value="<%=Parts_P_P_No%>"><%=Parts_P_P_No%></option>
<%
	RS1.MoveNext
loop
RS1.Close
%>
		</select>
	</td>
</tr>
<script language="javascript">
if(parent.divBOMPNO_Guide.style.display == "block")
{
	frmBOMPNOGuide.txtParts_P_P_No.focus();
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