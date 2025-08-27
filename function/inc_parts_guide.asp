<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->

<%
dim txtP_P_No
dim RS1
dim SQL

txtP_P_No = Request("txtP_P_No")

dim P_P_No
dim P_Spec
dim P_Spec_Short

if txtP_P_No <> "" then
	'SQL = "select P_P_No, P_Spec, P_Spec_Short from tbParts where P_P_No like '%"&txtP_P_No&"%' order by P_P_No"
	SQL = "select P_P_No=BOM_Parts_BP_PNO, P_Spec=BM_Spec from tblBOM_Mask where BOM_Parts_BP_PNO like '%"&txtP_P_No&"%' order by BOM_Parts_BP_PNO"
	'SQL = "select P_P_No=BOM_Parts_BP_PNO, P_Spec=BM_Spec from tblBOM_Mask order by BOM_Parts_BP_PNO"
else
	'SQL = "select P_P_No, P_Spec, P_Spec_Short from tbParts order by P_P_No"
	SQL = "select top 100 P_P_No=BOM_Parts_BP_PNO, P_Spec=BM_Spec from tblBOM_Mask order by BOM_Parts_BP_PNO"
end if
set RS1 = Server.CreateObject("ADODB.RecordSet")
%>

<script language="javascript">

function press_enter(strName)
{ 
	if(event.keyCode == 13) 
	{ 
		frmPartsGuide.submit();
	}
}
</script>

<table width=250px cellpadding=0 cellspacing=0 border=0>
<form name="frmPartsGuide" action="inc_parts_guide.asp" method="post">
<tr>
	<td align=left><img src="/img/blank.gif" width=1px height=1px><input type="text" name="txtP_P_No" value="<%=txtP_P_No%>" style="width:92%" onDblClick="javascript:parent.divParts_Guide.style.display='none';" onkeydown="javascript:press_enter('txtP_P_No')">&nbsp;<span style="cursor:hand;" onclick="javascript:parent.divParts_Guide.style.display='none';">¡å</span></td></td>
</tr>
<tr>
	<td>
		<select name="sltP_P_No" size=17 onDblClick="javascript:parent.OnDoubleClickParts(this.value)" style="width:100%;height:261px">
<%
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	P_P_No			= RS1("P_P_No")
	P_Spec			= RS1("P_Spec")
	'P_Spec_Short	= RS1("P_Spec_Short")
	
	'if P_Spec_Short <> "" then
		'P_Spec = P_Spec_Short
	'end if
%>
		<option value="<%=P_P_No%>"><%=P_P_No%> --- <%=P_Spec%></option>
<%
	RS1.MoveNext
loop
RS1.Close
%>
		</select>	
	</td>
</tr>
<script language="javascript">
if(parent.divParts_Guide.style.display == "block")
{
	frmPartsGuide.txtP_P_No.focus();
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