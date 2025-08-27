<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->

<%
dim txtM_P_No
dim RS1
dim SQL

txtM_P_No = Request("txtM_P_No")

dim M_P_No
dim M_Spec
dim M_Desc

if txtM_P_No <> "" then
	SQL = "select M_P_No, M_Desc, M_Spec from tbMaterial where M_P_No like '%"&txtM_P_No&"%' or M_Desc like '%"&txtM_P_No&"%' or M_Spec like '%"&txtM_P_No&"%' order by M_P_No"
else
	SQL = "select M_P_No, M_Desc, M_Spec from tbMaterial order by M_P_No"
end if
set RS1 = Server.CreateObject("ADODB.RecordSet")
%>

<script language="javascript">

function press_enter(strName)
{
	if(event.keyCode == 13)
	{
		frmMaterialGuide.submit();
	}
}
</script>

<table width=450px cellpadding=0 cellspacing=0 border=0>
<form name="frmMaterialGuide" action="inc_material_guide.asp" method="post">
<tr>
	<td align=left><img src="/img/blank.gif" width=1px height=1px><input type="text" name="txtM_P_No" value="<%=txtM_P_No%>" style="width:92%" onDblClick="javascript:parent.divMaterial_Guide.style.display='none';" onkeydown="javascript:press_enter('txtM_P_No')">&nbsp;<span style="cursor:hand;" onclick="javascript:parent.divMaterial_Guide.style.display='none';">¡å</span></td></td>
</tr>
<tr>
	<td>
		<select name="sltM_P_No" size=26 onDblClick="javascript:parent.OnDoubleClickMaterial(this.value)" style="width:100%;height:392px">
<%
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	M_P_No			= RS1("M_P_No")
	M_Spec			= RS1("M_Spec")
	M_Desc			= RS1("M_Desc")
%>
		<option value="<%=M_P_No%>"><%=M_P_No%> / <%=M_Desc%> / <%=M_Spec%></option>
<%
	RS1.MoveNext
loop
RS1.Close
%>
		</select>
	</td>
</tr>
<script language="javascript">
if(parent.divMaterial_Guide.style.display == "block")
{
	frmMaterialGuide.txtM_P_No.focus();
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