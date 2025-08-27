<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->

<%
dim txtB_D_No
dim RS1
dim SQL

txtB_D_No = Request("txtB_D_No")

dim B_Code
dim B_D_No
dim B_Version_Code
dim B_Version_Current_YN

SQL = "select "
SQL = SQL & "B_Code, "
SQL = SQL & "B_D_No, "
SQL = SQL & "B_Version_Code, B_Version_Current_YN "
SQL = SQL & " from  "
SQL = SQL & "tbBOM  "
SQL = SQL & " "
if txtB_D_No <> "" then
	SQL = SQL & " where B_D_No like '%"&txtB_D_No&"%' "	
end if
SQL = SQL & "order by B_D_No, B_Version_Date desc"

set RS1 = Server.CreateObject("ADODB.RecordSet")
%>

<script language="javascript">

function press_enter(strName)
{ 
	if(event.keyCode == 13) 
	{ 
		frmBOMSubVersionGuide.submit();
	}
}
</script>

<table width=250px cellpadding=0 cellspacing=0 border=0>
<form name="frmBOMSubVersionGuide" action="inc_bom_sub_version_guide.asp" method="post">
<tr>
	<td align=left><img src="/img/blank.gif" width=1px height=1px><input type="text" name="txtB_D_No" value="<%=txtB_D_No%>" style="width:92%" onDblClick="javascript:parent.divBOMSubVersion_Guide.style.display='none';" onkeydown="javascript:press_enter('txtB_D_No')">&nbsp;<span style="cursor:hand;" onclick="javascript:parent.divBOMSubVersion_Guide.style.display='none';">¡å</span></td></td>
</tr>
<tr>
	<td>
		<select name="sltB_D_No" size=17 onDblClick="javascript:parent.OnDoubleClickBOMSubVersion(this.value)" style="width:100%;height:261px">
<%
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	B_Code			= RS1("B_Code")
	B_D_No			= RS1("B_D_No")
	B_Version_Code	= RS1("B_Version_Code")
	B_Version_Current_YN	= RS1("B_Version_Current_YN")
%>
		<option value="<%=B_D_No%>&nbsp;/&nbsp;<%=B_Version_Code%>&nbsp;/&nbsp;<%=B_Code%>"><%=B_D_No%>-<%=B_Version_Code%>&nbsp;<%=B_Version_Current_YN%></option>
<%
	RS1.MoveNext
loop
RS1.Close
%>
		</select>	
	</td>
</tr>
<script language="javascript">
if(parent.divBOMSubVersion_Guide.style.display == "block")
{
	frmBOMSubVersionGuide.txtB_D_No.focus();
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