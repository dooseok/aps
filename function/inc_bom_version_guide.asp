<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->

<%
dim txtBS_D_No
dim RS1
dim SQL

txtBS_D_No = Request("txtBS_D_No")

dim BS_D_No
dim B_Version_Code
set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select "
SQL = SQL & "t1.BS_D_No, "
SQL = SQL & "t2.B_Version_Code "
SQL = SQL & " from  "
SQL = SQL & "tbBOM_Sub t1, tbBOM t2  "
SQL = SQL & "where t1.BOM_B_Code = t2.B_Code  "
if txtBS_D_No <> "" then
	SQL = SQL & " and t1.BS_D_No like '%"&txtBS_D_No&"%' "	
end if
SQL = SQL & "order by t1.BS_D_No, t2.B_Version_Date desc, t2.B_Code desc"
%>

<script language="javascript">



function press_enter(strName)
{ 
	if(event.keyCode == 13) 
	{ 	
		if(frmBOMVersionGuide.txtBS_D_No.value.length < 3)
		{
			alert("검색어를 3글자 이상 입력해주세요.");
			return false;
		}
		frmBOMVersionGuide.submit();
	}
}

function placeholding()
{
	if(frmBOMVersionGuide.txtBS_D_No.value == "검색어를 입력해주세요. (3글자 이상)")
	{
		turnOffPlaceholding();
	}
	else if(frmBOMVersionGuide.txtBS_D_No.value == "")
	{
		turnOnPlaceholding();
	}
}

function turnOnPlaceholding()
{
	document.getElementById("txtBS_D_No").style.color = "dimgray";
	frmBOMVersionGuide.txtBS_D_No.value = "검색어를 입력해주세요. (3글자 이상)";
}

function turnOffPlaceholding()
{
	document.getElementById("txtBS_D_No").style.color = "black";
	frmBOMVersionGuide.txtBS_D_No.value = "";
}
</script>

<table width=250px cellpadding=0 cellspacing=0 border=0>
<form name="frmBOMVersionGuide" action="inc_bom_version_guide.asp" method="post" onsubmit="return false;">
<tr>
	<td align=left><img src="/img/blank.gif" width=1px height=1px><input id="txtBS_D_No" type="text" name="txtBS_D_No" value="<%=txtBS_D_No%>" style="width:92%" onclick="turnOffPlaceholding();" onDblClick="javascript:parent.divBOMVersion_Guide.style.display='none';" onkeydown="javascript:press_enter('txtBS_D_No')">&nbsp;<span style="cursor:hand;" onclick="javascript:parent.divBOMVersion_Guide.style.display='none';">▼</span></td></td>
</tr>
<tr>
	<td>
		<select name="sltBS_D_No" size=17 onDblClick="javascript:parent.OnDoubleClickBOMVersion(this.value)" style="width:100%;height:261px">
<%
if txtBS_D_No <> "" then
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		BS_D_No			= RS1("BS_D_No")
		B_Version_Code	= RS1("B_Version_Code")
	%>
			<option value="<%=BS_D_No%>-<%=B_Version_Code%>"><%=BS_D_No%>-<%=B_Version_Code%></option>
	<%
		RS1.MoveNext
	loop
	RS1.Close
	
end if
%>
		</select>	
	</td>
</tr>
<script language="javascript">
placeholding();

</script>
</form>
</table>
<%
set RS1 = nothing
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->