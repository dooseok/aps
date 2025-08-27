<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->

<html>
<head>
<body topmargin=0 leftmargin=0>
<center>
<br>

<script language="javascript">
function Form_Check()
{
	frmXLS_Upload.target="_opener";
	frmXLS_Upload.submit();
	self.close();
}
</script>

<table width=670 cellpadding=5 cellspacing=1 border=0 bgcolor=#999999>
<form name="frmXLS_Upload" action="xls_upload_action<%if gM_ID="shindk" then%>2<%end if%>.asp" method="post" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="B_Code" value="<%=Request("B_Code")%>">
<tr>
	<td align=center bgcolor=#ffffff><input type="file" name="BOM_XLS" style="width:80%" class="input"></td>
</tr>
<tr bgcolor=#eeeeee>
	<td align=center>
		<table width=77px cellpadding=0 cellspacing=0 border=0>
		<tr>
			<td width=77px><%=Make_BTN("엑셀업로드","javascript:Form_Check()","")%></td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
</center>

</body>
</html>
<!-- #include virtual = "/header/session_check_tail.asp" -->