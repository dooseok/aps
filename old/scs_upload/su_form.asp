<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim CNT1
dim CNT2
%>
<script language="javascript">
function frmFileUpload_Check()
{
	frmFileUpload.submit();
}

function List2Excel()
{
	frmList2Excel.submit();
}
</script>

<table width=450px cellpadding=0 cellspacing=0 border=0>
<tr>
	<td align=center>
		<table cellpadding=0 cellspacing=0 border=0>
		<form name="frmFileUpload" action="su_action.asp" method="post" enctype="MULTIPART/FORM-DATA">
		<tr>
			<td width=100px align=right>�ֹ���Ȳ : </td>
			<td width=350px>
				<input type="file" name="strFile1" style="width:100%">
			</td>
		</tr>
		<tr>
			<td width=100px align=right>��ũ���� : </td>
			<td width=350px>
				<input type="file" name="strFile2" style="width:100%">
			</td>
		</tr>
		<tr>
			<td width=100px align=right>���ó�� : </td>
			<td width=350px>
				<input type="file" name="strFile3" style="width:100%">
			</td>
		</tr>
		<tr>
			<td width=100px align=right>�����Ȳ : </td>
			<td width=350px>
				<input type="file" name="strFile4" style="width:100%">
			</td>
		</tr>
		<tr>
			<td width=100px align=right>�԰��� : </td>
			<td width=350px>
				<input type="file" name="strFile5" style="width:100%">
			</td>
		</tr>
		<tr>
			<td colspan=2 align=center>
				<table width=77px cellpadding=0 cellspacing=0>
				<tr>
					<td width=77px>	
						<%=Make_BTN("���ϵ��","javascript:frmFileUpload_Check()","")%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
</table>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->