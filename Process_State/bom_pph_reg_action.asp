<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem ��������
dim SQL
dim RS1

dim BOM_Sub_BS_D_No
dim BP_PPH

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")

BOM_Sub_BS_D_No	= trim(Request("BOM_Sub_BS_D_No"))
BP_PPH			= trim(Request("BP_PPH"))

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select BOM_Sub_BS_D_No from tbBOM_PPH where BOM_Sub_BS_D_No='"&BOM_Sub_BS_D_No&"'"
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
	strError = strError & "* �űԵ�� ����\n["&BOM_Sub_BS_D_No&"]�� ������ ��Ʈ�ѹ��� �̹� ��ϵǾ��ֽ��ϴ�.\n"
end if
RS1.Close


rem �����޼����� ���� ��� ����ȵ�
if strError = "" then	
	
	SQL = "insert tbBOM_PPH (BOM_Sub_BS_D_No,BP_PPH) values "
	SQL = SQL & "	('"&BOM_Sub_BS_D_No&"', "
	if isnumeric(BP_PPH) then
	else
		BP_PPH = 0
	end if
	SQL = SQL & "	"&BP_PPH&") "
	sys_DBCon.execute(SQL)
end if

rem ��ü ����
Set RS1	= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="bom_pph_list.asp" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="bom_pph_list.asp" method=post>

</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>

<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->