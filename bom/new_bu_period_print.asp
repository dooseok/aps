<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<%
call usePrinter()
%>
<%
dim RS1
dim SQL
dim CNT1

dim nTotalPage
dim nRecord
dim nPage

dim BU_Code
dim BOM_B_D_No
dim BU_Content
dim BU_Receive_Date
dim BU_Apply_Date
dim BU_Reply_Date
dim BU_Request_Reply_Date

dim s_Date_1
dim s_Date_2

s_Date_1 = Request("s_Date_1")
s_Date_2 = Request("s_Date_2")

nRecord = 1
nPage = 1
%>

<script language="javascript">
function UsePrint()
{
	factory.printing.header				= "";
	factory.printing.footer				= "";
	factory.printing.portrait			= false;
	factory.printing.leftMargin			= 0.5;
	factory.printing.rightMargin		= 0.5;
	factory.printing.topMargin			= 1.5;
	factory.printing.bottomMargin		= 1.5;
	if(confirm("Ȯ���� Ŭ���Ͻ� �� ��ñ�ٸ��ø�\n�μ� ��ȭ���ڰ� ��ϴ�."))
	{
		factory.printing.print(true, window);
	}
}
</script>

<%
set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from vwBU_List where BU_Receive_Date between '"&s_Date_1&"' and '"&s_Date_2&"' order by BU_Receive_Date, BU_Code"
RS1.Open SQL,sys_DBCon,1

nTotalPage = int(RS1.RecordCount / 6)

if RS1.RecordCount mod 6 <> 0 then
	nTotalPage = nTotalPage + 1
end if

do until RS1.Eof

BU_Code					= RS1("BU_Code")
BOM_B_D_No				= RS1("BOM_B_D_No")
BU_Content				= RS1("BU_Content")
BU_Receive_Date			= RS1("BU_Receive_Date")
BU_Apply_Date			= RS1("BU_Apply_Date")
BU_Reply_Date			= RS1("BU_Reply_Date")
BU_Request_Reply_Date	= RS1("BU_Request_Reply_Date")
%>

<%
if nRecord = 1 then
	if nPage = 1 then
%>
<table width=1040px cellpadding=0 cellspacing=0 border=0 bordercolor=black style="table-layout:fixed;" style="font-size:12px;font-face:����">
<tr height=30px>
	<td align=center style="font-size:25px"><b>�ù泻�� ����Ʈ</b></td>
</tr>
<tr>
	<td align=right><b>��ȸ�Ⱓ : <%=s_Date_1%> - <%=s_Date_2%> / ����� : <%=date()%> / ������ : <%=nPage%>/<%=nTotalPage%></b></td>
</tr>
</table>
<%
	else
%>
<table width=1040px cellpadding=0 cellspacing=0 border=0 bordercolor=black style="table-layout:fixed;" style="font-size:12px;font-face:����">
<tr height=30px>
	<td align=center style="font-size:25px"><b>&nbsp;</b></td>
</tr>
<tr>
	<td align=right><b>��ȸ�Ⱓ : <%=s_Date_1%> - <%=s_Date_2%> / ����� : <%=date()%> / ������ : <%=nPage%>/<%=nTotalPage%></b></td>
</tr>
</table>
<%
	end if
%>
<table width=1040px cellpadding=0 cellspacing=0 border=1 bordercolor=black style="table-layout:fixed;" style="font-size:12px;font-face:����">
<tr height=30px>
	<td width=110px>������ȣ</td>
	<td width=80px>������</td>
	<td width=80px>����</td>
	<td width=80px>��Ʈ�ѹ�</td>
	<td width=80px>ǰ��</td>
	<td>����</td>
	<td width=80px>�ۼ���</td>
	<td width=80px>ȸ�ſ䱸��</td>
	<td width=80px>ȸ����</td>
	<td width=80px>������</td>
	<td width=60px>����Ȯ��</td>
</tr>
<%
end if
%>
<tr height=110px>
	<td><%=BU_Code%></td>
	<td><%=BU_Receive_Date%></td>
	<td>LG����(��)</td>
	<td><%=BOM_B_D_No%></td>
	<td>PCB ASS'Y</td>
	<td align=left valign=top style="font-size:10px"><%=BU_Content%></td>
	<td>���������</td>
	<td><%=BU_Request_Reply_Date%>&nbsp;</td>
	<td><%=BU_Reply_Date%>&nbsp;</td>
	<td><%=BU_Apply_Date%></td>
	<td>&nbsp;</td>
</tr>
<%
	nRecord = nRecord + 1
	if nRecord = 7 then
		nPage = nPage + 1
		nRecord = 1
	end if
	RS1.MoveNext
loop
set RS1 = nothing
%>
</table>


<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->