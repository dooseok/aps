<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

 
<%
dim RS1
dim SQL
dim CNT1

dim BU_Code

dim BOM_B_D_No
dim BU_Content
dim BU_Apply_Date
dim BU_Reply_Date
dim BU_Request_Reply_Date
dim BU_File_1
dim BU_File_2
dim BU_File_3
dim BU_Type

BU_Code = Request("BU_Code")

Set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbBOM_Update where BU_Code='"&BU_Code&"'"
RS1.Open SQL,sys_DBCon
BOM_B_D_No		= RS1("BOM_B_D_No")
BU_Content		= RS1("BU_Content")
BU_Apply_Date	= RS1("BU_Apply_Date")
BU_Reply_Date	= RS1("BU_Reply_Date")
BU_Request_Reply_Date	= RS1("BU_Request_Reply_Date")
BU_File_1		= RS1("BU_File_1")
BU_File_2		= RS1("BU_File_2")
BU_File_3		= RS1("BU_File_3")
BU_Type			= RS1("BU_Type")
RS1.Close
Set RS1 = Nothing
%>

<%
call usePrinter()
%>
<table width=720px cellpadding=0 cellspacing=0 border=0 bordercolor=black bgcolor="#ffffff" style="table-layout:fixed" style="font-face:����;font-size:14px;font-weight:bold;">
<tr bgcolor=white>
	<td>
		<table cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed">
		<tr height=40 bgcolor=white>
			<td style="font-size:30px"><b>�ù�(����) ���� ��û��</b></td>
		</tr>
		<tr height=20 bgcolor=white>
			<td style="font-face:����;font-size:14px">
				<input type=checkbox name='BU_Type_New' style="border:none;font-size:14pt;color:#ffffff;background-color:#FFFFFF" value='Y'<%if instr(BU_Type,"�ű�") > 0 then%> checked<%end if%>>�ű԰���
				<input type=checkbox name='BU_Type_Add' style="border:none;font-size:14pt;color:#ffffff;background-color:#FFFFFF" value='Y'<%if instr(BU_Type,"�߰�") > 0 then%> checked<%end if%>>�۾��߰�
				<input type=checkbox name='BU_Type_Update' style="border:none;font-size:14pt;color:#ffffff;background-color:#FFFFFF" value='Y'<%if instr(BU_Type,"�ù�") > 0 then%> checked<%end if%>>����ù�
			</td>
		</tr>
		</table>
	</td>
	<td width=300px align=right valign=bottom>
		<img src="/img/blank.gif" width=10px height=2px><br>
		<table width=300px cellpadding=0 cellspacing=0 border=1 bordercolor=black style="table-layout:fixed;">
		
		<tr bgcolor=white>
			<td style="font-size:12px" width=30px rowspan=2>��<br>��</td>
			<td style="font-size:12px">�� ��</td>
			<td style="font-size:12px">�� ��</td>
			<td style="font-size:12px">�� ��</td>
			<td style="font-size:12px">�� ��</td>
		</tr>
		<tr bgcolor=white height=60px>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<img src="/img/blank.gif" width=10px height=30px><br>
<table width=720px cellpadding=0 cellspacing=0 border=1 bordercolor=black style="font-face:����;font-size:14px;table-layout:fixed;">
<tr height=30px>
	<td width=100px>������ȣ</td>
	<td width=140px><%=BU_Code%></td>
	<td width=100px>ǰ��</td>
	<td width=140px>PCB ASS'Y</td>
	<td width=100px>�ۼ���</td>
	<td width=140px>���������</td>
</tr>
<tr height=30px>
	<td width=100px>������</td>
	<td width=140px><%=BU_Apply_Date%></td>
	<td width=100px>ȸ����</td>
	<td width=140px><%=BU_Reply_Date%>&nbsp;</td>
	<td width=100px>ȸ�ſ�û��</td>
	<td width=140px><%=BU_Request_Reply_Date%>&nbsp;</td>
</tr>
<tr height=30px>
	<td width=100px>��Ʈ�ѹ�</td>
	<td width=380px colspan=3 align=left><img src="/img/blank.gif" width=8px height=1px><%=BOM_B_D_No%></td>
	<td width=100px>����</td>
	<td width=140px>LG����(��)</td>
</tr>
<tr height=250px>
	<td width=100px>�ù泻��</td>
	<td colspan=5 align=center valign=center>
		<table width=600px height=230px cellpadding=0 cellspacing=0 border=0 bordercolor=black style="font-face:����;font-size:14px;table-layout:fixed;">
		<tr height=230px>
			<td width=600px align=left valign=top><pre><%=BU_Content%></pre></td></td>
		</tr>
		</table>
	</td>	
</tr>	
</table>
<img src="/img/blank.gif" width=10px height=30px><br>
<table width=720px cellpadding=0 cellspacing=0 border=1 bordercolor=black style="font-face:����;font-size:14px;table-layout:fixed;">
<tr height=30px>
	<td width=100px>�μ�</td>
	<td width=320px>���ΰ��� ���� �� �����ǰ�</td>
	<td width=100px>���</td>
	<td width=100px>��¥</td>
	<td width=100px>Ȯ��</td>
</tr>
<tr height=30px>
	<td>���������</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>����1��</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>����2��</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>ǰ����</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>����1��</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>����2��(IMD)</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>����2��(SMT)</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>������</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>������</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>��������</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr height=30px>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
</table>

<script language="javascript">
factory.printing.header				= "";
factory.printing.footer				= "";
factory.printing.portrait			= true;
factory.printing.leftMargin			= 0.5;
factory.printing.rightMargin		= 0.5;
factory.printing.topMargin			= 0.5;
factory.printing.bottomMargin		= 0.5;
factory.printing.print(true, window);
self.close();
</script>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->