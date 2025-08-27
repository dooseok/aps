<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
dim CNT1
dim RS1
dim RS2
dim SQL

dim s_lstBOM_Sub_BS_D_No
dim s_lstQty
dim arrBOM_Sub_BS_D_No
dim arrQty

dim BOM_Sub_BS_D_No
dim Qty

s_lstBOM_Sub_BS_D_No = request("s_lstBOM_Sub_BS_D_No")
s_lstQty = request("s_lstQty")

arrBOM_Sub_BS_D_No	= split(s_lstBOM_Sub_BS_D_No,chr(13)&chr(10))
arrQty				= split(s_lstQty,chr(13)&chr(10))
%>
<Script language="javascript">
function check_frmQuery()
{
	if(!frmQuery.s_lstBOM_Sub_BS_D_No.value)
	{
		alert("파트넘버를 한개 이상 입력해주십시오.");
		return false;
	}

	if(!frmQuery.s_lstQty.value)
	{
		alert("수량을 한개 이상 입력해주십시오.");
		return false;
	}

	frmQuery.submit();
}
</Script>
<table width=800px cellpadding=0 cellspacing=0 border=0>
<tr>
	<td><b>조회할 정보 입력</b></td>
</tr>
<tr>
	<td>
		<table width=50% height=200px cellpadding="0" cellspacing="1" border=0 bgcolor="black">
		<form name="frmQuery" action="j_list.asp" method="post">
		<tr bgcolor="white">
			<td width=40% align="center">
				파트넘버<br>
				<textarea style="width:90%;height=100%" name="s_lstBOM_Sub_BS_D_No"><%=ucase(request("s_lstBOM_Sub_BS_D_No"))%></textarea>
			</td>	
			<td width=40% align="center">
				수량<br>
				<textarea style="width:90%;height=100%" name="s_lstQty"><%=request("s_lstQty")%></textarea>
			</td>
			<td width=20%>
				<input type="button" value="조회" onclick="javascript:check_frmQuery();"><br><br>
				<input type="button" value="엑셀"><br><br>
				<input type="button" value="리셋" onclick="javascript:frmQuery.s_lstBOM_Sub_BS_D_No.value='';frmQuery.s_lstQty.value=''">
			</td>		
		</tr>
		</form>
		</table>	

	</td>
</tr>
<%
if s_lstBOM_Sub_BS_D_No <> "" then
%>
<tr>
	<td>&nbsp;</td>
</tr>
<tr>
	<td><b>재료비 리스트</b></td>
</tr>
<tr>
	<td>
		<table width=100% cellpadding="0" cellspacing="1" border=0 bgcolor="black">		
		<tr bgcolor="white">
			<td>파트넘버</td>
			<td>판가</td>
			<td>수량</td>
			<td>매출액</td>
		</tr> 
<%
	dim BQ_Qty
	dim BP_Price
	dim Parts_P_P_No
	dim Part_Price

	set RS1 = server.CreateObject("ADODB.RecordSet")
	set RS2 = server.CreateObject("ADODB.RecordSet")

	SQL = "select distinct Parts_P_P_No, M_Price = (select top 1 MO_Price from tbMaterial_Order where Material_M_P_No = Parts_P_P_No order by MO_Code desc),CP_Price = (select top 1 CP_Price from tbCOSP_Price where Material_M_P_No = Parts_P_P_No), IC_Price = (tbMaterial where ) from tbBOM_Qty where BOM_Sub_BS_D_No in ('"&ucase(replace(s_lstBOM_Sub_BS_D_No,chr(13)&chr(10),"','"))&"') and BQ_Qty > 0"
	response.write SQL

	for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)
		BOM_Sub_BS_D_No = ucase(trim(arrBOM_Sub_BS_D_No(CNT1)))
		Qty = trim(arrQty(CNT1))

		SQL = "select top 1 BP_Price from tbBOM_Price where BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"' order by BP_Code desc"
		RS1.Open SQL,sys_DBCon
		BP_Price = RS1(0)
		RS1.Close

		SQL = "select Parts_P_P_No, BQ_Qty from tbBOM_Qty where BOM_Sub_BS_D_No = '"&Bom_Sub_BS_D_No&"' and BQ_Qty > 0"
		RS1.Open SQL,sys_DBCon
		do until RS1.Eof
			Parts_P_P_No = RS1("Parts_P_P_No")
			BQ_Qty = RS1("BQ_Qty")
			Part_Price = 0

			'tbMaterial에서 조회'
			SQL = "select M_Price from tbMaterial_Order where Material_M_P_No = '"&Part_P_P_No&"' and MO_Order_Date "
			RS2.Open SQL,sys_DBCon
			if not(RS2.Eof or RS2.Bof) then
				M_Price = RS2("M_Price")
			end if
			RS2.Close

			RS1.MoveNext
		loop
		RS1.Close
%>
		<tr bgcolor="white">
			<td><%=BOM_Sub_BS_D_No%></td>
			<td><%=CustomFormatComma(BP_Price)%></td>
			<td><%=Qty%></td>
			<td><%=CustomFormatComma(BP_Price*Qty)%></td>
		</tr>
<%
	next
	set RS2 = nothing
	set RS1 = nothing
%>
		</table>
	</td>
</tr>
<%
end if
%>	
</table>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->