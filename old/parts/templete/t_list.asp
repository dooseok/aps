<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim RS1
dim RS2
dim SQL

dim s_MT_Name
dim s_Opener_Type
dim s_Opener_SubType
dim s_Opener_Code

s_MT_Name			= Request("s_MT_Name")
s_Opener_Type		= Request("s_Opener_Type")
s_Opener_SubType	= Request("s_Opener_SubType")
s_Opener_Code		= Request("s_Opener_Code")



if Request("newMT_Name") <> "" then
	SQL = "insert into tbMaterial_Templete (MT_Name) values ('"&Request("newMT_Name")&"')"
	sys_DBCon.execute(SQL)
end if

if Request("delMT_Name") <> "" then
	SQL = "delete tbMaterial_Templete where MT_Name = '"&Request("delMT_Name")&"'"
	sys_DBCon.execute(SQL)
end if

dim MT_Company
dim oldM_Qty

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")

if Request("applyMT_Name") <> "" then
	if s_Opener_Type = "Transaction" then
		SQL = "select Material_M_P_No, MTD_Count from tbMaterial_Templete_Detail where Material_Templete_MT_Name='"&Request("applyMT_Name")&"'"
		RS1.Open SQL,sys_DBCon
		do until RS1.Eof
			SQL = "insert into tbMaterial_Transaction_Detail (Material_Transaction_MT_Code,Material_M_P_No,MTD_Qty,MTD_Remark) values ('"&s_Opener_Code&"','"&RS1("Material_M_P_No")&"',"&RS1("MTD_Count")&",'')"
			sys_DBCon.execute(SQL)
			
			SQL = "select M_Qty from tbMaterial where M_Code='"&RS1("Material_M_P_No")&"'"
			RS2.Open SQL,sys_DBCon
			oldM_Qty = RS2("M_Qty")
			RS2.Close
			
			if s_Opener_SubType = "Ipgo" then
				SQL = "update tbMaterial set M_Qty = M_Qty + "&RS1("MTD_Count")&" where M_P_No = '"&RS1("Material_M_P_No")&"'"
				sys_DBCon.execute(SQL)
			else
				SQL = "update tbMaterial set M_Qty = M_Qty - "&RS1("MTD_Count")&" where M_P_No = '"&RS1("Material_M_P_No")&"'"
				sys_DBCon.execute(SQL)
			end if
			
			SQL = "select MT_Company from tbMaterial_Material where MT_Code = "&s_Opener_Code
			RS2.Open SQL,sys_DBCon
			MT_Company = RS2("MT_Company")
			RS2.Close
			
			SQL = "insert into tbMaterial_Stock_History (Material_M_P_No,MSH_Change_Stock,MSH_Applyed_Stock,MSH_Change_Type,MSH_Change_Date,MSH_Company) values ('"
			SQL = SQL &RS1("Material_M_P_No")&"',"
			if s_Opener_SubType = "Ipgo" then
				SQL = SQL &RS1("MTD_Count")&","
				SQL = SQL &oldM_Qty+RS1("MTD_Count")&",'"
				SQL = SQL &"사내입고','"
			else
				SQL = SQL &int(RS1("MTD_Count")) * -1 &","
				SQL = SQL &oldM_Qty-RS1("MTD_Count")&",'"
				SQL = SQL &"사내출고','"
			end if
			SQL = SQL &date()&"','"
			SQL = SQL &MT_Company&"')"
			if RS1("MTD_Count") <> 0 then
				sys_DBCon.execute(SQL)
			end if
			
			RS1.MoveNext
		loop
		RS1.Close
	elseif s_Opener_Type = "Order" then
		dim MOD_Price
		
		SQL = "select Material_M_P_No, MTD_Count from tbMaterial_Templete_Detail where Material_Templete_MT_Name='"&Request("applyMT_Name")&"'"
		RS1.Open SQL,sys_DBCon
		do until RS1.Eof
			MOD_Price = ""
			SQL = "select M_Price from tbMaterial where M_P_No ='"&RS1("Material_M_P_No")&"'"
			RS2.Open SQL,sys_DBCon
			if RS2.Eof or RS1.Bof then
				MOD_Price = 0
			else
				MOD_Price = RS2("M_Price")
			end if
			RS2.Close
			SQL = "insert into tbMaterial_Order_Detail (Material_Order_MO_Code,Material_M_P_No,MOD_Price,MOD_Qty,MOD_In_Qty,MOD_Remark) values ('"&s_Opener_Code&"','"&RS1("Material_M_P_No")&"',"&MOD_Price&","&RS1("MTD_Count")&",0,'')"
			sys_DBCon.execute(SQL)
			RS1.MoveNext
		loop
		RS1.Close
		
	end if
%>
<script language="javascript">
parent.opener.location.reload();
parent.self.close();
</script>
<%
else
%>

	<script language="javascript">
	function view_change(strMT_Name)
	{
		parent.tview.location.href='/material/templete/detail/mtd_list.asp?s_MT_Name=' + strMT_Name;
	}

	function frmNew_Check()
	{
		if (frmNew.newMT_Name.value == '')
		{
			alert("등록할 이름을 입력해주세요.\n예)6871A20133R_수삽");
			return false;
		}
		else
		{
			frmNew.submit();
			return true;
		}
	}

	function frmSearch_Check()
	{
		frmSearch.submit();
	}

	function frmDel_Check(strMT_Name)
	{
		if(confirm("삭제하시겠습니까?"))
		{
			location.href='t_list.asp?delMT_Name=' + strMT_Name + '&s_Opener_Type=<%=s_Opener_Type%>&s_Opener_SubType=<%=s_Opener_SubType%>&s_Opener_Code=<%=s_Opener_Code%>';
		}
	}

	function frmApply_Check(strMT_Name)
	{
		if(confirm("적용하시겠습니까?"))
		{
			location.href='t_list.asp?applyMT_Name=' + strMT_Name + '&s_Opener_Type=<%=s_Opener_Type%>&s_Opener_SubType=<%=s_Opener_SubType%>&s_Opener_Code=<%=s_Opener_Code%>';
		}
	}
	</script>

	<table width=200px cellpadding=3 cellspacing=1 border=0 bgcolor="black">
	<form name="frmNew" action="t_list.asp" method="post">
	<input type="hidden" name="s_Opener_Type" value="<%=s_Opener_Type%>">
	<input type="hidden" name="s_Opener_SubType" value="<%=s_Opener_SubType%>">
	<input type="hidden" name="s_Opener_Code" value="<%=s_Opener_Code%>">
	<tr>
		<td colspan=3 align=center bgcolor=white><input type="text" name="newMT_Name" value="">&nbsp;<input type="button" onclick="javascript:frmNew_Check()" value="등록"></td>
	</tr>
	</form>
	<form name="frmSearch" action="t_list.asp" method="post">
	<input type="hidden" name="s_Opener_Type" value="<%=s_Opener_Type%>">
	<input type="hidden" name="s_Opener_SubType" value="<%=s_Opener_SubType%>">
	<input type="hidden" name="s_Opener_Code" value="<%=s_Opener_Code%>">
	<tr>
		<td colspan=3 align=center bgcolor=white><input type="text" name="s_MT_Name" value="<%=s_MT_Name%>">&nbsp;<input type="button" onclick="javascript:frmSearch_Check()" value="검색"></td>
	</tr>
	</form>
	<%
	if trim(s_MT_Name) = "" then
		SQL = "select * from tbMaterial_Templete order by mt_name"
	else
		SQL = "select * from tbMaterial_Templete where MT_Name = '"&s_MT_Name&"' order by mt_name"
	end if
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
	%>
	<tr bgcolor=white>
		<td colspan=3>검색 결과가 없습니다.</td>
	</tr>
	<%
	else
		do until RS1.Eof
	%>
	<tr bgcolor=white onmouseover="this.style.backgroundColor='skyblue';" onmouseout="this.style.backgroundColor='#ffffff';">
		<td align=center width=130px><span onclick="javascript:view_change('<%=Server.URLEncode(RS1("MT_Name"))%>')" style="cursor:hand"><%=RS1("MT_Name")%></span></td>
		<td align=center width=35px><span onclick="javascript:frmApply_Check('<%=Server.URLEncode(RS1("MT_Name"))%>')" style="cursor:hand">적용</span></td>
		<td align=center width=35px><span onclick="javascript:frmDel_Check('<%=Server.URLEncode(RS1("MT_Name"))%>')" style="cursor:hand">삭제</span></td>
	</tr>
	<%
			RS1.MoveNext
		loop
	end if
	RS1.Close
	%>
	<table>
<%
end if

set RS1 = nothing
set RS2 = nothing
%>


<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->