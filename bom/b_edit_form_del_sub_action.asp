<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<% 
dim SQL
dim RS1

dim B_Code
dim B_D_No
dim BS_D_No

B_Code = request("B_Code")
B_D_No = request("B_D_No")
BS_D_No = request("BS_D_No")

dim strError

strError = ""

set RS1 = Server.CreateObject("ADODB.RecordSet")

if strError = "" then
	SQL = "select * from tbBOM_Sub where BOM_B_Code = "&B_Code&" and BS_D_No = '"&BS_D_No&"'"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		strError = "해당 품번은 존재하지 않습니다."
	end if
	RS1.Close
end if

if strError = "" then
	SQL = "select count(*) from tbBOM_Sub where BOM_B_Code = "&B_Code
	RS1.Open SQL,sys_DBCon
	if RS1(0) = 1 then
		strError = "유일한 품번은 삭제할 수 없습니다."
	end if
	RS1.Close
end if

set RS1 = nothing

if strError = "" then
	SQL = "delete tbBOM_Qty where BOM_B_Code = "&B_Code&" and BOM_Sub_BS_D_No = '"&BS_D_No&"'"
	sys_DBCon.execute(SQL)
	
	SQL = "delete tbBOM_Qty_Archive where BOM_B_Code = "&B_Code&" and BOM_Sub_BS_D_No = '"&BS_D_No&"'"
	sys_DBCon.execute(SQL)
	
	SQL = "delete tbBOM_Sub where BOM_B_Code = "&B_Code&" and BS_D_No = '"&BS_D_No&"'"
	sys_DBCon.execute(SQL)
	
	'SQL = "delete tbBOM where B_Code = "&B_Code
	'sys_DBCon.execute(SQL)
end if
%>
<form name="frmRedirect" action="b_edit_form.asp" method=post>
<input type="hidden" name="B_Code" value="<%=B_Code%>">
<input type="hidden" name="postback_yn" value="Y">
</form>
<script language="javascript">
<%
if strError = "" then
%>
alert("품번 [<%=BS_D_No%>]에 대한 정보가 삭제 되었습니다.");
<%
else
%>
alert("<%=strError%>");
<%
end if
%>
frmRedirect.submit();
</script>


<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->