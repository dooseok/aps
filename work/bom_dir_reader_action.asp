<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
Dim B_Code
dim CNT1
dim CNT2
dim UpLoad
dim objXLS
dim XLSConnection
dim arrBOM_XLS_File_Name
dim BOM_XLS_File_Name
dim temp
dim RS1
dim SQL
dim Sheet_Name
dim arrXLS

dim XLSDNO
dim DNO

dim QTY
dim Model_CNT

dim MODEL_LOC
dim QTY_LOC
dim PNO_LOC
dim DESCRIPTION_LOC
dim SPEC_LOC
dim REMARK_LOC
dim MAKER_LOC
dim LOC_LOC

set UpLoad	= Server.CreateObject("Dext.FileUpLoad")



temp = "d:\home\msekorea\admin\bom_src\" & Request("strFile")
BOM_XLS_File_Name		= Request("strFile")
XLSDNO					= left(BOM_XLS_File_Name,instr(BOM_XLS_File_Name,".")-1)

SQL = "insert into tbBOM (B_D_No,B_Current_YN,B_Issue_Date,B_Reg_Date,B_Edit_Date) values ('"&XLSDNO&"','Y','"&date()&"','"&date()&"','"&date()&"')"
sys_DBCon.execute(SQL)

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select max(B_Code) as B_Code from tbBOM"
RS1.Open SQL,sys_DBCon
B_Code = RS1("B_Code")
RS1.Close
set RS1 = nothing

set objXLS = Server.CreateObject("ADODB.Connection") 
XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""

objXLS.Open XLSConnection    
set RS1 = objXLS.OpenSchema(20)
Sheet_Name	= RS1(2)
Sheet_Name	= "["&replace(Sheet_Name,"'","")&"]"
set RS1 = nothing

set RS1 = Server.CreateObject("ADODB.RecordSet") 
SQL  = " select * from "&Sheet_Name
RS1.Open SQL,objXLS 
arrXLS = RS1.getRows()
RS1.close

SQL = "select B_D_No from tbBOM where B_Code='"&B_Code&"'"
RS1.Open SQL,sys_DBCon
DNO = RS1("B_D_No")
RS1.Close
set RS1 = nothing
%>
<html>
<head>
</head>
<body>
<form name="frmXLS_Upload" action="/bom/b_model_reg_form.asp" method="post">
<input type="hidden" name="B_Code" value="<%=B_Code%>">
<input type="hidden" name="DNO" value="<%=DNO%>"> <Br>

<%
Model_CNT = 0
for CNT1 = 0 to ubound(arrXLS, 2)
	for CNT2 = 0 to ubound(arrXLS, 1)
		//response.write arrXLS(CNT2, CNT1) & " / "
%>
		
		
<%
		if CNT1 = 0 and trim(arrXLS(CNT2, CNT1)) <> "" then
			QTY_LOC		= CNT2
			MODEL_LOC	= CNT2
			Model_CNT = Model_CNT + 1
%>
		<input type="hidden" name="DNOSUB" value="<%=arrXLS(CNT2, CNT1)%>"> <Br>
<%		
		elseif CNT1 = 1 then
			if instr(lcase(arrXLS(CNT2, CNT1)),"p/no") > 0 then
				PNO_LOC			= CNT2
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"desc") > 0 then
				DESCRIPTION_LOC	= CNT2			
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"spec") > 0 then
				SPEC_LOC		= CNT2
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"remark") > 0 then
				REMARK_LOC		= CNT2
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"maker") > 0 then
				MAKER_LOC		= CNT2
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"loc") > 0 then
				LOC_LOC			= CNT2
			end if
%>
<%
		elseif CNT1 >= 2 then
			if CNT2 = PNO_LOC then 
%>
		<input type="hidden" name="PNO_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%		
			elseif CNT2 = DESCRIPTION_LOC then
%>
		<input type="hidden" name="DESCRIPTION_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%					
			elseif CNT2 = SPEC_LOC then
%>
		<input type="hidden" name="SPEC_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%					
			elseif CNT2 = REMARK_LOC then
%>
		<input type="hidden" name="REMARK_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%					
			elseif CNT2 = MAKER_LOC then
%>
		<input type="hidden" name="MAKER_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%					
			elseif CNT2 = LOC_LOC then
%>
		<input type="hidden" name="NO_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%					
			else
				QTY = arrXLS(CNT2, CNT1)
				if QTY <> "" then
					QTY = replace(trim(QTY),"-","0")
				else
					QTY = "0"
				end if
%>
		<input type="hidden" name="QTY_<%=CNT1-1%>" value="<%=QTY%>">
<%							
			end if
		end if
	next
	response.write "<br>"
next

'call File_Delete(DefaultPath_BOM_XLS_Reader & BOM_XLS_File_Name)
%>
<input type="hidden" name="Parts_CNT" value="<%=ubound(arrXLS, 2)-1%>"> <Br>
<input type="hidden" name="Model_CNT" value="<%=Model_CNT%>"> <Br>
<input type="hidden" name="oldParts_CNT" value="<%=ubound(arrXLS, 2)-1%>"> <Br>
<input type="hidden" name="oldModel_CNT" value="<%=Model_CNT%>"> <Br>
</form>
</body>
</html>
<script language="javascript">
frmXLS_Upload.submit();
</script>
<%
objXLS.close
set objXLS = nothing
set UpLoad = nothing
%>
<!-- #include virtual = "/header/db_tail.asp" -->
