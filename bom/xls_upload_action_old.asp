<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<% 
dim strError

dim B_Code
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
dim DNOSUB

dim QTY

dim Model_CNT
dim Parts_CNT

dim MODEL_LOC
dim QTY_LOC
dim PNO_LOC
dim DESCRIPTION_LOC
dim SPEC_LOC
dim REMARK_LOC
dim TYPE_LOC
dim CHECKSUM_LOC
dim MAKER_LOC
dim LOC_LOC
dim PNO2_LOC

dim strBS_D_No

set UpLoad	= Server.CreateObject("Dext.FileUpLoad")
UpLoad.DefaultPath = DefaultPath_BOM_XLS_Reader

dim Diff_YN
Diff_YN = UpLoad("Diff_YN")

B_Code = UpLoad("B_Code")

arrBOM_XLS_File_Name	= split(UpLoad("BOM_XLS"),"\")
BOM_XLS_File_Name		= arrBOM_XLS_File_Name(ubound(arrBOM_XLS_File_Name))
XLSDNO					= left(BOM_XLS_File_Name,instr(BOM_XLS_File_Name,".")-1)

temp = UpLoad("BOM_XLS").SaveAs(DefaultPath_BOM_XLS_Reader & BOM_XLS_File_Name, False)

set objXLS = Server.CreateObject("ADODB.Connection") 
XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""

objXLS.Open XLSConnection    
set RS1 = objXLS.OpenSchema(20)
Sheet_Name	= RS1(2)
Sheet_Name	= "["&replace(Sheet_Name,"'","")&"]"
set RS1 = nothing

set RS1 = Server.CreateObject("ADODB.RecordSet") 
SQL  = " select * from [admin$]"
RS1.Open SQL,objXLS 
arrXLS = RS1.getRows()
RS1.close

SQL = "select B_D_No from tbBOM where B_Code='"&B_Code&"'"
RS1.Open SQL,sys_DBCon
DNO = RS1("B_D_No")
RS1.Close

strBS_D_No = ""
SQL = "select BS_D_No from tbBOM_Sub where BS_Confirm_YN = 'Y' and BOM_B_Code="&B_Code
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strBS_D_No = strBS_D_No & RS1("BS_D_No") & "-"
	RS1.MoveNext
loop
strBS_D_No = "-"&strBS_D_No
RS1.Close
set RS1 = nothing
%>
<html>
<head>
</head>
<body>
<form name="frmXLS_Upload" action="b_model_reg_form.asp" method="post">
<input type="hidden" name="B_Code" value="<%=B_Code%>">
<input type="hidden" name="DNO" value="<%=DNO%>"> <Br>

<%
Model_CNT = 0
Parts_CNT = 0

for CNT1 = 0 to ubound(arrXLS, 2)
	for CNT2 = 0 to ubound(arrXLS, 1)

		if CNT1 = 0 and trim(arrXLS(CNT2, CNT1)) <> "" then
			QTY_LOC		= CNT2
			MODEL_LOC	= CNT2
			Model_CNT = Model_CNT + 1
			
			DNOSUB = ""
			if len(trim(arrXLS(CNT2, CNT1))) > 2 then 
				DNOSUB = trim(arrXLS(CNT2, CNT1))
			elseif len(trim(arrXLS(CNT2, CNT1))) = 2 then 
				DNOSUB = DNO & trim(arrXLS(CNT2, CNT1))
			elseif len(trim(arrXLS(CNT2, CNT1))) = 1 then 
				if isnumeric(arrXLS(CNT2, CNT1)) then
					DNOSUB = DNO & "0" & trim(arrXLS(CNT2, CNT1))
				else
					DNOSUB = DNO & trim(arrXLS(CNT2, CNT1))
				end if
			end if
			
			if len(DNOSUB) = 0 then
				strError = "Upload Failed!\nError in [ModelName]"
			end if
%>
		<input type="hidden" name="DNOSUB" value="<%=ucase(DNOSUB)%>">
		<input type="hidden" name="DNOCONFIRM" value="<%if instr(strBS_D_No,"-"&ucase(DNOSUB)&"-") > 0 then%>Y<%else%>N<%end if%>"><Br>
<%		
		elseif CNT1 = 1 then
			if trim(lcase(arrXLS(CNT2, CNT1)))="p/no" then
				PNO_LOC			= CNT2
			elseif trim(lcase(arrXLS(CNT2, CNT1)))="p/no2" then
				PNO2_LOC			= CNT2
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"description") > 0 then
				DESCRIPTION_LOC	= CNT2			
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"spec") > 0 then
				SPEC_LOC		= CNT2
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"remark") > 0 then
				REMARK_LOC	= CNT2
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"checksum") > 0 then
				CHECKSUM_LOC	= CNT2
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"maker") > 0 then
				MAKER_LOC		= CNT2
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"loc") > 0 then
				LOC_LOC			= CNT2
			elseif instr(lcase(arrXLS(CNT2, CNT1)),"type") > 0 then
				TYPE_LOC		= CNT2
			end if
%>
<%
		elseif CNT1 >= 2 then
			if trim(arrXLS(PNO_LOC, CNT1)) <> "" then
				if CNT2 = PNO_LOC then
					Parts_CNT = Parts_CNT + 1
%>
		<input type="hidden" name="PNO_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%		 
				elseif CNT2 = PNO2_LOC then
%>
		<input type="hidden" name="PNO2_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
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
				elseif CNT2 = CHECKSUM_LOC then
%>
		<input type="hidden" name="CHECKSUM_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%					
				elseif CNT2 = MAKER_LOC then
%>
		<input type="hidden" name="MAKER_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%					
				elseif CNT2 = TYPE_LOC then
%>
		<input type="hidden" name="WORKTYPE_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%					
				elseif CNT2 = LOC_LOC then
%>
		<input type="hidden" name="NO_<%=CNT1-1%>" value="<%=arrXLS(CNT2, CNT1)%>">
<%					
				elseif CNT2 >= 11 then
					QTY = arrXLS(CNT2, CNT1)
					if QTY <> "" then
						QTY = replace(trim(QTY),"-","0")
					else
						QTY = "0"
					end if
					
					if not(isnumeric(QTY)) then
						strError = "업로드 실패!\n자재의 수량에는 0또는 공백이 허용됩니다."
					end if	
					
%>
		<input type="hidden" name="QTY_<%=CNT1-1%>" value="<%=QTY%>">
<%							
				end if
			end if
		end if
		if strError <> "" then
			Exit For
		end if
	next
	if strError <> "" then
		Exit For
	end if
next
objXLS.close
set objXLS = nothing
set UpLoad = nothing

if strError = "" then
%>
<input type="hidden" name="Parts_CNT" value="<%=Parts_CNT%>">
<input type="hidden" name="Model_CNT" value="<%=Model_CNT%>">
<input type="hidden" name="oldParts_CNT" value="<%=Parts_CNT%>">
<input type="hidden" name="oldModel_CNT" value="<%=Model_CNT%>">
<input type="hidden" name="Diff_YN" value="<%=Diff_YN%>">
</form>
</body>
</html>
<script language="javascript">
frmXLS_Upload.submit();
</script>
<%
else
%>
</form>
</body>
</html>
<form name="frmRedirect" action="db_load_action.asp" method=post>
<input type="hidden" name="B_Code" value="<%=B_Code%>">
<input type="hidden" name="DNO" value="<%=DNO%>">
</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->