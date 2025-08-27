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

dim xlsModel_CNT
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
dim STYPE_LOC
dim PNO2PinYN_LOC
dim LEVEL_LOC

dim remark

dim oldPNORemark
dim CNT3
dim tempBQ_Remark
dim tempREMARK_LOC

dim PNORemark_Match_YN
dim strBS_D_No
dim strP_P_No
dim strP_P_No2
dim strP_P_No2_PinYN
dim strBQ_P_Desc
dim strBQ_P_Spec
dim strBQ_Remark
dim strBQ_Checksum
dim strBQ_P_Maker
dim strP_Work_Type
dim strBQ_Order
dim strBQ_Qty
dim arrBS_D_No
dim arrP_P_No
dim arrP_P_No2
dim arrP_P_No2_PinYN
dim arrBQ_P_Desc
dim arrBQ_P_Spec
dim arrBQ_Remark
dim arrBQ_Checksum
dim arrBQ_P_Maker
dim arrP_Work_Type
dim arrBQ_Order
dim arrBQ_Qty
dim strAdded

dim strDNOSUB

set UpLoad	= Server.CreateObject("Dext.FileUpLoad")
UpLoad.DefaultPath = DefaultPath_BOM_XLS_Reader

B_Code = UpLoad("B_Code")

arrBOM_XLS_File_Name	= split(UpLoad("BOM_XLS"),"\")
BOM_XLS_File_Name		= arrBOM_XLS_File_Name(ubound(arrBOM_XLS_File_Name))
XLSDNO					= left(BOM_XLS_File_Name,instr(BOM_XLS_File_Name,".")-1)

dim Diff_YN
Diff_YN = UpLoad("Diff_YN")

temp = UpLoad("BOM_XLS").SaveAs(DefaultPath_BOM_XLS_Reader & BOM_XLS_File_Name, False)

set objXLS = Server.CreateObject("ADODB.Connection") 
XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""
'XLSConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & temp & "; Extended Properties=""Excel 12.0 Xml;HDR=YES"""

On Error Resume Next
objXLS.Open XLSConnection
If Err.Number <> 0 then
	strError = "[파일업로드 실패] 다음의 내용을 확인해주십시오.\n1. 업로드할 파일은 엑셀 97-2003 통합문서 형식이어야 합니다\n2. 시트명은 [admin], [bomViewList.xls], [BOM Explosion] 등이 가능합니다\n3. 파일의 이름은 파트넘버와 일치하여야 합니다. 예)"&DNO&"01(...).xls"
%>
	<form name="frmRedirect" action="db_load_action.asp" method=post>
	<input type="hidden" name="B_Code" value="<%=B_Code%>">
	<input type="hidden" name="DNO" value="<%=DNO%>">
	</form>
	<script language="javascript">
	alert("<%=strError%>");
	frmRedirect.submit();
	</script>
<%
	response.end
end if
On Error GoTo 0

set RS1 = objXLS.OpenSchema(20)
do until RS1.Eof
	Sheet_Name = RS1(2)
	RS1.MoveNext
loop
RS1.Close
set RS1 = nothing

set RS1 = Server.CreateObject("ADODB.RecordSet") 
SQL = "select B_D_No from tbBOM where B_Code="&B_Code
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

if instr(XLSDNO,DNO) = 0 then
	strError = "업로드할 파일의 이름은 파트넘버가 포함되어야 합니다.\n예)"&DNO&"01(...).xls"
end if
%>

<html>
<head>
</head>
<body>
<form name="frmXLS_Upload" action="b_model_reg_form.asp" method="post">
<input type="hidden" name="B_Code" value="<%=B_Code%>">
<input type="hidden" name="DNO" value="<%=DNO%>"> <Br>

<%
dim Log




if strError = "" then
	Sheet_Name = replace(Sheet_Name,"_","")
	Sheet_Name = replace(Sheet_Name,"'","")
	SQL  = " select * from ["&Sheet_Name&"]"
	
	RS1.Open SQL,objXLS 
	arrXLS = RS1.getRows()
	RS1.close
	if CheckField(arrXLS,0,"level|_|part name|_|description|_|part site specification|_|designator/split|_|maker|_|job explanation|_|seq|_|qty|_|supply type") then
		
		call xlsUpbomViewList(arrXLS, strBS_D_No)
	elseif CheckField(arrXLS,0,"level|_|part no|_|item desc|_|item spec|_|location info.|_|comments|_|no|_|component qty|_|parent part no|_|substitute item|_|supply type") then
		call xlsUpBOMExplosion(arrXLS, strBS_D_No)
	elseif CheckField(arrXLS,1,"p/no|_|description|_|spec|_|remark|_|maker|_|loc") then
		call xlsUpAdmin(arrXLS, strBS_D_No)
	else
		strError = "1. 업로드할 파일은 엑셀 97-2003 통합문서 형식이어야 합니다\n2. 시트명은 [admin], [bomViewList.xls], [BOM Explosion] 등이 가능합니다\n3. 파일의 이름은 파트넘버와 일치하여야 합니다. 예)"&DNO&"01(...).xls"
	end if
end if

function CheckField(arrXLS, nFieldRow, strFindField)
	dim CNT1
	dim strXLSField
	dim arrFindField
	dim bResult
	
	strXLSField = "-"
	for CNT1 = 0 to ubound(arrXLS, 1)
		strXLSField = strXLSField & trim(lcase(arrXLS(CNT1, nFieldRow))) &"-"
	next
	
	arrFindField = split(strFindField,"|_|")
	
	bResult = true
	for CNT1 = 0 to ubound(arrFindField)
		if instr(strXLSField,arrFindField(CNT1)) = 0 then
			bResult = false
		end if
	next
	
	CheckField = bResult
end function

sub xlsUpbomViewList(arrXLS, strBS_D_No)
	dim strLevel
	
	dim LEVEL_LOC 
	dim JOB_LOC
	dim strNO
	dim strMaker
	
	Model_CNT = 1
	Parts_CNT = 0
	
	for CNT1 = 0 to ubound(arrXLS, 1)
		if trim(lcase(arrXLS(CNT1, 0)))="level" then
			LEVEL_LOC		= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="part name" then
			PNO_LOC			= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="description" then
			DESCRIPTION_LOC	= CNT1			
		elseif trim(lcase(arrXLS(CNT1, 0)))="part site specification" then
			SPEC_LOC		= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="designator/split" then
			REMARK_LOC		= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="maker" then
			MAKER_LOC		= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="job explanation" then
			JOB_LOC			= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="seq" then
			LOC_LOC			= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="qty" then
			QTY_LOC			= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="supply type" then
			STYPE_LOC		= CNT1
		end if
	next
	
	for CNT1 = 1 to ubound(arrXLS, 2)
		strLevel = ""
		if isnull(arrXLS(LEVEL_LOC, CNT1)) then
			strLevel = ""
		else
			strLevel = trim(cstr(arrXLS(LEVEL_LOC, CNT1)))
		end if
	
		if strLevel = "0" then
			DNOSUB = arrXLS(PNO_LOC, CNT1)
			strDNOSUB = strDNOSUB & "'"&ucase(DNOSUB) &"',"
%>
		<input type="hidden" name="DNOSUB" value="<%=ucase(DNOSUB)%>">
		<input type="hidden" name="DNOCONFIRM" value="<%if instr(strBS_D_No,"-"&ucase(DNOSUB)&"-") > 0 then%>Y<%else%>N<%end if%>"><Br>
<%
		end if
			
			
		if instr(arrXLS(LEVEL_LOC, CNT1),".*Q*") = 0 and strLevel <> "0" and trim(arrXLS(PNO_LOC, CNT1)) <> "" then
			Parts_CNT = Parts_CNT + 1
		
			
			remark = trim(arrXLS(REMARK_LOC, CNT1))
			if not(isnull(remark)) then
				remark = replace(remark,chr(13),"")
			end if
			
			QTY = arrXLS(QTY_LOC, CNT1)
			if QTY <> "" then
				QTY = replace(trim(QTY),"-","0")
			else
				QTY = "0"
			end if
			
			if not(isnumeric(QTY)) then
				strError = "Check QTY [ "&QTY&" ] of PNO [ "&arrXLS(PNO_LOC, CNT1)&" ]\nLocation [ "&mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ",CNT2+1,1)&CNT1+1&" ]"
			end if					
			
			if instr(arrXLS(LEVEL_LOC, CNT1),".*S*") > 0 then
				strNO = "R"
				'QTY = ""
			else
				strNO = trim(arrXLS(LOC_LOC, CNT1))
			end if
			
			if trim(arrXLS(JOB_LOC, CNT1)) <> "" and trim(arrXLS(MAKER_LOC, CNT1)) <> "" then
				if trim(arrXLS(JOB_LOC, CNT1)) = trim(arrXLS(MAKER_LOC, CNT1)) then
					strMaker = trim(arrXLS(JOB_LOC, CNT1))
				else
					strMaker = trim(arrXLS(JOB_LOC, CNT1))&","&trim(arrXLS(MAKER_LOC, CNT1))
				end if
			elseif trim(arrXLS(JOB_LOC, CNT1)) <> "" and trim(arrXLS(MAKER_LOC, CNT1)) = "" then
				strMaker = trim(arrXLS(JOB_LOC, CNT1))
			elseif trim(arrXLS(JOB_LOC, CNT1)) = "" and trim(arrXLS(MAKER_LOC, CNT1)) <> "" then
				strMaker = trim(arrXLS(MAKER_LOC, CNT1))
			else
				strMaker = "-"
			end if
%>
		<input type="hidden" name="PNO_<%=Parts_CNT%>"			value="<%=trim(arrXLS(PNO_LOC, CNT1))%>">
		<input type="hidden" name="DESCRIPTION_<%=Parts_CNT%>"	value="<%=trim(arrXLS(DESCRIPTION_LOC, CNT1))%>">
		<input type="hidden" name="SPEC_<%=Parts_CNT%>"			value="<%=trim(arrXLS(SPEC_LOC, CNT1))%>">
		<input type="hidden" name="MAKER_<%=Parts_CNT%>"		value="<%=strMaker%>">
		<input type="hidden" name="NO_<%=Parts_CNT%>"			value="<%=strNO%>">
		<input type="hidden" name="REMARK_<%=Parts_CNT%>"		value="<%=remark%>">
		<input type="hidden" name="STYPE_<%=Parts_CNT%>"		value="<%=trim(arrXLS(SType_LOC, CNT1))%>">
		<%if strNO = "R" then%>
		<input type="hidden" name="QTY_<%=DNOSUB%>_<%=trim(arrXLS(PNO_LOC, CNT1))%>_<%=remark%>_R" value="<%=QTY%>"  >
		<%else%>
		<input type="hidden" name="QTY_<%=DNOSUB%>_<%=trim(arrXLS(PNO_LOC, CNT1))%>_<%=remark%>_X" value="<%=QTY%>"  >
		<%end if%>
<%							
		end if
	next
	
	xlsModel_CNT = Model_CNT

	strDNOSUB = left(strDNOSUB,len(strDNOSUB)-1)
end sub

sub xlsUpBOMExplosion(arrXLS, strBS_D_No)
	dim PARENT_LOC 
	dim SUB_LOC
	dim STYPE_LOC
	
	dim strLevel
	
	dim strSUB
	dim arrSUB
	dim strSpec
	dim strDesc
	dim strMaker
	dim strSType
	dim strNo
	
	dim oldQty
	
	Model_CNT = 1
	Parts_CNT = 0
	
	for CNT1 = 0 to ubound(arrXLS, 1)
		if trim(lcase(arrXLS(CNT1, 0)))="level" then
			LEVEL_LOC		= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="part no" then
			PNO_LOC			= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="item desc" then
			DESCRIPTION_LOC	= CNT1			
		elseif trim(lcase(arrXLS(CNT1, 0)))="item spec" then
			SPEC_LOC		= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="location info." then
			REMARK_LOC		= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="comments" then
			MAKER_LOC		= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="no" then
			LOC_LOC			= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="component qty" then
			QTY_LOC			= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="parent part no" then
			PARENT_LOC		= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="substitute item" then
			SUB_LOC		= CNT1
		elseif trim(lcase(arrXLS(CNT1, 0)))="supply type" then
			STYPE_LOC	= CNT1
		end if
	next
	
	oldQty = 0
	for CNT1 = 1 to ubound(arrXLS, 2)
	
		strLevel = ""
		if isnull(arrXLS(LEVEL_LOC, CNT1)) then
			strLevel = ""
		else
			strLevel = trim(cstr(arrXLS(LEVEL_LOC, CNT1)))
		end if
		
		if strLevel = "1" and DNOSUB = "" then
			DNOSUB = arrXLS(PARENT_LOC, CNT1)
			strDNOSUB = strDNOSUB & "'"&ucase(DNOSUB) &"',"
%>
		<input type="hidden" name="DNOSUB" value="<%=ucase(DNOSUB)%>">
		<input type="hidden" name="DNOCONFIRM" value="<%if instr(strBS_D_No,"-"&ucase(DNOSUB)&"-") > 0 then%>Y<%else%>N<%end if%>"><Br>
<%
		end if
		
		if strLevel <> "0" and trim(arrXLS(PNO_LOC, CNT1)) <> "" then
			Parts_CNT = Parts_CNT + 1
			
			remark = trim(arrXLS(REMARK_LOC, CNT1))
			if not(isnull(remark)) then
				remark = replace(remark,chr(13),"")
			end if
			
			QTY = arrXLS(QTY_LOC, CNT1)
			if QTY <> "" then
				if instr(QTY,"e") = 0 then
					QTY = replace(trim(QTY),"-","0")
				end if
			else
				QTY = "0"
			end if

			if ucase(arrXLS(LOC_LOC, CNT1)) = "R" or instr(ucase(arrXLS(Level_LOC, CNT1)),"S") > 0 then
				strNO = "R"
			'	QTY = oldQty
			else
				strNO = trim(arrXLS(LOC_LOC, CNT1))
			end if
			
			'oldQty = QTY
			
			if not(isnumeric(QTY)) then
				strError = "Check QTY [ "&QTY&" ] of PNO [ "&arrXLS(PNO_LOC, CNT1)&" ]\nLocation [ "&mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ",CNT2+1,1)&CNT1+1&" ]"
			end if									
%>
		<input type="hidden" name="PNO_<%=Parts_CNT%>"			value="<%=trim(arrXLS(PNO_LOC, CNT1))%>">
		<input type="hidden" name="DESCRIPTION_<%=Parts_CNT%>"	value="<%=trim(arrXLS(DESCRIPTION_LOC, CNT1))%>">
		<input type="hidden" name="SPEC_<%=Parts_CNT%>"			value="<%=trim(arrXLS(SPEC_LOC, CNT1))%>">
		<input type="hidden" name="MAKER_<%=Parts_CNT%>"		value="<%=trim(arrXLS(MAKER_LOC, CNT1))%>">
		<input type="hidden" name="STYPE_<%=Parts_CNT%>"		value="<%=trim(arrXLS(STYPE_LOC, CNT1))%>">
		<input type="hidden" name="NO_<%=Parts_CNT%>"			value="<%=strNO%>">
		<input type="hidden" name="REMARK_<%=Parts_CNT%>"		value="<%=remark%>">
		<%if strNO = "R" then%>
		<input type="hidden" name="QTY_<%=DNOSUB%>_<%=trim(arrXLS(PNO_LOC, CNT1))%>_<%=remark%>_R" value="<%=QTY%>"  >
		<%else%>
		<input type="hidden" name="QTY_<%=DNOSUB%>_<%=trim(arrXLS(PNO_LOC, CNT1))%>_<%=remark%>_X" value="<%=QTY%>"  >
		<%end if%>
<%				
		end if
	next
	
	xlsModel_CNT = Model_CNT

	strDNOSUB = left(strDNOSUB,len(strDNOSUB)-1)
end sub



sub xlsUpAdmin(arrXLS, strBS_D_No)
	Model_CNT = 0
	Parts_CNT = 0
	
	for CNT1 = 0 to ubound(arrXLS, 2)
		for CNT2 = 0 to ubound(arrXLS, 1)
	
			if CNT1 = 0 and trim(arrXLS(CNT2, CNT1)) <> "" then
				if QTY_LOC = "" then
					QTY_LOC		= CNT2
					MODEL_LOC	= CNT2
				end if
				
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
				
				strDNOSUB = strDNOSUB & "'"&ucase(DNOSUB) &"',"
%>
		<input type="hidden" name="DNOSUB" value="<%=ucase(DNOSUB)%>">
		<input type="hidden" name="DNOCONFIRM" value="<%if instr(strBS_D_No,"-"&ucase(DNOSUB)&"-") > 0 then%>Y<%else%>N<%end if%>"><Br>
<%		
			elseif CNT1 = 1 then
				if trim(lcase(arrXLS(CNT2, CNT1)))="p/no" then
					PNO_LOC			= CNT2
				elseif trim(lcase(arrXLS(CNT2, CNT1)))="p/no2" then
					PNO2_LOC			= CNT2
				elseif trim(lcase(arrXLS(CNT2, CNT1)))="p/no2pinyn" then
					PNO2PinYN_LOC			= CNT2
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
				elseif lcase(arrXLS(CNT2, CNT1))="type" then
					TYPE_LOC		= CNT2
				elseif lcase(arrXLS(CNT2, CNT1))="stype" then
					STYPE_LOC		= CNT2
				end if
%>
<%
			elseif CNT1 >= 2 then
				if trim(arrXLS(PNO_LOC, CNT1)) <> "" then
					if CNT2 = PNO_LOC then
						Parts_CNT = Parts_CNT + 1
%>
		<input type="hidden" name="PNO_<%=CNT1-1%>" value="<%=trim(arrXLS(CNT2, CNT1))%>">
<%		 
					elseif CNT2 = PNO2_LOC then
%>
		<input type="hidden" name="PNO2_<%=CNT1-1%>" value="<%=trim(arrXLS(CNT2, CNT1))%>">
<%		 
					elseif CNT2 = PNO2PinYN_LOC then
%>
		<input type="hidden" name="PNO2PinYN_<%=CNT1-1%>" value="<%=ucase(trim(arrXLS(CNT2, CNT1)))%>">
<%		 
					elseif CNT2 = DESCRIPTION_LOC then
%>
		<input type="hidden" name="DESCRIPTION_<%=CNT1-1%>" value="<%=trim(arrXLS(CNT2, CNT1))%>">
<%					
					elseif CNT2 = SPEC_LOC then
%>
		<input type="hidden" name="SPEC_<%=CNT1-1%>" value="<%=trim(arrXLS(CNT2, CNT1))%>">
<%					
					elseif CNT2 = REMARK_LOC then
						remark = trim(arrXLS(CNT2, CNT1))
						if not(isnull(remark)) then
							remark = replace(remark,chr(13),"")
						end if
%>
		<input type="hidden" name="REMARK_<%=CNT1-1%>" value="<%=remark%>">
<%					
					elseif CNT2 = CHECKSUM_LOC then
%>
		<input type="hidden" name="CHECKSUM_<%=CNT1-1%>" value="<%=trim(arrXLS(CNT2, CNT1))%>">
<%					
					elseif CNT2 = MAKER_LOC then
%>
		<input type="hidden" name="MAKER_<%=CNT1-1%>" value="<%=trim(arrXLS(CNT2, CNT1))%>">
<%					
					elseif CNT2 = TYPE_LOC then
%>
		<input type="hidden" name="WORKTYPE_<%=CNT1-1%>" value="<%=trim(arrXLS(CNT2, CNT1))%>">
<%					
					elseif CNT2 = STYPE_LOC then
%>
		<input type="hidden" name="STYPE_<%=CNT1-1%>" value="<%=trim(arrXLS(CNT2, CNT1))%>">
<%					
					elseif CNT2 = LOC_LOC then
%>
		<input type="hidden" name="NO_<%=CNT1-1%>" value="<%=trim(arrXLS(CNT2, CNT1))%>">
<%				
					elseif CNT2 >= QTY_LOC and trim(arrXLS(CNT2, 0)) <> "" then
						'if ucase(trim(arrXLS(CNT2, CNT1))) = "R" then
						'	QTY = ""
						'else
						'	QTY = arrXLS(CNT2, CNT1)
						'end if
						QTY = arrXLS(CNT2, CNT1)
						
						if QTY <> "" then
							QTY = replace(trim(QTY),"-","0")
						else
							QTY = "0"
						end if
						
						if not(isnumeric(QTY)) then
							strError = "Check QTY [ "&QTY&" ] of PNO [ "&arrXLS(PNO_LOC, CNT1)&" ]\nLocation [ "&mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ",CNT2+1,1)&CNT1+1&" ]"
						end if	
%>
		<%if ucase(arrXLS(LOC_LOC, CNT1)) = "R" then%>
		<input type="hidden" name="QTY_<%=trim(arrXLS(CNT2, 0))%>_<%=trim(arrXLS(PNO_LOC, CNT1))%>_<%=remark%>_R" value="<%=QTY%>"  >
		<%else%>
		<input type="hidden" name="QTY_<%=trim(arrXLS(CNT2, 0))%>_<%=trim(arrXLS(PNO_LOC, CNT1))%>_<%=remark%>_X" value="<%=QTY%>"  >
		<%end if%>
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
	xlsModel_CNT = Model_CNT

	strDNOSUB = left(strDNOSUB,len(strDNOSUB)-1)
end sub





if strError = "" then
	'위에서 엑셀파일에 존재하는 PNO들을 가져왔다면,
	'이번에는 엑셀에는 없지만 기존에 DB에 있는 자료들을 가져온다.
	SQL = "select * from tbBOM_Sub where BS_D_No not in ("&strDNOSUB&") and BOM_B_Code = "&B_Code&" order by BS_D_No asc"
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		Model_CNT = Model_CNT + 1
%>
	<input type="hidden" name="DNOSUB" value="<%=RS1("BS_D_No")%>">
	<input type="hidden" name="DNOCONFIRM" value="<%if instr(strBS_D_No,"-"&RS1("BS_D_No")&"-") > 0 then%>Y<%else%>N<%end if%>"><Br>
	
<%
		RS1.MoveNext
	loop
	RS1.Close
	
	
	SQL = ""
	SQL = SQL & "select "
	SQL = SQL & "	BOM_Sub_BS_D_No, "
	SQL = SQL & "	Parts_P_P_No, "
	SQL = SQL & "	Parts_P_P_No2, "
	SQL = SQL & "	Parts_P_P_No2_PinYN, "
	SQL = SQL & "	BQ_P_Desc, "
	SQL = SQL & "	BQ_P_Spec, "
	SQL = SQL & "	BQ_Remark, "
	SQL = SQL & "	BQ_Checksum, "
	SQL = SQL & "	BQ_P_Maker, "
	SQL = SQL & "	P_Work_Type = (select top 1 M_Process from tbMaterial where M_P_No = Parts_P_P_No2), "
	SQL = SQL & "	BQ_Order, "
	SQL = SQL & "	BQ_Qty "
	SQL = SQL & " from tbBOM_Qty where BOM_Sub_BS_D_No not in ("&strDNOSUB&") and BOM_B_Code = "&B_Code
	SQL = SQL & " order by BQ_Remark, BQ_Code, Parts_P_P_No, BOM_Sub_BS_D_No"
	RS1.Open SQL,sys_DBCon

	strBS_D_No = ""
	
	dim XLS_Order_R
	dim Uploaded_Order_R
	
	do until RS1.Eof
		PNORemark_Match_YN = "N"
	
		for CNT1 =0 to ubound(arrXLS, 2)
		
			tempBQ_Remark 	= RS1("BQ_Remark")
			tempREMARK_LOC	= arrXLS(REMARK_LOC, CNT1)
			if isnull(tempBQ_Remark) then
				tempBQ_Remark = ""
			end if
			if isnull(tempREMARK_LOC) then
				tempREMARK_LOC = ""
			end if
			
			'올리던게 R인지? 			
			if ucase(arrXLS(LOC_LOC, CNT1)) = "R" or instr(ucase(arrXLS(Level_LOC, CNT1)),"S") > 0 or instr(arrXLS(LEVEL_LOC, CNT1),".*S*") > 0 then
				XLS_Order_R = true
			else
				XLS_Order_R = false
			end if	
			
			'올라갔던게 R인지?
			if ucase(RS1("BQ_Order")) = "R" then
				Uploaded_Order_R = true
			else
				Uploaded_Order_R = false
			end if
			
			'동일한 리마크에 동일한 품번을 찾았다면,
			if lcase(RS1("Parts_P_P_No")) = lcase(trim(arrXLS(PNO_LOC, CNT1))) and lcase(trim(tempBQ_Remark)) = lcase(trim(tempREMARK_LOC)) and XLS_Order_R = Uploaded_Order_R then
%>
				<%if XLS_Order_R then%>
				<input type="hidden" name="QTY_<%=RS1("BOM_Sub_BS_D_No")%>_<%=RS1("Parts_P_P_No")%>_<%=RS1("BQ_Remark")%>_R" value="<%=RS1("BQ_Qty")%>" >
				<%else%>
				<input type="hidden" name="QTY_<%=RS1("BOM_Sub_BS_D_No")%>_<%=RS1("Parts_P_P_No")%>_<%=RS1("BQ_Remark")%>_X" value="<%=RS1("BQ_Qty")%>" >
				<%end if%>
<%
				PNORemark_Match_YN = "Y"
				exit for
				
			end if
		next
		
		if PNORemark_Match_YN = "N" then
			strBS_D_No		= strBS_D_No	& RS1("BOM_Sub_BS_D_No")& "||"
			strP_P_No		= strP_P_No		& RS1("Parts_P_P_No")	& "||"
			strP_P_No2		= strP_P_No2	& RS1("Parts_P_P_No2")	& "||"
			strP_P_No2_PinYN= strP_P_No2_PinYN	& RS1("Parts_P_P_No2_PinYN") & "||"
			strBQ_P_Desc	= strBQ_P_Desc	& RS1("BQ_P_Desc")		& "||"
			strBQ_P_Spec	= strBQ_P_Spec	& RS1("BQ_P_Spec")		& "||"
			strBQ_Remark	= strBQ_Remark	& RS1("BQ_Remark")		& "||"
			strBQ_Checksum	= strBQ_Checksum& RS1("BQ_Checksum")	& "||"
			strBQ_P_Maker	= strBQ_P_Maker	& RS1("BQ_P_Maker")		& "||"
			strP_Work_Type	= strP_Work_Type& RS1("P_Work_Type")	& "||"
			strBQ_Order		= strBQ_Order	& RS1("BQ_Order")		& "||"
			strBQ_Qty		= strBQ_Qty		& RS1("BQ_Qty")			& "||"
		end if
		RS1.MoveNext
	loop
	RS1.Close
	arrBS_D_No		= split(strBS_D_No,		"||")
	arrP_P_No		= split(strP_P_No,		"||")
	arrP_P_No2		= split(strP_P_No2,		"||")
	arrP_P_No2_PinYN= split(strP_P_No2_PinYN,"||")
	arrBQ_P_Desc	= split(strBQ_P_Desc,	"||")
	arrBQ_P_Spec	= split(strBQ_P_Spec,	"||")
	arrBQ_Remark	= split(strBQ_Remark,	"||")
	arrBQ_Checksum	= split(strBQ_Checksum,	"||")
	arrBQ_P_Maker	= split(strBQ_P_Maker,	"||")
	arrP_Work_Type	= split(strP_Work_Type,	"||")
	arrBQ_Order		= split(strBQ_Order,	"||")
	arrBQ_Qty		= split(strBQ_Qty,		"||")
	

	CNT2 = 0
	strAdded = ""
	for CNT1 = 0 to ubound(arrP_P_No)-1
		
		if (arrBQ_Order(CNT1) = "R" and instr(strAdded, arrP_P_No(CNT1)&"_"&arrBQ_Remark(CNT1)&"_R") = 0) or (arrBQ_Order(CNT1) <> "R" and instr(strAdded, arrP_P_No(CNT1)&"_"&arrBQ_Remark(CNT1)&"_X") = 0) then
%>
	<input type="hidden" name="PNO_<%=Parts_CNT+CNT2+1%>" value="<%=arrP_P_No(CNT1)%>">
	<input type="hidden" name="PNO2_<%=Parts_CNT+CNT2+1%>" value="<%=arrP_P_No2(CNT1)%>">
	<input type="hidden" name="PNO2PinYN_<%=Parts_CNT+CNT2+1%>" value="<%=arrP_P_No2(CNT1)%>">
	<input type="hidden" name="DESCRIPTION_<%=Parts_CNT+CNT2+1%>" value="<%=arrBQ_P_Desc(CNT1)%>">
	<input type="hidden" name="SPEC_<%=Parts_CNT+CNT2+1%>" value="<%=arrBQ_P_Spec(CNT1)%>">
	<input type="hidden" name="REMARK_<%=Parts_CNT+CNT2+1%>" value="<%=arrBQ_Remark(CNT1)%>">
	<input type="hidden" name="CHECKSUM_<%=Parts_CNT+CNT2+1%>" value="<%=arrBQ_Checksum(CNT1)%>">
	<input type="hidden" name="MAKER_<%=Parts_CNT+CNT2+1%>" value="<%=arrBQ_P_Maker(CNT1)%>">
	<input type="hidden" name="WORKTYPE_<%=Parts_CNT+CNT2+1%>" value="<%=arrP_Work_Type(CNT1)%>">
	<input type="hidden" name="NO_<%=Parts_CNT+CNT2+1%>" value="<%=arrBQ_Order(CNT1)%>">
	
<%
			CNT2 = CNT2 + 1
		end if
%>	
	<%if ucase(arrBQ_Order(CNT1)) = "R" then%>
	<input type="hidden" name="QTY_<%=arrBS_D_No(CNT1)%>_<%=arrP_P_No(CNT1)%>_<%=arrBQ_Remark(CNT1)%>_R" value="<%=arrBQ_Qty(CNT1)%>">
	<%else%>
	<input type="hidden" name="QTY_<%=arrBS_D_No(CNT1)%>_<%=arrP_P_No(CNT1)%>_<%=arrBQ_Remark(CNT1)%>_X" value="<%=arrBQ_Qty(CNT1)%>">
	<%end if%>
<%	
		if arrBQ_Order(CNT1) = "R" then
			strAdded = strAdded & arrP_P_No(CNT1) & "_" & arrBQ_Remark(CNT1) & "_R_"
		else
			strAdded = strAdded & arrP_P_No(CNT1) & "_" & arrBQ_Remark(CNT1) & "_X_"
		end if
	next
	
	Parts_CNT = Parts_CNT + CNT2
end if

objXLS.close
set objXLS = nothing
set UpLoad = nothing
set RS1 = nothing


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
