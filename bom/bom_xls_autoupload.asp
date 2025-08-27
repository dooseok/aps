<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->



<%
'업로드 룰 (파일명은 모델명으로 예)EBR391877
dim RS1
dim SQL

dim B_Code
dim objFSO
dim objFolder
dim objFiles

dim File
dim FileName
dim PNO
'db_load_action.asp
set objFSO		= server.CreateObject("Scripting.FileSystemObject")
set objFolder	= objFSO.GetFolder("d:\work\")
set objFiles	= objFolder.Files 

set RS1 = server.CreateOBject("ADODB.RecordSet")
'해당폴더의 이미지를 배열화 
for each File In objFiles
	FileName = lcase(File.Name)
	PNO		 = left(lcase(FileName),instr(FileName,".")-1)
	
	SQL = "select B_Code from tbBOM where B_D_No='"&PNO&"'"
	RS1.Open SQL,sys_DBCon,1
	B_Code = RS1(0)
	RS1.Close
	response.write "<a href='db_load_action.asp?b_code="&B_Code&"' target='_blank'>"&FileName & "</a><br>"
	'response.write B_Code & "." &FileName & "<br>"
	'call XLS2Data(PNO, "d:\work\"&FileName)
next

set RS1 = nothing
set objFiles	= nothing
set objFolder	= nothing
set objFSO		= nothing
%>

<%
sub XLS2Data(PNO, strFileName)
	dim objXLS
	dim XLSConnection
	
	dim RS1
	dim Sheet_Name

	dim SQL
	dim arrXLS
	
	dim CNT1, CNT2
	
	dim Model_LOC
	dim QTY_LOC
	
	dim PNO_LOC
	dim DESCRIPTION_LOC
	dim CHECKSUM_LOC
	dim SPEC_LOC
	dim MAKER_LOC
	dim REMARK_LOC
	dim LOC_LOC
	dim TYPE_LOC
	dim QTY
	
	dim strModel
	dim strPartsPNO
	dim strDesc
	dim strCheckSum
	dim strSpec
	dim strMaker
	dim strRemark
	dim strLOC
	dim strType
	dim strQty
	
	dim arrModel
	dim arrPartsPNO
	dim arrDesc
	dim arrCheckSum
	dim arrSpec
	dim arrMaker
	dim arrRemark
	dim arrLOC
	dim arrType
	dim arrQty
	
	set objXLS = Server.CreateObject("ADODB.Connection") 
	XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFileName & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""
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
	set RS1 = nothing

	objXLS.Close
	set objXLS		= nothing
	
	for CNT1 = 0 to ubound(arrXLS, 2)
		for CNT2 = 0 to ubound(arrXLS, 1)
			
			'첫 행이고, 무언가 값이 있다면! ex) A or 1
			if CNT1 = 0 and trim(arrXLS(CNT2, CNT1)) <> "" then
				if len(trim(arrXLS(CNT2, CNT1))) = 1 then
					strModel = strModel & arrXLS(CNT2, CNT1) & "|%|"	'strModel에 Model값 누적
				else
					strModel = strModel & PNO&arrXLS(CNT2, CNT1) & "|%|"	'strModel에 Model값 누적
				end if
			elseif CNT1 = 6 then
				if instr(lcase(arrXLS(CNT2, CNT1)),"p/no") > 0 then
					PNO_LOC			= CNT2
				elseif instr(lcase(arrXLS(CNT2, CNT1)),"description") > 0 then
					DESCRIPTION_LOC	= CNT2			
				elseif instr(lcase(arrXLS(CNT2, CNT1)),"checksum") > 0 then
					CHECKSUM_LOC	= CNT2
				elseif instr(lcase(arrXLS(CNT2, CNT1)),"spec") > 0 then
					SPEC_LOC		= CNT2
				elseif instr(lcase(arrXLS(CNT2, CNT1)),"maker") > 0 then
					MAKER_LOC		= CNT2
				elseif instr(lcase(arrXLS(CNT2, CNT1)),"remark") > 0 then
					REMARK_LOC		= CNT2
				elseif instr(lcase(arrXLS(CNT2, CNT1)),"loc") > 0 then
					LOC_LOC			= CNT2
				elseif instr(lcase(arrXLS(CNT2, CNT1)),"type") > 0 then
					TYPE_LOC		= CNT2
				end if
			elseif CNT1 >= 7 then
				if CNT2 = PNO_LOC then 
					strPartsPNO = strPartsPNO & arrXLS(CNT2, CNT1)
				elseif CNT2 = DESCRIPTION_LOC then
					strDesc = strDesc & arrXLS(CNT2, CNT1)
				elseif CNT2 = CHECKSUM_LOC then
					strCheckSum = strCheckSum & arrXLS(CNT2, CNT1)			
				elseif CNT2 = SPEC_LOC then
					strSpec = strSpec & arrXLS(CNT2, CNT1)		
				elseif CNT2 = MAKER_LOC then
					strMaker = strMaker & arrXLS(CNT2, CNT1)		
				elseif CNT2 = REMARK_LOC then
					strRemark = strRemark & arrXLS(CNT2, CNT1)
				elseif CNT2 = LOC_LOC then
					strLOC = strLOC & arrXLS(CNT2, CNT1)	
				elseif CNT2 = TYPE_LOC then
					strType = strType & arrXLS(CNT2, CNT1)
				elseif arrXLS(CNT2, 6) = "QTY" then
					QTY = arrXLS(CNT2, CNT1)
					if QTY <> "" then
						QTY = replace(trim(QTY),"-","0")
					else
						QTY = "0"
					end if
					strQty = strQty & QTY & "|%|"				
				end if
			end if
			
			
		next
		strPartsPNO = strPartsPNO & "|$|"
		strDesc = strDesc & "|$|"
		strCheckSum = strCheckSum & "|$|"
		strSpec = strSpec & "|$|"
		strMaker = strMaker & "|$|"
		strRemark = strRemark & "|$|"
		strLOC = strLOC & "|$|"
		strType = strType & "|$|"
		strQty = strQty & "|$|"
	next
	
	arrModel	= split(strModel,"|%|")
	arrPartsPNO	= split(strPartsPNO,"|$|")
	arrDesc		= split(strDesc,"|$|")
	arrCheckSum	= split(strCheckSum,"|$|")
	arrSpec		= split(strSpec,"|$|")
	arrMaker	= split(strMaker,"|$|")
	arrRemark	= split(strRemark,"|$|")
	arrLOC		= split(strLOC,"|$|")
	arrType		= split(strType,"|$|")
	arrQty		= split(strQty,"|$|")
	
	
	SQL = "select * from tbBOM_Sub where BOM_B_Code='"&B_Code&"'"
	RS1.Open SQL,sys_DBCon
	Do Until RS1.Eof
		SQL = "delete tbBOM_Qty where BOM_Sub_BS_Code='"&RS1("BS_Code")&"'"
		sys_DBCon.execute(SQL)
		RS1.MoveNext
	loop
	RS1.Close
	
	SQL = "select * from tbBOM_Sub where BOM_B_Code='"&B_Code&"'"
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		strBS_Info = strBS_Info & RS1("BS_D_No") &"|"& RS1("BS_IMD_Qty") &"|"& RS1("BS_SMD_Qty") &"|"& RS1("BS_MAN_Qty") &"|"& RS1("BS_ASM_Qty") &"|"& RS1("BS_IMD_Axial_Point") &"|"& RS1("BS_IMD_Radial_Point") &"//"
		RS1.MoveNext
	loop
	RS1.Close
	arrBS_Info = split(strBS_Info,"//")
	
	SQL = "delete tbBOM_Sub where BOM_B_Code='"&B_Code&"'"
	sys_DBCon.execute(SQL)
	
	
	
	
end sub
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->