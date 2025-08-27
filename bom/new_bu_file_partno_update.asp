<%
sub BU_File_PartNo_Update(BU_Code, BU_File_PartNo)
	dim CNT1
	dim RS1
	dim SQL
	dim strError
	
	dim objXLS
	dim objXLSRS
	dim strFilePath
	dim strXLSConnection
	
	dim strSheetName
	dim nSheetCount
	
	dim nRow
	dim nCol
	
	dim arrXLS
	
	dim strDB_PartNo
	dim strDB_Desc
	dim strFlag
	
	dim arrDB_PartNo
	dim arrDB_Desc
	dim arrFlag
	
	dim strXLS_PartNo
	dim strXLS_Desc
	
	dim bNewPartNo
	
	dim strBU_Apply_Date
	dim strBU_MSE_LG
	dim strBU_Link_YN
	
	'파일 삭제
	if BU_File_PartNo = "" then
		SQL = "delete from tbBOM_Update_PartNo where BOM_Update_BU_Code = '"&BU_Code&"'"
		sys_DBCon.execute(SQL)
		
	elseif instr(lcase(BU_File_PartNo),".xls") = 0 then
		strError = "첨부파일1 ( 품 번 )에는 엑셀 파일만 업로드해주세요."
%>
		
<%
	else
		set objXLS = Server.CreateObject("ADODB.Connection") 
		set RS1 = Server.CreateObject("ADODB.RecordSet")
		
		if instr(BU_File_PartNo, "\") = 0 then
			strFilePath = DefaultPath_BOM_Update & BU_File_PartNo
		else
			strFilePath = BU_File_PartNo
		end if
		if instr(BU_File_PartNo,".xlsx") > 0 then
			strXLSConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFilePath & "; Extended Properties=""Excel 12.0;HDR=No;IMEX=1"""
		elseif instr(BU_File_PartNo,".xls") > 0 then
			strXLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""
		end if
		objXLS.Open strXLSConnection    
	
		set objXLSRS = objXLS.OpenSchema(20)
		
		nSheetCount = 0
		do until objXLSRS.Eof
			if instr(objXLSRS.Fields("table_name").Value,"_xlnm") = 0 then
				strSheetName = objXLSRS.Fields("table_name").Value	
				nSheetCount = nSheetCount + 1
			end if
			objXLSRS.MoveNext
		loop
		objXLSRS.Close
		
		if nSheetCount <> 1 then
			strSheetName = ""
			strError = "업로드된 엑셀파일에 하나 이상의 시트가 있습니다."
		end if
		
		SQL  = "select * from ["&strSheetName&"]"
		RS1.Open SQL,objXLS 
		arrXLS = RS1.getRows()
		RS1.close
	
		objXLS.close
		
		strBU_Apply_Date = ""
		strBU_MSE_LG = ""
		strBU_Link_YN = ""
		SQL = "select * from tbBOM_Update_New where BU_Code = '"&BU_Code&"'"
		RS1.Open SQL,sys_DBCon
		if not(RS1.Eof or RS1.Bof) then
			strBU_Apply_Date = RS1("BU_Apply_Date")
			strBU_MSE_LG = RS1("BU_MSE_LG")
			strBU_Link_YN = RS1("BU_Link_YN")
		end if
		RS1.Close
		
		'기존 DB내용 배열화 
		SQL = "select * from tbBOM_Update_PartNo where BOM_Update_BU_Code = '"&BU_Code&"'"
		RS1.Open SQL,sys_DBCon
		do until RS1.Eof
			strDB_PartNo = strDB_PartNo &ucase(trim(RS1("BUP_PartNo")))& "|_|"
			strDB_Desc = strDB_Desc &trim(RS1("BUP_Desc"))& "|_|"
			strFlag = strFlag & "D|_|" '기본으로 삭제(D)로 해두자.
			
			RS1.MoveNext
		loop
		RS1.Close
		arrDB_PartNo = split(strDB_PartNo,"|_|")
		arrDB_Desc = split(strDB_Desc,"|_|")
		arrFlag = split(strFlag,"|_|")
		
		'엑셀 loop
		for nRow=2 to ubound(arrXLS,2)
			strXLS_PartNo = ucase(arrXLS(1,nRow))
			strXLS_Desc = trim(arrXLS(13,nRow))
			
			bNewPartNo = true
			for CNT1 = 0 to ubound(arrDB_PartNo)-1
				if arrDB_PartNo(CNT1) = strXLS_PartNo then 'DB에 있던 파트넘버와 일치하는 XLS파트넘버가 있다면 내용 업데이트.
					arrFlag(CNT1) = "U"
					
					SQL = "update tbBOM_Update_PartNo set "
					SQL = SQL & "	BUP_desc = '"&strXLS_Desc&"', "
					SQL = SQL & "	BOM_Update_BU_Apply_Date = '"&strBU_Apply_Date&"', "
					SQL = SQL & "	BOM_Update_BU_MSE_LG = '"&strBU_MSE_LG&"', "
					SQL = SQL & "	BOM_Update_BU_Link_YN = '"&strBU_Link_YN&"' "
					SQL = SQL & "where "
					SQL = SQL & "	BOM_Update_BU_Code = '"&BU_Code&"' and "
					SQL = SQL & "	BUP_PartNo = '"&arrDB_PartNo(CNT1)&"' "
					sys_DBCon.execute(SQL)
					
					bNewPartNo = false
					exit for
				end if		
			next
			
			'DB에 있던 파트넘버와 일치하는게 없다면, 추가등록
			if bNewPartNo then
				SQL = "insert into tbBOM_Update_PartNo "
				SQL = SQL & "(BOM_Update_BU_Code, BUP_PartNo, BUP_Desc, BOM_Update_BU_Apply_Date, BOM_Update_BU_MSE_LG, BOM_Update_BU_Link_YN) values ("
				SQL = SQL & "'"&BU_Code&"',"
				SQL = SQL & "'"&strXLS_PartNo&"',"
				SQL = SQL & "'"&strXLS_Desc&"',"
				SQL = SQL & "'"&strBU_Apply_Date&"',"
				SQL = SQL & "'"&strBU_MSE_LG&"',"
				SQL = SQL & "'"&strBU_Link_YN&"'"
				SQL = SQL & ") "
				sys_DBCon.execute(SQL)
			end if
		next
		
		'DB에 있던 것들 중, XLS에 없었던 것들은 삭제
		for CNT1 = 0 to ubound(arrDB_PartNo)-1
			if arrFlag(CNT1) = "D"  then 'DB에 있던 파트넘버와 일치하는 XLS파트넘버가 있다면 내용 업데이트.
				SQL = "delete tbBOM_Update_PartNo where "
				SQL = SQL & "	BOM_Update_BU_Code = '"&BU_Code&"' and "
				SQL = SQL & "	BUP_PartNo = '"&arrDB_PartNo(CNT1)&"' "
				sys_DBCon.execute(SQL)
			end if		
		next
		
		set objXLSRS = nothing
		
		set objXLS = nothing
		set RS1 = nothing
	end if
	
	if strError <> "" then
%>
<script>alert("<%=strError%>");</script>
<%
	end if
end sub
%>
