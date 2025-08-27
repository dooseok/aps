<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<% 
dim RS1
dim SQL
dim CNT1

dim PNO1
dim PNO2
dim arrPNO1
dim arrPNO2

dim B_D_No1
dim B_D_No2
dim B_Version_Code1
dim B_Version_Code2
dim B_Code1
dim B_Code2
dim B_Code1_YN
dim B_Code2_YN

PNO1 = Request("PNO1")
PNO2 = Request("PNO2")

arrPNO1 = split(PNO1,"&nbsp;/&nbsp;")
arrPNO2 = split(PNO2,"&nbsp;/&nbsp;")

B_D_No1			= trim(arrPNO1(0))
B_Version_Code1	= trim(arrPNO1(1))
B_Code1			= trim(arrPNO1(2))

B_D_No2			= trim(arrPNO2(0))
B_Version_Code2	= trim(arrPNO2(1))
B_Code2			= trim(arrPNO2(2))

Response.Buffer = true
Response.ContentType = "application/vnd.ms-excel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment;filename=BOM_New_Diff.xls"
response.write "DIFF"&vbtab
response.write "Assy PNO"&vbtab
response.write "Part PNO"&vbtab
response.write "W/O1"&vbtab
response.write "W/O2"&vbtab
response.write "QTY1"&vbtab
response.write "QTY2"&vbtab
response.write "Remark1"&vbtab
response.write "Remark2"&vbtab
response.write "Desc1"&vbtab
response.write "Desc2"&vbtab
response.write "Maker1"&vbtab
response.write "Maker2"&vbtab
response.write "Spec1"&vbtab
response.write "Spec2"&vbtab
response.write vbcrlf

set RS1 = server.CreateObject("ADODB.RecordSet")

dim strTable

SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&B_Code1	
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
else
	if RS1("B_Version_Current_YN") = "Y" then
		strTable = "tbBOM_Qty"
	else
		strTable = "tbBOM_Qty_Archive"
	end if
end if
RS1.Close
	
SQL = "select distinct BOM_Sub_BS_D_No from "&strTable&" where BOM_B_Code = '"&B_Code1&"' order by BOM_Sub_BS_D_No asc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	call BOM_Diff(RS1("BOM_Sub_BS_D_No"))
	RS1.MoveNext
loop
RS1.Close

SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&B_Code2
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
else
	if RS1("B_Version_Current_YN") = "Y" then
		strTable = "tbBOM_Qty"
	else
		strTable = "tbBOM_Qty_Archive"
	end if
end if
RS1.Close
	
SQL = "select distinct BOM_Sub_BS_D_No from "&strTable&" where BOM_B_Code = '"&B_Code2&"' order by BOM_Sub_BS_D_No asc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	call BOM_Diff(RS1("BOM_Sub_BS_D_No"))
	RS1.MoveNext
loop
RS1.Close

set RS1 = nothing

sub BOM_Diff(BOM_Sub_BS_D_No)
	dim SQL
	dim RS1
	dim RS2
	dim CNT1
	dim CNT2
	dim CNT3
	
	dim arrBOM1
	dim arrBOM2

	dim strFind_YN
	dim strMatched1
	dim strMatched2
	dim strDiff
	dim strDiffChr
	
	dim strTable
	
	SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&B_Code1	
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
	else
		if RS1("B_Version_Current_YN") = "Y" then
			strTable = "tbBOM_Qty"
		else
			strTable = "tbBOM_Qty_Archive"
		end if
	end if
	RS1.Close

	SQL = "select BQ_Code,BOM_Sub_BS_D_No,Parts_P_P_No,BQ_Order,BQ_Qty,BQ_Remark,BQ_P_Desc,BQ_P_Maker,BQ_P_Spec from "&strTable&" where BOM_B_Code = '"&B_Code1&"' and BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"' order by BQ_Code asc"
	set RS1 = sys_DBCon.Execute(SQL)
	if RS1.Eof or RS1.Bof then
		B_Code1_YN = "N"
	else
		B_Code1_YN = "Y"
		arrBOM1 = RS1.GetRows()
	end if
	set RS1 = nothing
	
	

	SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&B_Code2
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
	else
		if RS1("B_Version_Current_YN") = "Y" then
			strTable = "tbBOM_Qty"
		else
			strTable = "tbBOM_Qty_Archive"
		end if
	end if
	RS1.Close
	
	'			  0       1               2            3        4         5         6         7       8    
	SQL = "select BQ_Code,BOM_Sub_BS_D_No,Parts_P_P_No,BQ_Order,BQ_Qty,BQ_Remark,BQ_P_Desc,BQ_P_Maker,BQ_P_Spec from "&strTable&" where BOM_B_Code = '"&B_Code2&"' and BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"' order by BQ_Code asc"
	set RS1 = sys_DBCon.Execute(SQL)
	if RS1.Eof or RS1.Bof then
		B_Code2_YN = "N"
	else
		B_Code2_YN = "Y"
		arrBOM2 = RS1.GetRows()
	end if
	set RS1 = nothing
	
	strMatched1 = "-"
	strMatched2 = "-"
	strDiff = ""
	
	if B_Code1_YN = "Y" and B_Code2_YN = "Y" then
		for CNT1 = lbound(arrBOM1, 2) To ubound(arrBOM1, 2) '1번 BOM을 루프를 돈다
		
		
			strFind_YN = "N"
			strDiff = ""
			for CNT2 = lbound(arrBOM2, 2) To ubound(arrBOM2, 2) '1번 BOM의 파트넘버와 2번 BOM의 파트넘버를 비교한다.
				if arrBOM1(2,CNT1) = arrBOM2(2,CNT2) and instr(strMatched2,"-"&arrBOM2(0,CNT2)&"-") = 0 then '같은 파트넘버를 찾는다, 매칭된 건 제외)
					strFind_YN = "Y"
					strMatched1 = strMatched1 &arrBOM1(0,CNT1)&"-" '찾은 것을 매칭리스트1에 기록
					strMatched2 = strMatched2 &arrBOM2(0,CNT2)&"-" '찾은 것을 매칭리스트2에 기록
					for CNT3 = 4 to 8
						if arrBOM1(CNT3,CNT1) <> arrBOM2(CNT3,CNT2) then
							strDiff = strDiff &cstr(CNT3)&"-"
						end if
					next
					strDiffChr = replace(strDiff,"4","Q")
					strDiffChr = replace(strDiffChr,"5","R")
					strDiffChr = replace(strDiffChr,"6","D")
					strDiffChr = replace(strDiffChr,"7","M")
					strDiffChr = replace(strDiffChr,"8","S")
					strDiffChr = replace(strDiffChr,"-","")
					if strDiffChr = "" then
						strDiffChr = "same"
					else
						strDiffChr = "Diff("&strDiffChr&")"
					end if
					response.write strDiffChr &vbtab
					response.write arrBOM1(1,CNT1) &vbtab
					response.write arrBOM1(2,CNT1) &vbtab
					
					response.write arrBOM1(3,CNT1) &vbtab
					response.write arrBOM2(3,CNT2) &vbtab
					
					if cstr(arrBOM1(4,CNT1)) = "0" then
						arrBOM1(4,CNT1) = ""
					end if
					
					for CNT3 = 4 to 8
							response.write arrBOM1(CNT3,CNT1) &vbtab
							response.write arrBOM2(CNT3,CNT2) &vbtab
					next
					response.write vbcrlf
					exit for
				end if	
			next
		next
		
		for CNT1 = lbound(arrBOM1, 2) To ubound(arrBOM1, 2)
			if instr(strMatched1,"-"&arrBOM1(0,CNT1)&"-") = 0 then
				response.write "only1"&vbtab
				response.write arrBOM1(1,CNT1) &vbtab
				response.write arrBOM1(2,CNT1) &vbtab
				response.write arrBOM1(3,CNT1) &vbtab
				response.write vbtab
				
				if cstr(arrBOM1(4,CNT1)) = "0" then
					arrBOM1(4,CNT1) = ""
				end if
			
				for CNT3 = 4 to 8
					response.write arrBOM1(CNT3,CNT1) &vbtab
					response.write vbtab
				next
				response.write vbcrlf
			end if
		next
		
		for CNT1 = lbound(arrBOM2, 2) To ubound(arrBOM2, 2)
			if instr(strMatched2,"-"&arrBOM2(0,CNT1)&"-") = 0 then
				response.write "only2"&vbtab
				response.write arrBOM2(1,CNT1) &vbtab
				response.write arrBOM2(2,CNT1) &vbtab
				response.write vbtab
				response.write arrBOM2(3,CNT1) &vbtab
				
				if cstr(arrBOM2(4,CNT1)) = "0" then
					arrBOM2(4,CNT1) = ""
				end if
			
				for CNT3 = 4 to 8
					response.write vbtab
					response.write arrBOM2(CNT3,CNT1) &vbtab
				next
				response.write vbcrlf
			end if
		next
	elseif B_Code1_YN = "Y" and B_Code2_YN = "N" then
		for CNT1 = lbound(arrBOM1, 2) To ubound(arrBOM1, 2) '1번 BOM을 루프를 돈다
			response.write "only1"&vbtab
			response.write arrBOM1(1,CNT1) &vbtab
			response.write arrBOM1(2,CNT1) &vbtab
			
			response.write arrBOM1(3,CNT1) &vbtab
			response.write vbtab
			
			if cstr(arrBOM1(4,CNT1)) = "0" then
				arrBOM1(4,CNT1) = ""
			end if
			
			for CNT3 = 4 to 8
				response.write arrBOM1(CNT3,CNT1) &vbtab
				response.write vbtab
			next
			response.write vbcrlf
		next
	elseif B_Code1_YN = "N" and B_Code2_YN = "Y" then
		for CNT1 = lbound(arrBOM2, 2) To ubound(arrBOM2, 2) '2번 BOM을 루프를 돈다 matched 안된 것들을 출력
			response.write "only2"&vbtab
			response.write arrBOM2(1,CNT1) &vbtab
			response.write arrBOM2(2,CNT1) &vbtab
			response.write vbtab
			response.write arrBOM2(3,CNT1) &vbtab
			
			if cstr(arrBOM2(4,CNT1)) = "0" then
				arrBOM2(4,CNT1) = ""
			end if
		
			for CNT3 = 4 to 8
				response.write vbtab
				response.write arrBOM2(CNT3,CNT1) &vbtab
			next
			response.write vbcrlf
		next	
	end if
	
	set RS1 = nothing
	set RS2 = nothing
end sub
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->