<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<% 
'변수선언
dim CNT1
dim CNT2

dim DNO
dim arrDNOSUB
dim arrDNOCONFIRM
dim arrMODEL

dim SQL
dim RS1
dim RS2

dim B_Code
dim BS_Code

dim Parts_CNT
dim strQTY
dim strNO
dim strPNO
dim strDESCRIPTION
dim strWORKTYPE
dim strSPEC
dim strREMARK
dim strCHECKSUM
dim strMAKER
dim strSTYPE
dim strPNO2
dim strPNO2PinYN

dim arrQTY
dim arrNO
dim arrPNO
dim arrDESCRIPTION
dim arrWORKTYPE
dim arrSPEC
dim arrREMARK
dim arrCHECKSUM
dim arrMAKER
dim arrSTYPE
dim arrPNO2
dim arrPNO2PinYN

dim P_Code

dim P_No_of_BOM_Qty
dim P_No_of_Material
dim P_No_of_COSP

dim arrQTY_BY_DNO

dim strBS_Info
dim arrBS_Info
dim arrBS_Info_Sub

dim strConflict

dim strBQinsert
dim arrBQinsert
dim arr2BQinsert

'부품 수량
Parts_CNT = Request("Parts_CNT")

'객체 선언
set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")
B_Code = Request("B_Code")
DNO	= Request("DNO")
dim DNOSUB
dim DNOCONFIRM
dim MODEL

dim BMD_Sort
dim SortedValue
dim SortedListIDX
SortedListIDX = now()

DNOSUB		= Request("DNOSUB")
DNOCONFIRM	= Request("DNOCONFIRM")
MODEL		= Request("MODEL")

arrDNOSUB		= split(Request("DNOSUB"),", ")
arrDNOCONFIRM	= split(Request("DNOCONFIRM"),", ")
arrMODEL		= split(Request("MODEL"),", ")

'------------------------------------------------------------------정렬-
dim strBMD_Desc
dim strBMD_Sort
dim arrBMD_Desc
dim arrBMD_Sort

'표준품명과 정렬 정보를 가져와서 리스트화
SQL = "select BMD_Desc, BMD_Sort from tblBOM_Mask_Desc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof 
	strBMD_Desc = strBMD_Desc & RS1("BMD_Desc") &"|/|"
	strBMD_Sort = strBMD_Sort & RS1("BMD_Sort") &"|/|"
	RS1.MoveNext
loop
RS1.Close

'위에서 만든 품명/정렬 순서에 추가
SQL = ""
SQL = SQL & "select "
SQL = SQL & "	BMDD_Desc_BOM, "
SQL = SQL & "	BMD_Sort = isnull(("
SQL = SQL & "		select top 1 BMD_Sort "
SQL = SQL & "		from tblBOM_Mask_Desc "
SQL = SQL & "		where BMD_Desc = BOM_Mask_Desc_BMD_Desc "
SQL = SQL & "		),999) "
SQL = SQL & "from tblBOM_Mask_Desc_Detail "
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strBMD_Desc = strBMD_Desc & RS1("BMDD_Desc_BOM") &"|/|"
	strBMD_Sort = strBMD_Sort & RS1("BMD_Sort") &"|/|"
	RS1.MoveNext
loop
RS1.Close

arrBMD_Desc = split(strBMD_Desc,"|/|")
arrBMD_Sort = split(strBMD_Sort,"|/|")
'-----------------------------------------------------------------------


'Remark동일 항목들이 같은 순번을 가지도록 처리 시작
dim descSort
dim nSort
dim arrSortByRemark_Desc()
dim arrSortByRemark_Sort()
redim arrSortByRemark_Desc(Parts_CNT)
redim arrSortByRemark_Sort(Parts_CNT)
dim arrSortPos
arrSortPos = 0
'Remark동일 항목들이 같은 순번을 가지도록 처리 끝



'정렬을 위해 테이블에 넣었다가 뺌-------------------------------------------------
for CNT1 = 1 to Parts_CNT
	BMD_Sort = 999 '결정 순번
	descSort = 999 'Desc정규화에서 찾을 순번
	nSort = 999 'Sort목록에서 찾을 순번
	
	'Sort목록을 찾아서, Remark로 등록된 순번이 있다면 가져옴
	for CNT2 = 0 to ubound(arrSortByRemark_Desc)
		if arrSortByRemark_Desc(CNT2) <> "" then
			if arrSortByRemark_Desc(CNT2) = lcase(Request("REMARK_"&CNT1)) then '이번 파츠의 리마크와 일치하는 리마크가 있나?
				nSort = arrSortByRemark_Sort(CNT2) '그럼 소트는 그 리마크에서 가져오면 되겠음
				CNT2 = ubound(arrSortByRemark_Desc) '루프 종료
			end if
		end if
	next
	
	'Sort목록을 찾아서, Desc로 등록된 순번이 있다면 가져옴
	for CNT2=0 to ubound(arrBMD_Desc) - 1 
		if lcase(Request("DESCRIPTION_"&CNT1)) = lcase(arrBMD_Desc(CNT2)) then
			descSort = arrBMD_Sort(CNT2)
			CNT2 = ubound(arrBMD_Desc) - 1
		end if
	next

	if nSort = 999 and descSort < 999 then 'desc에서만 찾은 경우
		BMD_Sort = descSort
		
		'Sort목록에 추가한다.
		arrSortByRemark_Desc(arrSortPos) = lcase(Request("REMARK_"&CNT1))
		arrSortByRemark_Sort(arrSortPos) = BMD_Sort 
		arrSortPos = arrSortPos + 1
		
	elseif nSort < 999 and descSort = 999 then 'sort목록에서만 찾은 경우
		BMD_Sort = nSort
	elseif nSort < 999 and descSort < 999 then '둘다 찾은 경우
		'descSort값이 더 작다면 Sort목록을 업데이트한다.
		if nSort > descSort then
			'Sort목록을 찾아서, Remark로 등록된 순번이 있다면 갱신
			for CNT2 = 0 to ubound(arrSortByRemark_Desc)
				if arrSortByRemark_Desc(CNT2) <> "" then
					if arrSortByRemark_Desc(CNT2) = lcase(Request("REMARK_"&CNT1)) then
						arrSortByRemark_Sort(CNT2) = descSort
					end if
				end if
			next
			BMD_Sort = descSort
		else 'nSort값이 더 작다면 그대로 사용한다.
			BMD_Sort = nSort
		end if
	else '둘다 없는 경우
		BMD_Sort = 999
		
		'Sort목록에 추가한다.
		arrSortByRemark_Desc(arrSortPos) = lcase(Request("REMARK_"&CNT1))
		arrSortByRemark_Sort(arrSortPos) = BMD_Sort 
		arrSortPos = arrSortPos + 1
	end if
	
	'for CNT2=0 to ubound(arrBMD_Desc) - 1
	'	if Request("DESCRIPTION_"&CNT1) = arrBMD_Desc(CNT2) then
	'		BMD_Sort = arrBMD_Sort(CNT2)
	'		CNT2 = ubound(arrBMD_Desc) - 1
	'	end if
	'next
	'if BMD_Sort = "" then
	'	BMD_Sort = 999
	'end if
	'Remark동일 항목들이 같은 순번을 가지도록 처리 끝
	
	SortedValue = ""
	SortedValue = SortedValue & Request("PNO_"&CNT1) & "//"
	SortedValue = SortedValue & replace(Request("DESCRIPTION_"&CNT1),"'","''") & "//"
	SortedValue = SortedValue & Request("WORKTYPE_"&CNT1) & "//"
	SortedValue = SortedValue & replace(Request("SPEC_"&CNT1),"'","''") & "//"
	SortedValue = SortedValue & Request("MAKER_"&CNT1) & "//"
	SortedValue = SortedValue & Request("CHECKSUM_"&CNT1) & "//"
	SortedValue = SortedValue & Request("PNO2_"&CNT1) &"//"
	SortedValue = SortedValue & Request("PNO2PinYN_"&CNT1)&"//"
	SortedValue = SortedValue & Request("STYPE_"&CNT1)
	
	SQL = "insert into tbSortedList (SL_IDX,SL_Key,SL_Key2,SL_Key3,SL_Key4,SL_Value) values ('"&SortedListIDX&"',"&BMD_Sort&",'"&Request("REMARK_"&CNT1)&"','"&Request("QTY_"&CNT1)&"','"&Request("NO_"&CNT1)&"','"&SortedValue&"')"
	sys_DBCon.execute(SQL)
next
'품번순으로 / R이 아닌 것들은 remark명 순으로 / R은 code를 따라간다. / 수량은 제거.

dim arrRS
SQL = "select * from tbSortedList where SL_IDX = '"&SortedListIDX&"' order by SL_Key, SL_Key2, SL_Code"
RS1.Open SQL,sys_DBCon

'R행의 상위 수량을 R로 복사하기 위한 코드 시작 
'dim nQTY
'dim nOldQTY
'nOldQTY = 0
'dim nOldRemark
'R행의 상위 수량을 R로 복사하기 위한 코드 끝

do until RS1.Eof
	arrRS = split(RS1("SL_Value"),"//")
		
		'R행의 상위 수량을 R로 복사하기 위한 코드 시작 
		'nQTY = RS1("SL_Key3")
		'if RS1("SL_Key4") = "R" and nOldRemark=RS1("SL_Key2") then
		'	nQTY = nOldQTY
		'end if

		'nOldRemark = RS1("SL_Key2")
		'nOldQTY = nQTY
		'R행의 상위 수량을 R로 복사하기 위한 코드 끝
		
		
		'R행의 상위 수량을 R로 복사하기 위한 코드 시작 
		'strQTY			= strQTY			& nQTY	& "|/|"
		'R행의 상위 수량을 R로 복사하기 위한 코드 끝
		strQTY			= strQTY			& RS1("SL_Key3")	& "|/|"
		
		
		
		strNO			= strNO				& RS1("SL_Key4")	& "|/|"
		strPNO			= strPNO			& arrRS(0)			& "|/|"
		strDESCRIPTION	= strDESCRIPTION	& arrRS(1)			& "|/|"
		strWORKTYPE		= strWORKTYPE		& arrRS(2)			& "|/|"
		strSPEC			= strSPEC			& arrRS(3)			& "|/|"
		strMAKER		= strMAKER			& arrRS(4)			& "|/|"
		strREMARK		= strREMARK			& RS1("SL_Key2")	& "|/|"
		strCHECKSUM		= strCHECKSUM		& arrRS(5)			& "|/|"
		strPNO2			= strPNO2			& arrRS(6)			& "|/|"
		strPNO2PinYN	= strPNO2PinYN		& arrRS(7)			& "|/|"
		strSType		= strSType			& arrRS(8)			& "|/|"
		
		
	RS1.MoveNext
loop
RS1.Close

SQL = "delete tbSortedList where SL_IDX = '"&SortedListIDX&"'"
sys_DBCon.execute(SQL)

arrQTY			= split(strQTY,"|/|")
arrNO			= split(strNO,"|/|")
arrPNO			= split(strPNO,"|/|")
arrDESCRIPTION	= split(strDESCRIPTION,"|/|")
arrWORKTYPE		= split(strWORKTYPE,"|/|")
arrSPEC			= split(strSPEC,"|/|")
arrMAKER		= split(strMAKER,"|/|")
arrREMARK		= split(strREMARK,"|/|")
arrCHECKSUM		= split(strCHECKSUM,"|/|")
arrPNO2			= split(strPNO2,"|/|")
arrPNO2PinYN	= split(strPNO2PinYN,"|/|")
arrSType		= split(strSType,"|/|")
for CNT1=0 to ubound(arrQTY)
	arrQTY(CNT1) = trim(replace(arrQTY(CNT1),"'","''"))
	arrNO(CNT1) = trim(replace(arrNO(CNT1),"'","''"))
	arrPNO(CNT1) = trim(replace(arrPNO(CNT1),"'","''"))
	arrDESCRIPTION(CNT1) = trim(replace(arrDESCRIPTION(CNT1),"'","''"))
	arrWORKTYPE(CNT1) = trim(replace(arrWORKTYPE(CNT1),"'","''"))
	arrSPEC(CNT1) = trim(replace(arrSPEC(CNT1),"'","''"))
	arrMAKER(CNT1) = trim(replace(arrMAKER(CNT1),"'","''"))
	arrREMARK(CNT1) = trim(replace(arrREMARK(CNT1),"'","''"))
	arrCHECKSUM(CNT1) = trim(replace(arrCHECKSUM(CNT1),"'","''"))
	arrPNO2(CNT1) = trim(replace(arrPNO2(CNT1),"'","''"))
	arrPNO2PinYN(CNT1) = trim(replace(arrPNO2PinYN(CNT1),"'","''"))
	arrSType(CNT1) = trim(replace(arrSType(CNT1),"'","''"))
next
'-----------------------------------------------------------------------------

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

for CNT1 = 0 to ubound(arrDNOSUB)
	if trim(arrDNOSUB(CNT1)) <> "" Then
		SQL = "select BS_Code from tbBOM_Sub where BS_D_No = '"&arrDNOSUB(CNT1)&"' and BOM_B_Code = "&B_Code
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			SQL = "insert into tbBOM_Sub (BS_D_No,BOM_B_Code,BS_IMD_QTY,BS_SMD_QTY,BS_MAN_QTY,BS_ASM_QTY, BS_IMD_Axial_Point, BS_IMD_Radial_Point, BS_Confirm_YN) values ('"&arrDNOSUB(CNT1)&"','"&B_Code&"',0,0,0,0,0,0,'N')"
			sys_DBCon.execute(SQL)
		
			SQL = "select max(BS_Code) as BS_Code from tbBOM_Sub where BOM_B_Code = "&B_Code
			RS2.Open SQL,sys_DBCon
			BS_Code = RS2("BS_Code")
			RS2.Close	
		else
			BS_Code = RS1("BS_Code")
		end if
		RS1.Close

		for CNT2 = 0 to ubound(arrPNO)-1
			if arrPNO(CNT2) <> "" then
				strBQinsert = strBQinsert & B_Code & "/!/" '0
				strBQinsert = strBQinsert & BS_Code & "/!/" '1
				strBQinsert = strBQinsert & arrDNOSUB(CNT1) & "/!/" '2
				strBQinsert = strBQinsert & arrPNO(CNT2) & "/!/" '3
				
				if arrQTY(CNT2) = "" then			
					strBQinsert = strBQinsert & "0/!/" '4
				else
					arrQTY_BY_DNO	= split(arrQTY(CNT2),",")
					
					If trim(arrQTY_BY_DNO(CNT1)) = "" OR trim(arrQTY_BY_DNO(CNT1)) = "-" Then
						arrQTY_BY_DNO(CNT1) = 0
					End if
					strBQinsert = strBQinsert & trim(arrQTY_BY_DNO(CNT1)) & "/!/" '4
				end if
				
				strBQinsert = strBQinsert & arrNO(CNT2) & "/!/" '5
				strBQinsert = strBQinsert & arrREMARK(CNT2) & "/!/" '6
				strBQinsert = strBQinsert & arrCHECKSUM(CNT2) & "/!/" '7
				strBQinsert = strBQinsert & arrDESCRIPTION(CNT2) & "/!/" '8
				strBQinsert = strBQinsert & arrSPEC(CNT2) & "/!/" '9
				strBQinsert = strBQinsert & arrMAKER(CNT2) & "/!/" '10
				strBQinsert = strBQinsert & arrPNO2(CNT2) & "/!/" '11
				strBQinsert = strBQinsert & arrPNO2PinYN(CNT2) & "/!/" '12
				strBQinsert = strBQinsert & arrSType(CNT2) & "|!|" '13
				
				call updateBOM_Info(arrPNO(CNT2),arrPNO2(CNT2),arrPNO2PinYN(CNT2),arrSType(CNT2))
				
			end if
		next
	end if
next

dim strTable
SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&B_Code
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

SQL = "delete from "&strTable&" where BOM_B_Code = "&B_Code
sys_DBCon.execute(SQL)

arrBQinsert = split(strBQinsert,"|!|")


for CNT1 = 0 to ubound(arrBQinsert) - 1
	arr2BQinsert = split(arrBQinsert(CNT1),"/!/")
	
	

	SQL = "insert into "&strTable&" (BOM_B_Code, " '0
	SQL = SQL & "BOM_Sub_BS_Code, " '1
	SQL = SQL & "BOM_Sub_BS_D_No, " '2
	SQL = SQL & "Parts_P_P_No, " '3
	SQL = SQL & "BQ_Qty, " '4
	SQL = SQL & "BQ_Order, " '5
	SQL = SQL & "BQ_Remark, " '6
	SQL = SQL & "BQ_CHECKSUM, " '7
	SQL = SQL & "BQ_P_Desc, " '8
	SQL = SQL & "BQ_P_Spec, " '9
	SQL = SQL & "BQ_P_Maker, " '10
	SQL = SQL & "Parts_P_P_No2, " '11
	SQL = SQL & "Parts_P_P_No2_PinYN, " '12
	SQL = SQL & "BOM_B_D_No "
	SQL = SQL & ") values ("
	SQL = SQL & arr2BQinsert(0)&",'"&arr2BQinsert(1)&"','"&arr2BQinsert(2)&"','"&arr2BQinsert(3)&"',"&arr2BQinsert(4)&",'"&arr2BQinsert(5)&"','"&arr2BQinsert(6)&"','"&arr2BQinsert(7)&"','"&arr2BQinsert(8)&"','"&arr2BQinsert(9)&"','"&arr2BQinsert(10)&"','"&arr2BQinsert(11)&"','"&arr2BQinsert(12)&"','"&DNO&"')"
	sys_DBCon.execute(SQL)
	
	
	'if arr2BQinsert(3) = "EAE64483901" then
	'	SQL = "insert tblError_Temp (val1,val2,val3)values ('"&arr2BQinsert(2)&"','"&arr2BQinsert(3)&"','"&arr2BQinsert(4)&"')"
	'	sys_DBCon.execute(SQL)
	'end if
	
	if strTable = "tbBOM_Qty" then
		if isnumeric(arr2BQinsert(4)) then
			if arr2BQinsert(4) > 0 then
				SQL = "insert into tblBOM_Mask (BOM_Parts_BP_PNO, BM_Filter, RegDate, EditDate, M_ID,BM_SType_BOM,BM_Desc_BOM,BM_Spec_BOM,BM_Maker_BOM) values "
				SQL = SQL & "('"&arr2BQinsert(3)&"','_', getdate(), getdate(), '"&gM_ID&"','"&arr2BQinsert(13)&"','"&arr2BQinsert(8)&"','"&arr2BQinsert(9)&"','"&arr2BQinsert(10)&"')"
				
				'if arr2BQinsert(3) = "EAE64483901" then
				'	SQL = "insert tblError_Temp (val1)values ('"&replace(SQL,"'","''")&"')"
				'	sys_DBCon.execute(SQL)
				'end if
				
				on error resume next
				sys_DBCon.execute(SQL)
				on error goto 0
			end if
		end if
	end if
next

for CNT1 = 0 to ubound(arrDNOCONFIRM)
	SQL = "update tbBOM_Sub set BS_Confirm_YN = 'Y' where BS_D_No = '"&arrDNOCONFIRM(CNT1)&"' and BOM_B_Code = "&B_Code
	'response.write SQL
	sys_DBCon.execute(SQL)
next

for CNT2 = 0 to ubound(arrBS_Info)-1
	arrBS_Info_Sub = split(arrBS_Info(CNT2),"|")
	if isnull(arrBS_Info_Sub(1)) or not(isnumeric(arrBS_Info_Sub(1))) then
		arrBS_Info_Sub(1) = 0
	end if
	if isnull(arrBS_Info_Sub(2)) or not(isnumeric(arrBS_Info_Sub(2))) then
		arrBS_Info_Sub(2) = 0
	end if
	if isnull(arrBS_Info_Sub(3)) or not(isnumeric(arrBS_Info_Sub(3))) then
		arrBS_Info_Sub(3) = 0
	end if
	if isnull(arrBS_Info_Sub(4)) or not(isnumeric(arrBS_Info_Sub(4))) then
		arrBS_Info_Sub(4) = 0
	end if
	if isnull(arrBS_Info_Sub(5)) or not(isnumeric(arrBS_Info_Sub(5))) then
		arrBS_Info_Sub(5) = 0
	end if
	if isnull(arrBS_Info_Sub(6)) or not(isnumeric(arrBS_Info_Sub(6))) then
		arrBS_Info_Sub(6) = 0
	end if
		
	SQL = "update tbBOM_Sub set "
	SQL = SQL & "BS_IMD_Qty=" & arrBS_Info_Sub(1) &", "
	SQL = SQL & "BS_SMD_Qty=" & arrBS_Info_Sub(2) &", "
	SQL = SQL & "BS_MAN_Qty=" & arrBS_Info_Sub(3) &", "
	SQL = SQL & "BS_ASM_Qty=" & arrBS_Info_Sub(4) &", "
	SQL = SQL & "BS_IMD_Axial_Point=" & arrBS_Info_Sub(5) &", "
	SQL = SQL & "BS_IMD_Radial_Point=" & arrBS_Info_Sub(6) &" where BS_D_No = '" & arrBS_Info_Sub(0) & "' and BOM_B_Code = "&B_Code
	sys_DBCon.execute(SQL)
next

SQL = "update tbBOM set B_Opt_YN = 'Y' where B_Code = "&B_Code
sys_DBCon.execute(SQL)

if gM_ID = "shindk" then
	dim OPT_B_Code
	SQL = "select top 1 B_Code from tbBOM where B_OPT_YN <> 'Y' or B_OPT_YN is null order by B_Version_Current_YN desc"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		OPT_B_Code = ""
	else
		OPT_B_Code = RS1(0)
	end if
	RS1.Close
end if

call BOM_Level_Reset(B_Code)

set RS1 = nothing
set RS2 = nothing
%>

<%
sub BOM_Level_Reset(B_Code)
	dim RS1
	
	dim SQL
	dim B_Version_Code
	dim B_Version_Current_YN
			
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	
	'B코드로 버젼정보 조회
	SQL = "select B_Version_Code, B_Version_Current_YN from tbBOM where B_Code='"&B_Code&"'"
	RS1.Open SQL,sys_DBCon
	if not(RS1.Eof or RS1.Bof) then
		B_Version_Code = RS1("B_Version_Code")
		B_Version_Current_YN = RS1("B_Version_Current_YN")
	end if
	RS1.Close
	
	'현재적용중이라면
	if B_Version_Current_YN = "Y" then
		'bom을 루프돌면서 bs_code, bs_d_no 루핑
		SQL = "select BS_Code, BS_D_No from tbBOM_Sub where BOM_B_Code="&B_Code
		RS1.Open SQL,sys_DBCon
		do until RS1.Eof
	
			'기존 행 삭제
			SQL = "delete tblBOM_Level_Master where B_PARTNO_ASSY = '"&RS1("BS_D_No")&"'"
			sys_DBCon.execute(SQL)
			
			'재등록
			SQL = "insert into tblBOM_Level_Master (B_PARTNO_ASSY,B_LEVEL_READY_YN,B_LEVEL_Date) values ('"&RS1("BS_D_No")&"','N',getdate())"
			sys_DBCon.execute(SQL)
			
			RS1.MoveNext
		loop
		RS1.Close
	end if
	
	set RS1 = nothing
end sub
%>


<%
'if OPT_B_Code <> "" then
%>
<!--
<form name="frmRedirect" action="db_load_action.asp" method="post">
<input type="hidden" name="B_Code" value="<%=OPT_B_Code%>">
<input type="hidden" name="Diff_YN" value="N">
</form>
-->
<%
'else
%>
<form name="frmRedirect" action="db_load_action.asp" method="post">
<input type="hidden" name="B_Code" value="<%=B_Code%>">
<input type="hidden" name="Diff_YN" value="<%=Request("Diff_YN")%>">
</form>
<%
'end if
%>

<%
function updateBOM_Info(Parts_P_P_No,Parts_P_P_No2,Parts_P_P_No2_PinYN,BQI_SType)
	if (Parts_P_P_No2 <> "" and Parts_P_P_No2_PinYN <> "Y") or BQI_SType <> "" then
		SQL = "select top 1 * from tbBOM_QTY_Info where Parts_P_P_No='"&Parts_P_P_No&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			if Parts_P_P_No2 <> "" and BQI_SType <> "" then
				SQL = "insert into tbBOM_QTY_Info (Parts_P_P_No,Parts_P_P_No2,BQI_SType) values ('"&Parts_P_P_No&"','"&Parts_P_P_No2&"','"&BQI_SType&"')"
			elseif Parts_P_P_No2 <> "" and BQI_SType = "" then
				SQL = "insert into tbBOM_QTY_Info (Parts_P_P_No,Parts_P_P_No2,BQI_SType) values ('"&Parts_P_P_No&"','"&Parts_P_P_No2&"','')"
			elseif Parts_P_P_No2 = "" and BQI_SType <> "" then
				SQL = "insert into tbBOM_QTY_Info (Parts_P_P_No,Parts_P_P_No2,BQI_SType) values ('"&Parts_P_P_No&"','','"&BQI_SType&"')"
			end if
			sys_DBCon.execute(SQL)
		else
			SQL = "update tbBOM_QTY_Info set "
			if Parts_P_P_No2 <> "" and BQI_SType <> "" then
				SQL = SQL & " Parts_P_P_No2 = '"&Parts_P_P_No2&"', BQI_SType = '"&BQI_SType&"'"
			elseif Parts_P_P_No2 <> "" and BQI_SType = "" then
				SQL = SQL & " Parts_P_P_No2 = '"&Parts_P_P_No2&"'"
			elseif Parts_P_P_No2 = "" and BQI_SType <> "" then
				SQL = SQL & " BQI_SType = '"&BQI_SType&"'"
			end if
			SQL = SQL & " where Parts_P_P_No = '"&Parts_P_P_No&"'" 
			sys_DBCon.execute(SQL)
		end if
		RS1.Close
	end if
end function 	
%>

<script language="javascript">
alert('저장이 완료되었습니다.');
frmRedirect.submit();
</script>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->