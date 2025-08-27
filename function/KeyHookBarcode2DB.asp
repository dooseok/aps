<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
dim RS1

dim SQL
dim strLine
dim strProcess
dim strKeys

dim strBarcode
dim strBarcode1
dim strBarcode2
dim strBarcode3
dim strBarcode4
dim strBarcode5
dim strBarcode6

dim strNow
dim strDate
dim strTime
dim bExist

strBarcode = ""
strBarcode1 = ""
strBarcode2 = ""
strBarcode3 = ""
strBarcode4 = ""
strBarcode5 = ""
strBarcode6 = ""

strLine = ucase(Request("strLine"))
strProcess = ucase(Request("strProcess"))
strKeys = ucase(Request("strKeys"))

strNow = now()
strDate = FormatDateTime(strNow,2)
strTime = replace(FormatDateTime(strNow,4),":","")


strKeys = right(strKeys,23)
if instr("-EBR-687-ABQ-499-","-"&left(strKeys,3)&"-") > 0 then '2D바코드라면,
	strBarcode = ucase(strKeys)
else
	strKeys = right(strKeys,15)
	if left(strKeys,1) = "1" then
		strBarcode1=(Convert36to10(mid(strKeys,2,1))*36*36+Convert36to10(mid(strKeys,3,1))*36+Convert36to10(mid(strKeys,4,1)))
		strBarcode1=StringExtender(strBarcode1,4)
		
		strBarcode2=(Convert36to10(mid(strKeys,5,1))*36*36+Convert36to10(mid(strKeys,6,1))*36+Convert36to10(mid(strKeys,7,1)))
		strBarcode2=StringExtender(strBarcode2,4)
		
		strBarcode3=Convert36to10(mid(strKeys,10,1))
		strBarcode3=StringExtender(strBarcode3,2)
		
		strBarcode4=Convert36to10(mid(strKeys,11,1))
		strBarcode4=StringExtender(strBarcode4,2)
		
		strBarcode5=Convert36to10(mid(strKeys,12,1))
		strBarcode5=StringExtender(strBarcode5,2)
		
		strBarcode6=(Convert36to10(mid(strKeys,13,1))*36*36+Convert36to10(mid(strKeys,14,1))*36+Convert36to10(mid(strKeys,15,1)))
		strBarcode6=StringExtender(strBarcode6,4)
		
		strBarcode="EBR"&strBarcode1&strBarcode2&"13"&strBarcode3&strBarcode4&strBarcode5&strBarcode6
	elseif left(strKeys,1)="G" then
		strBarcode1=(Convert36to10(mid(strKeys,2,1))*36*36+Convert36to10(mid(strKeys,3,1))*36+Convert36to10(mid(strKeys,4,1)))
		strBarcode1=StringExtender(strBarcode1,4)
		
		strBarcode2=(Convert36to10(mid(strKeys,5,1))*36*36+Convert36to10(mid(strKeys,6,1))*36+Convert36to10(mid(strKeys,7,1)))
		strBarcode2=StringExtender(strBarcode2,4)
		
		strBarcode3=Convert36to10(mid(strKeys,10,1))
		strBarcode3=StringExtender(strBarcode3,2)
		
		strBarcode4=Convert36to10(mid(strKeys,11,1))
		strBarcode4=StringExtender(strBarcode4,2)
		
		strBarcode5=Convert36to10(mid(strKeys,12,1))
		strBarcode5=StringExtender(strBarcode5,2)
		
		strBarcode6=(Convert36to10(mid(strKeys,13,1))*36*36+Convert36to10(mid(strKeys,14,1))*36+Convert36to10(mid(strKeys,15,1)))
		strBarcode6=StringExtender(strBarcode6,4)

		strBarcode="ABQ"&strBarcode1&strBarcode2&"13"&strBarcode3&strBarcode4&strBarcode5&strBarcode6
	elseif left(strKeys,2)="2A" then
		strBarcode1=(Convert36to10(mid(strKeys,5,1))*36+Convert36to10(mid(strKeys,6,1)))
		strBarcode1=StringExtender(strBarcode1,3)
		
		strBarcode3=Convert36to10(mid(strKeys,10,1))
		strBarcode3=StringExtender(strBarcode3,2)
		
		strBarcode4=Convert36to10(mid(strKeys,11,1))
		strBarcode4=StringExtender(strBarcode4,2)
		
		strBarcode5=Convert36to10(mid(strKeys,12,1))
		strBarcode5=StringExtender(strBarcode5,2)
		
		strBarcode6=(Convert36to10(mid(strKeys,13,1))*36*36+Convert36to10(mid(strKeys,14,1))*36+Convert36to10(mid(strKeys,15,1)))
		strBarcode6=StringExtender(strBarcode6,4)
		
		strBarcode="6871A"&mid(strKeys,3,2)&strBarcode1&mid(strKeys,7,1)&"13"&strBarcode3&strBarcode4&strBarcode5&strBarcode6
	elseif left(strKeys,2)="CA" then	
		strBarcode1=(Convert36to10(mid(strKeys,5,1))*36+Convert36to10(mid(strKeys,6,1)))
		strBarcode1=StringExtender(strBarcode1,3)
		
		strBarcode3=Convert36to10(mid(strKeys,10,1))
		strBarcode3=StringExtender(strBarcode3,2)
		
		strBarcode4=Convert36to10(mid(strKeys,11,1))
		strBarcode4=StringExtender(strBarcode4,2)
		
		strBarcode5=Convert36to10(mid(strKeys,12,1))
		strBarcode5=StringExtender(strBarcode5,2)
		
		strBarcode6=(Convert36to10(mid(strKeys,13,1))*36*36+Convert36to10(mid(strKeys,14,1))*36+Convert36to10(mid(strKeys,15,1)))
		strBarcode6=StringExtender(strBarcode6,4)
		
		strBarcode="4995A"&mid(strKeys,3,2)&strBarcode1&mid(strKeys,7,1)&"13"&strBarcode3&strBarcode4&strBarcode5&strBarcode6
	end if
end if




'!!!!!!!!!!!!!!!!!!!!!!!!!!!에러 트래킹용 코드 시작 
'SQL = "insert into tbPWS_Raw_Data_Test values ('"&strKeys&"','"&strBarcode&"','"&strProcess&"','"&strLine&"','"&strDate&" "&strTime&"')"
'sys_DBCon.execute(SQL)
'!!!!!!!!!!!!!!!!!!!!!!!!!!!에러 트래킹용 코드 끝
if strBarcode <> "" then
	
	if strProcess = "INPUT" then
		SQL = "insert tbPWS_Raw_Data (PRD_Input_Date, PRD_Input_Time, PRD_Line, PRD_Barcode, PRD_PartNo) values ('" & strDate & "','" & strTime & "','" & strLine & "','" & strBarcode & "','" & left(strBarcode, 11) & "')"
		sys_DBCon.execute(SQL)
	elseif strProcess = "BOX" then
		set RS1 = server.CreateObject("ADODB.RecordSet")
		SQL = "select top 1 PRD_Code from tbPWS_Raw_Data where PRD_Barcode = '" & strBarcode & "'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			bExist=false
		else
			bExist=true
		end if
		RS1.Close
		set RS1 = nothing
		
		if bExist = false then
			SQL = "insert tbPWS_Raw_Data (PRD_Input_Date, PRD_Input_Time, PRD_BOX_Date, PRD_BOX_Time, PRD_Line, PRD_Barcode, PRD_PartNo, PRD_ByHook_YN) values "
			SQL = SQL & "('" & strDate & "','" & strTime & "','" & strDate & "','" & strTime & "','" & strLine & "pp','" & strBarcode & "','" & left(strBarcode, 11) & "','Y')"
			'response.write SQL
			sys_DBCon.execute(SQL)     
				      
		elseif bExist = true then
			SQL = "update tbPWS_Raw_Data set "
        	SQL = SQL & "PRD_BOX_Date = '" & strDate & "', "
        	SQL = SQL & "PRD_BOX_Time = '" & strTime & "', "
        	SQL = SQL & "PRD_Line = '" & strLine & "', "
        	SQL = SQL & "PRD_ByHook_YN = 'Y' "
        	SQL = SQL & "where PRD_Barcode = '" & strBarcode & "'"
        	'response.write SQL
        	sys_DBCon.execute(SQL)
		end if
		
		
	end if
end if

function StringExtender(strSrc, nSize)    
	dim CNT1
    dim strResult
    strResult = strSrc
    for CNT1 = 1 to nSize - len(strSrc)
    	strResult = "0" & strResult
	next
	
    StringExtender = strResult
end function

function Convert36to10(strValue)
	dim nResult
	
	strValue = cstr(strValue)
	
    if strValue = "0" then
        nResult = 0
    elseif strValue = "1" then
        nResult = 1
    elseif strValue = "2" then
        nResult = 2
    elseif strValue = "3" then
        nResult = 3
    elseif strValue = "4" then
        nResult = 4
    elseif strValue = "5" then
        nResult = 5
    elseif strValue = "6" then
        nResult = 6
    elseif strValue = "7" then
        nResult = 7
    elseif strValue = "8" then
        nResult = 8
    elseif strValue = "9" then
        nResult = 9
    elseif strValue = "A" then
        nResult = 10
    elseif strValue = "B" then
        nResult = 11
    elseif strValue = "C" then
        nResult = 12
    elseif strValue = "D" then
        nResult = 13
    elseif strValue = "E" then
        nResult = 14
    elseif strValue = "F" then
        nResult = 15
    elseif strValue = "G" then
        nResult = 16
    elseif strValue = "H" then
        nResult = 17
    elseif strValue = "I" then
        nResult = 18
    elseif strValue = "J" then
        nResult = 19
    elseif strValue = "K" then
        nResult = 20
    elseif strValue = "L" then
        nResult = 21
    elseif strValue = "M" then
        nResult = 22
    elseif strValue = "N" then
        nResult = 23
    elseif strValue = "O" then
        nResult = 24
    elseif strValue = "P" then
        nResult = 25
    elseif strValue = "Q" then
        nResult = 26
    elseif strValue = "R" then
        nResult = 27
    elseif strValue = "S" then
        nResult = 28
    elseif strValue = "T" then
        nResult = 29
    elseif strValue = "U" then
        nResult = 30
    elseif strValue = "V" then
        nResult = 31
    elseif strValue = "W" then
        nResult = 32
    elseif strValue = "X" then
        nResult = 33
    elseif strValue = "Y" then
        nResult = 34
    elseif strValue = "Z" then
        nResult = 35
	end if
	
	Convert36to10 = nResult

end function
	



%>
<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->