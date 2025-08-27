<%
dim DNO
dim Model_CNT
dim Parts_CNT
dim arrDNOSUB
dim arrDNOCONFIRM
dim arrBQ_Order
dim arrParts_P_P_No
dim arrBQ_P_Desc
dim arrBQ_P_Spec
dim arrBQ_P_Maker
dim arrP_Work_Type
dim arrBQI_SType
dim arrBQ_Remark
dim arrBQ_CheckSum
dim arrParts_P_P_No2
dim arrParts_P_P_No2_PinYN

call B_Model_Reg_Form()
sub B_Model_Reg_Form()

	dim B_Code
	dim BS_Code
	dim BS_D_No
	
	dim strDNOSUB
	dim strDNOCONFIRM
	
	dim strBQ_Order
	dim strParts_P_P_No
	dim strBQ_P_Desc
	dim strBQ_P_Spec
	dim strBQ_P_Maker
	dim strP_Work_Type
	dim strBQI_SType
	dim strBQ_Remark
	dim strBQ_CheckSum
	dim strParts_P_P_No2
	dim strParts_P_P_No2_PinYN
	
	dim BQA_Key
	
	strBQ_Order				= "._."
	strParts_P_P_No			= "._."
	strBQ_P_Desc			= "._."
	strBQ_P_Spec			= "._."
	strBQ_P_Maker			= "._."
	strP_Work_Type			= "._."
	strBQI_SType			= "._."
	strBQ_Remark			= "._."
	strBQ_CheckSum			= "._."
	strParts_P_P_No2		= "._."
	strParts_P_P_No2_PinYN	= "._."
	
	dim BU_Code '시방번호
	B_Code 	= Request("B_Code") '도번
	BS_Code = Request("BS_Code") '옵션PNO
	BU_Code	= Request("BU_Code") '시방번호
	
	dim CNT1
	dim CNT2
	dim B_D_No
	
	dim Parts_P_P_NO2
	dim BQI_SType
	
	dim RS1
	dim RS2
	dim SQL
	
	set RS1 = Server.CreateObject("ADODB.RecordSet") 
	set RS2 = Server.CreateObject("ADODB.RecordSet") 
	
	'옵션PNO만 있을 때 도번 가져오기
	if B_Code = "" and BS_Code <> "" then
		B_Code = getB_Code(BS_Code)
	end if
	
	'BU_Code(시방코드)로 B_Code 가져오기
	if BU_Code <> "" then
		SQL = "select top 1 BOM_B_D_No, BU_Sibang_No from tbBOM_Update_NEW where BU_Code = '"&BU_Code&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			response.write "해당되는 도면을 찾을 수 없습니다."
			response.end
		else
			SQL = "select top 1 B_Code from tbBOM where B_D_No = '"&RS1("BOM_B_D_No")&"' and B_Version_Code = '"&RS1("BU_Sibang_No")&"'"
			RS2.Open SQL,sys_DBCon
			if RS2.Eof or RS2.Bof then
				response.write "해당되는 도면을 찾을 수 없습니다."
				response.end
			else
				B_Code = RS2("B_Code")
			end if
			RS2.Close
		end if
		RS1.Close
	end if
	
	'DNO 가져오기
	SQL  = "select top 1 B_D_No from tbBOM where B_Code="&request("B_Code")
	RS1.Open SQL,sys_DBCon
	DNO = RS1("B_D_No")
	RS1.close
	
	SQL = "select * from tbBOM_Sub where BOM_B_Code='"&B_Code&"' order by BS_D_No desc"
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		BS_D_No = RS1("BS_D_No")
		strDNOSUB = strDNOSUB & RS1("BS_D_No") &"._."
		strDNOCONFIRM = strDNOCONFIRM & RS1("BS_Confirm_YN") &"._."
		RS1.MoveNext
	loop
	RS1.Close
	arrDNOSUB		= split(strDNOSUB,"._.")
	arrDNOCONFIRM	= split(strDNOCONFIRM,"._.")
	if strDNOSUB = "" then
		Model_CNT = 10
	else
		Model_CNT = ubound(arrDNOSUB)
	end if
	if Model_CNT = 0 then
		Model_CNT = 1
	end if
	
	SQL = "select "
	SQL = SQL & "BQ_Order, "
	SQL = SQL & "Parts_P_P_No, "
	SQL = SQL & "Parts_P_P_No2, "
	SQL = SQL & "Parts_P_P_No2_PinYN, "
	SQL = SQL & "BQ_P_Desc, "
	SQL = SQL & "BQ_P_Spec, "
	SQL = SQL & "BQ_P_Maker, "
	SQL = SQL & "P_Work_Type = (select top 1 P_Work_Type from tbParts where P_P_No = t1.Parts_P_P_No), "
	SQL = SQL & "BQ_Remark, "
	SQL = SQL & "BQ_CheckSum "
	SQL = SQL & "from tbBOM_Qty t1 where BOM_B_Code="&B_Code&" and BOM_Sub_BS_D_No='"&BS_D_No&"' order by BQ_Code"
	RS1.Open SQL,sys_DBCon
	Do Until RS1.Eof
		BQI_SType = ""
		Parts_P_P_No2 = ""
		SQL = "select top 1 BQI_SType, Parts_P_P_No2 from tbBOM_QTY_Info where Parts_P_P_No = '"&RS1("Parts_P_P_No")&"'"
		RS2.Open SQL,sys_DBCon
		if not(RS2.Eof or RS2.Bof) then
			BQI_SType = RS2("BQI_SType")
			Parts_P_P_No2 = RS2("Parts_P_P_No2")
		end if
		RS2.Close
	
		'해당 파트넘버가 핀이 꽂혀있다면, PNO2는 도면DB에서 가져온다.
		if RS1("Parts_P_P_No2_PinYN") = "Y" then
			Parts_P_P_No2 = RS1("Parts_P_P_No2")
		end if
		
		strBQ_Order			= strBQ_Order		& RS1("BQ_Order")		& "._."
		strParts_P_P_No		= strParts_P_P_No	& RS1("Parts_P_P_No")	& "._."
		strBQ_P_Desc		= strBQ_P_Desc		& RS1("BQ_P_Desc")		& "._."
		strBQ_P_Spec		= strBQ_P_Spec		& RS1("BQ_P_Spec")		& "._."
		strBQ_P_Maker		= strBQ_P_Maker		& RS1("BQ_P_Maker")		& "._."
		strP_Work_Type		= strP_Work_Type	& RS1("P_Work_Type")	& "._."
		strBQI_SType		= strBQI_SType		& BQI_SType				& "._."
		strBQ_Remark		= strBQ_Remark		& RS1("BQ_Remark")		& "._."
		strBQ_CheckSum		= strBQ_CheckSum	& RS1("BQ_CheckSum")	& "._."
		strParts_P_P_No2	= strParts_P_P_No2	& Parts_P_P_No2			& "._."
		
		if RS1("Parts_P_P_No2_PinYN") = "" or isnull(RS1("Parts_P_P_No2_PinYN")) then
			strParts_P_P_No2_PinYN = strParts_P_P_No2_PinYN & "N._." 
		else
			strParts_P_P_No2_PinYN = strParts_P_P_No2_PinYN & "Y._."
		end if
		RS1.MoveNext
	Loop
	RS1.Close
	arrBQ_Order				= split(strBQ_Order,"._.")
	arrParts_P_P_No			= split(strParts_P_P_No,"._.")
	arrBQ_P_Desc			= split(strBQ_P_Desc,"._.")
	arrBQ_P_Spec			= split(strBQ_P_Spec,"._.")
	arrBQ_P_Maker			= split(strBQ_P_Maker,"._.")
	arrP_Work_Type			= split(strP_Work_Type,"._.")
	arrBQI_SType			= split(strBQI_SType,"._.")
	arrBQ_Remark			= split(strBQ_Remark,"._.")
	arrBQ_CheckSum			= split(strBQ_CheckSum,"._.")
	arrParts_P_P_No2		= split(strParts_P_P_No2,"._.")
	arrParts_P_P_No2_PinYN	= split(strParts_P_P_No2_PinYN,"._.")
	if strParts_P_P_No = "" then
		Parts_CNT = 30
	else
		Parts_CNT = ubound(arrParts_P_P_No)-1
	end if
	if Parts_CNT = 0 then
		Parts_CNT = 1
	end if
	
	set RS2 = nothing
	set RS1 = Nothing
end sub
%>

<%
function getB_Code(BS_Code)
	dim RS1
	dim SQL
	dim B_Code
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select BOM_B_Code from tbBOM_Sub where BS_Code = '"&BS_Code&"'"
	RS1.Open SQL,sys_DBCon
	B_Code = RS1("BOM_B_Code")
	RS1.Close
	set RS1 = nothing
	getB_Code = B_Code
end function 
%>