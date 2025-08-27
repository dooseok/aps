<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
dim CNT1
dim RS1
dim SQL
set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select B_Code from tbBOM order by B_Code desc"
'SQL = "select top 1 B_Code from tbBOM order by B_Code"
RS1.Open SQL,sys_DBCon
'CNT1 = 1 
do until RS1.Eof
	'response.write "call fB_Code("&RS1("B_Code")&"_"&CNT1&")<br>"
	call fB_Code(RS1("B_Code"))
	'CNT1 = CNT1 + 1
	RS1.MoveNext
loop
RS1.Close
set RS1 = nothing
%>


<% '도번을 돌면서,
sub fB_Code(B_Code)
	dim RS1
	dim SQL
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select top 1 BS_D_No from tbBOM_Sub where BOM_B_Code = "&B_Code '대표작업을 하나 선택 
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		'response.write "call fBS_Code("&RS1("BS_D_No")&")<br>"
		call fBS_Code(RS1("BS_D_No"),B_Code) '
		RS1.MoveNext
	loop
	RS1.Close
	set RS1 = nothing
end sub
%>
	
	
<%
sub fBS_Code(BS_D_No,B_Code) '한 도번의 대표작업을 분석
	dim strRemark
	dim oldRemark
	dim oldDesc
	dim cntPNOinSameRemark
	dim bR
	dim strOrder
	dim strPNO
	dim arrPNO
	dim arrPNO2
	
	dim RS1
	dim SQL
	bR = ""
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select * from tbBOM_QTY where BOM_B_Code = "&B_Code&" and BOM_Sub_BS_D_No ='"&BS_D_No&"'"
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		
		strRemark = RS1("BQ_Remark") '현재 리마크
		strOrder = RS1("BQ_Order") '현재 작업번호
		if strOrder = "R" then 'R이면
			bR = "R" 'R켜고
		else
			strOrder = "" 'R아니면 공백으로
		end if
		
		if oldRemark = strRemark then '리마크 동일함
			cntPNOinSameRemark = cntPNOinSameRemark + 1
			strPNO = strPNO & RS1("Parts_P_P_No") &"/_/"& strOrder  &"/|/"
		else '리마크 새로움
			if cntPNOinSameRemark = 2 then
				
				arrPNO = split(strPNO,"/|/")
				
				if instr(arrPNO(0),"/_/R") > 0 then
					for CNT1 = 0 to ubound(arrPNO)-1
						response.write B_Code&"?"
						arrPNO2 = split(arrPNO(CNT1),"/_/")
						if isnumeric(left(BS_D_No,3)) then
							response.write left(BS_D_No,10)&"?"
						else
							response.write left(BS_D_No,9)&"?"
						end if
						response.write arrPNO2(0)&"?"
						response.write arrPNO2(1)&"?"
						response.write oldDesc&"?"&oldremark&"?"&cntPNOinSameRemark
						response.write "<br>"
					next
				end if
			end if
			
			strPNO = RS1("Parts_P_P_No")&"/_/"& strOrder &"/|/"
			
			cntPNOinSameRemark = 1
			
			bR = ""
		end if
		
		oldremark = strRemark '이전 리마크를 메모
		oldDesc = RS1("BQ_P_Desc")
		RS1.MoveNext
	loop
	RS1.Close
	set RS1 = nothing
	
	if cntPNOinSameRemark = 2 then
		arrPNO = split(strPNO,"/|/")
		
		if instr(arrPNO(0),"/_/R") > 0 then
			for CNT1 = 0 to ubound(arrPNO)-1
				response.write B_Code&"?"
				arrPNO2 = split(arrPNO(CNT1),"/_/")
				if isnumeric(left(BS_D_No,3)) then
					response.write left(BS_D_No,10)&"?"
				else
					response.write left(BS_D_No,9)&"?"
				end if
				response.write arrPNO2(0)&"?"
				response.write arrPNO2(1)&"?"
				response.write oldDesc&"?"&oldremark&"?"&cntPNOinSameRemark
				response.write "<br>"
			next
		end if
	end if
end sub
%>
	
<!-- #include virtual = "/header/db_tail.asp" -->
