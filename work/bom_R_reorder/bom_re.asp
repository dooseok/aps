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


<% '������ ���鼭,
sub fB_Code(B_Code)
	dim RS1
	dim SQL
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select top 1 BS_D_No from tbBOM_Sub where BOM_B_Code = "&B_Code '��ǥ�۾��� �ϳ� ���� 
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
sub fBS_Code(BS_D_No,B_Code) '�� ������ ��ǥ�۾��� �м�
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
		
		strRemark = RS1("BQ_Remark") '���� ����ũ
		strOrder = RS1("BQ_Order") '���� �۾���ȣ
		if strOrder = "R" then 'R�̸�
			bR = "R" 'R�Ѱ�
		else
			strOrder = "" 'R�ƴϸ� ��������
		end if
		
		if oldRemark = strRemark then '����ũ ������
			cntPNOinSameRemark = cntPNOinSameRemark + 1
			strPNO = strPNO & RS1("Parts_P_P_No") &"/_/"& strOrder  &"/|/"
		else '����ũ ���ο�
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
		
		oldremark = strRemark '���� ����ũ�� �޸�
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
