<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
Response.Buffer = False


dim CNT1
dim RS1
dim SQL
set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select B_Code from tbBOM where B_Code between 0 and 5000 order by B_Code desc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	call fB_Code(RS1("B_Code"))
	RS1.MoveNext
loop
RS1.Close
set RS1 = nothing
response.write "done!"
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
	dim RS2
	set RS2 = Server.CreateObject("ADODB.RecordSet")
	dim b_d_no
	if isnumeric(left(BS_D_No,3)) then
		b_d_no = left(BS_D_No,10)
	else
		b_d_no = left(BS_D_No,9)
	end if
	
	dim strRemark
	dim oldRemark
	dim oldDesc
	dim cntPNOinSameRemark
	dim bR
	dim strOrder
	dim strPNO
	dim arrPNO
	dim arrPNO2
	dim BQ_Code
	dim oldBQ_Code
	dim strDesc
	
	dim etc
	
	dim RS1
	dim SQL
	bR = ""
	dim strTable
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select * from tbBOM_QTY where BOM_B_Code = "&B_Code&" and BOM_Sub_BS_D_No ='"&BS_D_No&"' order by BQ_Remark, BQ_Code asc"
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		
		strRemark = RS1("BQ_Remark") '���� ����ũ
		strOrder = RS1("BQ_Order") '���� �۾���ȣ
		BQ_Code = RS1("BQ_Code")
		strDesc = RS1("BQ_P_Desc")
		
		if strOrder = "R" then 'R�̸�
			bR = "R" 'R�Ѱ�
		else
			strOrder = "X" 'R�ƴϸ� ��������
		end if
		
		if oldRemark = strRemark then '����ũ ������
			cntPNOinSameRemark = cntPNOinSameRemark + 1
			strPNO = strPNO & RS1("Parts_P_P_No") &"/_/"& strOrder &"/_/"& BQ_Code &"/_/"& strDesc &"/_/"& strRemark &"/|/"
		else '����ũ ���ο�
			
			if cntPNOinSameRemark >= 2 then
				arrPNO = split(strPNO,"/|/")
				arrPNO2 = split(arrPNO(0),"/_/")
				
				if int(arrPNO2(2)) + cntPNOinSameRemark - 1 = oldBQ_Code then
					etc = "Y"
				else
					etc = "N"
				end if
				
				if etc = "N" then
					for CNT1 = 0 to ubound(arrPNO)-1
						arrPNO2 = split(arrPNO(CNT1),"/_/")
					
						SQL = "select top 1 strTable from tbBOM_Table where "
						SQL = SQL & "dno = '"&b_d_no&"' and "
						SQL = SQL & "pno = '"&arrPNO2(0)&"' and "
						SQL = SQL & "bq_remark = '"&arrPNO2(4)&"' "
						RS2.Open SQL,sys_DBCon
						if RS2.Eof or RS2.Bof then
							strTable = ""
						else
							strTable = RS2("strTable")
						end if
						RS2.Close
					
						if strTable <> "" then
							response.write B_Code&"|_|"
							response.write b_d_no&"|_|"
							response.write arrPNO2(0)&"|_|"
							response.write arrPNO2(1)&"|_|"
							response.write arrPNO2(2)&"|_|"
							response.write arrPNO2(3)&"|_|"
							response.write arrPNO2(4)&"|_|"
							response.write cntPNOinSameRemark&"|_|"
							response.write etc&"|_|"
							response.write replace(strTable,"/_/","|_|")
							
							response.write "///<br>"
						end if
					next
				end if
			end if
			
			strPNO = RS1("Parts_P_P_No")&"/_/"& strOrder &"/_/"& BQ_Code &"/_/"& strDesc &"/_/"& strRemark &"/|/"
			
			cntPNOinSameRemark = 1
			
			bR = ""
		end if
		
		oldBQ_Code = BQ_Code
		oldremark = strRemark '���� ����ũ�� �޸�
		oldDesc = RS1("BQ_P_Desc")
		RS1.MoveNext
	loop
	RS1.Close
	set RS1 = nothing
	
		
	if cntPNOinSameRemark >= 2 then
		arrPNO = split(strPNO,"/|/")
		arrPNO2 = split(arrPNO(0),"/_/")
	
		if int(arrPNO2(2)) + cntPNOinSameRemark - 1 = oldBQ_Code then
			etc = "Y"
		else
			etc = "N"
		end if
		
		if etc = "N" then
			for CNT1 = 0 to ubound(arrPNO)-1
				arrPNO2 = split(arrPNO(CNT1),"/_/")
			
				SQL = "select top 1 strTable from tbBOM_Table where "
				SQL = SQL & "dno = '"&b_d_no&"' and "
				SQL = SQL & "pno = '"&arrPNO2(0)&"' and "
				SQL = SQL & "bq_remark = '"&arrPNO2(4)&"' "
				RS2.Open SQL,sys_DBCon
				if RS2.Eof or RS2.Bof then
					strTable = ""
				else
					strTable = RS2("strTable")
				end if
				RS2.Close
			
				if strTable <> "" then
					response.write B_Code&"|_|"
					response.write b_d_no&"|_|"
					response.write arrPNO2(0)&"|_|"
					response.write arrPNO2(1)&"|_|"
					response.write arrPNO2(2)&"|_|"
					response.write arrPNO2(3)&"|_|"
					response.write arrPNO2(4)&"|_|"
					response.write cntPNOinSameRemark&"|_|"
					response.write etc&"|_|"
					response.write replace(strTable,"/_/","|_|")
					
					response.write "///<br>"
				end if
			next
		end if
	end if
	
	set RS2 = nothing
end sub
%>


	
	
<!-- #include virtual = "/header/db_tail.asp" -->
