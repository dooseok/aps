<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
dim RS1
dim SQL

dim CNT1
dim CNT2

dim strAll_Work_Order
dim arrAll_Work_Order

dim strLP_IMD_Complete_Qty
dim strLP_SMD_Complete_Qty
dim strLP_MAN_Complete_Qty
dim strLP_ASM_Complete_Qty
dim strLP_DLV_Complete_Qty

dim arrLP_IMD_Complete_Qty
dim arrLP_SMD_Complete_Qty
dim arrLP_MAN_Complete_Qty
dim arrLP_ASM_Complete_Qty
dim arrLP_DLV_Complete_Qty

dim strOLD_LP_IMD_Complete_Qty
dim strOLD_LP_SMD_Complete_Qty
dim strOLD_LP_MAN_Complete_Qty
dim strOLD_LP_ASM_Complete_Qty
dim strOLD_LP_DLV_Complete_Qty

dim arrOLD_LP_IMD_Complete_Qty
dim arrOLD_LP_SMD_Complete_Qty
dim arrOLD_LP_MAN_Complete_Qty
dim arrOLD_LP_ASM_Complete_Qty
dim arrOLD_LP_DLV_Complete_Qty

dim arrValueString
dim strWork_Order
dim strAmount
dim strBOM_Sub_BS_D_No
dim arrBOM_Sub_BS_D_No

dim arrLPE_Code

dim NEW_YN

strAll_Work_Order = request("strAll_Work_Order") &", "
arrAll_Work_Order = split(strAll_Work_Order,", ")

strLP_IMD_Complete_Qty		= Request("strLP_IMD_Complete_Qty")		&", "
strLP_SMD_Complete_Qty		= Request("strLP_SMD_Complete_Qty")		&", "
strLP_MAN_Complete_Qty		= Request("strLP_MAN_Complete_Qty")		&", "
strLP_ASM_Complete_Qty		= Request("strLP_ASM_Complete_Qty")		&", "
strLP_DLV_Complete_Qty		= Request("strLP_DLV_Complete_Qty")		&", "

arrLP_IMD_Complete_Qty		= split(strLP_IMD_Complete_Qty,", ")
arrLP_SMD_Complete_Qty		= split(strLP_SMD_Complete_Qty,", ")
arrLP_MAN_Complete_Qty		= split(strLP_MAN_Complete_Qty,", ")
arrLP_ASM_Complete_Qty		= split(strLP_ASM_Complete_Qty,", ")
arrLP_DLV_Complete_Qty		= split(strLP_DLV_Complete_Qty,", ")

strOLD_LP_IMD_Complete_Qty	= Request("strOLD_LP_IMD_Complete_Qty")	&", "
strOLD_LP_SMD_Complete_Qty	= Request("strOLD_LP_SMD_Complete_Qty")	&", "
strOLD_LP_MAN_Complete_Qty	= Request("strOLD_LP_MAN_Complete_Qty")	&", "
strOLD_LP_ASM_Complete_Qty	= Request("strOLD_LP_ASM_Complete_Qty")	&", "
strOLD_LP_DLV_Complete_Qty	= Request("strOLD_LP_DLV_Complete_Qty")	&", "

arrOLD_LP_IMD_Complete_Qty	= split(strOLD_LP_IMD_Complete_Qty,", ")
arrOLD_LP_SMD_Complete_Qty	= split(strOLD_LP_SMD_Complete_Qty,", ")
arrOLD_LP_MAN_Complete_Qty	= split(strOLD_LP_MAN_Complete_Qty,", ")
arrOLD_LP_ASM_Complete_Qty	= split(strOLD_LP_ASM_Complete_Qty,", ")
arrOLD_LP_DLV_Complete_Qty	= split(strOLD_LP_DLV_Complete_Qty,", ")

'받은 제번을 쭉 돌면서
for CNT1 = 0 to Ubound(arrAll_Work_Order) - 1

	'데이터를 제번, 수량, BOM으로 나눈다.
	arrValueString		= split(arrAll_Work_Order(CNT1),"//")
	strWork_Order		= arrValueString(0)
	strAmount			= arrValueString(1)
	strBOM_Sub_BS_D_No	= arrValueString(2)
	arrBOM_Sub_BS_D_No	= split(strBOM_Sub_BS_D_No,"<br>")
	
	for CNT2 = 0 to Ubound(arrBOM_Sub_BS_D_No)
	
		if left(arrBOM_Sub_BS_D_No(CNT2),3) = "EBR" or left(arrBOM_Sub_BS_D_No(CNT2),4) = "6871" then
			if instr(strWork_Order,"_") > 0 then
				arrLPE_Code = split(strWork_Order,"_")
				SQL = "update tbLGE_Plan_ETC set LPE_IMD_Complete_Qty = "&arrLP_IMD_Complete_Qty(CNT1)& " where LPE_Code = '"&arrLPE_Code(1)&"'"
				sys_DBCon.execute(SQL)
			else
				SQL = "update tbLGE_Plan set LP_IMD_Complete_Qty = "&arrLP_IMD_Complete_Qty(CNT1)& " where LP_Work_Order = '"&strWork_Order&"'"
				sys_DBCon.execute(SQL)
			end if
			
			'이 제번이 새로 체크된 제번임을 확인한다.
			NEW_YN = "N"
			if int(arrLP_IMD_Complete_Qty(CNT1)) > int(arrOLD_LP_IMD_Complete_Qty(CNT1)) then
				NEW_YN = "Y"
			end if
			
			'오직 새로 체크된 경우에 한해서
			if NEW_YN = "Y" then	
				SQL = 		"insert into tbProcess_Record (BOM_Sub_BS_D_No,PR_Work_Order,PR_WorkType,PR_Process,PR_Line,PR_Amount,PR_Worker_CNT,PR_Work_Date,PR_Start_Time,PR_End_Time,PR_Memo) values ("
				SQL = SQL & "'"&arrBOM_Sub_BS_D_No(CNT2)&"',"
				SQL = SQL & "'"&strWork_Order&"',"
				SQL = SQL & "'작업',"
				SQL = SQL & "'IMD',"
				SQL = SQL & "'',"
				SQL = SQL & arrLP_IMD_Complete_Qty(CNT1) - arrOLD_LP_IMD_Complete_Qty(CNT1)&","
				SQL = SQL & "0,"
				SQL = SQL & "'"&date()&"',"
				SQL = SQL & "'',"
				SQL = SQL & "'',"
				SQL = SQL & "'')"		
				sys_DBCon.execute(SQL)

				'입력된 실적에 해당하는 모델파트넘버의 해당공정 재고를 +시킴
				call Process_Qty_BOM_Sub_Plus(arrBOM_Sub_BS_D_No(CNT2),"IMD",arrLP_IMD_Complete_Qty(CNT1) - arrOLD_LP_IMD_Complete_Qty(CNT1))
									
				'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 -시킴
				call Process_Qty_BOM_Sub_Before_Minus(arrBOM_Sub_BS_D_No(CNT2),"IMD",arrLP_IMD_Complete_Qty(CNT1) - arrOLD_LP_IMD_Complete_Qty(CNT1))
		
				'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 -시킴
				call Process_Qty_Parts_Minus(arrBOM_Sub_BS_D_No(CNT2),"IMD",arrLP_IMD_Complete_Qty(CNT1) - arrOLD_LP_IMD_Complete_Qty(CNT1))

			end if
			
			
			
			
			if instr(strWork_Order,"_") > 0 then
				arrLPE_Code = split(strWork_Order,"_")
				SQL = "update tbLGE_Plan_ETC set LPE_SMD_Complete_Qty = "&arrLP_SMD_Complete_Qty(CNT1)& " where LPE_Code = '"&arrLPE_Code(1)&"'"
				sys_DBCon.execute(SQL)
			else
				SQL = "update tbLGE_Plan set LP_SMD_Complete_Qty = "&arrLP_SMD_Complete_Qty(CNT1)& " where LP_Work_Order = '"&strWork_Order&"'"
				sys_DBCon.execute(SQL)
			end if
			
			'이 제번이 새로 체크된 제번임을 확인한다.
			NEW_YN = "N"
			if int(arrLP_SMD_Complete_Qty(CNT1)) > int(arrOLD_LP_SMD_Complete_Qty(CNT1)) then
				NEW_YN = "Y"
			end if
		
			'오직 새로 체크된 경우에 한해서
			if NEW_YN = "Y" then				
				SQL = 		"insert into tbProcess_Record (BOM_Sub_BS_D_No,PR_Work_Order,PR_WorkType,PR_Process,PR_Line,PR_Amount,PR_Worker_CNT,PR_Work_Date,PR_Start_Time,PR_End_Time,PR_Memo) values ("
				SQL = SQL & "'"&arrBOM_Sub_BS_D_No(CNT2)&"',"
				SQL = SQL & "'"&strWork_Order&"',"
				SQL = SQL & "'작업',"
				SQL = SQL & "'SMD',"
				SQL = SQL & "'',"
				SQL = SQL & arrLP_SMD_Complete_Qty(CNT1) - arrOLD_LP_SMD_Complete_Qty(CNT1)&","
				SQL = SQL & "0,"
				SQL = SQL & "'"&date()&"',"
				SQL = SQL & "'',"
				SQL = SQL & "'',"
				SQL = SQL & "'')"		
				sys_DBCon.execute(SQL)
	
				'입력된 실적에 해당하는 모델파트넘버의 해당공정 재고를 +시킴
				call Process_Qty_BOM_Sub_Plus(arrBOM_Sub_BS_D_No(CNT2),"SMD",arrLP_SMD_Complete_Qty(CNT1) - arrOLD_LP_SMD_Complete_Qty(CNT1))
									
				'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 -시킴
				call Process_Qty_BOM_Sub_Before_Minus(arrBOM_Sub_BS_D_No(CNT2),"SMD",arrLP_SMD_Complete_Qty(CNT1) - arrOLD_LP_SMD_Complete_Qty(CNT1))
		
				'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 -시킴
				call Process_Qty_Parts_Minus(arrBOM_Sub_BS_D_No(CNT2),"SMD",arrLP_SMD_Complete_Qty(CNT1) - arrOLD_LP_SMD_Complete_Qty(CNT1))
			end if
			
			
			
			
			if instr(strWork_Order,"_") > 0 then
				arrLPE_Code = split(strWork_Order,"_")
				SQL = "update tbLGE_Plan_ETC set LPE_MAN_Complete_Qty = "&arrLP_MAN_Complete_Qty(CNT1)& " where LPE_Code = '"&arrLPE_Code(1)&"'"
				sys_DBCon.execute(SQL)
			else
				SQL = "update tbLGE_Plan set LP_MAN_Complete_Qty = "&arrLP_MAN_Complete_Qty(CNT1)& " where LP_Work_Order = '"&strWork_Order&"'"
				sys_DBCon.execute(SQL)
			end if
			
			'이 제번이 새로 체크된 제번임을 확인한다.
			NEW_YN = "N"
			if int(arrLP_MAN_Complete_Qty(CNT1)) > int(arrOLD_LP_MAN_Complete_Qty(CNT1)) then
				NEW_YN = "Y"
			end if
		
			'오직 새로 체크된 경우에 한해서
			if NEW_YN = "Y" then
				SQL = 		"insert into tbProcess_Record (BOM_Sub_BS_D_No,PR_Work_Order,PR_WorkType,PR_Process,PR_Line,PR_Amount,PR_Worker_CNT,PR_Work_Date,PR_Start_Time,PR_End_Time,PR_Memo) values ("
				SQL = SQL & "'"&arrBOM_Sub_BS_D_No(CNT2)&"',"
				SQL = SQL & "'"&strWork_Order&"',"
				SQL = SQL & "'작업',"
				SQL = SQL & "'MAN',"
				SQL = SQL & "'',"
				SQL = SQL & arrLP_MAN_Complete_Qty(CNT1) - arrOLD_LP_MAN_Complete_Qty(CNT1)&","
				SQL = SQL & "0,"
				SQL = SQL & "'"&date()&"',"
				SQL = SQL & "'',"
				SQL = SQL & "'',"
				SQL = SQL & "'')"		
				sys_DBCon.execute(SQL)

				'입력된 실적에 해당하는 모델파트넘버의 해당공정 재고를 +시킴
				call Process_Qty_BOM_Sub_Plus(arrBOM_Sub_BS_D_No(CNT2),"MAN",arrLP_MAN_Complete_Qty(CNT1) - arrOLD_LP_MAN_Complete_Qty(CNT1))
									
				'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 -시킴
				call Process_Qty_BOM_Sub_Before_Minus(arrBOM_Sub_BS_D_No(CNT2),"MAN",arrLP_MAN_Complete_Qty(CNT1) - arrOLD_LP_MAN_Complete_Qty(CNT1))
		
				'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 -시킴
				call Process_Qty_Parts_Minus(arrBOM_Sub_BS_D_No(CNT2),"MAN",arrLP_MAN_Complete_Qty(CNT1) - arrOLD_LP_MAN_Complete_Qty(CNT1))
			end if
		else
			if instr(strWork_Order,"_") > 0 then
				arrLPE_Code = split(strWork_Order,"_")
				SQL = "update tbLGE_Plan_ETC set LPE_ASM_Complete_Qty = "&arrLP_ASM_Complete_Qty(CNT1)& " where LPE_Code = '"&arrLPE_Code(1)&"'"
				sys_DBCon.execute(SQL)
			else
				SQL = "update tbLGE_Plan set LP_ASM_Complete_Qty = "&arrLP_ASM_Complete_Qty(CNT1)& " where LP_Work_Order = '"&strWork_Order&"'"
				sys_DBCon.execute(SQL)
			end if
			
			'이 제번이 새로 체크된 제번임을 확인한다.
			NEW_YN = "N"
			if int(arrLP_ASM_Complete_Qty(CNT1)) > int(arrOLD_LP_ASM_Complete_Qty(CNT1)) then
				NEW_YN = "Y"
			end if
		
			'오직 새로 체크된 경우에 한해서
			if NEW_YN = "Y" then
				SQL = 		"insert into tbProcess_Record (BOM_Sub_BS_D_No,PR_Work_Order,PR_WorkType,PR_Process,PR_Line,PR_Amount,PR_Worker_CNT,PR_Work_Date,PR_Start_Time,PR_End_Time,PR_Memo) values ("
				SQL = SQL & "'"&arrBOM_Sub_BS_D_No(CNT2)&"',"
				SQL = SQL & "'"&strWork_Order&"',"
				SQL = SQL & "'작업',"
				SQL = SQL & "'ASM',"
				SQL = SQL & "'',"
				SQL = SQL & arrLP_ASM_Complete_Qty(CNT1) - arrOLD_LP_ASM_Complete_Qty(CNT1)&","
				SQL = SQL & "0,"
				SQL = SQL & "'"&date()&"',"
				SQL = SQL & "'',"
				SQL = SQL & "'',"
				SQL = SQL & "'')"		
				sys_DBCon.execute(SQL)

				'입력된 실적에 해당하는 모델파트넘버의 해당공정 재고를 +시킴
				call Process_Qty_BOM_Sub_Plus(arrBOM_Sub_BS_D_No(CNT2),"ASM",arrLP_ASM_Complete_Qty(CNT1) - arrOLD_LP_ASM_Complete_Qty(CNT1))
									
				'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 -시킴
				call Process_Qty_BOM_Sub_Before_Minus(arrBOM_Sub_BS_D_No(CNT2),"ASM",arrLP_ASM_Complete_Qty(CNT1) - arrOLD_LP_ASM_Complete_Qty(CNT1))
		
				'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 -시킴
				call Process_Qty_Parts_Minus(arrBOM_Sub_BS_D_No(CNT2),"ASM",arrLP_ASM_Complete_Qty(CNT1) - arrOLD_LP_ASM_Complete_Qty(CNT1))
			end if
		end if
		
		
		
		
		if instr(strWork_Order,"_") > 0 then
			arrLPE_Code = split(strWork_Order,"_")
			SQL = "update tbLGE_Plan_ETC set LPE_DLV_Complete_Qty = "&arrLP_DLV_Complete_Qty(CNT1)& " where LPE_Code = '"&arrLPE_Code(1)&"'"
			sys_DBCon.execute(SQL)
		else
			SQL = "update tbLGE_Plan set LP_DLV_Complete_Qty = "&arrLP_DLV_Complete_Qty(CNT1)& " where LP_Work_Order = '"&strWork_Order&"'"
			sys_DBCon.execute(SQL)
		end if
			
		'이 제번이 새로 체크된 제번임을 확인한다.
		NEW_YN = "N"
		if int(arrLP_DLV_Complete_Qty(CNT1)) > int(arrOLD_LP_DLV_Complete_Qty(CNT1)) then
			NEW_YN = "Y"
		end if
	
		'오직 새로 체크된 경우에 한해서
		if NEW_YN = "Y" then
			SQL = 		"insert into tbProcess_Record (BOM_Sub_BS_D_No,PR_Work_Order,PR_WorkType,PR_Process,PR_Line,PR_Amount,PR_Worker_CNT,PR_Work_Date,PR_Start_Time,PR_End_Time,PR_Memo) values ("
			SQL = SQL & "'"&arrBOM_Sub_BS_D_No(CNT2)&"',"
			SQL = SQL & "'"&strWork_Order&"',"
			SQL = SQL & "'작업',"
			SQL = SQL & "'DLV',"
			SQL = SQL & "'',"
			SQL = SQL & arrLP_DLV_Complete_Qty(CNT1) - arrOLD_LP_DLV_Complete_Qty(CNT1)&","
			SQL = SQL & "0,"
			SQL = SQL & "'"&date()&"',"
			SQL = SQL & "'',"
			SQL = SQL & "'',"
			SQL = SQL & "'')"	
			sys_DBCon.execute(SQL)

			'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 -시킴
			call Process_Qty_BOM_Sub_Before_Minus(arrBOM_Sub_BS_D_No(CNT2),"DLV",arrLP_DLV_Complete_Qty(CNT1) - arrOLD_LP_DLV_Complete_Qty(CNT1))

		end if
	next
next
%>

<%
dim Request_Fields
dim strRequestForm
dim strRequestQueryString
for each Request_Fields in Request.Form
	if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	end if
next
for each Request_Fields in Request.QueryString
	if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	end if
next

%>
<form name="frmRedirect" action="lp_view.asp" method=post>
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>


<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->