<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem 변수선언
dim RS1
dim SQL
dim CNT1

dim strError
dim strError_Temp

dim arrID_All
dim arrPR_Work_Order
dim arrPR_Work_Date
dim arrPR_Line
dim arrPR_Worker_CNT
dim arrPR_WorkType
dim arrBOM_Sub_BS_D_No
dim arrPR_Amount
dim arrPR_Amount_NG
dim arrPR_Start_Time
dim arrPR_End_Time
dim arrPR_Loss_Time
dim arrPR_Rest_Time
dim arrPR_Memo

dim newPR_WorkType

dim PR_WorkType
dim BOM_Sub_BS_D_No
dim PR_Process
dim PR_Amount
dim PR_Work_Date

dim PR_ST
dim PR_Point

dim arrTemp

arrID_All				= split(Request("strID_All")&" "		,", ")
arrPR_Work_Order		= split(Request("PR_Work_Order")&" "	,", ")
arrPR_Work_Date			= split(Request("PR_Work_Date")&" "		,", ")
arrPR_Line				= split(Request("PR_Line")&" "			,", ")
arrPR_Worker_CNT		= split(Request("PR_Worker_CNT")&" "	,", ")
arrPR_WorkType			= split(Request("PR_WorkType")&" "		,", ")
arrBOM_Sub_BS_D_No		= split(Request("BOM_Sub_BS_D_No")&" "	,", ")
arrPR_Amount			= split(Request("PR_Amount")&" "		,", ")
arrPR_Amount_NG			= split(Request("PR_Amount_NG")&" "		,", ")
arrPR_Start_Time		= split(Request("PR_Start_Time")&" "	,", ")
arrPR_End_Time			= split(Request("PR_End_Time")&" "		,", ")
arrPR_Loss_Time			= split(Request("PR_Loss_Time")&" "		,", ")
arrPR_Rest_Time			= split(Request("PR_Rest_Time")&" "		,", ")
arrPR_Memo				= split(Request("PR_Memo")&" "			,", ")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrPR_Work_Order(CNT1)		= trim(arrPR_Work_Order(CNT1))
	arrPR_Work_Date(CNT1)		= trim(arrPR_Work_Date(CNT1))
	arrPR_Line(CNT1)			= trim(arrPR_Line(CNT1))
	arrPR_Worker_CNT(CNT1)		= trim(arrPR_Worker_CNT(CNT1))
	arrPR_WorkType(CNT1)		= trim(arrPR_WorkType(CNT1))
	arrBOM_Sub_BS_D_No(CNT1)	= trim(arrBOM_Sub_BS_D_No(CNT1))
	arrPR_Amount(CNT1)			= trim(arrPR_Amount(CNT1))
	arrPR_Amount_NG(CNT1)		= trim(arrPR_Amount_NG(CNT1))
	arrPR_Start_Time(CNT1)		= trim(arrPR_Start_Time(CNT1))
	arrPR_End_Time(CNT1)		= trim(arrPR_End_Time(CNT1))
	
	if arrPR_End_Time(CNT1) - arrPR_Start_Time(CNT1) < 1 then
		strError = strError & arrID_All(CNT1) & "번 실적의 작업시간이 0분입니다. 1분 이상이 되도록 입력하여 주십시오\n"
	end if
	
	arrPR_Loss_Time(CNT1)		= trim(arrPR_Loss_Time(CNT1))
	arrPR_Rest_Time(CNT1)		= trim(arrPR_Rest_Time(CNT1))
	arrPR_Memo(CNT1)			= trim(arrPR_Memo(CNT1))
next



set RS1 = Server.CreateObject("ADODB.RecordSet")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""		
	
		if strError_Temp = "" then
			
			SQL = "select top 1 PR_WorkType,BOM_Sub_BS_D_No,PR_Process,PR_Amount,PR_Work_Date from tbProcess_Record where PR_Code = '"&arrID_All(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			if not(RS1.Eof or RS1.Bof) then
				PR_WorkType		= RS1("PR_WorkType")
				BOM_Sub_BS_D_No	= RS1("BOM_Sub_BS_D_No")
				PR_Process		= RS1("PR_Process")
				PR_Amount		= RS1("PR_Amount")
				PR_Work_Date	= RS1("PR_Work_Date")

				newPR_WorkType	= arrPR_WorkType(CNT1)
				
				if instr(arrPR_Line(CNT1),"RH") > 0 then
					if BOM_Sub_BS_D_No = arrBOM_Sub_BS_D_No(CNT1) then	'동일 도번
					
						if PR_WorkType = "작업" and newPR_WorkType <> "작업" then '작업이었다가 다른 타입으로 변환 -> 원래 수량만큼 차감
							'입력된 실적에 해당하는 모델 파트넘버의 해당공정 재고를 -시킴
							call Process_Qty_BOM_Sub_Minus(BOM_Sub_BS_D_No,PR_Process,PR_Amount,PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 +시킴
							call Process_Qty_BOM_Sub_Before_Plus(BOM_Sub_BS_D_No,PR_Process,PR_Amount,PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 +시킴
							call Process_Qty_Parts_Plus(BOM_Sub_BS_D_No,PR_Process,PR_Amount)
							
						elseif PR_WorkType = "작업" and newPR_WorkType = "작업" then '계속 작업 -> 이전수량을 차감, 새로운 수량을 증가
							'입력된 실적에 해당하는 모델파트넘버의 해당공정 재고를 +시킴
							call Process_Qty_BOM_Sub_Plus(BOM_Sub_BS_D_No,PR_Process,int(arrPR_Amount(CNT1)) - int(PR_Amount),PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 -시킴
							call Process_Qty_BOM_Sub_Before_Minus(BOM_Sub_BS_D_No,PR_Process,int(arrPR_Amount(CNT1)) - int(PR_Amount),PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 -시킴
							call Process_Qty_Parts_Minus(BOM_Sub_BS_D_No,PR_Process,int(arrPR_Amount(CNT1)) - int(PR_Amount))
										
						elseif PR_WorkType <> "작업" and newPR_WorkType = "작업" then '작업이 아니다가 작업으로 변환 -> 새로운 수량을 증가
							'입력된 실적에 해당하는 모델파트넘버의 해당공정 재고를 +시킴
							call Process_Qty_BOM_Sub_Plus(arrBOM_Sub_BS_D_No(CNT1),PR_Process,arrPR_Amount(CNT1),PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 -시킴
							call Process_Qty_BOM_Sub_Before_Minus(arrBOM_Sub_BS_D_No(CNT1),PR_Process,arrPR_Amount(CNT1),PR_Work_Date)
					
							'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 -시킴
							call Process_Qty_Parts_Minus(arrBOM_Sub_BS_D_No(CNT1),PR_Process,arrPR_Amount(CNT1))
						end if		
					else	'다른 도번
						if PR_WorkType = "작업" and newPR_WorkType <> "작업" then '작업이었다가 다른 타입으로 변환
							'입력된 실적에 해당하는 모델 파트넘버의 해당공정 재고를 -시킴
							call Process_Qty_BOM_Sub_Minus(BOM_Sub_BS_D_No,PR_Process,PR_Amount,PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 +시킴
							call Process_Qty_BOM_Sub_Before_Plus(BOM_Sub_BS_D_No,PR_Process,PR_Amount,PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 +시킴
							call Process_Qty_Parts_Plus(BOM_Sub_BS_D_No,PR_Process,PR_Amount)
						elseif PR_WorkType = "작업" and newPR_WorkType = "작업" then '계속 작업 -> 이전수량을 차감
							'입력된 실적에 해당하는 모델 파트넘버의 해당공정 재고를 -시킴
							call Process_Qty_BOM_Sub_Minus(BOM_Sub_BS_D_No,PR_Process,PR_Amount,PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 +시킴
							call Process_Qty_BOM_Sub_Before_Plus(BOM_Sub_BS_D_No,PR_Process,PR_Amount,PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 +시킴
							call Process_Qty_Parts_Plus(BOM_Sub_BS_D_No,PR_Process,PR_Amount)
							
							'입력된 실적에 해당하는 모델파트넘버의 해당공정 재고를 +시킴
							call Process_Qty_BOM_Sub_Plus(arrBOM_Sub_BS_D_No(CNT1),PR_Process,arrPR_Amount(CNT1),PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 -시킴
							call Process_Qty_BOM_Sub_Before_Minus(arrBOM_Sub_BS_D_No(CNT1),PR_Process,arrPR_Amount(CNT1),PR_Work_Date)
					
							'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 -시킴
							call Process_Qty_Parts_Minus(arrBOM_Sub_BS_D_No(CNT1),PR_Process,arrPR_Amount(CNT1))
						elseif PR_WorkType <> "작업" and newPR_WorkType = "작업" then '작업이 아니다가 작업으로 변환
							'입력된 실적에 해당하는 모델파트넘버의 해당공정 재고를 +시킴
							call Process_Qty_BOM_Sub_Plus(arrBOM_Sub_BS_D_No(CNT1),PR_Process,arrPR_Amount(CNT1),PR_Work_Date)
							
							'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 -시킴
							call Process_Qty_BOM_Sub_Before_Minus(arrBOM_Sub_BS_D_No(CNT1),PR_Process,arrPR_Amount(CNT1),PR_Work_Date)
					
							'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 -시킴
							call Process_Qty_Parts_Minus(arrBOM_Sub_BS_D_No(CNT1),PR_Process,arrPR_Amount(CNT1))
						end if
					end if
				end if
			end if
			RS1.Close
			
			if newPR_WorkType <> "작업" then
				arrPR_Work_Order(CNT1) = ""
			end if
			
			'SQL = "select top 1 sum(PR_Amount) from tbProcess_Record where PR_Work_Order='"&arrPR_Work_Order(CNT1)&"' and PR_Code <> '"&arrID_All(CNT1)&"' and PR_Process = '"&PR_Process&"'"
			'RS1.Open SQL,sys_DBCon
			'if RS1.Eof or RS1.Bof then
				'if arrPR_Work_Order(CNT1) <> "" then
					'SQL = "update tbLGE_Plan set LP_"&PR_Process&"_Complete_YN = '' where LP_Work_Order='"&arrPR_Work_Order(CNT1)&"'"
					'sys_DBCon.execute(SQL)
				'end if
			'else
				'if arrPR_Work_Order(CNT1) <> "" and RS1(0) > 0 then
					'SQL = "update tbLGE_Plan set LP_"&PR_Process&"_Complete_YN = 'Y' where LP_Work_Order='"&arrPR_Work_Order(CNT1)&"'"
					'sys_DBCon.execute(SQL)
				'end if
			'end if	
			'RS1.Close
			
			SQL = "select * from tbBOM where B_Code in (select BOM_B_Code from tbBOM_Sub where BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"')"
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				PR_ST		= 3
			else
				if PR_Process = "MAN" then
					PR_ST		= RS1("B_ST")
				else
					PR_ST		= RS1("B_ST_Assm")
				end if
			end if
			RS1.Close
			
			if PR_Process = "IMD" then
				
				if instr(lcase(arrPR_Line(CNT1)),"rh") > 0 then
					if isnumeric(left(arrBOM_Sub_BS_D_No(CNT1),3)) then
						SQL = "select avg(bs_imd_radial_point) from tbBOM_Sub where left(BS_D_No,10) = '"&arrBOM_Sub_BS_D_No(CNT1)&"'"
					else
						SQL = "select avg(bs_imd_radial_point) from tbBOM_Sub where left(BS_D_No,9) = '"&arrBOM_Sub_BS_D_No(CNT1)&"'"
					end if
				else
					if isnumeric(left(arrBOM_Sub_BS_D_No(CNT1),3)) then
						SQL = "select avg(bs_imd_axial_point) from tbBOM_Sub where left(BS_D_No,10) = '"&arrBOM_Sub_BS_D_No(CNT1)&"'"
					else
						SQL = "select avg(bs_imd_axial_point) from tbBOM_Sub where left(BS_D_No,9) = '"&arrBOM_Sub_BS_D_No(CNT1)&"'"
					end if
				end if
				RS1.Open SQL,sys_DBCon
				PR_Point = RS1(0)
				RS1.Close
			elseif PR_Process = "SMD" and trim(arrBOM_Sub_BS_D_No(CNT1)) <> "" then
				SQL = "select sum(BQ_Qty) from tbBOM_Qty where BOM_Sub_BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"' and Parts_P_P_No in (select P_P_No from tbParts where P_Work_Type = 'SMD')"
				RS1.Open SQL,sys_DBCon
				PR_Point = RS1(0)
				RS1.Close
			end if
			
			if isnumeric(PR_Point) and PR_Point <> "" and not(ISNULL(PR_Point)) then
			else
				PR_Point = 20
			end if
			
			arrPR_Start_Time(CNT1)	= left(arrPR_Start_Time(CNT1),2) * 60 + right(arrPR_Start_Time(CNT1),2) - 500
			arrPR_End_Time(CNT1)	= left(arrPR_End_Time(CNT1),2) * 60 + right(arrPR_End_Time(CNT1),2) - 500
			
			SQL = "update tbProcess_Record set "
			SQL = SQL & "	PR_Work_Order='"&arrPR_Work_Order(CNT1)&"', "
			SQL = SQL & "	PR_Work_Date='"&arrPR_Work_Date(CNT1)&"', "
			SQL = SQL & "	PR_Line='"&arrPR_Line(CNT1)&"', "
			SQL = SQL & "	PR_Worker_CNT="&arrPR_Worker_CNT(CNT1)&", "
			SQL = SQL & "	PR_WorkType='"&arrPR_WorkType(CNT1)&"', "
			SQL = SQL & "	BOM_Sub_BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"', "
			SQL = SQL & "	PR_ST="&PR_ST&", "
			SQL = SQL & "	PR_Point="&PR_Point&", "
			SQL = SQL & "	PR_Amount="&arrPR_Amount(CNT1)&", "
			SQL = SQL & "	PR_Amount_NG="&arrPR_Amount_NG(CNT1)&", "
			SQL = SQL & "	PR_Start_Time='"&arrPR_Start_Time(CNT1)&"', "
			SQL = SQL & "	PR_End_Time='"&arrPR_End_Time(CNT1)&"', "
			SQL = SQL & "	PR_Loss_Time="&arrPR_Loss_Time(CNT1)&", "
			SQL = SQL & "	PR_Rest_Time="&arrPR_Rest_Time(CNT1)&", "
			SQL = SQL & "	PR_Memo='"&arrPR_Memo(CNT1)&"' "			
			SQL = SQL & "where PR_Code='"&arrID_All(CNT1)&"' "
			sys_DBCon.execute(SQL)
		end if
		
		strError = strError & strError_Temp
		
		SQL = "select sum(PR_Amount) from tbProcess_Record where PR_Work_Order='"&arrPR_Work_Order(CNT1)&"' and PR_Process='"&PR_Process&"'"
		RS1.Open SQL,sys_DBCon
		if instr(arrPR_Work_Order(CNT1),"_") > 0 then
			arrTemp = split(arrPR_Work_Order(CNT1),"_")
			SQL = "update tbLGE_Plan_ETC set LPE_"&PR_Process&"_Complete_Qty = "&RS1(0)&" where LPE_Type='"&arrTemp(1)&"' and LPE_Code='"&arrTemp(0)&"'"
		else
			SQL = "update tbLGE_Plan set LP_"&PR_Process&"_Complete_Qty = "&RS1(0)&" where LP_Work_Order='"&arrPR_Work_Order(CNT1)&"'"
		end if
		'sys_DBCon.execute(SQL)
		RS1.Close
	next
end if

set RS1 = nothing
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
if strError = "" then
%>
<form name="frmRedirect" action="pr_simple_list.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
	'strError = strError & "* 일부의 수정이 취소되었습니다."
%>
<form name="frmRedirect" action="pr_simple_list.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>



<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->