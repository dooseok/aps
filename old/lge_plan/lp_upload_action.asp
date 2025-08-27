<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
dim CNT1
dim CNT2
dim UpLoad
dim objXLS
dim XLSConnection
dim temp
dim strProperties

dim RS1
dim SQL
dim Sheet_Name
dim arrXLS

dim strFile
dim arrFile
dim File_Name

dim s_Min_LPD_Input_Date
dim s_Diff_LPD_Input_Date
dim s_LM_Company
dim S_Order_By_1
dim S_Order_By_2
dim S_Order_By_3
dim S_Order_By_4

dim Exist_YN

set UpLoad	= Server.CreateObject("Dext.FileUpLoad")
UpLoad.DefaultPath = DefaultPath_SCS_XLS_Reader

s_Min_LPD_Input_Date	= UpLoad("s_Min_LPD_Input_Date")
s_Diff_LPD_Input_Date	= UpLoad("s_Diff_LPD_Input_Date")
s_LM_Company			= UpLoad("s_LM_Company")
S_Order_By_1			= UpLoad("S_Order_By_1")
S_Order_By_2			= UpLoad("S_Order_By_2")
S_Order_By_3			= UpLoad("S_Order_By_3")
S_Order_By_4			= UpLoad("S_Order_By_4")

strFile 	= UpLoad("strFile")
arrFile		= split(strFile,"\")
File_Name	= lcase(arrFile(ubound(arrFile)))

temp = UpLoad("strFile").SaveAs(DefaultPath_SCS_XLS_Reader & File_Name, False)
temp = replace(temp,"\","/")

set objXLS = Server.CreateObject("ADODB.Connection") 
XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""

objXLS.Open XLSConnection    
set RS1 = objXLS.OpenSchema(20)
Sheet_Name	= RS1(2)
Sheet_Name	= "["&replace(Sheet_Name,"'","")&"]"
set RS1 = nothing

set RS1 = Server.CreateObject("ADODB.RecordSet") 
SQL  = " select * from "&Sheet_Name
RS1.Open SQL,objXLS 
arrXLS = RS1.getRows()
RS1.close
set RS1 = nothing

if instr(File_Name,"daily") > 0 then
	call Daily_Plan()
elseif instr(File_Name,"changed") > 0 then
	call Changed_Plan()
elseif instr(File_Name,"scschb01") > 0 then
	call Monthly_Order_Old()
elseif instr(File_Name,"pur10") > 0 then
	call Monthly_Order()
end if
%>
<form name="frmRedirect" action="lp_view.asp" method="post">
<input type="hidden" name="s_Min_LPD_Input_Date"	value="<%=s_Min_LPD_Input_Date%>">
<input type="hidden" name="s_Diff_LPD_Input_Date"	value="<%=s_Diff_LPD_Input_Date%>">
<input type="hidden" name="S_Order_By_1"			value="<%=S_Order_By_1%>">
<input type="hidden" name="S_Order_By_2"			value="<%=S_Order_By_2%>">
<input type="hidden" name="S_Order_By_3"			value="<%=S_Order_By_3%>">
<input type="hidden" name="S_Order_By_4"			value="<%=S_Order_By_4%>">
<input type="hidden" name="s_LM_Company"			value="<%=s_LM_Company%>">
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
objXLS.close
set objXLS = nothing
set UpLoad = nothing
%>

<%
sub Monthly_Order()
	dim RS1
	dim SQL
	
	dim D_No
	dim BOM_Model_BM_D_Sub_No
	
	dim Work_Order
	dim LM_Code
	dim LM_Name
	dim LM_Company
	dim BOM_Sub_BS_D_No_1
	dim BOM_Sub_BS_D_No_2
	dim BOM_Sub_BS_D_No_3
	dim BOM_Sub_BS_D_No_4

	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	
	for CNT1 = 1 to ubound(arrXLS, 2)
		Work_Order = arrXLS(15,CNT1)
		if instr(Work_Order,"-") > 0 then
			Work_Order = left(Work_Order,instr(Work_Order,"-")+1)
		end if

		SQL = "select LP_Model from tbLGE_Plan where LP_Work_Order = '"&Work_Order&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			LM_Name = "N/A"
		else
			LM_Name = RS1("LP_Model")
		end if
		RS1.Close
		
		'if instr(LM_Name,".") > 0 then
			'LM_Name = left(LM_Name,instr(LM_Name,".")-1)
		'end if	
		
		Exist_YN				= ""
		LM_Code					= ""
		
		if LM_Name <> "N/A" then
			SQL = "select * from tbLGE_Model where LM_Name='"&LM_Name&"'"
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				Exist_YN = "N"
			else
				Exist_YN = "Y"
				LM_Code	= RS1("LM_Code")
			end if
			RS1.Close
			
			if Exist_YN = "N" then	'기존에 이런 모델이 없었을 경우.
				SQL = "insert into tbLGE_Model (LM_Company,LM_Name) values "
				SQL = SQL & "('MSE','"&LM_Name&"')"
				sys_DBCon.execute(SQL)
			else					'기존에 있던 모델명인 경우
				SQL = "update tbLGE_Model set LM_Company='MSE', BOM_Sub_BS_D_No_1='',BOM_Sub_BS_D_No_2='', BOM_Sub_BS_D_No_3='',BOM_Sub_BS_D_No_4='' where LM_Code='"&LM_Code&"'"
				sys_DBCon.execute(SQL)			
			end if
		end if
	next
	
	for CNT1 = 1 to ubound(arrXLS, 2)
		D_No		= arrXLS(1,CNT1)
		Work_Order	= arrXLS(15,CNT1)
		
		SQL = "select LP_Model from tbLGE_Plan where LP_Work_Order = '"&Work_Order&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			LM_Name = "N/A"
		else
			LM_Name = RS1("LP_Model")
		end if
		RS1.Close
		
		'if instr(LM_Name,".") > 0 then
			'LM_Name = left(LM_Name,instr(LM_Name,".")-1)
		'end if
		
		Exist_YN			= ""
		LM_Code				= ""
		LM_Company			= ""
		BOM_Sub_BS_D_No_1	= ""
		BOM_Sub_BS_D_No_2	= ""
		BOM_Sub_BS_D_No_3	= ""
		BOM_Sub_BS_D_No_4	= ""
		
		if LM_Name <> "N/A" then
			SQL = "select * from tbLGE_Model where LM_Name='"&LM_Name&"'"
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				Exist_YN = "N"
			else
				Exist_YN = "Y"
				LM_Code				= RS1("LM_Code")
				LM_Company			= RS1("LM_Company")
				BOM_Sub_BS_D_No_1	= RS1("BOM_Sub_BS_D_No_1")
				BOM_Sub_BS_D_No_2	= RS1("BOM_Sub_BS_D_No_2")
				BOM_Sub_BS_D_No_3	= RS1("BOM_Sub_BS_D_No_3")
				BOM_Sub_BS_D_No_4	= RS1("BOM_Sub_BS_D_No_4")
			end if
			RS1.Close
			
			if	(BOM_Sub_BS_D_No_1 = D_No) or (BOM_Sub_BS_D_No_2 = D_No) or (BOM_Sub_BS_D_No_3 = D_No) or (BOM_Sub_BS_D_No_4 = D_No) then
			else
				if BOM_Sub_BS_D_No_1 = "" or isnull(BOM_Sub_BS_D_No_1) then
					SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_1='"&D_No&"' where LM_Code='"&LM_Code&"'"
				elseif BOM_Sub_BS_D_No_2 = "" or isnull(BOM_Sub_BS_D_No_2) then
					if BOM_Sub_BS_D_No_1 < D_No then
						SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_2=BOM_Sub_BS_D_No_1, BOM_Sub_BS_D_No_1='"&D_No&"' where LM_Code='"&LM_Code&"'"
					else
						SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_2='"&D_No&"' where LM_Code='"&LM_Code&"'"
					end if
				elseif BOM_Sub_BS_D_No_3 = "" or isnull(BOM_Sub_BS_D_No_3) then
					if BOM_Sub_BS_D_No_1 < D_No then
						SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_3=BOM_Sub_BS_D_No_2, BOM_Sub_BS_D_No_2=BOM_Sub_BS_D_No_1, BOM_Sub_BS_D_No_1='"&D_No&"' where LM_Code='"&LM_Code&"'"
					elseif BOM_Sub_BS_D_No_2 < D_No then
						SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_3=BOM_Sub_BS_D_No_2, BOM_Sub_BS_D_No_2='"&D_No&"' where LM_Code='"&LM_Code&"'"
					else
						SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_3='"&D_No&"' where LM_Code='"&LM_Code&"'"
					end if
				elseif BOM_Sub_BS_D_No_4 = "" or isnull(BOM_Sub_BS_D_No_4) then
					if BOM_Sub_BS_D_No_1 < D_No then
						SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_4=BOM_Sub_BS_D_No_3, BOM_Sub_BS_D_No_3=BOM_Sub_BS_D_No_2, BOM_Sub_BS_D_No_2=BOM_Sub_BS_D_No_1, BOM_Sub_BS_D_No_1='"&D_No&"' where LM_Code='"&LM_Code&"'"
					elseif BOM_Sub_BS_D_No_2 < D_No then
						SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_4=BOM_Sub_BS_D_No_3, BOM_Sub_BS_D_No_3=BOM_Sub_BS_D_No_2, BOM_Sub_BS_D_No_2='"&D_No&"' where LM_Code='"&LM_Code&"'"
					elseif BOM_Sub_BS_D_No_3 < D_No then
						SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_4=BOM_Sub_BS_D_No_3, BOM_Sub_BS_D_No_3='"&D_No&"' where LM_Code='"&LM_Code&"'"
					else
						SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_4='"&D_No&"' where LM_Code='"&LM_Code&"'"
					end if
				end if
			end if
			sys_DBCon.execute(SQL)
		end if
	next
	
	set RS1 = nothing
end sub
%>

<%
sub Daily_Plan()
	dim RS1
	dim SQL
	
	dim Ubound_Date
	
	dim LP_Line
	dim LP_Work_Order
	dim LP_Model
	dim LP_Suffix
	dim LP_Buyer
	dim LP_Tool
	dim LP_Input_Time
	dim LP_Lot
	dim LP_Lot_Remain
	
	dim LPD_Input_Date
	dim LPD_Input_Qty
	
	dim DateField
	dim strToday
	dim strDate
	dim arrDate
	dim StartDate
	dim EndDate
	
	dim locLine
	dim locWork_Order
	dim locModel
	dim locSuffix
	dim locTool
	dim locInput_Time
	dim locLot
	dim locLot_Remain
	
	dim locStartDate
	
	strToday = date() '2008-01-04'
	
	for CNT1 = 0 to ubound(arrXLS, 1)
		if isnumeric(arrXLS(CNT1, 0)) and locStartDate = 0 then
			if int(arrXLS(CNT1, 0)) = int(right(date(),2)) then
				locStartDate = CNT1
			end if
		end if
	next
	
	for CNT1 = 0 to 36
		if instr(lcase(arrXLS(CNT1, 0)),"line") > 0 then
			locLine = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"w/o") > 0 then
			locWork_Order = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"model") > 0 then
			locModel = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"suffix") > 0 then
			locSuffix = CNT1
		end if

		if instr(lcase(arrXLS(CNT1, 0)),"tool") > 0 then
			locTool = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"input") > 0 then
			locInput_Time = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"lot") > 0 then
			locLot = CNT1
		end if
		if instr(arrXLS(CNT1, 0),"잔량") > 0 or instr(arrXLS(CNT1, 0),"계획") > 0 then
			locLot_Remain = CNT1
		end if
	next
	
	for CNT1 = locStartDate to ubound(arrXLS, 1)	'26번째 칸부터 끝까지 루프
		DateField = arrXLS(CNT1,0)		'날짜필드에, 최상단 값을 받음
		if (isNumeric(DateField)) then	'날짜 필드라면
			if CNT1 > locLot_Remain + 1 and cint(DateField) = 1 then	'1자가 나온다면
				strToday = dateadd("m",1,strToday)			'오늘변수를 한달 연장
			end if
			strDate = strDate & left(strToday,8)&DateField & "|%|"	'날짜문자열에 오늘의 달, 날짜열을 추가한다.
		end if
	next
	
	arrDate		= split(strDate,"|%|")
	
	Ubound_Date = locStartDate + ubound(arrDate)-1
	
	StartDate	= arrDate(0)
	EndDate		= arrDate(ubound(arrDate)-1)
	
	SQL = "delete tbLGE_Plan_Date where LPD_Input_Date between '"&StartDate&"' and '"&EndDate&"'"
	sys_DBCon.execute(SQL)
	'SQL = "delete tbLGE_Plan_Date"
	'sys_DBCon.execute(SQL)
	'SQL = "delete tbLGE_Plan"
	'sys_DBCon.execute(SQL)
	'SQL = "dbcc checkident('tbLGE_Plan_Date',reseed,0)"
	'sys_DBCon.execute(SQL)
	'SQL = "dbcc checkident('tbLGE_Plan',reseed,0)"
	'sys_DBCon.execute(SQL)
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	
	for CNT1 = 2 to ubound(arrXLS, 2)
		if instr(arrXLS(0, CNT1),"요약") = 0 then
			
			LP_Line			= trim(arrXLS(locLine, CNT1))
			LP_Work_Order	= trim(arrXLS(locWork_Order, CNT1))
			LP_Model		= trim(arrXLS(locModel, CNT1))
			LP_Suffix		= trim(arrXLS(locSuffix, CNT1))
			LP_Tool			= trim(arrXLS(locTool, CNT1))
			LP_Input_Time	= trim(arrXLS(locInput_Time, CNT1))
			LP_Lot			= trim(arrXLS(locLot, CNT1))
			LP_Lot_Remain	= trim(arrXLS(locLot_Remain, CNT1))
			
			if isnull(LP_Tool) then
				LP_Tool = ""
			else
				LP_Tool = replace(LP_Tool,"'","''")
			end if
			
			if len(LP_Input_Time) = 4 then
				LP_Input_Time="0"&LP_Input_Time
			end if
			
			SQL = "select top 1 LP_Work_Order from tbLGE_Plan where LP_Work_Order='"&LP_Work_Order&"'"
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				Exist_YN = "N"
			else
				Exist_YN = "Y"
			end if
			RS1.Close
			
			if LP_Lot = "" then
				LP_Lot = 0
			end if
			
			if LP_Lot_Remain = "" then
				LP_Lot_Remain = 0
			end if

			if Exist_YN = "N" then
				SQL = "insert into tbLGE_Plan "
				SQL = SQL & "(LP_Line, LP_Work_Order, LP_Model, LP_Suffix, LP_Buyer, LP_Tool, LP_Input_Time, LP_Lot, LP_Lot_Remain) "
				SQL = SQL & "values " 
				SQL = SQL & "('"&LP_Line&"', "
				SQL = SQL & "'"&LP_Work_Order&"', "
				SQL = SQL & "'"&LP_Model&"', "
				SQL = SQL & "'"&LP_Suffix&"', "
				SQL = SQL & "'', "
				SQL = SQL & "'"&LP_Tool&"', "
				SQL = SQL & "'"&LP_Input_Time&"', "
				SQL = SQL & "'"&LP_Lot&"', "
				SQL = SQL & "'"&LP_Lot_Remain&"')"
				sys_DBCon.execute(SQL)
			else
				SQL = "update tbLGE_Plan set "
				SQL = SQL & "LP_Line='"&LP_Line&"', "
				SQL = SQL & "LP_Model='"&LP_Model&"', "
				SQL = SQL & "LP_Suffix='"&LP_Suffix&"', "
				SQL = SQL & "LP_Buyer='', "
				SQL = SQL & "LP_Tool='"&LP_Tool&"', "
				SQL = SQL & "LP_Input_Time='"&LP_Input_Time&"', "
				SQL = SQL & "LP_Lot='"&LP_Lot&"', "
				SQL = SQL & "LP_Lot_Remain='"&LP_Lot_Remain&"' "
				SQL = SQL & "Where LP_Work_Order = '"&LP_Work_Order&"'"
				sys_DBCon.execute(SQL)
			end if
			
			for CNT2 = locStartDate to Ubound_Date
				LPD_Input_Date	= arrDate(CNT2-locStartDate)
				LPD_Input_Qty	= trim(arrXLS(CNT2, CNT1))
				if isNumeric(LPD_Input_Qty) then
					if LPD_Input_Qty > 0 then
						SQL = "insert into tbLGE_Plan_Date "
						SQL = SQL & "(LGE_Plan_LP_Work_Order,LGE_Plan_LP_Model,LPD_Input_Date, LPD_Input_Qty) "
						SQL = SQL & "values " 
						SQL = SQL & "('"&LP_Work_Order&"', " 
						SQL = SQL & "'"&LP_Model&"', " 
						SQL = SQL & "'"&LPD_Input_Date&"', " 
						SQL = SQL & LPD_Input_Qty&")" 
						sys_DBCon.execute(SQL)
					end if
				end if
			next
			
			SQL = "select top 1 LM_Code from tbLGE_Model where LM_Name='"&LP_Model&"'"
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				SQL = "insert into tbLGE_Model (LM_Company,LM_Name) values ('미분류','"&LP_Model&"')"
				sys_DBCon.execute(SQL)
			end if
			RS1.Close
		end if
	next
	
	set RS1 = nothing
end sub
%>

<%
sub Changed_Plan()
	dim RS1
	dim SQL
	
	dim Ubound_Date
	
	dim LP_Line
	dim LP_Work_Order
	dim LP_Model
	dim LP_Suffix
	dim LP_Buyer
	dim LP_Tool
	dim LP_Input_Time
	dim LP_Lot
	dim LP_Lot_Remain
	
	dim LPD_Input_Date
	dim LPD_Input_Qty
	
	dim DateField
	dim strToday
	dim strDate
	dim arrDate
	dim StartDate
	dim EndDate
	
	dim locLine
	dim locWork_Order
	dim locModel
	dim locSuffix
	dim locTool
	dim locInput_Time
	dim locLot
	dim locLot_Remain
	
	dim locStartDate
	
	locStartDate = 0
	
	strToday = date() '2008-01-02'
	
	for CNT1 = 0 to ubound(arrXLS, 1)
		if isnumeric(arrXLS(CNT1, 0)) and locStartDate = 0 then
			if int(arrXLS(CNT1, 0)) = int(right(date(),2)) then
				locStartDate = CNT1
			end if
		end if
	next
	
	for CNT1 = 0 to 36
		if instr(lcase(arrXLS(CNT1, 0)),"line") > 0 then
			locLine = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"w/o") > 0 then
			locWork_Order = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"model") > 0 then
			locModel = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"suffix") > 0 then
			locSuffix = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"tool") > 0 then
			locTool = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"input") > 0 then
			locInput_Time = CNT1
		end if
		if instr(lcase(arrXLS(CNT1, 0)),"lot") > 0 then
			locLot = CNT1
		end if
		if instr(arrXLS(CNT1, 0),"계획") > 0 then
			locLot_Remain = CNT1
		end if
	next
	
	for CNT1 = locStartDate to ubound(arrXLS, 1)	'26번째 칸부터 끝까지 루프
		DateField = arrXLS(CNT1,0)		'날짜필드에, 최상단 값을 받음
		if (isNumeric(DateField)) then	'날짜 필드라면
			if CNT1 > locLot_Remain + 1 and cint(DateField) = 1 then	'1자가 나온다면
				strToday = dateadd("m",1,strToday)			'오늘변수를 한달 연장
			end if
			strDate = strDate & left(strToday,8)&DateField & "|%|"	'날짜문자열에 오늘의 달, 날짜열을 추가한다.
		end if
	next

	arrDate		= split(strDate,"|%|")
	
	Ubound_Date = locStartDate + ubound(arrDate)-1
	
	StartDate	= arrDate(0)
	EndDate		= arrDate(ubound(arrDate)-1)

	SQL = "delete tbLGE_Plan_Date where LPD_Input_Date between '"&StartDate&"' and '"&EndDate&"'"
	sys_DBCon.execute(SQL)
	'SQL = "delete tbLGE_Plan_Date"
	'sys_DBCon.execute(SQL)
	'SQL = "delete tbLGE_Plan"
	'sys_DBCon.execute(SQL)
	'SQL = "dbcc checkident('tbLGE_Plan_Date',reseed,0)"
	'sys_DBCon.execute(SQL)
	'SQL = "dbcc checkident('tbLGE_Plan',reseed,0)"
	'sys_DBCon.execute(SQL)
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")	

	for CNT1 = 2 to ubound(arrXLS, 2)
		if arrXLS(locLine, CNT1) <> "" and arrXLS(locWork_Order, CNT1) <> "" and arrXLS(locModel, CNT1) <> "" then
						
			LP_Line			= trim(arrXLS(locLine, CNT1))
			LP_Work_Order	= trim(arrXLS(locWork_Order, CNT1))
			LP_Model		= trim(arrXLS(locModel, CNT1))
			LP_Suffix		= trim(arrXLS(locSuffix, CNT1))
			'LP_Buyer		= trim(arrXLS(4, CNT1))
			LP_Tool			= trim(arrXLS(locTool, CNT1))
			LP_Input_Time	= trim(arrXLS(locInput_Time, CNT1))
			LP_Lot			= trim(arrXLS(locLot, CNT1))
			LP_Lot_Remain	= trim(arrXLS(locLot_Remain, CNT1))
			
			if isnull(LP_Tool) then
				LP_Tool = ""
			else
				LP_Tool = replace(LP_Tool,"'","''")
			end if
			
			if len(LP_Input_Time) = 4 then
				LP_Input_Time="0"&LP_Input_Time
			end if
			
			SQL = "select top 1 LP_Work_Order from tbLGE_Plan where LP_Work_Order='"&LP_Work_Order&"'"
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				Exist_YN = "N"
			else
				Exist_YN = "Y"
			end if
			RS1.Close
			
			if LP_Lot = "" then
				LP_Lot = 0
			end if
			
			if LP_Lot_Remain = "" then
				LP_Lot_Remain = 0
			end if
			
			if Exist_YN = "N" then
				SQL = "insert into tbLGE_Plan "
				SQL = SQL & "(LP_Line, LP_Work_Order, LP_Model, LP_Suffix, LP_Tool, LP_Input_Time, LP_Lot, LP_Lot_Remain) "
				SQL = SQL & "values " 
				SQL = SQL & "('"&LP_Line&"', "
				SQL = SQL & "'"&LP_Work_Order&"', "
				SQL = SQL & "'"&LP_Model&"', "
				SQL = SQL & "'"&LP_Suffix&"', "
				SQL = SQL & "'"&LP_Tool&"', "
				SQL = SQL & "'"&LP_Input_Time&"', "
				SQL = SQL & "'"&LP_Lot&"', "
				SQL = SQL & "'"&LP_Lot_Remain&"')"				
				sys_DBCon.execute(SQL)
			else
				SQL = "update tbLGE_Plan set "
				SQL = SQL & "LP_Line='"&LP_Line&"', "
				SQL = SQL & "LP_Model='"&LP_Model&"', "
				SQL = SQL & "LP_Suffix='"&LP_Suffix&"', "
				SQL = SQL & "LP_Tool='"&LP_Tool&"', "
				SQL = SQL & "LP_Input_Time='"&LP_Input_Time&"', "
				SQL = SQL & "LP_Lot='"&LP_Lot&"', "
				SQL = SQL & "LP_Lot_Remain='"&LP_Lot_Remain&"' "
				SQL = SQL & "Where LP_Work_Order = '"&LP_Work_Order&"'"
				sys_DBCon.execute(SQL)
			end if
			
			for CNT2 = locStartDate to Ubound_Date
				LPD_Input_Date	= arrDate(CNT2-locStartDate)
				LPD_Input_Qty	= trim(arrXLS(CNT2, CNT1))
				
				if isNumeric(LPD_Input_Qty) then
					if LPD_Input_Qty > 0 then
						SQL = "insert into tbLGE_Plan_Date "
						SQL = SQL & "(LGE_Plan_LP_Work_Order,LGE_Plan_LP_Model,LPD_Input_Date, LPD_Input_Qty) "
						SQL = SQL & "values " 
						SQL = SQL & "('"&LP_Work_Order&"', " 
						SQL = SQL & "'"&LP_Model&"', "
						SQL = SQL & "'"&LPD_Input_Date&"', " 
						SQL = SQL & LPD_Input_Qty&")" 
						sys_DBCon.execute(SQL)
					end if
				end if
			next
			
			SQL = "select top 1 LM_Code from tbLGE_Model where LM_Name='"&LP_Model&"'"
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				SQL = "insert into tbLGE_Model (LM_Company,LM_Name) values ('미분류','"&LP_Model&"')"
				sys_DBCon.execute(SQL)
			end if
			RS1.Close
		end if
	next
	
	set RS1 = nothing
end sub
%>

<%
sub Monthly_Order_Old()
	dim RS1
	dim SQL
	
	dim D_No
	dim BOM_Model_BM_D_Sub_No
	
	dim LM_Code
	dim LM_Name
	dim LM_Company
	dim BOM_Sub_BS_D_No_1
	dim BOM_Sub_BS_D_No_2
	dim BOM_Sub_BS_D_No_3
	dim BOM_Sub_BS_D_No_4

	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	
	for CNT1 = 1 to ubound(arrXLS, 2)
		LM_Name	= arrXLS(2,CNT1)
		
		if instr(LM_Name,".") > 0 then
			LM_Name = left(LM_Name,instr(LM_Name,".")-1)
		end if	
		
		Exist_YN				= ""
		LM_Code					= ""
		
		SQL = "select * from tbLGE_Model where LM_Name='"&LM_Name&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			Exist_YN = "N"
		else
			Exist_YN = "Y"
			LM_Code	= RS1("LM_Code")
		end if
		RS1.Close
		
		if Exist_YN = "N" then	'기존에 이런 모델이 없었을 경우.
			SQL = "insert into tbLGE_Model (LM_Company,LM_Name) values "
			SQL = SQL & "('MSE','"&LM_Name&"')"
			sys_DBCon.execute(SQL)
		else					'기존에 있던 모델명인 경우
			SQL = "update tbLGE_Model set LM_Company='MSE', BOM_Sub_BS_D_No_1='',BOM_Sub_BS_D_No_2='', BOM_Sub_BS_D_No_3='',BOM_Sub_BS_D_No_4='' where LM_Code='"&LM_Code&"'"
			sys_DBCon.execute(SQL)			
		end if
	next
	
	for CNT1 = 1 to ubound(arrXLS, 2)
		D_No	= arrXLS(3,CNT1)
		LM_Name	= arrXLS(2,CNT1)
		
		if instr(LM_Name,".") > 0 then
			LM_Name = left(LM_Name,instr(LM_Name,".")-1)
		end if
		
		Exist_YN			= ""
		LM_Code				= ""
		LM_Company			= ""
		BOM_Sub_BS_D_No_1	= ""
		BOM_Sub_BS_D_No_2	= ""
		BOM_Sub_BS_D_No_3	= ""
		BOM_Sub_BS_D_No_4	= ""
		
		SQL = "select * from tbLGE_Model where LM_Name='"&LM_Name&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			Exist_YN = "N"
		else
			Exist_YN = "Y"
			LM_Code				= RS1("LM_Code")
			LM_Company			= RS1("LM_Company")
			BOM_Sub_BS_D_No_1	= RS1("BOM_Sub_BS_D_No_1")
			BOM_Sub_BS_D_No_2	= RS1("BOM_Sub_BS_D_No_2")
			BOM_Sub_BS_D_No_3	= RS1("BOM_Sub_BS_D_No_3")
			BOM_Sub_BS_D_No_4	= RS1("BOM_Sub_BS_D_No_4")
		end if
		RS1.Close
		
		if	(BOM_Sub_BS_D_No_1 = D_No) or (BOM_Sub_BS_D_No_2 = D_No) or (BOM_Sub_BS_D_No_3 = D_No) or (BOM_Sub_BS_D_No_4 = D_No) then
		else
			if BOM_Sub_BS_D_No_1 = "" or isnull(BOM_Sub_BS_D_No_1) then
				SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_1='"&D_No&"' where LM_Code='"&LM_Code&"'"
			elseif BOM_Sub_BS_D_No_2 = "" or isnull(BOM_Sub_BS_D_No_2) then
				if BOM_Sub_BS_D_No_1 < D_No then
					SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_2=BOM_Sub_BS_D_No_1, BOM_Sub_BS_D_No_1='"&D_No&"' where LM_Code='"&LM_Code&"'"
				else
					SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_2='"&D_No&"' where LM_Code='"&LM_Code&"'"
				end if
			elseif BOM_Sub_BS_D_No_3 = "" or isnull(BOM_Sub_BS_D_No_3) then
				if BOM_Sub_BS_D_No_1 < D_No then
					SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_3=BOM_Sub_BS_D_No_2, BOM_Sub_BS_D_No_2=BOM_Sub_BS_D_No_1, BOM_Sub_BS_D_No_1='"&D_No&"' where LM_Code='"&LM_Code&"'"
				elseif BOM_Sub_BS_D_No_2 < D_No then
					SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_3=BOM_Sub_BS_D_No_2, BOM_Sub_BS_D_No_2='"&D_No&"' where LM_Code='"&LM_Code&"'"
				else
					SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_3='"&D_No&"' where LM_Code='"&LM_Code&"'"
				end if
			elseif BOM_Sub_BS_D_No_4 = "" or isnull(BOM_Sub_BS_D_No_4) then
				if BOM_Sub_BS_D_No_1 < D_No then
					SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_4=BOM_Sub_BS_D_No_3, BOM_Sub_BS_D_No_3=BOM_Sub_BS_D_No_2, BOM_Sub_BS_D_No_2=BOM_Sub_BS_D_No_1, BOM_Sub_BS_D_No_1='"&D_No&"' where LM_Code='"&LM_Code&"'"
				elseif BOM_Sub_BS_D_No_2 < D_No then
					SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_4=BOM_Sub_BS_D_No_3, BOM_Sub_BS_D_No_3=BOM_Sub_BS_D_No_2, BOM_Sub_BS_D_No_2='"&D_No&"' where LM_Code='"&LM_Code&"'"
				elseif BOM_Sub_BS_D_No_3 < D_No then
					SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_4=BOM_Sub_BS_D_No_3, BOM_Sub_BS_D_No_3='"&D_No&"' where LM_Code='"&LM_Code&"'"
				else
					SQL = "update tbLGE_Model set BOM_Sub_BS_D_No_4='"&D_No&"' where LM_Code='"&LM_Code&"'"
				end if
			end if
		end if
		sys_DBCon.execute(SQL)
	next
	
	set RS1 = nothing
end sub
%>


<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
