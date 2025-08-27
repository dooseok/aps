<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<%
dim FileName
FileName = right(replace(date(),"-",""),6) & "_lp_view"

dim RS1
dim RS2
dim CNT1

dim SQL

dim s_diff_LPD_Input_Date
dim s_Min_LPD_Input_Date

dim LM_Company
dim LP_Line
dim LP_Work_Order
dim LP_Model
dim LP_Suffix
dim LP_Tool
dim LP_Tool_Type
dim LP_Input_time
dim LP_LOT
dim LP_LOT_Remain
dim LP_IMD_Complete_QTY
dim LP_SMD_Complete_QTY
dim LP_MAN_Complete_QTY
dim LP_ASM_Complete_QTY
dim LP_DLV_Complete_QTY
dim BOM_Sub_BS_D_No
dim BOM_Sub_BS_D_No_1
dim BOM_Sub_BS_D_No_2
dim BOM_Sub_BS_D_No_3
dim BOM_Sub_BS_D_No_4
dim BOM_Sub_BS_D_No_Str

dim LPD_Input_Qty

SQL							= Request("SQL")

s_diff_LPD_Input_Date		= Request("s_diff_LPD_Input_Date")
s_Min_LPD_Input_Date		= Request("s_Min_LPD_Input_Date")

set RS2 = Server.CreateObject("ADODB.RecordSet") 
set RS1 = Server.CreateObject("ADODB.RecordSet") 
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
%>
<script language="javascript">
alert("조회결과가 없습니다.")
window.close();
</script>
<%
else

	Response.Buffer = false
	Response.Expires = 0
	Response.ContentType = "application/vnd.ms-excel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition","attachment;filename="&FileName&".xls"

	response.write "COMP"
	response.write vbtab
	response.write "LINE"
	response.write vbtab
	response.write "W/O"
	response.write vbtab
	response.write "MODEL"
	response.write vbtab
	response.write "PART NO"
	response.write vbtab
	response.write "SUFFIX"
	response.write vbtab
	response.write "TOOL"
	response.write vbtab
	response.write "TYPE"
	response.write vbtab
	response.write "INPUT"
	response.write vbtab
	response.write "LOT"
	response.write vbtab
	response.write "PLAN"
	response.write vbtab
	response.write "IMD"
	response.write vbtab
	response.write "SMD"
	response.write vbtab
	response.write "MAN"
	response.write vbtab
	response.write "ASM"
	response.write vbtab
	response.write "DLV"
	response.write vbtab

	for CNT1 = 0 to s_diff_LPD_Input_Date
		response.write Right(dateadd("d",CNT1,s_Min_LPD_Input_Date),2)
		response.write vbtab
	next
	response.write vbcrlf
	
	do until RS1.Eof
		LM_Company				= RS1("LM_Company")
		LP_Line					= RS1("LP_Line")
		LP_Work_Order			= RS1("LP_Work_Order")
		LP_Model				= RS1("LP_Model")
	
		LP_Suffix				= RS1("LP_Suffix")
		LP_Tool					= RS1("LP_Tool")
		LP_Tool_Type			= RS1("LP_Tool_Type")
		LP_Input_time			= RS1("LP_Input_time")
		LP_LOT					= RS1("LP_LOT")
		LP_LOT_Remain			= RS1("LP_LOT_Remain")
		
		LP_IMD_Complete_QTY		= RS1("LP_IMD_Complete_QTY")
		LP_SMD_Complete_QTY		= RS1("LP_SMD_Complete_QTY")
		LP_MAN_Complete_QTY		= RS1("LP_MAN_Complete_QTY")
		LP_ASM_Complete_QTY		= RS1("LP_ASM_Complete_QTY")
		LP_DLV_Complete_QTY		= RS1("LP_DLV_Complete_QTY")
		
		BOM_Sub_BS_D_No			= RS1("BOM_Sub_BS_D_No")
		
		BOM_Sub_BS_D_No_1		= RS1("BOM_Sub_BS_D_No_1")
		BOM_Sub_BS_D_No_2		= RS1("BOM_Sub_BS_D_No_2")
		BOM_Sub_BS_D_No_3		= RS1("BOM_Sub_BS_D_No_3")
		BOM_Sub_BS_D_No_4		= RS1("BOM_Sub_BS_D_No_4")
		
		if LP_IMD_Complete_QTY = "Y" then
			LP_IMD_Complete_QTY = "V"
		end if
		if LP_SMD_Complete_QTY = "Y" then
			LP_SMD_Complete_QTY = "V"
		end if
		if LP_MAN_Complete_QTY = "Y" then
			LP_MAN_Complete_QTY = "V"
		end if
		if LP_ASM_Complete_QTY = "Y" then
			LP_ASM_Complete_QTY = "V"
		end if
		if LP_DLV_Complete_QTY = "Y" then
			LP_DLV_Complete_QTY = "V"
		end if
		
		BOM_Sub_BS_D_No_Str	 	= ""
		BOM_Sub_BS_D_No_Str = RS1("BOM_Sub_BS_D_No")
		if not(ISNULL(BOM_Sub_BS_D_No_Str)) then
			BOM_Sub_BS_D_No_Str = replace(BOM_Sub_BS_D_No_Str,"<br><br><br>","")
			BOM_Sub_BS_D_No_Str = replace(BOM_Sub_BS_D_No_Str,"<br><br>","")
		
			if right(BOM_Sub_BS_D_No_Str,4) = "<br>" then
				BOM_Sub_BS_D_No_Str = left(BOM_Sub_BS_D_No_Str,len(BOM_Sub_BS_D_No_Str)-4)
			end if
			BOM_Sub_BS_D_No_Str = replace(BOM_Sub_BS_D_No_Str,"<br>",",")
		end if

		response.write LM_Company
		response.write vbtab
		response.write LP_Line
		response.write vbtab
		response.write LP_Work_Order
		response.write vbtab
		response.write LP_Model
		response.write vbtab
		response.write BOM_Sub_BS_D_No_Str
		response.write vbtab
		response.write LP_Suffix
		response.write vbtab
		response.write LP_Tool
		response.write vbtab
		response.write LP_Tool_Type
		response.write vbtab
		response.write LP_Input_Time
		response.write vbtab
		response.write LP_LOT
		response.write vbtab
		response.write LP_LOT_Remain
		response.write vbtab
		response.write LP_IMD_Complete_QTY
		response.write vbtab
		response.write LP_SMD_Complete_QTY
		response.write vbtab
		response.write LP_MAN_Complete_QTY
		response.write vbtab
		response.write LP_ASM_Complete_QTY
		response.write vbtab
		response.write LP_DLV_Complete_QTY
		response.write vbtab
	
		for CNT1 = 0 to s_diff_LPD_Input_Date
			LPD_Input_Qty = ""		
			LPD_Input_Qty = RS1("DATE_QTY_"&CNT1)
			
			response.write LPD_Input_Qty
			response.write vbtab
			
		next
		response.write vbcrlf
		RS1.MoveNext
	loop
end if
RS1.close
set RS1 = nothing
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->