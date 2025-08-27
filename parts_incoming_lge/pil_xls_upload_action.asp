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

dim Exist_YN

set UpLoad	= Server.CreateObject("Dext.FileUpLoad")
UpLoad.DefaultPath = DefaultPath_SAGUP_XLS_Reader

strFile 	= UpLoad("strFile")
arrFile		= split(strFile,"\")
File_Name	= lcase(arrFile(ubound(arrFile)))

temp = UpLoad("strFile").SaveAs(DefaultPath_SAGUP_XLS_Reader & File_Name, False)
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

if instr(File_Name,"osp50") > 0 then
	call XLS_Upload()
end if
%>
<form name="frmRedirect" action="pil_list.asp" method="post">
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
sub XLS_Upload()
	dim RS1
	dim RS2
	dim SQL
	
	dim P_Code
	
	dim PIL_Code
	dim Parts_P_P_No
	dim Parts_P_Desc
	dim Parts_P_Spec
	dim PIL_In_Date
	dim PIL_Qty
	dim PIL_Price
	dim PIL_WorkOrder
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
			
	for CNT1 = 2 to ubound(arrXLS, 2)
			
		Parts_P_P_No	= trim(arrXLS(1, CNT1))
		Parts_P_Desc	= trim(arrXLS(12, CNT1))
		Parts_P_Spec	= trim(arrXLS(13, CNT1))
		PIL_In_Date		= trim(arrXLS(7, CNT1))
		PIL_In_Date		= left(PIL_In_Date,4) &"-"& mid(PIL_In_Date,5,2) &"-"&right(PIL_In_Date,2)
		PIL_Qty			= trim(arrXLS(8, CNT1))
		PIL_Price		= trim(arrXLS(9, CNT1))
		PIL_WorkOrder	= trim(arrXLS(11, CNT1))
		
		'파츠 DB에 있는지 조회
		SQL = "select top 1 P_Code from tbParts where P_P_No = '"&Parts_P_P_No&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then '없다면
			'파트넘버 추가
			SQL = 		"insert into tbParts "&vbcrlf
			SQL = SQL & "(P_P_No,P_Desc,P_Spec,P_Maker,P_CO_Price,P_LGE_Price,P_MSE_Price,Partner_P_Code,Partner_P_Name,P_Work_Type,P_Qty,P_Safe_Qty,P_Same_Code) "&vbcrlf
			SQL = SQL & "values "&vbcrlf
			SQL = SQL & "('"&Parts_P_P_No&"','"&Parts_P_Desc&"','"&Parts_P_Spec&"','',0,0,0,'','','',0,0,-1) "&vbcrlf
			sys_DBCon.execute(SQL)
			
			'추가한 파트넘버의 일련번호 조회
			SQL = "select top 1 P_Code from tbParts where P_P_No = '"&Parts_P_P_No&"'"
			RS2.Open SQL,sys_DBCon
			P_Code = RS2("P_Code")
			RS2.Close
		end if
		RS1.Close
		
		'동일한 일련번호의 사급가 및 인증가를 업데이트
		SQL = "update tbParts set P_CO_Price = '"&PIL_Price&"', P_LGE_Price = '"&PIL_Price&"' where P_Code = '"&P_Code&"'"
		sys_DBCon.execute(SQL)
		
		'사급출고일과 파트넘버, 워크오더까지 일치하는거 조회	
		SQL = "select PIL_Qty from tbParts_Incoming_LGE where Parts_P_P_No = '"&Parts_P_P_No&"' and PIL_In_Date = '"&PIL_In_Date&"' and PIL_WorkOrder = '"&PIL_WorkOrder&"'"
		RS1.Open SQL,sys_DBCon
		'없다면
		if RS1.Eof or RS1.Bof then
			'등록
			SQL = "insert into tbParts_Incoming_LGE (Parts_P_P_No, PIL_In_Date, PIL_Price, PIL_Qty, PIL_WorkOrder) "
			SQL = SQL & "values ('"&Parts_P_P_No&"','"&PIL_In_Date&"',"&PIL_Price&","&PIL_Qty&",'"&PIL_WorkOrder&"')"
			sys_DBCon.execute(SQL)
			
			'자재 수량에 반영 (수량은 중복 반영되면 안되기 때문에)
			SQL = "update tbParts set P_Qty = P_Qty + "&PIL_Qty&" where P_P_No = '"&Parts_P_P_No&"'"
			sys_DBCon.execute(SQL)
			
			SQL = "insert tbParts_Transaction (Parts_P_P_No,PT_Qty,PT_Type,PT_Description,Member_M_ID,PT_Date) values ('"&Parts_P_P_No&"',"&PIL_Qty*-1&",'사급입고','재고변경','"&gM_ID&"','"&PIL_In_Date&"')"
			sys_DBCon.execute(SQL)
		'있다면
		else
			SQL = "update tbParts_Incoming_LGE set "
			SQL = SQL & "PIL_Price = "&PIL_Price&","
			SQL = SQL & "PIL_Qty = "&PIL_Qty&" where Parts_P_P_No = '"&Parts_P_P_No&"' and PIL_In_Date = '"&PIL_In_Date&"' and PIL_WorkOrder = '"&PIL_WorkOrder&"'"
			sys_DBCon.execute(SQL)
			
			'자재 수량에 반영했던 건, 수정
			SQL = "update tbParts set P_Qty = P_Qty - "&PIL_Qty&" + "&RS1("PIL_Qty")&" where P_P_No = '"&Parts_P_P_No&"'"
			sys_DBCon.execute(SQL)
			
			if PIL_Qty - RS1("PIL_Qty") <> 0 then
				SQL = "insert tbParts_Transaction (Parts_P_P_No,PT_Qty,PT_Type,PT_Description,Member_M_ID,PT_Date) values ('"&Parts_P_P_No&"',"&PIL_Qty*-1 - RS1("PIL_Qty")*-1&",'사급입고','재고변경','"&gM_ID&"','"&PIL_In_Date&"')"
				sys_DBCon.execute(SQL)
			end if
		end if
		RS1.Close
	next
	
	set RS1 = nothing
	set RS2 = nothing
end sub
%>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->


