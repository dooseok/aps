<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
'변수선언
dim CNT1
dim CNT2
dim UpLoad
dim objXLS
dim XLSConnection
dim temp
dim strProperties

dim BP_Division

dim RS1
dim SQL
dim Sheet_Name
dim arrXLS

dim strFile
dim arrFile
dim File_Name

dim strDesc
dim Exist_YN

dim strError
set UpLoad	= Server.CreateObject("Dext.FileUpLoad")
UpLoad.DefaultPath = DefaultPath_BOM_Update
strFile 	= UpLoad("strFile")
arrFile		= split(strFile,"\")
File_Name	= lcase(arrFile(ubound(arrFile)))

'파일명 체크
if instr("-CLZ-DGZ-DMZ-SRJ-","-"&left(ucase(File_Name),3)&"-") > 0 then
	BP_Division			= left(ucase(File_Name),3)
else
	strError			= "파일명을 사업부로 변경하여 주십시오. 예) DGZ.xls"
end if


temp = UpLoad("strFile").SaveAs(DefaultPath_BOM_Update & File_Name, False)
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
if strError = "" then
	call XLS_Upload()
%>
<form name="frmRedirect" action="b_price_list.asp" method="post">
</form>
<script language="javascript">
alert("<%=left(ucase(File_Name),3)%> 판가 업데이트가 완료되었습니다.");
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="b_price_list.asp" method=post>
</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
objXLS.close
set objXLS = nothing
set UpLoad = nothing
%>

<%
sub XLS_Upload()
	dim RS1
	dim RS2
	dim SQL
	
	dim BOM_Sub_BS_D_No
	dim BP_Market
	dim BP_Currency
	dim BP_Price
	dim BP_Creation_Date
	dim BP_Start_Date
	dim BP_End_Date
	dim BP_Type
	dim BP_Gap
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
	
	'SQL = "delete from tbBOM_Price"
	'sys_DBCon.execute(SQL)
	'SQL = "dbcc checkident('tbBOM_Price',reseed,0)"
	'sys_DBCon.execute(SQL)
			
	for CNT1 = 2 to ubound(arrXLS, 2)

		BOM_Sub_BS_D_No		= trim(arrXLS(1, CNT1))
		BP_Market			= trim(arrXLS(3, CNT1))
		BP_Currency			= trim(arrXLS(6, CNT1))
		BP_Price			= replace(trim(arrXLS(7, CNT1)),",","")
		BP_Creation_Date	= trim(arrXLS(9, CNT1))
		BP_Start_Date		= trim(arrXLS(10, CNT1))
		BP_End_Date			= trim(arrXLS(11, CNT1))		

		SQL = "select top 1 BP_Price from tbBOM_Price where "
		SQL = SQL & "BP_Division		= '"&BP_Division&"' 		and "
		SQL = SQL & "BOM_Sub_BS_D_No	= '"&BOM_Sub_BS_D_No&"' 	and "
		SQL = SQL & "BP_Market			= '"&BP_Market&"' 			and "
		SQL = SQL & "BP_Currency		= '"&BP_Currency&"' 		and "
		SQL = SQL & "BP_Price			= "&BP_Price&"		 		and "
		SQL = SQL & "BP_Creation_Date	= '"&BP_Creation_Date&"' 	and "
		SQL = SQL & "BP_Start_Date		= '"&BP_Start_Date&"' 		and "
		SQL = SQL & "BP_End_Date		= '"&BP_End_Date&"' "
		SQL = SQL & "order by BP_Code desc"
		
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			SQL = "select top 1 * from tbBOM_Price where "
			SQL = SQL & "BP_Division		= '"&BP_Division&"' 		and "
			SQL = SQL & "BOM_Sub_BS_D_No	= '"&BOM_Sub_BS_D_No&"' 	and "
			SQL = SQL & "BP_Market			= '"&BP_Market&"' 			and "
			SQL = SQL & "BP_Currency		= '"&BP_Currency&"' order by BP_Code desc"
			RS2.Open SQL,sys_DBCon
			if RS2.Eof or RS2.Bof then
				BP_Type = "신규"
				BP_Gap = 0
			else
				BP_Gap = BP_Price-RS2("BP_Price")
				if BP_Gap > 0 then
					BP_Type = "인상"
				elseif BP_Gap = 0 then
					BP_Type = "기타"
				else
					BP_Type = "인하"
				end if
			end if
			RS2.Close
			SQL = "insert into tbBOM_Price (BP_Division,BOM_Sub_BS_D_No,BP_Market,BP_Currency,BP_Type,BP_Gap,BP_Price,BP_Creation_Date,BP_Start_Date,BP_End_Date,BP_Update_Date,BP_Desc) values "
			SQL = SQL & "('"&BP_Division&"','"&BOM_Sub_BS_D_No&"','"&BP_Market&"','"&BP_Currency&"','"&BP_Type&"',"&BP_Gap&","&BP_Price&",'"&BP_Creation_Date&"','"&BP_Start_Date&"','"&BP_End_Date&"','"&date()&"','')"
			sys_DBCon.execute(SQL)
		end if
		RS1.Close
		
		SQL = "select top 1 BSP_D_No from tbBOM_Sub_Priced where BSP_D_No='"&arrXLS(1, CNT1)&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			if isnull(arrXLS(12, CNT1)) then
				strDesc = ""
			else
				strDesc = replace(arrXLS(12, CNT1),"'","''")
			end if
			SQL = "insert into tbBOM_Sub_Priced (BSP_D_No,BSP_Desc) values ('"&arrXLS(1, CNT1)&"','"&strDesc&"')"
			sys_DBCon.execute(SQL)
		end if
		RS1.Close

	next
	'MSE 제품이 아닌 가격 정보는 삭제
	'SQL = "delete from tbBOM_Price where BOM_Sub_BS_D_No not in (select BS_D_No from tbBOM_Sub)"
	'sys_DBCon.execute(SQL)
	
	SQL = "insert into tbMacro_Log (ML_Item,ML_UploadDate) values ('Price_"&BP_Division&"',getdate())"
	sys_DBCon.execute(SQL)
	
	set RS1 = nothing
	set RS2 = nothing
end sub
%>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
