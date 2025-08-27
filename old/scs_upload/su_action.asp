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

dim strProperties

dim RS1
dim SQL
dim Sheet_Name

dim temp1
dim arrXLS1
dim strFile1
dim arrFile1
dim File_Name1
dim temp2
dim arrXLS2
dim strFile2
dim arrFile2
dim File_Name2
dim temp3
dim arrXLS3
dim strFile3
dim arrFile3
dim File_Name3
dim temp4
dim arrXLS4
dim strFile4
dim arrFile4
dim File_Name4
dim temp5
dim arrXLS5
dim strFile5
dim arrFile5
dim File_Name5

dim Exist_YN

set UpLoad	= Server.CreateObject("Dext.FileUpLoad")
UpLoad.DefaultPath = DefaultPath_SCS_XLS_Reader

strFile1 		= UpLoad("strFile1")
strFile2 		= UpLoad("strFile2")
strFile3 		= UpLoad("strFile3")
strFile4 		= UpLoad("strFile4")
strFile5 		= UpLoad("strFile5")

set objXLS = Server.CreateObject("ADODB.Connection") 

if strFile1 <> "" then
	arrFile1		= split(strFile1,"\")
	File_Name1	= lcase(arrFile1(ubound(arrFile1)))
	
	temp1 = UpLoad("strFile1").SaveAs(DefaultPath_SCS_Upload_Update & File_Name1, False)
	temp1 = replace(temp1,"\","/")
	
	XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp1 & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""
	objXLS.Open XLSConnection    
	set RS1 = objXLS.OpenSchema(20)
	Sheet_Name	= RS1(2)
	Sheet_Name	= "["&replace(Sheet_Name,"'","")&"]"
	set RS1 = nothing
	
	set RS1 = Server.CreateObject("ADODB.RecordSet") 
	SQL  = " select * from "&Sheet_Name
	RS1.Open SQL,objXLS 
	arrXLS1 = RS1.getRows()
	RS1.close
	set RS1 = nothing
	
	call XLS_Upload1(arrXLS1)
	objXLS.close
end if

if strFile2 <> "" then
	arrFile2		= split(strFile2,"\")
	File_Name2	= lcase(arrFile2(ubound(arrFile2)))
	
	temp2 = UpLoad("strFile2").SaveAs(DefaultPath_SCS_Upload_Update & File_Name2, False)
	temp2 = replace(temp2,"\","/")

	XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp2 & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""
	objXLS.Open XLSConnection    
	set RS1 = objXLS.OpenSchema(20)
	Sheet_Name	= RS1(2)
	Sheet_Name	= "["&replace(Sheet_Name,"'","")&"]"
	set RS1 = nothing
	
	set RS1 = Server.CreateObject("ADODB.RecordSet") 
	SQL  = " select * from "&Sheet_Name
	RS1.Open SQL,objXLS 
	arrXLS2 = RS1.getRows()
	RS1.close
	set RS1 = nothing
	
	call XLS_Upload2(arrXLS2)
	objXLS.close
end if

if strFile3 <> "" then

	arrFile3		= split(strFile3,"\")
	File_Name3	= lcase(arrFile3(ubound(arrFile3)))
	
	temp3 = UpLoad("strFile3").SaveAs(DefaultPath_SCS_Upload_Update & File_Name3, False)
	temp3 = replace(temp3,"\","/")

	XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp3 & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""
	objXLS.Open XLSConnection    
	set RS1 = objXLS.OpenSchema(20)
	Sheet_Name	= RS1(2)
	Sheet_Name	= "["&replace(Sheet_Name,"'","")&"]"
	set RS1 = nothing
	
	set RS1 = Server.CreateObject("ADODB.RecordSet") 
	SQL  = " select * from "&Sheet_Name
	RS1.Open SQL,objXLS 
	arrXLS3 = RS1.getRows()
	RS1.close
	set RS1 = nothing
	
	call XLS_Upload3(arrXLS3)
	objXLS.close
end if

if strFile4 <> "" then

	arrFile4		= split(strFile4,"\")
	File_Name4	= lcase(arrFile4(ubound(arrFile4)))
	
	temp4 = UpLoad("strFile").SaveAs(DefaultPath_SCS_Upload_Update & File_Name4, False)
	temp4 = replace(temp4,"\","/")

	XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp4 & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""
	objXLS.Open XLSConnection    
	set RS1 = objXLS.OpenSchema(20)
	Sheet_Name	= RS1(2)
	Sheet_Name	= "["&replace(Sheet_Name,"'","")&"]"
	set RS1 = nothing
	
	set RS1 = Server.CreateObject("ADODB.RecordSet") 
	SQL  = " select * from "&Sheet_Name
	RS1.Open SQL,objXLS 
	arrXLS4 = RS1.getRows()
	RS1.close
	set RS1 = nothing
	
	call XLS_Upload4(arrXLS4)
	objXLS.close
end if

if strFile5 <> "" then

	arrFile5		= split(strFile5,"\")
	File_Name5	= lcase(arrFile5(ubound(arrFile5)))
	
	temp5 = UpLoad("strFile").SaveAs(DefaultPath_SCS_Upload_Update & File_Name5, False)
	temp5 = replace(temp5,"\","/")
	
	XLSConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp5 & "; Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""
	objXLS.Open XLSConnection    
	set RS1 = objXLS.OpenSchema(20)
	Sheet_Name	= RS1(2)
	Sheet_Name	= "["&replace(Sheet_Name,"'","")&"]"
	set RS1 = nothing
	
	set RS1 = Server.CreateObject("ADODB.RecordSet") 
	SQL  = " select * from "&Sheet_Name
	RS1.Open SQL,objXLS 
	arrXLS5 = RS1.getRows()
	RS1.close
	set RS1 = nothing
	
	call XLS_Upload5(arrXLS5)
	objXLS.close
end if

set objXLS = nothing
set UpLoad = nothing
%>

<form name="frmRedirect" action="su_form.asp" method="post">
</form>
<script language="javascript">
//frmRedirect.submit();
</script>

<%
sub XLS_Upload1(arrXLS)
	dim RS1
	dim RS2
	dim SQL
	
	SQL = "delete from tbSCS_PurchaseOrder"
	sys_DBCon.execute(SQL)
				
	for CNT1 = 2 to ubound(arrXLS, 2)
		SQL = "insert into tbSCS_PurchaseOrder values ("
		SQL = SQL & "'"&arrXLS()&"',"
		SQL = SQL & "'"&BOM_Sub_BS_D_No&"',"
		SQL = SQL & "'"&BOM_Sub_BS_D_No&"',"
		SQL = SQL & "'"&BOM_Sub_BS_D_No&"',"
		SQL = SQL & "'"&BOM_Sub_BS_D_No&"',"
		SQL = SQL & "'"&BOM_Sub_BS_D_No&"',"
		SQL = SQL & "'"&BOM_Sub_BS_D_No&"',"
		SQL = SQL & "'"&BOM_Sub_BS_D_No&"',"
		SQL = SQL & "'"&BOM_Sub_BS_D_No&"'"
		sys_DBCon.execute(SQL)
	next
	
	set RS1 = nothing
	set RS2 = nothing
end sub
%>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
