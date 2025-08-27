<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->

<form name='frmErrorRedirect' action='b_list.asp' method=post></form>
<% 
rem 변수선언
dim SQL
dim RS1
dim UpLoad

dim B_Code
dim B_D_No
dim B_Version_Code
dim B_Version_Date
dim B_Version_Current_YN
dim B_Version_Current_YN_old
dim B_Class
dim B_Tool
dim B_Desc
dim B_Spec
dim B_File_1
dim B_File_2
dim B_File_3
dim B_File_4
dim B_State
dim B_Memo
dim B_Issue_Date
dim B_Reg_Date
dim B_Edit_Date

dim oldB_File_1
dim oldB_File_2
dim oldB_File_3
dim oldB_File_4
dim oldB_File_5
dim oldB_File_6

dim temp
dim strError
dim URL_Prev
dim URL_Next

Dim strDelete

rem 객체선언
Set RS1		= Server.CreateObject("ADODB.RecordSet")
Set UpLoad	= Server.CreateObject("Dext.FileUpLoad")

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

rem 업로드 될 물리적 경로지정
UpLoad.DefaultPath = DefaultPath_BOM
UpLoad.MaxFileLen = (1024 * 1024 * 10)

URL_Prev	= UpLoad("URL_Prev")
URL_Next	= UpLoad("URL_Next")

strDelete	= UpLoad("strDelete")

rem 업로드 될 파일 체크
if trim(UpLoad("B_File_1")) <> "" then
	if UpLoad("B_File_1").FileLen > UpLoad.MaxFileLen then '10메가 이하인지 체크
		strError = "파일1은 10메가까지 업로드 가능합니다.\n"
		%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("B_File_2")) <> "" then
	if UpLoad("B_File_2").FileLen > UpLoad.MaxFileLen then '10메가 이하인지 체크
		strError = "파일2는 10메가까지 업로드 가능합니다.\n"
		%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("B_File_3")) <> "" then
	if UpLoad("B_File_3").FileLen > UpLoad.MaxFileLen then '10메가 이하인지 체크
		strError = "파일3은 10메가까지 업로드 가능합니다.\n"
		%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("B_File_4")) <> "" then
	if UpLoad("B_File_4").FileLen > UpLoad.MaxFileLen then '10메가 이하인지 체크
		strError = "파일4는 10메가까지 업로드 가능합니다.\n"
		%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	B_File_1	= Trim(UpLoad("B_File_1"))
	oldB_File_1	= DefaultPath_BOM & Trim(UpLoad("oldB_File_1"))
	B_File_2	= Trim(UpLoad("B_File_2"))
	oldB_File_2	= DefaultPath_BOM & Trim(UpLoad("oldB_File_2"))
	B_File_3	= Trim(UpLoad("B_File_3"))
	oldB_File_3	= DefaultPath_BOM & Trim(UpLoad("oldB_File_3"))
	B_File_4	= Trim(UpLoad("B_File_4"))
	oldB_File_4	= DefaultPath_BOM & Trim(UpLoad("oldB_File_4"))

	If B_File_1 <> "" then
		If oldB_File_1 <> "" Then
			File_Delete(oldB_File_1)
		End If
		B_File_1 = UpLoad("B_File_1").Save(,False)

	Else 
		If oldB_File_1 <> "" Then
			If InStr(strDelete, "B_File_1") > 0 Then
				File_Delete(oldB_File_1)
				B_File_1 = ""
			Else 
				B_File_1 = oldB_File_1
			End If 
		Else 
			B_File_1 = ""
		End If
	End If 

	If B_File_2 <> "" then
		If oldB_File_2 <> "" Then
			File_Delete(oldB_File_2)
		End If
		B_File_2 = UpLoad("B_File_2").Save(,False)
	Else 
		If oldB_File_2 <> "" Then
			If InStr(strDelete, "B_File_2") > 0 Then
				File_Delete(oldB_File_2)
				B_File_2 = ""
			Else 
				B_File_2 = oldB_File_2
			End If 
		Else 
			B_File_2 = ""
		End If
	End If 

	If B_File_3 <> "" then
		If oldB_File_3 <> "" Then
			File_Delete(oldB_File_3)
		End If
		B_File_3 = UpLoad("B_File_3").Save(,False)
	Else 
		If oldB_File_3 <> "" Then
			If InStr(strDelete, "B_File_3") > 0 Then
				File_Delete(oldB_File_3)
				B_File_3 = ""
			Else 
				B_File_3 = oldB_File_3
			End If 
		Else 
			B_File_3 = ""
		End If
	End If 

	If B_File_4 <> "" then
		If oldB_File_4 <> "" Then
			File_Delete(oldB_File_4)
		End If
		B_File_4 = UpLoad("B_File_4").Save(,False)
	Else 
		If oldB_File_4 <> "" Then
			If InStr(strDelete, "B_File_4") > 0 Then
				File_Delete(oldB_File_4)
				B_File_4 = ""
			Else 
				B_File_4 = oldB_File_4
			End If 
		Else 
			B_File_4 = ""
		End If
	End If

	B_File_1 = Replace(B_File_1,DefaultPath_BOM,"")
	B_File_2 = Replace(B_File_2,DefaultPath_BOM,"")
	B_File_3 = Replace(B_File_3,DefaultPath_BOM,"")
	B_File_4 = Replace(B_File_4,DefaultPath_BOM,"")

	B_Code					= UpLoad("B_Code")
	B_D_No					= UpLoad("B_D_No")
	B_Version_Code			= trim(Upload("B_Version_Code"))
	B_Version_Current_YN	= Upload("B_Version_Current_YN")
	B_Version_Current_YN_old= Upload("B_Version_Current_YN_old")
	B_Version_Date			= Upload("B_Version_Date")
	
	B_Memo			= UpLoad("B_Memo")
	B_Issue_Date	= UpLoad("B_Issue_Date")
	B_Edit_Date		= date()
	B_Class			= Trim(UpLoad("B_Class"))
	B_Tool			= Trim(UpLoad("B_Tool"))
	B_Desc			= Trim(UpLoad("B_Desc"))
	B_Spec			= Trim(UpLoad("B_Spec"))
	
	SQL = "update tbBOM set "
	SQL = SQL & "B_Version_Code = '"&B_Version_Code&"', "
	SQL = SQL & "B_Version_Current_YN = '"&B_Version_Current_YN&"', "
	SQL = SQL & "B_Version_Date = '"&B_Version_Date&"', "
	SQL = SQL & "B_Class = '"&B_Class&"', "
	SQL = SQL & "B_Tool = '"&B_Tool&"', "
	SQL = SQL & "B_Desc = '"&B_Desc&"', "
	SQL = SQL & "B_Spec = '"&B_Spec&"', "
	SQL = SQL & "B_File_1 = '"&B_File_1&"', "
	SQL = SQL & "B_File_2 = '"&B_File_2&"', "
	SQL = SQL & "B_File_3 = '"&B_File_3&"', "
	SQL = SQL & "B_File_4 = '"&B_File_4&"', "
	SQL = SQL & "B_Memo = '"&B_Memo&"', "
	SQL = SQL & "B_Edit_Date = '"&B_Edit_Date&"', "
	SQL = SQL & "B_Issue_Date = '"&B_Issue_Date&"' "
	SQL = SQL & "where B_Code = "&B_Code
	sys_DBCon.execute(SQL)
	
end if


'BOM LEVEL - 현재적용이 N에서 Y로 변경시
if B_Version_Current_YN = "Y" and B_Version_Current_YN_old = "N" then
	'tbBOM_Qty_Archive에서 tbBOM_Qty로 이동
	call moveBOM_Qty(B_Code, "N_to_Y")

	call BOM_Level_Reset(B_Code)
	
	call regist_BOM_Mask(B_Code)
end if

if B_Version_Current_YN = "Y" then
	SQL = "select B_Code from tbBOM where B_D_No = '"&B_D_No&"' and B_Code <> "&B_Code
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		'tbBOM_Qty에서 tbBOM_Qty_Archive로 이동
		call moveBOM_Qty(RS1("B_Code"), "Y_to_N")
		RS1.MoveNext
	loop
	RS1.Close
	
	SQL = "update tbBOM set B_Version_Current_YN = 'N' where B_D_No = '"&B_D_No&"' and B_Code <> "&B_Code
	sys_DBCon.execute(SQL)
end if

'객체 해제
set UpLoad	= nothing
Set RS1		= nothing
%>

<%
sub moveBOM_Qty(B_Code, strMove)
	dim SQL
	dim strTableFrom
	dim strTableTo
	
	if strMove = "N_to_Y" then
		strTableFrom= "tbBOM_Qty_Archive"
		strTableTo	= "tbBOM_Qty"
	elseif strMove = "Y_to_N" then
		strTableFrom= "tbBOM_Qty"
		strTableTo	= "tbBOM_Qty_Archive"
	end if
	
	'tbBOM_Qty_Archive에서 tbBOM_Qty로 복사
	SQL = "insert into "&strTableTo&" select " 
	SQL = SQL & B_Code&", "
	SQL = SQL & "BOM_Sub_BS_D_No, "
	SQL = SQL & "BOM_Sub_BS_Code, "
	SQL = SQL & "Parts_P_P_No, "
	SQL = SQL & "Parts_P_P_No2, "
	SQL = SQL & "Parts_P_P_No2_PinYN, "
	SQL = SQL & "BQ_Qty, "
	SQL = SQL & "BQ_Use_YN, "
	SQL = SQL & "BQ_Order, "
	SQL = SQL & "BQ_Remark, "
	SQL = SQL & "BQ_CheckSum, "
	SQL = SQL & "BQ_P_Desc, "
	SQL = SQL & "BQ_P_Spec, "
	SQL = SQL & "BQ_P_Maker, "
	SQL = SQL & "BOM_B_D_No "
	SQL = SQL & "from "&strTableFrom&" where BOM_B_Code = "&B_Code&" order by BQ_Code asc"
	sys_DBCon.execute(SQL)
	
	'tbBOM_Qty_Archive에서 삭제
	SQL = "delete from "&strTableFrom&" where BOM_B_Code = "&B_Code
	sys_DBCon.execute(SQL)
end sub
%>

<%
sub BOM_Level_Reset(B_Code)
	dim RS1
	
	dim SQL
	dim B_Version_Code
	dim B_Version_Current_YN
			
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	
	'B코드로 버젼정보 조회
	SQL = "select B_Version_Code, B_Version_Current_YN from tbBOM where B_Code='"&B_Code&"'"
	RS1.Open SQL,sys_DBCon
	if not(RS1.Eof or RS1.Bof) then
		B_Version_Code = RS1("B_Version_Code")
		B_Version_Current_YN = RS1("B_Version_Current_YN")
	end if
	RS1.Close
	
	'현재적용중이라면
	if B_Version_Current_YN = "Y" then
		'bom을 루프돌면서 bs_code, bs_d_no 루핑
		SQL = "select BS_Code, BS_D_No from tbBOM_Sub where BOM_B_Code="&B_Code
		RS1.Open SQL,sys_DBCon
		do until RS1.Eof
	
			'기존 행 삭제
			SQL = "delete tblBOM_Level_Master where B_PARTNO_ASSY = '"&RS1("BS_D_No")&"'"
			sys_DBCon.execute(SQL)
			
			'재등록
			SQL = "insert into tblBOM_Level_Master (B_PARTNO_ASSY,B_LEVEL_READY_YN,B_LEVEL_Date) values ('"&RS1("BS_D_No")&"','N',getdate())"
			sys_DBCon.execute(SQL)
			
			RS1.MoveNext
		loop
		RS1.Close
	end if
	
	set RS1 = nothing
end sub
%>

<%
sub regist_BOM_Mask(B_Code)
	dim RS1
	dim RS2
	dim SQL
	
	dim BQI_SType
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
	
	SQL = "select distinct Parts_P_P_No, BQ_P_Desc, BQ_P_Spec, BQ_P_Maker, BQ_Qty from tbBOM_Qty where BOM_B_Code = "&B_Code&" and BQ_Qty > 0 "
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		if isnumeric(RS1("BQ_Qty")) then
			if RS1("BQ_Qty") > 0 then
				
				SQL = "select top 1 BQI_SType from tbBOM_Qty_Info where Parts_P_P_No = '"&RS1("Parts_P_P_No")&"'"
				RS2.Open SQL,sys_DBCon
				if RS2.Eof or RS2.Bof then
					BQI_SType = ""
				else
					BQI_SType = RS2("BQI_SType")
				end if
				RS2.Close
				
				SQL = "insert into tblBOM_Mask (BOM_Parts_BP_PNO, BM_Filter, RegDate, EditDate, M_ID,BM_SType_BOM,BM_Desc_BOM,BM_Spec_BOM,BM_Maker_BOM) values "
				SQL = SQL & "('"&RS1("Parts_P_P_No")&"','_', getdate(), getdate(), '"&gM_ID&"','"&BQI_SType&"','"&RS1("BQ_P_Desc")&"','"&RS1("BQ_P_Spec")&"','"&RS1("BQ_P_Maker")&"')"
				on error resume next
				sys_DBCon.execute(SQL)
				on error goto 0
			end if
		end if
		
		RS1.MoveNext
	loop
	RS1.Close
	set RS1 = nothing
end sub
%>

<form name="frmRedirect" action="<%=URL_Next%>" method=post>
<input type="hidden" name="B_Code" value="<%=B_Code%>">
<input type="hidden" name="B_D_No" value="<%=B_D_No%>">

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
alert("수정이 완료되었습니다.");
frmRedirect.submit();
</script>



<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->