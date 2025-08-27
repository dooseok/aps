<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
 
<%
rem 변수선언
dim SQL
dim RS1
dim UpLoad

dim BU_Code
dim BOM_B_D_No
dim BU_Content
dim BU_Receive_Date
dim BU_Apply_Date
dim BU_Reply_Date
dim BU_Request_Reply_Date
dim BU_File_1
dim BU_File_2
dim BU_File_3
dim Member_M_ID
dim BU_Type_New
dim BU_Type_Add
dim BU_Type_Update
dim BU_Type

dim arrBU_Code
dim cntBU_Code

dim temp
dim strError
dim URL_Prev
dim URL_Next

rem 객체선언
Set RS1		= Server.CreateObject("ADODB.RecordSet")
Set UpLoad	= Server.CreateObject("Dext.FileUpLoad")

rem 업로드 될 물리적 경로지정
UpLoad.DefaultPath = DefaultPath_BOM_Update

URL_Prev	= UpLoad("URL_Prev")
URL_Next	= UpLoad("URL_Next")

rem 업로드 될 파일 체크
if trim(UpLoad("BU_File_1")) <> "" then
	if UpLoad("BU_File_1").FileLen > (1024 * 1024 * 10) then '10메가 이하인지 체크
		strError = "파일1은 10메가까지 업로드 가능합니다.\n"
	end if
end if
if trim(UpLoad("BU_File_2")) <> "" then
	if UpLoad("BU_File_2").FileLen > (1024 * 1024 * 10) then '10메가 이하인지 체크
		strError = "파일2은 10메가까지 업로드 가능합니다.\n"
	end if
end if
if trim(UpLoad("BU_File_3")) <> "" then
	if UpLoad("BU_File_3").FileLen > (1024 * 1024 * 10) then '10메가 이하인지 체크
		strError = "파일3은 10메가까지 업로드 가능합니다.\n"
	end if
end if


rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

rem 업로드작업
	if trim(UpLoad("BU_File_1")) <> "" then
		BU_File_1 = UpLoad("BU_File_1").Save(,False)
	end if
	if trim(UpLoad("BU_File_2")) <> "" then
		BU_File_2 = UpLoad("BU_File_2").Save(,False)
	end if
	if trim(UpLoad("BU_File_3")) <> "" then
		BU_File_3 = UpLoad("BU_File_3").Save(,False)
	end if
	
	BOM_B_D_No		= Trim(UpLoad("BOM_B_D_No"))
	BU_Content		= Trim(UpLoad("BU_Content"))
	BU_Receive_Date	= Trim(UpLoad("BU_Receive_Date"))
	BU_Apply_Date	= Trim(UpLoad("BU_Apply_Date"))
	BU_Reply_Date	= Trim(UpLoad("BU_Reply_Date"))
	BU_Request_Reply_Date	= Trim(UpLoad("BU_Request_Reply_Date"))
	BU_File_1		= Replace(BU_File_1,DefaultPath_BOM_Update,"")
	BU_File_2		= Replace(BU_File_2,DefaultPath_BOM_Update,"")
	BU_File_3		= Replace(BU_File_3,DefaultPath_BOM_Update,"")
	Member_M_ID		= gM_ID
	BU_Type_New	= Trim(UpLoad("BU_Type_New"))
	BU_Type_Add	= Trim(UpLoad("BU_Type_Add"))
	BU_Type_Update	= Trim(UpLoad("BU_Type_Update"))
	
	if BU_Type_New = "Y" then
		BU_Type = BU_Type & "신규-"
	end if
	if BU_Type_Add = "Y" then
		BU_Type = BU_Type & "추가-"
	end if
	if BU_Type_Update = "Y" then
		BU_Type = BU_Type & "시방-"
	end if

	SQL = "select max(convert(integer,right(bu_code,3))) from tbBOM_Update where '20'+substring(BU_Code,3,5)+'-'+substring(BU_Code,8,2) = '"&date()&"'"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		cntBU_Code = 0
	elseif isnull(RS1(0)) then
		cntBU_Code = 0
	elseif not(isnumeric(RS1(0))) then
		cntBU_code = 0
	else	
		cntBU_Code = RS1(0)
	end if
	RS1.Close
	cntBU_Code = cntBU_Code + 1
	if len(cntBU_Code)=1 then
		cntBU_Code = "00" & cntBU_Code
	elseif len(cntBU_Code)=2 then
		cntBU_Code = "0" & cntBU_Code
	end if
	
	BU_Code = "MS" & mid(date(),3,5) & right(date(),2) & "-" & cntBU_Code
	rem DB 업데이트
	RS1.Open "tbBOM_Update",sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("BU_Code")			= BU_Code
		.Fields("BOM_B_D_No")		= BOM_B_D_No
		.Fields("BU_Content")		= BU_Content
		If BU_Receive_Date <> "" then
			.Fields("BU_Receive_Date")	= BU_Receive_Date
		End If
		If BU_Apply_Date <> "" then
			.Fields("BU_Apply_Date")	= BU_Apply_Date
		End If
		If BU_Reply_Date <> "" then
			.Fields("BU_Reply_Date")	= BU_Reply_Date
		End If
		If BU_Request_Reply_Date <> "" then
			.Fields("BU_Request_Reply_Date")	= BU_Request_Reply_Date
		End if
		.Fields("BU_File_1")		= BU_File_1
		.Fields("BU_File_2")		= BU_File_2
		.Fields("BU_File_3")		= BU_File_3
		.Fields("BU_Type")			= BU_Type
		.Fields("Member_M_ID")		= Member_M_ID
		
		.Update
		.Close
	end with
	
	rem DB 업데이트
	RS1.Open "tbNotice",sys_DBConString,3,2,2
	with RS1
		.AddNew
		if BOM_B_D_No <> "" then
			.Fields("N_Title")			= BOM_B_D_No & "에 대한 시방이 등록되었습니다."
			.Fields("N_Content")		= BU_Content
		else
			.Fields("N_Title")			= "시방이 등록되었습니다."
			.Fields("N_Content")		= BU_Content
		end if
		
		.Fields("N_Reg_Date")		= date()
		.Fields("N_Edit_Date")		= date()
		.Fields("N_File_1")			= ""
		.Fields("N_File_2")			= ""
		.Fields("N_File_3")			= ""
		.Fields("Member_M_ID")		= Member_M_ID
		
		.Update
		.Close
	end with
end if



rem 객체 해제
set UpLoad	= nothing
Set RS1		= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="<%=URL_Next%>" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="<%=URL_Prev%>" method=post>

</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>

<!-- #include Virtual = "/header/db_tail.asp" -->