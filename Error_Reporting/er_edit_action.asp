<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem 변수선언
dim SQL
dim RS1
dim UpLoad

dim ER_Code
dim ER_Title
dim ER_Content
dim ER_Edit_Date
dim ER_File_1
dim ER_File_2
dim ER_File_3
dim Member_M_ID

dim oldER_File_1
dim oldER_File_2
dim oldER_File_3

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
UpLoad.DefaultPath = DefaultPath_error_reporting

URL_Prev	= UpLoad("URL_Prev")
URL_Next	= UpLoad("URL_Next")

strDelete	= UpLoad("strDelete")

SQL = "select Member_M_ID from tberror_reporting where ER_Code = '"&UpLoad("ER_Code")&"'"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	strError = strError & "*작성자정보를 회원DB에서 찾을 수 없습니다.\n*관리자에게 문의하여 주십시오.\n"
else
	if lcase(RS1("Member_M_ID")) <> lcase(gM_ID) then
		strError = strError & "*작성자 본인이 요청내용을 수정할 수 있습니다.\n"
	end if
end if
RS1.Close

rem 업로드 될 파일 체크
if trim(UpLoad("ER_File_1")) <> "" then
	if UpLoad("ER_File_1").FileLen > (1024 * 1024 * 10) then '10메가 이하인지 체크
		strError = "파일1은 10메가까지 업로드 가능합니다.\n"
	end if
end if
if trim(UpLoad("ER_File_2")) <> "" then
	if UpLoad("ER_File_2").FileLen > (1024 * 1024 * 10) then '10메가 이하인지 체크
		strError = "파일2은 10메가까지 업로드 가능합니다.\n"
	end if
end if
if trim(UpLoad("ER_File_3")) <> "" then
	if UpLoad("ER_File_3").FileLen > (1024 * 1024 * 10) then '10메가 이하인지 체크
		strError = "파일3은 10메가까지 업로드 가능합니다.\n"
	end if
end if

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	ER_File_1	= Trim(UpLoad("ER_File_1"))
	oldER_File_1	= DefaultPath_error_reporting & Trim(UpLoad("oldER_File_1"))
	ER_File_2	= Trim(UpLoad("ER_File_2"))
	oldER_File_2	= DefaultPath_error_reporting & Trim(UpLoad("oldER_File_2"))
	ER_File_3	= Trim(UpLoad("ER_File_3"))
	oldER_File_3	= DefaultPath_error_reporting & Trim(UpLoad("oldER_File_3"))
	
	If ER_File_1 <> "" then
		If oldER_File_1 <> "" Then
			File_Delete(oldER_File_1)
		End If
		ER_File_1 = UpLoad("ER_File_1").Save(,False)

	Else 
		If oldER_File_1 <> "" Then
			If InStr(strDelete, "ER_File_1") > 0 Then
				File_Delete(oldER_File_1)
				ER_File_1 = ""
			Else 
				ER_File_1 = oldER_File_1
			End If 
		Else 
			ER_File_1 = ""
		End If
	End If 

	If ER_File_2 <> "" then
		If oldER_File_2 <> "" Then
			File_Delete(oldER_File_2)
		End If
		ER_File_2 = UpLoad("ER_File_2").Save(,False)
	Else 
		If oldER_File_2 <> "" Then
			If InStr(strDelete, "ER_File_2") > 0 Then
				File_Delete(oldER_File_2)
				ER_File_2 = ""
			Else 
				ER_File_2 = oldER_File_2
			End If 
		Else 
			ER_File_2 = ""
		End If
	End If 

	If ER_File_3 <> "" then
		If oldER_File_3 <> "" Then
			File_Delete(oldER_File_3)
		End If
		ER_File_3 = UpLoad("ER_File_3").Save(,False)
	Else 
		If oldER_File_3 <> "" Then
			If InStr(strDelete, "ER_File_3") > 0 Then
				File_Delete(oldER_File_3)
				ER_File_3 = ""
			Else 
				ER_File_3 = oldER_File_3
			End If 
		Else 
			ER_File_3 = ""
		End If
	End If 

	ER_Code		= UpLoad("ER_Code")
	ER_Title		= UpLoad("ER_Title")
	ER_Content	= UpLoad("ER_Content")
	ER_Edit_Date	= now()
	ER_File_1 	= Replace(lcase(ER_File_1),DefaultPath_error_reporting,"")
	ER_File_2 	= Replace(lcase(ER_File_2),DefaultPath_error_reporting,"")
	ER_File_3 	= Replace(lcase(ER_File_3),DefaultPath_error_reporting,"")
	
	rem DB 업데이트
	SQL = "select * from tberror_reporting where ER_Code = '"&ER_Code&"'"
	RS1.Open SQL,sys_DBconString,3,2,&H0001
	with RS1		
		.Fields("ER_Title")		= ER_Title
		.Fields("ER_Content")	= ER_Content
		.Fields("ER_Edit_Date")	= ER_Edit_Date
		.Fields("ER_File_1")		= ER_File_1
		.Fields("ER_File_2")		= ER_File_2
		.Fields("ER_File_3")		= ER_File_3
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
<input type="hidden" name="ER_Code" value="<%=ER_Code%>">

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="<%=URL_Prev%>" method=post>
<input type="hidden" name="B_Code" value="<%=B_Code%>">

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