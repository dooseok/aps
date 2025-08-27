<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1
dim UpLoad

dim N_Code

dim N_Title
dim N_Content
dim N_Reg_Date
dim N_Edit_Date
dim N_File_1
dim N_File_2
dim N_File_3
dim Member_M_ID

dim temp
dim strError
dim URL_Prev
dim URL_Next

rem 객체선언
Set RS1		= Server.CreateObject("ADODB.RecordSet")
Set UpLoad	= Server.CreateObject("Dext.FileUpLoad")

rem 업로드 될 물리적 경로지정
UpLoad.DefaultPath = DefaultPath_Notice

URL_Prev	= UpLoad("URL_Prev")
URL_Next	= UpLoad("URL_Next")

rem 업로드 될 파일 체크
if trim(UpLoad("N_File_1")) <> "" then
	if UpLoad("N_File_1").FileLen > (1024 * 1024 * 10) then '10메가 이하인지 체크
		strError = "파일1은 10메가까지 업로드 가능합니다.\n"
	end if
end if
if trim(UpLoad("N_File_2")) <> "" then
	if UpLoad("N_File_2").FileLen > (1024 * 1024 * 10) then '10메가 이하인지 체크
		strError = "파일2은 10메가까지 업로드 가능합니다.\n"
	end if
end if
if trim(UpLoad("N_File_3")) <> "" then
	if UpLoad("N_File_3").FileLen > (1024 * 1024 * 10) then '10메가 이하인지 체크
		strError = "파일3은 10메가까지 업로드 가능합니다.\n"
	end if
end if


rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

rem 업로드작업
	if trim(UpLoad("N_File_1")) <> "" then
		N_File_1 = UpLoad("N_File_1").Save(,False)
	end if
	if trim(UpLoad("N_File_2")) <> "" then
		N_File_2 = UpLoad("N_File_2").Save(,False)
	end if
	if trim(UpLoad("N_File_3")) <> "" then
		N_File_3 = UpLoad("N_File_3").Save(,False)
	end if
	
	N_Title			= Trim(UpLoad("N_Title"))
	N_Content		= Trim(UpLoad("N_Content"))
	N_Reg_Date		= now()
	N_Edit_Date		= now()
	N_File_1		= Replace(N_File_1,DefaultPath_Notice,"")
	N_File_2		= Replace(N_File_2,DefaultPath_Notice,"")
	N_File_3		= Replace(N_File_3,DefaultPath_Notice,"")
	Member_M_ID		= gM_ID
	
	rem DB 업데이트
	RS1.Open "tbNotice",sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("N_Title")			= N_Title
		.Fields("N_Content")		= N_Content
		.Fields("N_Reg_Date")		= N_Reg_Date
		.Fields("N_Edit_Date")		= N_Edit_Date
		.Fields("N_File_1")			= N_File_1
		.Fields("N_File_2")			= N_File_2
		.Fields("N_File_3")			= N_File_3
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