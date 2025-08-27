<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
 
<%
dim Part_filePrefix
dim Part_Title
dim Part_Title_Eng
Part_filePrefix = "bpmd"
Part_Title		= "개발"
Part_Title_Eng	= "Dev"

rem 변수선언
dim SQL
dim RS1

dim BPM_PartNo
dim BPM_StartDate
dim BPM_EndDate
dim BPM_Memo

dim temp
dim strError
dim URL_Prev
dim URL_Next

rem 객체선언
Set RS1		= Server.CreateObject("ADODB.RecordSet")

rem 업로드 될 물리적 경로지정
URL_Prev	= Request("URL_Prev")
URL_Next	= Request("URL_Next")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

rem 업로드작업
	BPM_StartDate	= Trim(Request("BPM_StartDate"))
	BPM_EndDate		= Trim(Request("BPM_EndDate"))
	BPM_Memo		= Trim(Request("BPM_Memo"))
	BPM_PartNo		= Trim(Request("BPM_PartNo"))
	
	rem DB 업데이트
	RS1.Open "tbBOM_PartsOutSheet_Memo_"&Part_Title_Eng,sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("BPM_PartNo")	= BPM_PartNo
		.Fields("BPM_StartDate")= BPM_StartDate
		.Fields("BPM_EndDate")	= BPM_EndDate
		.Fields("BPM_Memo")		= BPM_Memo
		
		.Update
		.Close
	end with
end if

rem 객체 해제
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