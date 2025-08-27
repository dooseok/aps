<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->

<script language="javascript">
<%
dim CNT1
dim FSO
dim objFolder
dim objFile

dim strFile

set FSO			= Server.CreateObject("Scripting.FileSystemObject")
set objFolder	= FSO.GetFolder("d:\home\msekorea\admin\bom_src\")
set objFile		= objFolder.Files

CNT1 = 1
for Each strFile in objFile
%>
	setTimeout("fRun('bom_dir_reader_action.asp?strFile=<%=strFile.name%>')",<%=CNT1 * 7000%>);
<%
	CNT1=CNT1+1
next

set FSO = nothing
%>

function fRun(strURL)
{
	window.open(strURL);	
}
</script>


<!-- #include virtual = "/header/db_tail.asp" -->