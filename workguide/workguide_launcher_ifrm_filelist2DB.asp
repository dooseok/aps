<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
dim SQL

dim objFSO

dim objFolder
dim objSubFolders
dim subFolder

dim objFolder2
dim objSubFolders2
dim subFolder2

dim objFolder3
dim objFiles3
dim File3

dim arrWI_PartNo
dim WI_PartNo
dim WI_PartNo_Alt
dim WI_ProcessNumber
dim WI_Line

set objFSO = server.CreateObject("Scripting.FileSystemObject")
set objFolder = objFSO.GetFolder(DefaultPath_workguide_img)
set objSubFolders = objFolder.subFolders

for each subFolder in objSubFolders 
	set objFolder2		= objFSO.GetFolder(DefaultPath_workguide_img & subFolder.name)
	set objSubFolders2	= objFolder2.subFolders
	
	arrWI_PartNo = split(subFolder.name,"=")
	if instr(subFolder.name,"=") = 0 then
		WI_PartNo = arrWI_PartNo(0)
		WI_PartNo_Alt = ""
	else
		WI_PartNo = arrWI_PartNo(0)
		WI_PartNo_Alt = arrWI_PartNo(1)
	end if
	
	for each subFolder2 in objSubFolders2
		if isnumeric(left(subFolder2.Name,2)) then '앞의 두자리는 무조건 숫자여야 함
			WI_ProcessNumber = cint(left(subFolder2.Name,2))
			
			if instr(subFolder2.Name,"@") > 0 then
				WI_Line = right(subFolder2.Name,len(subFolder2.Name)-instr(subFolder2.Name,"@"))
			else
				WI_Line = ""
			end if			
			
			set objFolder3	= objFSO.GetFolder(DefaultPath_workguide_img & subFolder.name & "\" & subFolder2.Name)
			set objFiles3	= objFolder3.Files  

			for each File3 In objFiles3
				if right(lcase(File3.name),5) = ".jpeg" or instr("-.jpg-.png-.gif-","-"&right(lcase(File3.Name),4)&"-") > 0 then
					SQL = "insert into tbWorkGuideImage (WI_PartNo, WI_ProcessNumber, WI_ImageFileName, WI_PartNo_Alt, WI_ProcessName, WI_Line, WI_Temp_YN) values "
					SQL = SQL & "('"&WI_PartNo&"',"&WI_ProcessNumber&",'"&lcase(File3.name)&"','"&WI_PartNo_Alt&"','"&subFolder2.Name&"','"&WI_Line&"', 'Y')"
					sys_DBCon.execute(SQL)
				end if
			next
			set objFiles3		= nothing
			set objFolder3		= nothing
		end if
	next
	
	set objSubFolders2	= nothing	
	set objFolder2		= nothing
next

set objSubFolders	= nothing
set objFolder		= nothing

SQL = "delete from tbWorkGuideImage where WI_Temp_YN <> 'Y'"
sys_DBCon.execute(SQL)

SQL = "update tbWorkGuideImage set WI_Temp_YN = 'N'"
sys_DBCon.execute(SQL)
%>

<form name="frmRedirect" action="about:blank" method=post>
</form>
<script language="javascript">
alert("이미지가 DB에 업데이트 되었습니다.\n라인별로 별도로 할 필요는 없습니다.");
$('html,body', parent.document).css('cursor','default');
frmRedirect.submit();
</script>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->