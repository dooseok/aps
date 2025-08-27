<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
response.write getImageURL_FSO("EBR81767925", 10, "AUTO2")&"<BR>"
%>

<%
function getImageURL_FSO(strWI_PartNo, nProcessNo, strWI_Line)
	dim strImageURL
	dim strProcess
	dim strProcessName
	
	dim arrImageURL
	
	dim CNT1
	dim SQL
	dim RS1
	
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
	dim WI_ProcessNumber
	dim WI_Line
	
	dim strSpecificLines
	strSpecificLines = "|"
	
	'set RS1 = server.CreateObject("ADODB.RecordSet")
	set objFSO = server.CreateObject("Scripting.FileSystemObject")

	if objFSO.FolderExists(DefaultPath_workguide_img & strWI_PartNo) = true then

		set objFolder2		= objFSO.GetFolder(DefaultPath_workguide_img & strWI_PartNo)
		set objSubFolders2	= objFolder2.subFolders
		
		for each subFolder2 in objSubFolders2
			strProcessName = subFolder2.Name
			
			if isnumeric(left(strProcessName,2)) then '앞의 두자리는 무조건 숫자여야 함
				
				if cint(nProcessNo) = cint(left(strProcessName,2)) then
					
					if instr(lcase(strProcessName),"@") > 0 then
						strSpecificLines = strSpecificLines & right(strProcessName,len(strProcessName)-instr(strProcessName,"@")) & "|"
					end if
					
					WI_ProcessNumber = cint(left(strProcessName,2))
					
					set objFolder3	= objFSO.GetFolder(DefaultPath_workguide_img & strWI_PartNo & "\" & strProcessName)
					set objFiles3	= objFolder3.Files  
		
					for each File3 In objFiles3
					
						if right(lcase(File3.name),5) = ".jpeg" or instr("-.jpg-.png-.gif-","-"&right(lcase(File3.Name),4)&"-") > 0 then
							
							strImageURL = strImageURL & strProcessName & "\" & lcase(File3.name) & "|%|"

							'SQL = "select top 1 WI_Code from tbWorkGuideImage where "
							'SQL = SQL & "WI_PartNo = '"&strWI_PartNo&"' and "
							'SQL = SQL & "WI_ProcessNumber = "&WI_ProcessNumber&" and "
							'SQL = SQL & "WI_ImageFileName = '"&lcase(File3.name)&"' and "
							'SQL = SQL & "WI_Temp_YN = 'N' and "
							'SQL = SQL & "WI_PartNo_Alt = '' and "
							'SQL = SQL & "WI_ProcessName = '"&subFolder2.Name&"' and "
							'SQL = SQL & "WI_Line = '"&WI_Line&"' "
							'RS1.Open SQL, sys_DBCon
							'if RS1.Eof or RS1.Bof then
							'	SQL = "insert into tbWorkGuideImage (WI_PartNo, WI_ProcessNumber, WI_ImageFileName, WI_PartNo_Alt, WI_ProcessName, WI_Line, WI_Temp_YN) values "
							'	SQL = SQL & "('"&strWI_PartNo&"',"&WI_ProcessNumber&",'"&lcase(File3.name)&"','','"&subFolder2.Name&"','"&WI_Line&"', 'N')"
							'	sys_DBCon.execute(SQL)
							'	
								'response.write SQL &"<BR>"
							'else
								'response.write "exists<BR>"
							'end if
							'RS1.Close
						end if
					next
					
					set objFiles3		= nothing
					set objFolder3		= nothing
				end if				
			end if	
		next
		
		set objSubFolders2	= nothing
		set objFolder2		= nothing
	end if
	
	set objFSO = nothing
	'set RS1 = nothing
	
	'정리
	
	arrImageURL = split(strImageURL,"|%|")
	strImageURL = ""
	
	'case1. 일치하는 라인이 없다면, 특정라인명이 없는 경로만 재배열 한다.
	'특정라인명이 없는 경로만 재배열한다.
	if strWI_Line = "" or instr(strSpecificLines,"|"&strWI_Line&"|") = 0 then
		for CNT1 = 0 to ubound(arrImageURL)-1
			if instr(lcase(arrImageURL(CNT1)), "@") = 0 then
				strImageURL = arrImageURL(CNT1) & "|%|"
			end if
		next
	else
		for CNT1 = 0 to ubound(arrImageURL)-1
			if instr(lcase(arrImageURL(CNT1)), "@"&lcase(strWI_Line)) > 0 then
				strImageURL = arrImageURL(CNT1) & "|%|"
			end if
		next
	end if
	
	getImageURL_FSO = strImageURL
end function 
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->