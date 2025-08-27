<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<%
dim s_PRD_Line
dim WG_Pos
dim WG_SlideDelay
dim WG_SlideDelay_Main
dim strPartNo
dim WG_ResX
dim WG_ResY

dim strImageURL
dim WI_ProcessName
dim strWI_ImageFileName
dim arrWI_ImageFileName

dim nDelayUnitValue

dim IDX
dim SlideCNT

dim myImg
dim iWidth
dim iheight

dim nSlideDelay

nDelayUnitValue = 10

s_PRD_Line			= request("s_PRD_Line")
WG_Pos				= request("WG_Pos")
WG_SlideDelay		= request("WG_SlideDelay")
WG_SlideDelay_Main	= request("WG_SlideDelay_Main")
strPartNo			= request("strPartNo")
WG_ResX				= request("WG_ResX")
WG_ResY				= request("WG_ResY")

'슬라이드 번호 
IDX				= request("IDX")
if IDX = "" then
	IDX = 0
end if
'몇번째 슬라이드인지
SlideCNT		= request("SlideCNT")

dim RS1
dim SQL

'DB에서 해당파트넘버와 공정의 이미지를 읽어온다.
set RS1 = Server.CreateObject("ADODB.RecordSet")

strImageURL = ""

'SQL = "select * from tbWorkGuideImage where WI_PartNo = '"&strPartNo&"' and " &vbcrlf
'SQL = SQL & "WI_ProcessNumber = "&WG_Pos&" and WI_Temp_YN='N' and WI_Line = '"&s_PRD_Line&"' and " &vbcrlf
'SQL = SQL & "(right(WI_ImageFileName,5) = '.jpeg' or right(WI_ImageFileName,4) in ('.jpg','.png','.gif')) " &vbcrlf
'SQL = SQL & " order by WI_ImageFileName asc" &vbcrlf
'RS1.Open SQL,sys_DBCon
'if not(RS1.Eof or RS1.Bof) then
'	WI_ProcessName = RS1("WI_ProcessName")
'	do until RS1.Eof
'		strImageURL = strImageURL & RS1("WI_ImageFileName") & "|%|"
'		RS1.MoveNext
'	loop
'end if
'if strImageURL = "" then
'	SQL = "select * from tbWorkGuideImage where WI_PartNo = '"&strPartNo&"' and " &vbcrlf
'	SQL = SQL & "WI_ProcessNumber = "&WG_Pos&" and WI_Temp_YN='N' and " &vbcrlf
'	SQL = SQL & "(right(WI_ImageFileName,5) = '.jpeg' or right(WI_ImageFileName,4) in ('.jpg','.png','.gif')) " &vbcrlf
'	SQL = SQL & " order by WI_ImageFileName asc" &vbcrlf
'	RS1.Open SQL,sys_DBCon
	
	'루프를 돌며, 경로를 배열화 한다.
'	if not(RS1.Eof or RS1.Bof) then
'		WI_ProcessName = RS1("WI_ProcessName")
'		do until RS1.Eof
'			strImageURL = strImageURL & RS1("WI_ImageFileName") & "|%|"
'			RS1.MoveNext
'		loop
'	end if
'	RS1.Close
'end if

if strImageURL = "" then
	strImageURL = getImageURL_FSO(strPartNo, WG_Pos, s_PRD_Line)
end if

strWI_ImageFileName = strImageURL
response.write strWI_ImageFileName&"<BR>"
arrWI_ImageFileName = split(strImageURL,"|%|")

if strImageURL <> "" then
	'단일 파일인 경우 그대로 할당 
	if ubound(arrWI_ImageFileName) = 1 then
		strImageURL = DefaultPath_workguide_img & strPartNo & "\" & arrWI_ImageFileName(0)
	'파일이 복수인 경우 
	else
		
		'첫 슬라이드라면 nSlideDelay에 주 슬라이드간격을 할당
		if cint(IDX) = 0 then
			nSlideDelay = WG_SlideDelay_Main
		else
			nSlideDelay = WG_SlideDelay
		end if
		
		'만약 슬라이드 전환속도가 슬라이드번호와 슬라이드누적카운트가 같다면 슬라이드누적카운트는 0로 하고	
		if cint(nSlideDelay) = cint(SlideCNT) + nDelayUnitValue then 
			SlideCNT = 0
			
			'다음 이미지가 있다면 다음이미지로, 없다면 첫슬라이드로 설정한다.
			if cint(IDX) < ubound(arrWI_ImageFileName)-1 then 
				IDX = IDX + 1
			else 
				IDX = 0
			end if
		else
			SlideCNT = SlideCNT + nDelayUnitValue
		end if
		
		'해당공정에 맞는 파일중 IDX값에 해당하는 걸 할당 
		strImageURL = DefaultPath_workguide_img & strPartNo & "\" & arrWI_ImageFileName(IDX)	
	end if
end if
set RS1 = nothing

if strImageURL <> "" then
	set myImg = loadpicture(strImageURL)
	iWidth = round(myImg.width / 26.4583)
	iheight = round(myImg.height / 26.4583)
	set myImg = nothing
	strImageURL = replace(strImageURL,"d:\my_website\msekorea\admin","")
	strImageURL = replace(strImageURL,"\","/")
else
	iWidth = 1
	iHeight = 1
	strImageURL = "/img/blank.gif"
end if

'Response.write iWidth&"_"&iheight
%>

<script language="javascript">
parent.imgWorkGuide.src = "<%=strImageURL%>?<%=replace(replace(replace(replace(replace(now()," ",""),"오후","PM"),"오전","AM"),"-",""),":","")%>";
parent.imgWorkGuide.height = <%=WG_ResY-78%>;
parent.imgWorkGuide.width = parseInt(<%=WG_ResY-78%> * parseFloat(<%=iWidth%> / <%=iheight%>));

function reload_handle()
{
	location.href="workguide_viewer_ifrm.asp?WG_Pos=<%=WG_Pos%>&s_PRD_Line=<%=s_PRD_Line%>&WG_SlideDelay=<%=WG_SlideDelay%>&WG_SlideDelay_Main=<%=WG_SlideDelay_Main%>&strPartNo=<%=strPartNo%>&WG_ResX=<%=WG_ResX%>&WG_ResY=<%=WG_ResY%>&IDX=<%=IDX%>&SlideCNT=<%=SlideCNT%>";
}
</script>

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
				strImageURL = strImageURL & arrImageURL(CNT1) & "|%|"
			end if
		next
	else
		for CNT1 = 0 to ubound(arrImageURL)-1
			if instr(lcase(arrImageURL(CNT1)), "@"&lcase(strWI_Line)) > 0 then
				strImageURL = strImageURL & arrImageURL(CNT1) & "|%|"
			end if
		next
	end if
	
	getImageURL_FSO = strImageURL
end function 
%>



<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->