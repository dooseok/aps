<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
'���� ����!
dim RS1
dim SQL
dim arrRecordSet

dim WI_ProcessName
dim strWI_ImageFileName(15)
dim arrWI_ImageFileName

dim CNT1, CNT2
dim strWG_Pos
dim strWG_ResX
dim strWG_ResY
dim strWG_MCDelay
dim strWG_SlideDelay
dim strWG_SlideDelay_Main
dim strWG_Auto_YN
dim arrWG_Pos
dim arrWG_ResX
dim arrWG_ResY
dim arrWG_MCDelay
dim arrWG_SlideDelay
dim arrWG_SlideDelay_Main
dim arrWG_Auto_YN

dim s_PRD_Line

dim strPrePartNo
dim arrPrePartNo
dim strIDX
dim arrIDX
dim strSlideCNT
dim arrSlideCNT

'�̹��� ������¡
dim myImg
dim iWidth(15)
dim iheight(15)

dim strImageURL(15)

dim PRD_PartNo
dim arrPRD_PartNo(15)

dim nDelayUnitValue

dim dbtime
dim mctime

dim AMPM
dim arrHHMMSS

dim nSlideDelay

nDelayUnitValue = 10

'����
s_PRD_Line		= request("s_PRD_Line")

'���� ��Ʈ�ѹ�
strPrePartNo	= request("strPrePartNo")
arrPrePartNo	= split(strPrePartNo,",")

'�����̵� ��ȣ 
strIDX			= request("strIDX")
arrIDX			= split(strIDX,",")

'���° �����̵�����
strSlideCNT		= request("strSlideCNT")
arrSlideCNT		= split(strSlideCNT,",")

set RS1 = server.CreateObject("ADODB.RecordSet")

'SQL = "insert into tbTest_setinterval (ts_Work,ts_Desc,ts_Now,ts_Diff) values ('WorkGuide','"&s_PRD_Line&"',getdate(),0)"
'sys_DBCon.execute(SQL)

'DB���� ����� �������� ������ 
SQL = "select WG_Pos, WG_ResX, WG_ResY, WG_MCDelay, WG_SlideDelay, WG_SlideDelay_Main, WG_Auto_YN from tbWorkguide where PRD_Line='"&s_PRD_Line&"' order by WG_Pos asc"
RS1.Open SQL,sys_DBCon

strWG_Pos				= ""
strWG_ResX				= ""
strWG_ResY				= ""
strWG_MCDelay			= ""
strWG_SlideDelay		= ""
strWG_SlideDelay_Main	= ""
strWG_Auto_YN			= ""
do until RS1.Eof
	strWG_Pos				= strWG_Pos				& RS1("WG_Pos")	& ","
	strWG_ResX				= strWG_ResX			& RS1("WG_ResX")& ","
	strWG_ResY				= strWG_ResY			& RS1("WG_ResY")& ","
	strWG_MCDelay			= strWG_MCDelay			& RS1("WG_MCDelay")& ","
	strWG_SlideDelay		= strWG_SlideDelay		& RS1("WG_SlideDelay")& ","
	strWG_SlideDelay_Main	= strWG_SlideDelay_Main	& RS1("WG_SlideDelay_Main")& ","
	strWG_Auto_YN			= strWG_Auto_YN			& RS1("WG_Auto_YN")& ","
	RS1.MoveNext
loop
RS1.Close
arrWG_Pos				= split(strWG_Pos,",")
arrWG_ResX				= split(strWG_ResX,",")
arrWG_ResY				= split(strWG_ResY,",")
'��ü���� ������ 
arrWG_MCDelay			= split(strWG_MCDelay,",")
'�����̵� ��ȯ�ӵ�
arrWG_SlideDelay		= split(strWG_SlideDelay,",")
arrWG_SlideDelay_Main	= split(strWG_SlideDelay_Main,",")
arrWG_Auto_YN			= split(strWG_Auto_YN,",")

'for CNT1 = 0 to 14 
'	SQL = "select top 1 PRD_PartNo from tbPWS_Raw_Data "&vbcrlf
'	SQL = SQL & "where "&vbcrlf
'	SQL = SQL & "	PRD_Line = '"&s_PRD_Line&"' and "&vbcrlf
'	SQL = SQL & "	DATEADD(s,"&-1*arrWG_MCDelay(CNT1)&",getdate()) "&vbcrlf
'	SQL = SQL & "	> "&vbcrlf
'	SQL = SQL & "	convert(datetime,convert(char(10),PRD_Input_date) + ' ' + convert(char(8),PRD_Input_Time_Detail)) "&vbcrlf
'	SQL = SQL & "order by PRD_Code desc "&vbcrlf
'	RS1.Open SQL,sys_DBCon
'	if RS1.Eof or RS1.Bof then
'		arrPRD_PartNo(CNT1) = ""
'	else
'		arrPRD_PartNo(CNT1) = RS1("PRD_PartNo")
'	end if
'	RS1.Close
'next
'[��Ȳ�� ���� ��------------------------------
'SQL = "select top 900 PRD_PartNo, PRD_Input_Date, PRD_Input_Time_Detail from tbPWS_Raw_Data "&vbcrlf
'SQL = SQL & "where "&vbcrlf
'SQL = SQL & "	PRD_Line = '"&s_PRD_Line&"' and "&vbcrlf
'SQL = SQL & " 	(PRD_byHook_YN is null or PRD_byHook_YN = 'Y') "&vbcrlf
'SQL = SQL & "order by PRD_Input_Date desc, PRD_Input_Time_Detail desc "&vbcrlf
'��Ȳ�� ���� ��------------------------------]
'[��Ȳ�� ���� ��------------------------------
SQL = "select top 900 SML_PartNo, SML_Date, strhhmmss = LEFT(SML_Time,2)+':'+SUBSTRING(SML_Time,3,2)+':'+right(SML_Time,2) from tblStatus_Monitor_Line where "
SQL = SQL & "SML_Line='"&s_PRD_Line&"' and "
SQL = SQL & "SML_Type in ('N','F','T') and "
SQL = SQL & "SML_Process = 'START' " 
SQL = SQL & "order by SML_Code desc "
'response.write SQL
'��Ȳ�� ���� ��------------------------------]
RS1.Open SQL,sys_DBCon
dim ResultYN
if not(RS1.Eof or RS1.Bof) then
	ResultYN = "Y"
	arrRecordSet = RS1.GetRows()
end if
RS1.Close	


'������ ����ͻ��� ��ü���� �����̿� ���缭 �� ����ͺ� ��Ʈ�ѹ��� ������ 

CNT2 = 1

'[��Ȳ�� ���� ��------------------------------
if ResultYN = "Y" then
'��Ȳ�� ���� ��------------------------------]
	for CNT1 = 0 to ubound(arrRecordSet,2)
		
		dbtime = arrRecordSet(1,CNT1) &" "& arrRecordSet(2,CNT1)
		arrHHMMSS = split(formatDateTime(dateadd("s",-1*arrWG_MCDelay(CNT2),Time), vbLongTime),":")
		AMPM = left(arrHHMMSS(0),2)
		arrHHMMSS(0) = replace(arrHHMMSS(0),AMPM&" ","")
		
		'�����̰� 12�ö��
		if AMPM = "����" and cint(arrHHMMSS(0)) = 12 then
			arrHHMMSS(0) = 0
		elseif AMPM = "����" and cint(arrHHMMSS(0)) < 12 then
			arrHHMMSS(0) = cint(arrHHMMSS(0)) + 12
		end if
		if arrHHMMSS(0) < 10 then
			 arrHHMMSS(0) = "0" & cstr(arrHHMMSS(0))
		end if

		mctime = date()&" "& cstr(arrHHMMSS(0))&":"&cstr(arrHHMMSS(1))&":"&cstr(arrHHMMSS(2))
		if datediff("s",dbtime,mctime) >= 0 then
		
			arrPRD_PartNo(CNT2) = arrRecordSet(0,CNT1)
			'response.write arrPRD_PartNo(CNT2) &":"& CNT2 &"<br>"
			CNT1 = CNT1 - 1
			CNT2 = CNT2 + 1
			
		end if
		
		if CNT2 = 15 then
			exit for
		end if
	next
'[��Ȳ�� ���� ��------------------------------
end if
'��Ȳ�� ���� ��------------------------------]

'[��Ȳ�� ���� ��------------------------------
'SQL = ""
'SQL = SQL & " select top 1 PRD_PartNo "
'SQL = SQL & " from tbPWS_Raw_Data "
'SQL = SQL & " where "
'SQL = SQL & " 	PRD_Line='"&s_PRD_Line&"' and "
'SQL = SQL & " 	PRD_Input_Date = '"&date()&"' and "
'SQL = SQL & " 	(PRD_byHook_YN is null or PRD_byHook_YN = 'Y') "
'SQL = SQL & " order by PRD_Input_Time_Detail desc "
'RS1.Open SQL,sys_DBCon

'PRD_PartNo = ""
'if not(RS1.Eof or RS1.Bof) then
'	PRD_PartNo = RS1("PRD_PartNo")
'end if
'RS1.Close

'if PRD_PartNo = "" then
'	SQL = ""
'	SQL = SQL & " select top 1 PRD_PartNo "
'	SQL = SQL & " from tbPWS_Raw_Data "
'	SQL = SQL & " where "
'	SQL = SQL & " 	PRD_Line='"&s_PRD_Line&"' and "
'	SQL = SQL & " 	(PRD_byHook_YN is null or PRD_byHook_YN = 'Y') "
'	SQL = SQL & " order by "
'	SQL = SQL & " 	PRD_Input_Date desc, "
'	SQL = SQL & " 	PRD_Input_Time_Detail desc "
'	RS1.Open SQL,sys_DBCon
'	if not(RS1.Eof or RS1.Bof) then
'		PRD_PartNo = RS1("PRD_PartNo")
'	end if
'	RS1.Close
'end if
'��Ȳ�� ���� ��------------------------------]
'[��Ȳ�� ���� ��------------------------------
'PRD_PartNo = application(s_PRD_Line&"_Last")
if PRD_PartNo = "" then
	SQL = ""
	SQL = SQL & "select top 1 SML_PartNo from tblStatus_Monitor_Line where "
	SQL = SQL & "SML_Line='"&s_PRD_Line&"' and "
	SQL = SQL & "SML_Type in ('N','F','T') and "
	SQL = SQL & "SML_Process = 'START' "  
	SQL = SQL & "order by SML_Code desc "
	RS1.Open SQL,sys_DBCon
	if not(RS1.Eof or RS1.Bof) then
		PRD_PartNo = RS1("SML_PartNo")
	end if
	RS1.Close
	'if application(s_PRD_Line&"_Last") = "" then
	'	application(s_PRD_Line&"_Last")=PRD_PartNo
	'end if
end if

'��Ȳ�� ���� ��------------------------------]

'���� ��ü���� �ӵ��� 0��� �׳� ���� �ֱٵ����͸� ������
for CNT1 = 0 to 14 
	if arrPRD_PartNo(CNT1) = "" or cint(arrWG_MCDelay(CNT1)) = 0 then
		arrPRD_PartNo(CNT1) = PRD_PartNo
	end if
next

'������ ����
for CNT1 = 0 to 14
	strImageURL(CNT1) = ""
	
	'�ش���Ʈ�ѹ� �� �ش������ ��ġ�ϴ� �̹������ϸ���� �����´�.
	'Ư�����ο� �ش��ϴ°� �ִ��� ���� üũ
	
	'SQL = "select * from tbWorkGuideImage where WI_PartNo = '"&arrPRD_PartNo(CNT1)&"' and " &vbcrlf
	'SQL = SQL & "WI_ProcessNumber = "&cint(CNT1)+1&" and WI_Temp_YN='N' and WI_Line = '"&s_PRD_Line&"' and " &vbcrlf
	'SQL = SQL & "(right(WI_ImageFileName,5) = '.jpeg' or right(WI_ImageFileName,4) in ('.jpg','.png','.gif')) " &vbcrlf
	'SQL = SQL & " order by WI_ImageFileName asc" &vbcrlf
	'RS1.Open SQL,sys_DBCon
	'if not(RS1.Eof or RS1.Bof) then
	'	WI_ProcessName = RS1("WI_ProcessName")
	'	do until RS1.Eof
	'		strImageURL(CNT1) = strImageURL(CNT1) & RS1("WI_ImageFileName") & "|%|"
	'		RS1.MoveNext
	'	loop
	'end if
	'RS1.Close
	
	'if strImageURL(CNT1) = "" then
		'�ش���Ʈ�ѹ� �� �ش������ ��ġ�ϴ� �̹������ϸ���� �����´�.
	'	SQL = "select * from tbWorkGuideImage where WI_PartNo = '"&arrPRD_PartNo(CNT1)&"' and " &vbcrlf
	'	SQL = SQL & "WI_ProcessNumber = "&cint(CNT1)+1&" and WI_Temp_YN='N' and " &vbcrlf
	'	SQL = SQL & "(right(WI_ImageFileName,5) = '.jpeg' or right(WI_ImageFileName,4) in ('.jpg','.png','.gif')) " &vbcrlf
	'	SQL = SQL & " order by WI_ImageFileName asc" &vbcrlf
	'	RS1.Open SQL,sys_DBCon
		'response.write CNT1 &":"& arrPRD_PartNo(CNT1) &")<Br>"
		
	'	if not(RS1.Eof or RS1.Bof) then
	'		WI_ProcessName = RS1("WI_ProcessName")
	'		do until RS1.Eof
	'			strImageURL(CNT1) = strImageURL(CNT1) & RS1("WI_ImageFileName") & "|%|"
	'			RS1.MoveNext
	'		loop
	'	end if
	'	RS1.Close
	'end if
	
	'�� ã���� ���Ͻý��ۿ��� ã�Ƽ� DB�� �ִ´�.
	if strImageURL(CNT1) = "" then
		strImageURL(CNT1) = getImageURL_FSO(arrPRD_PartNo(CNT1), cint(CNT1)+1, s_PRD_Line)
	end if
	
	strWI_ImageFileName(CNT1) = strImageURL(CNT1)
	arrWI_ImageFileName = split(strImageURL(CNT1),"|%|")
	
	if strImageURL(CNT1) <> "" then
		'���� ������ ��� �״�� �Ҵ� 
		if ubound(arrWI_ImageFileName) = 1 then
			strImageURL(CNT1) = DefaultPath_workguide_img & arrPRD_PartNo(CNT1) & "\" & arrWI_ImageFileName(0)
		'������ ������ ��� 
		else
			'�̹��� ��ü������ �Ǿ��ٸ�, �����̵� ��ȣ�� �ٽ� ó������
			if arrPRD_PartNo(CNT1) <> arrPrePartNo(CNT1) then  
				arrIDX(CNT1) = 0 'ù�����̵�� ����
				arrSlideCNT(CNT1) = 0 '�����̵� ī��Ʈ�� ó������
				
			'��ü������ �� ���� �ƴ϶��
			else
				'ù �����̵��� nSlideDelay�� �� �����̵尣���� �Ҵ�
				if cint(arrIDX(CNT1)) = 0 then
					nSlideDelay = arrWG_SlideDelay_Main(CNT1)
				else
					nSlideDelay = arrWG_SlideDelay(CNT1)
				end if
				
				'���� �����̵尣���� �����̵��ȣ�� �����̵崩��ī��Ʈ�� ���ٸ� �����̵崩��ī��Ʈ�� 0�� �ϰ�
				if cint(nSlideDelay) = cint(arrSlideCNT(CNT1)) + nDelayUnitValue then
					arrSlideCNT(CNT1) = 0
					
					'���� �̹����� �ִٸ� �����̹�����, ���ٸ� ù�����̵�� �����Ѵ�.	
					if cint(arrIDX(CNT1)) < ubound(arrWI_ImageFileName)-1 then 
						arrIDX(CNT1) = arrIDX(CNT1) + 1
					else 
						arrIDX(CNT1) = 0
					end if
				else
					arrSlideCNT(CNT1) = arrSlideCNT(CNT1) + nDelayUnitValue
				end if
			end if
			
			'�ش������ �´� ������ arrIDX���� �ش��ϴ� �� �Ҵ� 
			strImageURL(CNT1) = DefaultPath_workguide_img & arrPRD_PartNo(CNT1) & "\" & arrWI_ImageFileName(arrIDX(CNT1))
			
		end if
	end if
next

set RS1 = nothing

strPrePartNo	= ""
strIDX 			= ""
strSlideCNT		= ""
for CNT1 = 0 to 14
	'response.write arrPRD_PartNo(CNT1)&"<BR>"
	strPrePartNo	= strPrePartNo	& arrPRD_PartNo(CNT1)	&","
	strIDX			= strIDX		& arrIDX(CNT1)			&","
	strSlideCNT		= strSlideCNT	& arrSlideCNT(CNT1)		&","
next
for CNT1 = 0 to 14
	if strImageURL(CNT1) <> "" then
		
		'[loadpicture���� ������--------------------------------------------
		'set myImg = loadpicture(strImageURL(CNT1))
		'iWidth(CNT1) = round(myImg.width / 26.4583)
		'iheight(CNT1) = round(myImg.height / 26.4583)
		'set myImg = nothing
		'loadpicture���� ������-------------------------------------------]
		strImageURL(CNT1) = replace(strImageURL(CNT1),DefaultPath_workguide_img,"\workguide\workguide_img\")
		strImageURL(CNT1) = replace(strImageURL(CNT1),"\","/")
	else
		iWidth(CNT1) = 1
		iHeight(CNT1) = 1
		strImageURL(CNT1) = "/img/blank.gif"
	end if
	'response.write CNT1 &":"&strImageURL(CNT1) &"<br>"
next
%>

<script language="javascript">
<%
for CNT1 = 0 to 14
	if arrWG_Auto_YN(CNT1) = "Y" then
%>
if(typeof(parent.arrWorkGuide_VW[<%=CNT1%>])=='object')
{	
	if(typeof(parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide)=='object')
	{
		parent.arrWorkGuide_VW[<%=CNT1%>].document.title = "<%=arrPRD_PartNo(CNT1)%> (<%=CNT1+1%>)";
		parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide.src = "<%=strImageURL(CNT1)%>?<%=replace(replace(replace(replace(replace(now()," ",""),"����","PM"),"����","AM"),"-",""),":","")%>";

		//[loadpicture���� ������--------------------------------------------
		//var nWidth = parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide.width;
        //var nHeight = parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide.height; 
        //console.log(nWidth+'-'+nHeight);
        //loadpicture���� ������--------------------------------------------]
        
		//parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide.height = <%=arrWG_ResY(CNT1)-78%>;
		
		//[loadpicture���� ������--------------------------------------------
		//parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide.width = parsecint(<%=arrWG_ResY(CNT1)-78%> * parseFloat(<%=iWidth(CNT1)%> / <%=iheight(CNT1)%>));
		//loadpicture���� ������--------------------------------------------]
		
		//[loadpicture���� ������--------------------------------------------
		//parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide.width = parsecint(<%=arrWG_ResY(CNT1)-78%> * parseFloat( nWidth / nHeight ));
		//loadpicture���� ������--------------------------------------------]
	}
}
<%
	end if
next
%>

function reload_handle()
{
	location.href="workguide_launcher_ifrm.asp?s_PRD_Line=<%=s_PRD_Line%>&strIDX=<%=server.urlencode(strIDX)%>&strSlideCNT=<%=server.urlencode(strSlideCNT)%>&strPrePartNo=<%=server.urlencode(strPrePartNo)%>&";
}

/*
function fRun()
{
	if(document.readyState == "complete")
	{
		location.href="workguide_launcher_ifrm.asp?s_PRD_Line=<%=s_PRD_Line%>&strIDX=<%=server.urlencode(strIDX)%>&strSlideCNT=<%=server.urlencode(strSlideCNT)%>&strPrePartNo=<%=server.urlencode(strPrePartNo)%>&";
	}
	else
	{
		setTimeout("fRun()",<%=nDelayUnitValue%>000);
	}
}

fRun();
*/
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
			
			if isnumeric(left(strProcessName,2)) then '���� ���ڸ��� ������ ���ڿ��� ��
				
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
	
	'����
	
	arrImageURL = split(strImageURL,"|%|")
	strImageURL = ""
	
	'case1. ��ġ�ϴ� ������ ���ٸ�, Ư�����θ��� ���� ��θ� ��迭 �Ѵ�.
	'Ư�����θ��� ���� ��θ� ��迭�Ѵ�.
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