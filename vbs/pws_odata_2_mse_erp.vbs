On Error Resume Next

'DB오픈
Set objDBCon = CreateObject("ADODB.Connection")
Set wShell = CreateObject("Shell.Application")

'strMSG = "[혼입] cell라인 KR0512827758에 6871A20572Q(03-15포장), 6871A20572A(03-25포장)"
'wShell.Open "http://192.168.123.17:1080/vbs/sms_send.asp?strPhone=01067555740&strMSG="&strMSG

'윤성권;양수성;조기찬
strPhone = "01696775540;01055114409;01093330072"
DB_Err_Sent_YN = "N"

DBCon_YN = "N"
Do Until DBCon_YN = "Y"
	objDBCon.Open("Provider=SQLOLEDB;User ID=sa;Password=78;server=localhost,1011;database=spstest")
	If Err = 0 Then
		DBCon_YN = "Y"
	Else
		If DB_Err_Sent_YN = "N" Then
			Call DB_Err_Notice(Err.Source)
			DB_Err_Sent_YN = "Y"
		End If
		Err.clear()
	End If
Loop

Call PWS_Data_Handling()

'DB해제
objDBCon.Close
Set objDBCon		= Nothing


Sub PWS_Data_Handling()
	'필요한 객체 선언
	Set objWMIService	= getObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set objDBRecord		= CreateObject("ADODB.RecordSet")
	Set objFSO 			= CreateObject("Scripting.FileSystemObject")

	'인풋을 위한 변수들 선언
	Dim PRD_Code
	Dim PRD_PartNo
	Dim PRD_Barcode
	Dim PI_ICT_Date
	Dim PI_FCT_Date
	Dim PI_BOX_Date
	Dim arrLine
	
	Dim Result1
	Dim Result2

	'오늘로부터 5일 전까지 데이터를 읽기 위해 루핑
	For nDateCNT = 0 To -5 step -1
	
		'가져올 날짜
		strDate = DateAdd("d",nDateCNT,Date())
		
		'가져올 폴더 설정
		strICT_Folder_Location = "D:\LG-PWS\TDATA\ICT\ODATA\" & replace(strDate,"-","")
		strFCT_Folder_Location = "D:\LG-PWS\TDATA\FCT\ODATA\" & replace(strDate,"-","")
		strBox_Folder_Location = "D:\LG-PWS\TDATA\BOX\ODATA\" & replace(strDate,"-","")
'----------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------	
		''우선 ICT부터 처리한다.
'		If objFSO.FolderExists(strICT_Folder_Location) Then
'			'sent 폴더 생성
'			if Not objFSO.FolderExists(strICT_Folder_Location&"\sent") then
'				objFSO.CreateFolder(strICT_Folder_Location&"\sent")
'			end If
'
'			'대상 폴더를 객체화
'			Set objFolder	= objFSO.GetFolder(strICT_Folder_Location)
'			'해당 폴더의 파일들을 객체배열화
'			Set objFiles	= objFolder.Files
'			
'			'파일 배열을 루핑										
'			For each strFile in objFiles
'				Set objFile = objFSO.getFile(strFile)
'				
'				'파일명 읽기
'				strFileName = objFSO.GetFileName(strFile)
'				
'				'파일을 열어서 ICT검사결과 확인
'				Set objFile2 = objFSO.OpenTextFile(strFile,1) 			 '파일 열기
'				strLine	= objFile2.readLine					'라인 읽기	
'				arrLine	= split(strLine,"|")
'				Result1 = arrLine(19)
'				strLine	= objFile2.readLine					'라인 읽기	
'				arrLine	= split(strLine,"|")
'				Result2 = arrLine(19)
'				Set objFile2 = Nothing			
'				
'				'파일명 검증 및 ICT검사결과가 OK인 경우..
'				If InStr("-687-EBR-",Left(strFileName,3)) > 0 And Len(strFileName) = 42 And Result1 = "OK" And Result2 = "OK" Then
'					'파일 명으로 부터 정보 추출
'					PRD_PartNo	= Left(strFileName,11)
'					PRD_Barcode	= Left(strFileName,23)
'					PRD_ICT_Date	= strDate
'					PRD_ICT_Time	= Mid(strFileName,33,4)
'					
'					'Owner정보(=Line정보) 읽기
'					Set colItems = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_LogicalFileSecuritySetting='"&strICT_Folder_Location&"\"&strFileName&"'} Where AssocClass=Win32_LogicalFileOwner ResultRole=Owner")
'					For each objItem in colItems
'						PRD_Line = objItem.AccountName
'					Next
'					Set colItems = Nothing	
'					
'					'DB에 해당 바코드가 존재 하는지 체크
'					SQL = "select top 1 PRD_Code from tbPWS_Raw_Data where PRD_Barcode = '"&PRD_Barcode&"'"	
'					objDBRecord.Open SQL,objDBCon
'					If objDBRecord.Eof Or objDBRecord.Bof Then
'						Exist_YN = "N"	
'					Else
'						Exist_YN = "Y"
'					End If
'					objDBRecord.Close
'					
'					'DB에 존재 하지 않는다면> 새로 추가, 존재한다면> 날짜와 시간을 업데이트
'					If Exist_YN = "N" Then
'						SQL = "insert tbPWS_Raw_Data (PRD_PartNo,PRD_Barcode,PRD_ICT_Date,PRD_ICT_Time,PRD_Line) values "
'						SQL = SQL & "('"&PRD_PartNo&"','"&PRD_Barcode&"','"&PRD_ICT_Date&"','"&PRD_ICT_Time&"','"&PRD_Line&"')"
'					Else
'						SQL = "update tbPWS_Raw_Data set PRD_ICT_Date = '"&PRD_ICT_Date&"', PRD_ICT_Time = '"&PRD_ICT_Time&"', PRD_Line = '"&PRD_Line&"' where PRD_Barcode = '"&PRD_Barcode&"'"
'					End If
'					objDBCon.execute(SQL)
'					
'				End If
'				
'				'sent'폴더에 파일이 존재 한다면
'				If objFSO.FileExists(strICT_Folder_Location&"\"&strFileName) and objFSO.FileExists(strICT_Folder_Location&"\sent\"&strFileName) Then
'					Set objFile2 = objFSO.getFile(strICT_Folder_Location&"\sent\"&strFileName)
'					
'					'본파일의 수정날짜가 sent안의 파일의 수정날짜보다 이후이면.
'					If objFile.DateLastModified > objFile2.DateLastModified Then
'						'sent안의 파일 삭제.
'						objFSO.DeleteFile (strICT_Folder_Location&"\sent\"&strFileName)
'						'파일 이동처리
'						objFile.Move (strICT_Folder_Location&"\sent\"&strFileName)
'					Else
'						'sent안의 파일이 본파일보다 수정날짜가 같거나 이후이면 본파일 삭제 처리
'						objFSO.DeleteFile (strICT_Folder_Location&"\"&strFileName)
'						
'					End If
'					Set objFile2 = Nothing
'				Else
'					
'					'sent폴더안에 파일이 없는 경우 > sent폴더 안에 파일을 삭제'
'					objFile.Move (strICT_Folder_Location&"\sent\"&strFileName)
'				End If
'				
'				Set objFile = Nothing				
'			Next
'			
'			Set objFiles	= Nothing
'			Set objFolder	= Nothing
'		
'		End If
'----------------------------------------------------------------------------------------------------------------------------------	
'----------------------------------------------------------------------------------------------------------------------------------	
	'우선 FCT부터 처리한다.
		If objFSO.FolderExists(strFCT_Folder_Location) Then
			'sent 폴더 생성
			if Not objFSO.FolderExists(strFCT_Folder_Location&"\sent") then
				objFSO.CreateFolder(strFCT_Folder_Location&"\sent")
			end If

			'대상 폴더를 객체화
			Set objFolder	= objFSO.GetFolder(strFCT_Folder_Location)
			'해당 폴더의 파일들을 객체배열화
			Set objFiles	= objFolder.Files
			
			'파일 배열을 루핑										
			For each strFile in objFiles
				Set objFile = objFSO.getFile(strFile)
				
				'파일명 읽기
				strFileName = objFSO.GetFileName(strFile)
				
				'파일을 열어서 FCT검사결과 확인
				Set objFile2 = objFSO.OpenTextFile(strFile,1) 			 '파일 열기
				strLine	= objFile2.readLine					'라인 읽기	
				Result1 = Right(strLine,2)
				Set objFile2 = Nothing		
				
				'파일명 검증 및 FCT검사결과가 OK인 경우..
				If InStr("-687-EBR-",Left(strFileName,3)) > 0 And Len(strFileName) = 42 And Result1 = "OK" Then
					'파일 명으로 부터 정보 추출
					PRD_PartNo	= Left(strFileName,11)
					PRD_Barcode	= Left(strFileName,23)
					PRD_FCT_Date	= strDate
					PRD_FCT_Time	= Mid(strFileName,33,4)
					
					'Owner정보(=Line정보) 읽기
					Set colItems = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_LogicalFileSecuritySetting='"&strFCT_Folder_Location&"\"&strFileName&"'} Where AssocClass=Win32_LogicalFileOwner ResultRole=Owner")
					For each objItem in colItems
						PRD_Line = objItem.AccountName
					Next
					Set colItems = Nothing	
					
					'DB에 해당 바코드가 존재 하는지 체크
					SQL = "select top 1 PRD_Code from tbPWS_Raw_Data where PRD_Barcode = '"&PRD_Barcode&"'"	
					objDBRecord.Open SQL,objDBCon
					If objDBRecord.Eof Or objDBRecord.Bof Then
						Exist_YN = "N"	
					Else
						Exist_YN = "Y"
					End If
					objDBRecord.Close
					
					'FCT컬럼 생성 실패 > 사용안하는 BOX컬럼 사용.
					'DB에 존재 하지 않는다면> 새로 추가, 존재한다면> 날짜와 시간을 업데이트
					If Exist_YN = "N" Then
						SQL = "insert tbPWS_Raw_Data (PRD_PartNo,PRD_Barcode,PRD_BOX_Date,PRD_BOX_Time,PRD_Line) values "
						SQL = SQL & "('"&PRD_PartNo&"','"&PRD_Barcode&"','"&PRD_FCT_Date&"','"&PRD_FCT_Time&"','"&PRD_Line&"')"
					Else
						SQL = "update tbPWS_Raw_Data set PRD_BOX_Date = '"&PRD_FCT_Date&"', PRD_BOX_Time = '"&PRD_FCT_Time&"', PRD_Line = '"&PRD_Line&"' where PRD_Barcode = '"&PRD_Barcode&"'"
					End If
					objDBCon.execute(SQL)
					
				End If
				
				'sent'폴더에 파일이 존재 한다면
				If objFSO.FileExists(strFCT_Folder_Location&"\"&strFileName) and objFSO.FileExists(strFCT_Folder_Location&"\sent\"&strFileName) Then
					Set objFile2 = objFSO.getFile(strFCT_Folder_Location&"\sent\"&strFileName)
					
					'본파일의 수정날짜가 sent안의 파일의 수정날짜보다 이후이면.
					If objFile.DateLastModified > objFile2.DateLastModified Then
						'sent안의 파일 삭제.
						objFSO.DeleteFile (strFCT_Folder_Location&"\sent\"&strFileName)
						'파일 이동처리
						objFile.Move (strFCT_Folder_Location&"\sent\"&strFileName)
					Else
						'sent안의 파일이 본파일보다 수정날짜가 같거나 이후이면 본파일 삭제 처리
						objFSO.DeleteFile (strFCT_Folder_Location&"\"&strFileName)
						
					End If
					Set objFile2 = Nothing
				Else
					
					'sent폴더안에 파일이 없는 경우 > sent폴더 안에 파일을 삭제'
					objFile.Move (strFCT_Folder_Location&"\sent\"&strFileName)
				End If
				
				Set objFile = Nothing				
			Next
			
			Set objFiles	= Nothing
			Set objFolder	= Nothing
		
		End If	
'----------------------------------------------------------------------------------------------------------------------------------	
'----------------------------------------------------------------------------------------------------------------------------------	
		''다음 BOX 차례
		'If objFSO.FolderExists(strBOX_Folder_Location) Then
'			'sent 폴더 생성
			'if Not objFSO.FolderExists(strBOX_Folder_Location&"\sent") then
'				objFSO.CreateFolder(strBOX_Folder_Location&"\sent")
			'end If
'
			''대상 폴더를 객체화
			'Set objFolder	= objFSO.GetFolder(strBOX_Folder_Location)
			''해당 폴더의 파일들을 객체배열화
			'Set objFiles	= objFolder.Files
'			
			''파일 배열을 루핑										
			'For each strFile in objFiles
'				'파일명 읽기
				'strFileName = objFSO.GetFileName(strFile)
'							
				''Owner정보(=Line정보) 읽기
				'Set objFile = objFSO.getFile(strFile)
				'Set colItems = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_LogicalFileSecuritySetting='"&strBOX_Folder_Location&"\"&strFileName&"'} Where AssocClass=Win32_LogicalFileOwner ResultRole=Owner")
				'For each objItem in colItems
'					PRD_Line = objItem.AccountName
				'Next
				'Set colItems = Nothing	
				'Set objFile = Nothing	
'				
				''파일을 열어서 로그에 없다면, 수량 증가
				'Set objFile = objFSO.OpenTextFile(strFile,1) 			'파일 열기
'				
				'SMS_Send_Error_YN = "N"
				'strPRD_PartNo = ""
				'Do While objFile.AtEndOfStream <> True					'파일 내용을 루프
'					strLine			= objFile.readLine					'라인 읽기	
					'arrLine			= split(strLine,"|")
					'PRD_Barcode		= arrLine(0)
					'PRD_PartNo		= Left(PRD_Barcode,11)
					'PRD_BOX_Date	= strDate
					'PRD_BOX_Time	= replace(Mid(arrLine(2),12,5),":","")
'					
					'If InStr(strPRD_PartNo,PRD_PartNo) = 0 Then
'						strPRD_PartNo = strPRD_PartNo & PRD_PartNo & ","
					'End If
'					
					''DB에 해당 바코드가 존재 하는지 체크
					'SQL = "select top 1 PRD_Box_Date from tbPWS_Raw_Data where PRD_Barcode = '"&PRD_Barcode&"'"	
					'objDBRecord.Open SQL,objDBCon
					''레코드가 없다. 고로 신규등록이 필요함
					'If objDBRecord.Eof Or objDBRecord.Bof Then
'						Exist_YN = "N"	
						''SQL = "update tbBOM_Sub set BS_MAN_Qty = BS_MAN_Qty + 1 where BS_D_No = '"&PRD_PartNo&"'"
						''objDBCon.execute(SQL)
					''투입은 되었으나, 포장은 안되었음. 이번에 포장 처리 할 것임
					'ElseIf IsNull(objDBRecord("PRD_Box_Date")) Then
'						Exist_YN = "Y"
						''SQL = "update tbBOM_Sub set BS_MAN_Qty = BS_MAN_Qty + 1 where BS_D_No = '"&PRD_PartNo&"'"
						''objDBCon.execute(SQL)
					''이미 포장 된 것들
					'Else
'						Exist_YN = "Y"
					'End If
					'objDBRecord.Close
'					
					''DB에 존재 하지 않는다면> 새로 추가, 존재한다면> 날짜와 시간을 업데이트
					'If Exist_YN = "N" Then
'						SQL = "insert tbPWS_Raw_Data (PRD_PartNo,PRD_Barcode,PRD_BOX_Date,PRD_BOX_Time,PRD_Line) values "
						'SQL = SQL & "('"&PRD_PartNo&"','"&PRD_Barcode&"','"&PRD_BOX_Date&"','"&PRD_BOX_Time&"','"&PRD_Line&"')"
						'objDBCon.execute(SQL)
					'Else
'						SQL = "update tbPWS_Raw_Data set PRD_BOX_Date = '"&PRD_BOX_Date&"', PRD_BOX_Time = '"&PRD_BOX_Time&"', PRD_Line = '"&PRD_Line&"', PRD_ByHook_YN='N' where PRD_Barcode = '"&PRD_Barcode&"'"
						'objDBCon.execute(SQL)
					'End If
'					
				'Loop
'				
				'strDuplicatedPCB = ""
				'arrLine(1) = Left(arrLine(1),12)
				'arrPRD_PartNo = split(strPRD_PartNo,",")
				'If UBound(arrPRD_PartNo) > 1 Then 'ubound가 2개 이상이면 혼입
'					strDuplicatedPCB = strDuplicatedPCB & arrPRD_PartNo(0) & "," & arrPRD_PartNo(1)
					'SMS_Send_Error_YN = "Y"
				'Else '반일 파일에 혼입이 없었다면
'					'기존에 박스-PCB 체크 자료가 있는지 확인
					'SQL = "select PBPC_PCB, PBPC_Update_Date from tbPWS_Boxing_PCB_Check where PBPC_Box='"&arrLine(1)&"'"
					'objDBRecord.Open SQL,objDBCon
					'If objDBRecord.Eof Or objDBRecord.Bof Then '기존에 없다면, 레코드 추가
'						SQL = "insert into tbPWS_Boxing_PCB_Check (PBPC_BOX,PBPC_PCB,PBPC_Update_Date) values ('"&arrLine(1)&"','"&arrPRD_PartNo(0)&"','"&Date()&"')"
						'objDBCon.execute(SQL)
					'Else '기존에 있다면
'						PBPC_Update_Date = objDBRecord("PBPC_Update_Date")
						'strDuplicatedPCB = strDuplicatedPCB & objDBRecord("PBPC_PCB") & "(" & Mid(PBPC_Update_Date,6,5) & "포장), " & arrPRD_PartNo(0) & "(" & Right(strDate,5) & "포장)"
						'If objDBRecord("PBPC_PCB") <> arrPRD_PartNo(0) And DateDiff("d",PBPC_Update_Date,strDate) <= 5 Then '기존 PCB와 다르고, 날짜 차이가 5일 이상이라면
'							SMS_Send_Error_YN = "Y"
						'End If
					'End If
					'objDBRecord.Close
				'End If
'				
				'If SMS_Send_Error_YN = "Y" Then
'					strMSG = "[혼입]" & replace(PRD_Line,"pwsbox","")&"라인 "&arrLine(1)&"에 "&strDuplicatedPCB
					'wShell.Open "http://192.168.123.17:1080/vbs/sms_send.asp?strPhone="&strPhone&"&strMSG="&strMSG
				'End If
'
				'Set objFile = Nothing			
				'Set objFile = objFSO.getFile(strFile)
'				
				''Owner정보(=Line정보) 읽기
				'Set colItems = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_LogicalFileSecuritySetting='"&strBOX_Folder_Location&"\"&strFileName&"'} Where AssocClass=Win32_LogicalFileOwner ResultRole=Owner")
				'For each objItem in colItems
'					PRD_Line = objItem.AccountName
				'Next
				'Set colItems = Nothing	
'				
				''파일이동
				'Set objFile = objFSO.getFile(strFile)
				''MsgBox(strBOX_Folder_Location&"\sent\"&strFileName)
				'objFile.Move (strBOX_Folder_Location&"\sent\"&strFileName)
				'Set objFile = Nothing		
			'Next
'			
			'Set objFiles	= Nothing
			'Set objFolder	= Nothing
'		
		'End If
'----------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------
	Next
	
	Set wShell			= Nothing
	Set objFSO			= Nothing
	Set objDBRecord		= Nothing
	Set objWMIService	= Nothing
	
	SQL = "delete from tbPWS_Boxing_PCB_Check where PBPC_Update_Date < '"&DateAdd("m",-1,Date())&"'"
	objDBCon.execute(SQL)
	
	If Err = 0 Then
		DBCon_YN = "Y"
	Else
		Call Log_Err_Notice(Err.Source)
	End If
End Sub


Sub DB_Err_Notice(strErrSource)
	strPhone = ""
	strMSG = "프로그램에["&strErrSource&"]에러발생"
	'wShell.Open "http://192.168.123.17:1080/vbs/sms_send.asp?strPhone="&strPhone&"&strMSG="&strMSG
End Sub
