On Error Resume Next

'DB����
Set objDBCon = CreateObject("ADODB.Connection")
Set wShell = CreateObject("Shell.Application")

'strMSG = "[ȥ��] cell���� KR0512827758�� 6871A20572Q(03-15����), 6871A20572A(03-25����)"
'wShell.Open "http://192.168.123.17:1080/vbs/sms_send.asp?strPhone=01067555740&strMSG="&strMSG

'������;�����;������
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

'DB����
objDBCon.Close
Set objDBCon		= Nothing


Sub PWS_Data_Handling()
	'�ʿ��� ��ü ����
	Set objWMIService	= getObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set objDBRecord		= CreateObject("ADODB.RecordSet")
	Set objFSO 			= CreateObject("Scripting.FileSystemObject")

	'��ǲ�� ���� ������ ����
	Dim PRD_Code
	Dim PRD_PartNo
	Dim PRD_Barcode
	Dim PI_ICT_Date
	Dim PI_FCT_Date
	Dim PI_BOX_Date
	Dim arrLine
	
	Dim Result1
	Dim Result2

	'���÷κ��� 5�� ������ �����͸� �б� ���� ����
	For nDateCNT = 0 To -5 step -1
	
		'������ ��¥
		strDate = DateAdd("d",nDateCNT,Date())
		
		'������ ���� ����
		strICT_Folder_Location = "D:\LG-PWS\TDATA\ICT\ODATA\" & replace(strDate,"-","")
		strFCT_Folder_Location = "D:\LG-PWS\TDATA\FCT\ODATA\" & replace(strDate,"-","")
		strBox_Folder_Location = "D:\LG-PWS\TDATA\BOX\ODATA\" & replace(strDate,"-","")
'----------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------	
		''�켱 ICT���� ó���Ѵ�.
'		If objFSO.FolderExists(strICT_Folder_Location) Then
'			'sent ���� ����
'			if Not objFSO.FolderExists(strICT_Folder_Location&"\sent") then
'				objFSO.CreateFolder(strICT_Folder_Location&"\sent")
'			end If
'
'			'��� ������ ��üȭ
'			Set objFolder	= objFSO.GetFolder(strICT_Folder_Location)
'			'�ش� ������ ���ϵ��� ��ü�迭ȭ
'			Set objFiles	= objFolder.Files
'			
'			'���� �迭�� ����										
'			For each strFile in objFiles
'				Set objFile = objFSO.getFile(strFile)
'				
'				'���ϸ� �б�
'				strFileName = objFSO.GetFileName(strFile)
'				
'				'������ ��� ICT�˻��� Ȯ��
'				Set objFile2 = objFSO.OpenTextFile(strFile,1) 			 '���� ����
'				strLine	= objFile2.readLine					'���� �б�	
'				arrLine	= split(strLine,"|")
'				Result1 = arrLine(19)
'				strLine	= objFile2.readLine					'���� �б�	
'				arrLine	= split(strLine,"|")
'				Result2 = arrLine(19)
'				Set objFile2 = Nothing			
'				
'				'���ϸ� ���� �� ICT�˻����� OK�� ���..
'				If InStr("-687-EBR-",Left(strFileName,3)) > 0 And Len(strFileName) = 42 And Result1 = "OK" And Result2 = "OK" Then
'					'���� ������ ���� ���� ����
'					PRD_PartNo	= Left(strFileName,11)
'					PRD_Barcode	= Left(strFileName,23)
'					PRD_ICT_Date	= strDate
'					PRD_ICT_Time	= Mid(strFileName,33,4)
'					
'					'Owner����(=Line����) �б�
'					Set colItems = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_LogicalFileSecuritySetting='"&strICT_Folder_Location&"\"&strFileName&"'} Where AssocClass=Win32_LogicalFileOwner ResultRole=Owner")
'					For each objItem in colItems
'						PRD_Line = objItem.AccountName
'					Next
'					Set colItems = Nothing	
'					
'					'DB�� �ش� ���ڵ尡 ���� �ϴ��� üũ
'					SQL = "select top 1 PRD_Code from tbPWS_Raw_Data where PRD_Barcode = '"&PRD_Barcode&"'"	
'					objDBRecord.Open SQL,objDBCon
'					If objDBRecord.Eof Or objDBRecord.Bof Then
'						Exist_YN = "N"	
'					Else
'						Exist_YN = "Y"
'					End If
'					objDBRecord.Close
'					
'					'DB�� ���� ���� �ʴ´ٸ�> ���� �߰�, �����Ѵٸ�> ��¥�� �ð��� ������Ʈ
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
'				'sent'������ ������ ���� �Ѵٸ�
'				If objFSO.FileExists(strICT_Folder_Location&"\"&strFileName) and objFSO.FileExists(strICT_Folder_Location&"\sent\"&strFileName) Then
'					Set objFile2 = objFSO.getFile(strICT_Folder_Location&"\sent\"&strFileName)
'					
'					'�������� ������¥�� sent���� ������ ������¥���� �����̸�.
'					If objFile.DateLastModified > objFile2.DateLastModified Then
'						'sent���� ���� ����.
'						objFSO.DeleteFile (strICT_Folder_Location&"\sent\"&strFileName)
'						'���� �̵�ó��
'						objFile.Move (strICT_Folder_Location&"\sent\"&strFileName)
'					Else
'						'sent���� ������ �����Ϻ��� ������¥�� ���ų� �����̸� ������ ���� ó��
'						objFSO.DeleteFile (strICT_Folder_Location&"\"&strFileName)
'						
'					End If
'					Set objFile2 = Nothing
'				Else
'					
'					'sent�����ȿ� ������ ���� ��� > sent���� �ȿ� ������ ����'
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
	'�켱 FCT���� ó���Ѵ�.
		If objFSO.FolderExists(strFCT_Folder_Location) Then
			'sent ���� ����
			if Not objFSO.FolderExists(strFCT_Folder_Location&"\sent") then
				objFSO.CreateFolder(strFCT_Folder_Location&"\sent")
			end If

			'��� ������ ��üȭ
			Set objFolder	= objFSO.GetFolder(strFCT_Folder_Location)
			'�ش� ������ ���ϵ��� ��ü�迭ȭ
			Set objFiles	= objFolder.Files
			
			'���� �迭�� ����										
			For each strFile in objFiles
				Set objFile = objFSO.getFile(strFile)
				
				'���ϸ� �б�
				strFileName = objFSO.GetFileName(strFile)
				
				'������ ��� FCT�˻��� Ȯ��
				Set objFile2 = objFSO.OpenTextFile(strFile,1) 			 '���� ����
				strLine	= objFile2.readLine					'���� �б�	
				Result1 = Right(strLine,2)
				Set objFile2 = Nothing		
				
				'���ϸ� ���� �� FCT�˻����� OK�� ���..
				If InStr("-687-EBR-",Left(strFileName,3)) > 0 And Len(strFileName) = 42 And Result1 = "OK" Then
					'���� ������ ���� ���� ����
					PRD_PartNo	= Left(strFileName,11)
					PRD_Barcode	= Left(strFileName,23)
					PRD_FCT_Date	= strDate
					PRD_FCT_Time	= Mid(strFileName,33,4)
					
					'Owner����(=Line����) �б�
					Set colItems = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_LogicalFileSecuritySetting='"&strFCT_Folder_Location&"\"&strFileName&"'} Where AssocClass=Win32_LogicalFileOwner ResultRole=Owner")
					For each objItem in colItems
						PRD_Line = objItem.AccountName
					Next
					Set colItems = Nothing	
					
					'DB�� �ش� ���ڵ尡 ���� �ϴ��� üũ
					SQL = "select top 1 PRD_Code from tbPWS_Raw_Data where PRD_Barcode = '"&PRD_Barcode&"'"	
					objDBRecord.Open SQL,objDBCon
					If objDBRecord.Eof Or objDBRecord.Bof Then
						Exist_YN = "N"	
					Else
						Exist_YN = "Y"
					End If
					objDBRecord.Close
					
					'FCT�÷� ���� ���� > �����ϴ� BOX�÷� ���.
					'DB�� ���� ���� �ʴ´ٸ�> ���� �߰�, �����Ѵٸ�> ��¥�� �ð��� ������Ʈ
					If Exist_YN = "N" Then
						SQL = "insert tbPWS_Raw_Data (PRD_PartNo,PRD_Barcode,PRD_BOX_Date,PRD_BOX_Time,PRD_Line) values "
						SQL = SQL & "('"&PRD_PartNo&"','"&PRD_Barcode&"','"&PRD_FCT_Date&"','"&PRD_FCT_Time&"','"&PRD_Line&"')"
					Else
						SQL = "update tbPWS_Raw_Data set PRD_BOX_Date = '"&PRD_FCT_Date&"', PRD_BOX_Time = '"&PRD_FCT_Time&"', PRD_Line = '"&PRD_Line&"' where PRD_Barcode = '"&PRD_Barcode&"'"
					End If
					objDBCon.execute(SQL)
					
				End If
				
				'sent'������ ������ ���� �Ѵٸ�
				If objFSO.FileExists(strFCT_Folder_Location&"\"&strFileName) and objFSO.FileExists(strFCT_Folder_Location&"\sent\"&strFileName) Then
					Set objFile2 = objFSO.getFile(strFCT_Folder_Location&"\sent\"&strFileName)
					
					'�������� ������¥�� sent���� ������ ������¥���� �����̸�.
					If objFile.DateLastModified > objFile2.DateLastModified Then
						'sent���� ���� ����.
						objFSO.DeleteFile (strFCT_Folder_Location&"\sent\"&strFileName)
						'���� �̵�ó��
						objFile.Move (strFCT_Folder_Location&"\sent\"&strFileName)
					Else
						'sent���� ������ �����Ϻ��� ������¥�� ���ų� �����̸� ������ ���� ó��
						objFSO.DeleteFile (strFCT_Folder_Location&"\"&strFileName)
						
					End If
					Set objFile2 = Nothing
				Else
					
					'sent�����ȿ� ������ ���� ��� > sent���� �ȿ� ������ ����'
					objFile.Move (strFCT_Folder_Location&"\sent\"&strFileName)
				End If
				
				Set objFile = Nothing				
			Next
			
			Set objFiles	= Nothing
			Set objFolder	= Nothing
		
		End If	
'----------------------------------------------------------------------------------------------------------------------------------	
'----------------------------------------------------------------------------------------------------------------------------------	
		''���� BOX ����
		'If objFSO.FolderExists(strBOX_Folder_Location) Then
'			'sent ���� ����
			'if Not objFSO.FolderExists(strBOX_Folder_Location&"\sent") then
'				objFSO.CreateFolder(strBOX_Folder_Location&"\sent")
			'end If
'
			''��� ������ ��üȭ
			'Set objFolder	= objFSO.GetFolder(strBOX_Folder_Location)
			''�ش� ������ ���ϵ��� ��ü�迭ȭ
			'Set objFiles	= objFolder.Files
'			
			''���� �迭�� ����										
			'For each strFile in objFiles
'				'���ϸ� �б�
				'strFileName = objFSO.GetFileName(strFile)
'							
				''Owner����(=Line����) �б�
				'Set objFile = objFSO.getFile(strFile)
				'Set colItems = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_LogicalFileSecuritySetting='"&strBOX_Folder_Location&"\"&strFileName&"'} Where AssocClass=Win32_LogicalFileOwner ResultRole=Owner")
				'For each objItem in colItems
'					PRD_Line = objItem.AccountName
				'Next
				'Set colItems = Nothing	
				'Set objFile = Nothing	
'				
				''������ ��� �α׿� ���ٸ�, ���� ����
				'Set objFile = objFSO.OpenTextFile(strFile,1) 			'���� ����
'				
				'SMS_Send_Error_YN = "N"
				'strPRD_PartNo = ""
				'Do While objFile.AtEndOfStream <> True					'���� ������ ����
'					strLine			= objFile.readLine					'���� �б�	
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
					''DB�� �ش� ���ڵ尡 ���� �ϴ��� üũ
					'SQL = "select top 1 PRD_Box_Date from tbPWS_Raw_Data where PRD_Barcode = '"&PRD_Barcode&"'"	
					'objDBRecord.Open SQL,objDBCon
					''���ڵ尡 ����. ��� �űԵ���� �ʿ���
					'If objDBRecord.Eof Or objDBRecord.Bof Then
'						Exist_YN = "N"	
						''SQL = "update tbBOM_Sub set BS_MAN_Qty = BS_MAN_Qty + 1 where BS_D_No = '"&PRD_PartNo&"'"
						''objDBCon.execute(SQL)
					''������ �Ǿ�����, ������ �ȵǾ���. �̹��� ���� ó�� �� ����
					'ElseIf IsNull(objDBRecord("PRD_Box_Date")) Then
'						Exist_YN = "Y"
						''SQL = "update tbBOM_Sub set BS_MAN_Qty = BS_MAN_Qty + 1 where BS_D_No = '"&PRD_PartNo&"'"
						''objDBCon.execute(SQL)
					''�̹� ���� �� �͵�
					'Else
'						Exist_YN = "Y"
					'End If
					'objDBRecord.Close
'					
					''DB�� ���� ���� �ʴ´ٸ�> ���� �߰�, �����Ѵٸ�> ��¥�� �ð��� ������Ʈ
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
				'If UBound(arrPRD_PartNo) > 1 Then 'ubound�� 2�� �̻��̸� ȥ��
'					strDuplicatedPCB = strDuplicatedPCB & arrPRD_PartNo(0) & "," & arrPRD_PartNo(1)
					'SMS_Send_Error_YN = "Y"
				'Else '���� ���Ͽ� ȥ���� �����ٸ�
'					'������ �ڽ�-PCB üũ �ڷᰡ �ִ��� Ȯ��
					'SQL = "select PBPC_PCB, PBPC_Update_Date from tbPWS_Boxing_PCB_Check where PBPC_Box='"&arrLine(1)&"'"
					'objDBRecord.Open SQL,objDBCon
					'If objDBRecord.Eof Or objDBRecord.Bof Then '������ ���ٸ�, ���ڵ� �߰�
'						SQL = "insert into tbPWS_Boxing_PCB_Check (PBPC_BOX,PBPC_PCB,PBPC_Update_Date) values ('"&arrLine(1)&"','"&arrPRD_PartNo(0)&"','"&Date()&"')"
						'objDBCon.execute(SQL)
					'Else '������ �ִٸ�
'						PBPC_Update_Date = objDBRecord("PBPC_Update_Date")
						'strDuplicatedPCB = strDuplicatedPCB & objDBRecord("PBPC_PCB") & "(" & Mid(PBPC_Update_Date,6,5) & "����), " & arrPRD_PartNo(0) & "(" & Right(strDate,5) & "����)"
						'If objDBRecord("PBPC_PCB") <> arrPRD_PartNo(0) And DateDiff("d",PBPC_Update_Date,strDate) <= 5 Then '���� PCB�� �ٸ���, ��¥ ���̰� 5�� �̻��̶��
'							SMS_Send_Error_YN = "Y"
						'End If
					'End If
					'objDBRecord.Close
				'End If
'				
				'If SMS_Send_Error_YN = "Y" Then
'					strMSG = "[ȥ��]" & replace(PRD_Line,"pwsbox","")&"���� "&arrLine(1)&"�� "&strDuplicatedPCB
					'wShell.Open "http://192.168.123.17:1080/vbs/sms_send.asp?strPhone="&strPhone&"&strMSG="&strMSG
				'End If
'
				'Set objFile = Nothing			
				'Set objFile = objFSO.getFile(strFile)
'				
				''Owner����(=Line����) �б�
				'Set colItems = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_LogicalFileSecuritySetting='"&strBOX_Folder_Location&"\"&strFileName&"'} Where AssocClass=Win32_LogicalFileOwner ResultRole=Owner")
				'For each objItem in colItems
'					PRD_Line = objItem.AccountName
				'Next
				'Set colItems = Nothing	
'				
				''�����̵�
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
	strMSG = "���α׷���["&strErrSource&"]�����߻�"
	'wShell.Open "http://192.168.123.17:1080/vbs/sms_send.asp?strPhone="&strPhone&"&strMSG="&strMSG
End Sub
