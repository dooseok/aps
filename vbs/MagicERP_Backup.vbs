Set objFSO = CreateObject ("Scripting.FileSystemObject")
If Not objFSO.FolderExists("d:\My_Backup\MagicERP\"&date()) then
	objFSO.CreateFolder("d:\My_Backup\MagicERP\"&date())
end If

Set srcFolder = objFSO.GetFolder("g:\e\")
srcFolder.Copy("d:\My_Backup\MagicERP\"&Date())
Set srcFolder = Nothing
Set objFSO = Nothing