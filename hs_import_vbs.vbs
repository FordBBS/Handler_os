Function hs_import_vbs(ByVal pathDir)
	'### hs_import_vbs, BBS ########################################################################
	'*** History ***********************************************************************************
	' 2020/11/04, BBS:	- First release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' 	Imports all VBS script file provided in 'pathDir' via 'ExecuteGlobal'
	' 	If any 'directory' is found, all VBS under that directory will be imported too.
	'
	'***********************************************************************************************

	On Error Resume Next
	hs_import_vbs = False

	'*** Initialization ****************************************************************************
	Dim objFSO, objFile, thisPath, arrPath, arrFiles, thisFile
	Dim flg_do, cntExe, cntRet

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	cntExe   = 0
	cntRet   = 0

	'*** Operations ********************************************************************************
	'--- Path's array preparation ------------------------------------------------------------------
	If IsArray(pathDir) Then
		arrPath = pathDir
	Else
		arrPath = Array(CStr(pathDir))
	End If

	'--- Importing ---------------------------------------------------------------------------------
	For Each thisPath in arrPath
		flg_do = True

		' File Path, no need to go deeper
		If objFSO.FileExists(thisPath) Then
			arrFiles = Array(thisPath)

		' Folder Path, need to collect every file under it
		ElseIf objFSO.FolderExists(thisPath) Then
			arrFiles = hs_getlist_under_path(thisPath, 2)
			Err.Clear

		Else
			flg_do = False
		End If

		If flg_do Then
			For Each thisFile in arrFiles
				On Error Resume Next

				Set objFile = objFSO.OpenTextFile(thisFile)
				ExecuteGlobal objFile.ReadAll()
				objFile.Close

				cntExe = cntExe + 1

				If Err.Number = 0 Then
					cntRet = cntRet + 1
				End If

				Err.Clear
			Next
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	If cntRet = cntExe Then
		hs_import_vbs = True
	End If

	'*** Error Handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function