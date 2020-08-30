Function hs_getlist_under_path(ByVal strRootPath, ByVal getMode)
	'*** History ***********************************************************************************
	' 2020/08/29, BBS:	- First Release
	' 2020/08/30, BBS:	- Implemented 'getMode' feature
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Get a list of all materials under 'strRootPath' directory path, materials to be collected
	'	depends on 'getMode' where 0: All materials, 1: Empty Folder only, 2: File only
	'
	' 	If 'strRootPath' is a file path then 'strRootPath' is returned
	'	If there is nothing under 'strRootPath' then "" is returned
	'
	'***********************************************************************************************
	
	On Error Resume Next
	hs_getlist_under_path = ""

	'*** Pre-Validation ****************************************************************************
	strRootPath = CStr(strRootPath)
	If len(strRootPath) < 1 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim objFSO, objFolder, objThis, chr_join, thisPath
	Dim RetVal, arrFolder, arrFile, arrRes()
	Redim Preserve arrRes(0)

	Set objFSO  = CreateObject("Scripting.FileSystemObject")
	chr_join    = "\"
	strRootPath = Replace(strRootPath, "/", chr_join)

	If Not IsNumeric(getMode) Then
		getMode = 0
	Else
		getMode = CInt(getMode)
	End If

	If getMode < 0 or getMode > 2 Then getMode = 0

	'*** Operations ********************************************************************************
	'--- Check whether 'strRootPath' is a file path ------------------------------------------------
	If objFSO.FileExists(strRootPath) Then
		If getMode <> 1 Then hs_getlist_under_path = strRootPath
		Exit Function
	End If

	'--- Check if 'strRootPath' does exist ---------------------------------------------------------
	If Not objFSO.FolderExists(strRootPath) Then
		Exit Function
	End If

	'--- Get materials under 'strRootPath' ---------------------------------------------------------
	If Right(strRootPath, 1) = "\" Then
		strRootPath = Left(strRootPath, len(strRootPath) - 1)
	End If

	Set objFolder = objFSO.GetFolder(strRootPath)
	Set arrFolder = objFolder.SubFolders
	Set arrFile   = objFolder.Files
	
	' Collect 'SubFolder' and append to Result array - recursive
	For each objThis in arrFolder
		RetVal = hs_getlist_under_path(objThis.Path, getMode)
		
		For each thisPath in RetVal
			If (objFSO.FileExists(thisPath) and getMode <> 1) _
				or (Not objFSO.FileExists(thisPath) and getMode <> 2) Then
				Call hs_arr_append(arrRes, thisPath)
			End If
		Next
	Next

	' Collect files and append to Result array
	If getMode <> 1 Then
		For each objThis in arrFile
		    Call hs_arr_append(arrRes, Join(Array(strRootPath, objThis.Name), chr_join))
		Next
	End If

	' Append 'strRootPath' in case it has no subFile or subFolder anymore (end of path)
	If getMode <> 2 Then
		If arrFile.count + arrFolder.count = 0 Then
			Call hs_arr_append(arrRes, strRootPath)
		End If
	End If

	'--- Release -----------------------------------------------------------------------------------
	hs_getlist_under_path = arrRes

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function