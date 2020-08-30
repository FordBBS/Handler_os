'*** History ***************************************************************************************
' 2020/08/29, BBS:	- First Release
' 					- Imported all mandatory materials
'
'***************************************************************************************************

'*** Imported Materials ****************************************************************************
'--- Documentation ---------------------------------------------------------------------------------
' (Version 2020/08/29) hs_import_vbs
' (Version 2020/08/30) hs_getlist_under_path
' (Version 2020/08/25) hs_arr_append
'
'---------------------------------------------------------------------------------------------------

Function hs_import_vbs(ByVal inpPath)
	'*** History ***********************************************************************************
	' 2020/08/29, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Imports all VBS script file provided in 'inpPath' via 'ExecuteGlobal'
	' 	If any 'directory' is found, all VBS under that directory will be imported too.
	'
	'***********************************************************************************************
	
	On Error Resume Next
	hs_import_vbs = False

	'*** Initialization ****************************************************************************
	Dim objFSO, objFile, arrPath, arrVbs, thisPath, thisFile, thisVbs, strMemoVbs
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strMemoVbs = ""

	'*** Operations ********************************************************************************
	'--- Path's array preparation ------------------------------------------------------------------
	If InStr(LCase(TypeName(inpPath)), "variant") = 0 Then
		arrPath = Array(CStr(inpPath))
	Else
		arrPath = inpPath
	End If

	'--- Importing ---------------------------------------------------------------------------------
	For Each thisPath in arrPath
		If objFSO.FileExists(thisPath) Then
			RetVal = Array(thisPath)
		Else
			RetVal = hs_getlist_under_path(thisPath, 2)
		End If

		For Each thisFile in RetVal
			If InStr(thisFile, ".vbs") > 0 Then
				thisVbs = objFSO.GetFile(thisFile).Name
				thisVbs = "%" & thisVbs & "%"

				If InStr(strMemoVbs, thisVbs) = 0 Then
					Set objFile = objFSO.OpenTextFile(thisFile, 1)
					ExecuteGlobal objFile.ReadAll()
					objFile.Close()
					Set objFile = Nothing

					If strMemoVbs = "" Then
						strMemoVbs = thisVbs
					Else
						strMemoVbs = strMemoVbs & ";" & thisVbs
					End If
				End If
			End If
		Next
	Next

	'--- Release -----------------------------------------------------------------------------------
	Set objFSO  = Nothing

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	Else
		hs_import_vbs = True
	End If
End Function

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

Function hs_arr_append(ByRef arrInput, ByVal tarValue)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First release
	' 2020/08/25, BBS:  - Implemented handler for Non-Array 'arrInput'
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Append 'tarValue' to target array provided as 'arrInput', 'arrInput' can be only a single
	'	column array only
	'
	'	Argument(s)
	'	<Array>  arrInput, Base array to be appended 'tarValue'
	'	<Any> 	 tarValue, Desire value to be appended to 'arrInput'
	'
	'***********************************************************************************************
	
	On Error Resume Next

	'*** Initialization ****************************************************************************
	' Nothing to be initialized

	'*** Operations ********************************************************************************
	'--- Ensure 'arrInput' is Array type before doing appending ------------------------------------
	If InStr(LCase(TypeName(arrInput)), "variant") = 0 Then
		arrInput = Array(arrInput)
	End If

	'--- Appending ---------------------------------------------------------------------------------
	If Not (UBound(arrInput) = 0 and LCase(TypeName(arrInput(0))) = "empty") Then
		Redim Preserve arrInput(UBound(arrInput) + 1)
	End If

	arrInput(UBound(arrInput)) = tarValue

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
'***************************************************************************************************

'*** Local Material ********************************************************************************
' No local material yet
'***************************************************************************************************
