Function hs_getlist_under_path(ByVal strRootPath)
	'*** History ***********************************************************************************
	' 2020/08/29, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Get a list of all materials under 'strRootPath' directory path
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

	'*** Operations ********************************************************************************
	'--- Check whether 'strRootPath' is a file path ------------------------------------------------
	If objFSO.FileExists(strRootPath) Then
		hs_getlist_under_path = strRootPath
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
		RetVal = hs_getlist_under_path(objThis.Path)
		
		For each thisPath in RetVal
			Call hs_arr_append(arrRes, thisPath)
		Next
	Next

	' Collect files and append to Result array
	For each objThis in arrFile
	    Call hs_arr_append(arrRes, Join(Array(strRootPath, objThis.Name), chr_join))
	Next

	' Append 'strRootPath' in case it has no subFile or subFolder anymore (end of path)
	If arrFile.count + arrFolder.count = 0 Then
		Call hs_arr_append(arrRes, strRootPath)
	End If

	'--- Release -----------------------------------------------------------------------------------
	hs_getlist_under_path = arrRes

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
