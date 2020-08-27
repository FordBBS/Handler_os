Function hs_read_text_file(ByVal strPathFile)
	'*** History ***********************************************************************************
	' 2020/08/27, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return a content exists in the file at 'strPathFile' in String type
	'
	'***********************************************************************************************
	On Error Resume Next
	hs_read_text_file = ""

	'*** Pre-Validation ****************************************************************************
	strPathFile = CStr(strPathFile)
	If len(strPathFile) < 1 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim objFSO, objFile, strContent
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'*** Operations ********************************************************************************
	'--- Validate the existence of target file -----------------------------------------------------
	If Not objFSO.FileExists(strPathFile) Then Exit Function

	'--- Read target file --------------------------------------------------------------------------
	Set objFile = objFSO.OpenTextFile(strPathFile, 1)
	strContent  = objFile.ReadAll()
	
	'--- Release -----------------------------------------------------------------------------------
	hs_read_text_file = strContent
	Set objFile = Nothing
	Set objFSO  = Nothing

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function