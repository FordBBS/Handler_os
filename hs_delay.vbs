Option Explicit

Function hs_delay(ByVal intSec)
	'*** History ***********************************************************************************
	' 2020/11/04, BBS:	- First release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' Delay the process for a specific second(s)
	'  
	'***********************************************************************************************

	'*** Initialization ****************************************************************************
	Dim wshShell, strCmd
	Set wshShell = CreateObject( "WScript.Shell" )
	
	intSec = CStr(CInt(intSec))
	strCmd = wshShell.ExpandEnvironmentStrings( "%COMSPEC% /C (Timeout.EXE /T " & intSec & " /NOBREAK)" )

	'*** Operations ********************************************************************************
	wshShell.Run strCmd, 0, 1
	
	'--- Release -----------------------------------------------------------------------------------
	Set wshShell = Nothing
	hs_delay     = True
End Function
