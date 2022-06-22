' FUNCTION   - Change the last modified date of the outlook VBA file
' PURPOSE    - Prevent loss of data when outlooks VBA file is overwritten
' BACKGROUND - the last modified date of Outlook's VBA file is not consistently updated when changes are made and saved
'              this is problematic when using roaming pofiles in a hosted environment because the most recent file can
'              be overwritten by an older file and result in lost data and irretrevable data
' MORE INFO  - Outlook cannot be the VBA host to update the last modified date because the file cannot modified while it
'              it is open; so this script is called with a shell command when Outlooks "Application_Quit" event fires.

' example call
'   Private Sub Application_Quit()
'       Dim scriptPath as String
'       scriptPath = Chr(34) & Environ("userprofile") & "\scripts\ChangeLastModifiedDate.vbs" & Chr(34) & Chr(34) & Chr(34)
'       Shell "wscript " & scriptPath
'   End Sub

Option Explicit

    Dim Shell
    Set Shell = WScript.CreateObject("WScript.Shell")
    
    Dim strAppData
    strAppData = shell.ExpandEnvironmentStrings("%APPDATA%")

	Dim longName, loopLimit, sleepTime
	longName = strAppData & "\Microsoft\Outlook\VbaProject.OTM"
	loopLimit = 50
	sleepTime = 500

	Dim fsoObject, fsoFile
	Set fsoObject = CreateObject("Scripting.FileSystemObject")
	Set fsoFile = fsoObject.GetFile(longName)
	
	Dim shellApp, shellFolder, shellFile
	Set shellApp = CreateObject("Shell.Application")
	Set shellFolder = shellApp.NameSpace("" & fsoFile.ParentFolder & "")
	Set shellFile = shellFolder.ParseName(fsoFile.Name)
	
	Dim timeStamp, loopCounter
	timeStamp = Now
	loopCounter = 0
	
	Do
		shellFile.ModifyDate = timeStamp
		wScript.Sleep sleepTime
		loopCounter = loopCounter + 1
	Loop Until fsoFile.DateLastModified = timeStamp Or loopCounter > loopLimit

	If fsoFile.DateLastModified <> timeStamp Then
		wScript.Echo "Not changed!!  Loop counter = " & loopCounter & "/" & loopLimit
	Else
		wScript.Echo "OTM date stamp changed to " & fsoFile.DateLastModified
	End If

    Set shell = Nothing
    Set fsoObject = Nothing
    Set fsoFile = Nothing
    Set shellApp = Nothing
    Set shellFolder = Nothing
    Set shellFile = Nothing
    
WScript.Quit
