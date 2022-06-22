' FUNCTION:    prompts to open the most recently modified excel workbook in a designated folder
' OUTPUT:      displays a list of invalid (file.attributes 1,2,4,16) and non Excel files
' PURPOSE:     quickly open the current period journals and monitors folder contents

' File.Attributes
' 	Constant		    Value	    Description
' 	Normal			    0		    Normal file. No attributes are set.
' 	ReadOnly		    1		    Read-only file. Attribute is read/write.
' 	Hidden			    2		    Hidden file. Attribute is read/write.
' 	System			    4		    System file. Attribute is read/write.
' 	Volume			    8		    Disk drive volume label. Attribute is read-only.
' 	Directory	    	    16		    Folder or directory. Attribute is read-only.
' 	Archive			    32		    File has changed since last backup. Attribute is read/write.
' 	Alias		      	    1024	    Link or shortcut. Attribute is read-only.
' 	Compressed	   	    2048	    Compressed file. Attribute is read-only.

Option Explicit 

Dim fso, file, recentFile            ' file system objects
Dim objExcel, objWorkbook            ' Excel objects
Dim objWord                          ' Word object ( to test if objWorkbook is open )
Dim result                           ' user response to yes/no box
Dim msgText                          ' string holding ignored/invalid file names
Dim invalidFileCount, invalidName    ' files with attributes: 1, 2, 4, 16
Dim fileCount                        ' Excel workbooks
Dim ignoredCount, ignoredName        ' files that are not Excel workbooks and not invalid
Dim shell                            ' shell object to get %USERPROFILE% path
Dim strProfile                       ' user profile path (path - strFolder)
Dim strFolder                        ' short path to folder (path - strProfile)
Dim path                             ' full folder path (strProfile + strFolder)

strFolder = "\Documents\Workbooks\Batches"
strProfile = shell.ExpandEnvironmentStrings("%USERPROFILE%")
path = strProfile & strFolder
Set shell = Nothing

Set fso = CreateObject("Scripting.FileSystemObject")

Set recentFile = Nothing
fileCount = 0
invalidFileCount = 0
ignoredCount = 0

For Each file in fso.GetFolder(path).Files

	If file.Attributes And 32 Then
		If (file.Attributes And 1) Or (file.Attributes And 2) Or (file.Attributes And 4) Or (file.Attributes And 16) Then
			invalidFileCount = invalidFileCount + 1
			invalidName = file.Name & vbNewLine & invalidName
		Else
			If InStr(file.Type, "Microsoft Excel") = 0 Then
				ignoredCount = ignoredCount + 1
				ignoredName = file.Name & vbNewLine & ignoredName
			Else
				If (recentFile is Nothing) Then
					Set recentFile = file
				ElseIf (file.DateLastModified > recentFile.DateLastModified) Then
					Set recentFile = file
				End If
				fileCount = fileCount + 1
			End If
		End If
	Else
		invalidFileCount = invalidFileCount + 1
		invalidName = file.Name & vbNewLine & invalidName
	End If
Next

msgText = vbNewLine & vbNewLine

If ignoredCount > 0 Then
	msgText = msgText & "List of " & ignoredCount & " files that are not workbooks:" & vbNewLine & ignoredName
	msgText = msgText & vbNewLine & vbNewLine
End If

If invalidFileCount > 0 Then
	msgText = msgText & "List of " & invalidFileCount & " invalid files:" & vbNewLine & invalidName
End If

If recentFile Is Nothing Then
	Msgbox msgText, vbInformation, "No Valid Files Found" 
Else
	Set objWord = CreateObject("Word.Application")
	If objWord.Tasks.Exists(recentFile.name) Then 
		Msgbox "File is already open." & msgText, vbExclamation, "Operation Aborted"
	Else
		result = Msgbox(recentFile.Name & vbNewLine & vbNewLine & path _
				& vbNewLine & vbNewLine & fileCount & " files found." & msgText, vbYesNo + vbQuestion, "Open File")
		Select Case result
			Case vbYes
				Set objExcel = CreateObject("Excel.Application")
				objExcel.Visible = True
				Set objWorkbook = objExcel.Workbooks.Open(path & "\" & recentFile.Name)
				If Not objWord.Tasks.Exists(recentFile.name) Then
					Msgbox "Unable to verify file is currently open.", vbExclamation, "Error"
				End If
			Case vbNo
				Msgbox "User quit without opening file.", vbInformation, "Operation Aborted"
		End Select
	End If
	
	objWord.Quit
	
End If
