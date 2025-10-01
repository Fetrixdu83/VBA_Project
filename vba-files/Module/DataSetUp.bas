Attribute VB_Name = "DataSetUp"
Option Explicit

Private mSelectedFolderPath As String

'Function to set up the folder path and create sheets for each file in the folder
Public Sub DataSetUp_SelectFolder()
	Dim dlg As FileDialog
	Dim result As Long
	Dim folderPath As String

	Set dlg = Application.FileDialog(msoFileDialogFolderPicker)

	With dlg
		.Title = "Select a folder"
		.AllowMultiSelect = False

		result = .Show

		If result <> -1 Then
			MsgBox "No folder was selected.", vbInformation, "Folder Picker"
			Exit Sub
		End If

		folderPath = .SelectedItems(1)
	End With

	If Right$(folderPath, 1) <> "\" Then
		folderPath = folderPath & "\"
	End If

	mSelectedFolderPath = folderPath

	MsgBox "Selected folder path:" & vbCrLf & folderPath, vbInformation, "Folder Picker"

	GetFileInFolder
End Sub
' Sub function to get files in the selected folder and create sheets for each file
Private Sub GetFileInFolder()
	Dim folderPath As String
	Dim fileName As String
	Dim row As Integer
	folderPath = GetSelectedFolderPath()
	fileName = Dir(folderPath & "*.xl??")

	row = 1
	do While fileName <> ""
		CreateSheet fileName
		row = row + 1
		fileName = Dir()
	Loop
End Sub

Private Function GetSelectedFolderPath() As String
	if mSelectedFolderPath = "" Then
		MsgBox "No folder has been selected yet. Please select a folder first.", vbExclamation, "Folder Not Selected"
		Exit Function
	End If
	GetSelectedFolderPath = mSelectedFolderPath
End Function

Private Sub CreateSheet(sheet As String)
	On Error Resume Next
	Dim ws As Worksheet
	Set ws = ThisWorkbook.Sheets(sheet)
	
	If ws Is Nothing Then
		ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = sheet
	End If
	
	On Error GoTo 0
End Sub

Private Sub CopyDataFileToSheet(filename As String)

	
End Sub