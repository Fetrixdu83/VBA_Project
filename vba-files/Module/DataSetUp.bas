Attribute VB_Name = "DataSetUp"
Option Explicit

Private mSelectedFolderPath As String

'Function to set up the folder path and create sheets for each file in the folder
Public Sub Fetch_Data_FromFolder()
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


	LoopFileInFolder
End Sub
' Sub function to loop through files in the selected folder and perform actions
Private Sub LoopFileInFolder()
	Dim folderPath As String
	Dim fileName As String
	Dim row As Integer
	folderPath = GetSelectedFolderPath()
	fileName = Dir(folderPath & "*.xl??")

	row = 1
	do While fileName <> ""
		CreateSheet fileName ' Create a new sheet for each file
		CopyDataFileToSheet folderPath & fileName ' Copy data from the file to the sheet
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
	Dim sheetName As String
	Dim ws As Worksheet
	Dim sourceWB As Workbook
	Dim sourceWS As Worksheet
	Dim sourceRange As Range
	Dim lastRow As Long
	Dim lastCol As Long

	sheetName = filename
	If InStrRev(filename, "\") > 0 Then ' Get the file name without the path
		sheetName = Mid$(filename, InStrRev(filename, "\") + 1)
	End If

	On Error Resume Next
	Set ws = ThisWorkbook.Sheets(sheetName)
	On Error GoTo 0

	If ws Is Nothing Then
		MsgBox "Destination sheet '" & sheetName & "' was not found.", vbExclamation, "Import Error"
		Exit Sub
	End If

	On Error GoTo CleanFail
	Set sourceWB = Workbooks.Open(filename, ReadOnly:=True) ' Open the source workbook as read-only
	Set sourceWS = sourceWB.Worksheets(1) ' Assuming data is in the first sheet
	
	lastRow = sourceWS.Cells(sourceWS.Rows.Count, "A").End(xlUp).Row
	lastCol = sourceWS.Cells(1, sourceWS.Columns.Count).End(xlToLeft).Column

	If lastRow < 1 Or lastCol < 1 Then GoTo CleanFail

	Set sourceRange = sourceWS.Range(sourceWS.Cells(1, 1), sourceWS.Cells(lastRow, lastCol)) ' Define the range to copy

	ws.Cells.Clear ' Clear existing data in the destination sheet
	sourceRange.Copy Destination:=ws.Range("A1") ' Copy data to the destination sheet

CleanExit: ' Ensure the source workbook is closed
	If Not sourceWB Is Nothing Then sourceWB.Close SaveChanges:=False
	On Error GoTo 0
	Exit Sub
		
CleanFail: ' Handle failure in copying data
	MsgBox "Unable to copy data from file: " & filename, vbExclamation, "Import Error"
	Resume CleanExit
End Sub

