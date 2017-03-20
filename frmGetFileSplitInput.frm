VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGetFileSplitInput 
   Caption         =   "Enter Split Options"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8520
   OleObjectBlob   =   "frmGetFileSplitInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGetFileSplitInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================
'Name: btnCancel
'Purpose: Cancels the process and unloads the form.
'==============================
Private Sub btnCancel_Click()
Unload Me
End Sub
'==============================
'Name btnOk_Click
'Purpose: Launches then main SplitData subroutine
'when the Ok button is clicked. Triggers form validation
'code before proceeding.
'==============================
Private Sub btnOk_Click()
Dim strDataSourceFileName As String
Dim booApplyBasicFormatting As Boolean
Dim booWarnBeforeOverwritingFiles As Boolean
Dim intSplitColumnNumber As Integer
Dim strFilesOrWorksheets As String ' F or W as valid options
Dim strOutputFolder As String
Dim strDataSourceSheetName As String

If optSheets = True Then
    strFilesOrWorksheets = "S"
Else: strFilesOrWorksheets = "F"
End If

If FolderExistenceCheck(txtOutputFolder.Value) = True Then
    strOutputFolder = txtOutputFolder.Value
Else
    MsgBox ("The output folder is not a valid folder.")
    Exit Sub
End If

If FileExistenceCheck(txtDataSourceFile.Value) = True Then
    strDataSourceFileName = txtDataSourceFile.Value
Else
    MsgBox ("The input file cannot be found.")
    Exit Sub
End If

If FileIsExcelFileCheck(txtDataSourceFile.Value) = False Then
    MsgBox (txtDataSourceFile.Value & " is not an Excel file.")
    Exit Sub
End If

If cmbSheetName = "" Then
    MsgBox ("Enter the name of the sheet you wish to split")
    Exit Sub
Else: strDataSourceSheetName = cmbSheetName.Value
End If

booApplyBasicFormatting = chkApplyBasicFormatting.Value
booWarnBeforeOverwritingFiles = chkWarnBeforeOverwritingFiles.Value

If IsValidInt(txtSplitColumn) = False Then
    MsgBox ("The split column number must be an integer larger than 0.")
    Exit Sub
Else
intSplitColumnNumber = txtSplitColumn
End If

Unload Me

Call SplitData(strDataSourceFileName, _
                            booApplyBasicFormatting, _
                            intSplitColumnNumber, _
                            strFilesOrWorksheets, _
                            strOutputFolder, _
                            strDataSourceSheetName, _
                            booWarnBeforeOverwritingFiles _
                            )
                            
End Sub
'==============================
'Name: txtDataSourceFile_Enter
'Purpose: Launch a file picker when the data source
'text box is entered.
'==============================
Private Sub txtDataSourceFile_Enter()
Dim fDialog As FileDialog
Dim arrSheetNames() As Variant
Dim i As Integer

Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
     
fDialog.AllowMultiSelect = False
fDialog.Title = "Select a file"
fDialog.InitialFileName = "C:\"
fDialog.Filters.Clear
fDialog.Filters.Add "Excel files", "*.xlsx"
fDialog.Filters.Add "All files", "*.*"
 
If fDialog.Show = -1 Then  '-1=Ok, 0=Cancel
    txtDataSourceFile.Value = fDialog.SelectedItems(1)
        
    If FileIsExcelFileCheck(txtDataSourceFile.Value) = False Then 'Check file is Excel file before getting sheet names
        MsgBox (txtDataSourceFile.Value & " is not an Excel file.")
        Exit Sub
    End If
        
    'Todo: Refactor: move get sheet names into its own subroutine
    cmbSheetName.Clear
    arrSheetNames = GetSheetNames(txtDataSourceFile.Value)
    
    For i = LBound(arrSheetNames) To UBound(arrSheetNames)
        cmbSheetName.AddItem (arrSheetNames(i))
    Next i
    
End If

End Sub
'===========================
'Name: GetSheetNames
'Purpose: Get the names of the sheets on the selected data source workbook.
'==============================
Private Function GetSheetNames(ByVal strWorkbookFullPath As String) As Variant
Application.ScreenUpdating = False
Dim wbWorkbook As Workbook
Dim i As Integer
Dim arrTemp() As Variant

Set wbWorkbook = Application.Workbooks.Open(strWorkbookFullPath)
ReDim arrTemp(wbWorkbook.Worksheets.Count - 1)

For i = 1 To wbWorkbook.Worksheets.Count
    arrTemp(i - 1) = wbWorkbook.Worksheets(i).Name
Next i

GetSheetNames = arrTemp
wbWorkbook.Close
Application.ScreenUpdating = True

End Function
'==============================
'Name: txtOutputFolder_Enter
'Purpose: Launch a Folder Picker when the
'output folder text box is entered.
'==============================
Private Sub txtOutputFolder_Enter()
Dim fDialog As FileDialog

Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
fDialog.AllowMultiSelect = False
fDialog.Title = "Select a Folder"
fDialog.InitialFileName = "C:\"
fDialog.Filters.Clear
 
If fDialog.Show = -1 Then   '-1 =Ok, 0=Cancel
    txtOutputFolder.Value = fDialog.SelectedItems(1)
End If

End Sub
'==============================
'Name: UserForm_Initialize
'Purpose: Initialise the form. Set the checkbox defaults.
'==============================
Private Sub UserForm_Initialize()

optFiles.Value = True
chkWarnBeforeOverwritingFiles = True

End Sub

