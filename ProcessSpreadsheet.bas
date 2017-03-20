Attribute VB_Name = "ProcessSpreadsheet"
Option Explicit

Public Sub LoadForm()
frmGetFileSplitInput.Show
End Sub

'Todo: Warn if empty file- perhaps

'=================================
'Name: SplitData
'Purpose: Main subroutine which passes the form data to other subs and functions
'which do the actual processing.
'=================================

Public Sub SplitData(ByVal strFileName As String, _
                        ByVal booApplyBasicFormatting As Boolean, _
                        ByVal intSplitColumnNumber As Integer, _
                        ByVal strFilesOrWorksheets As String, _
                        ByVal strOutputFolder As String, _
                        ByVal strDataSourceSheetName As String, _
                        ByVal booWarnBeforeOverwritingFiles As Boolean _
                        )

On Error GoTo errHandler
Application.EnableCancelKey = xlErrorHandler

If booWarnBeforeOverwritingFiles = False Then
    Application.DisplayAlerts = False
Else:
    Application.DisplayAlerts = True
End If

Application.ScreenUpdating = False

Dim booCancelMacro As Boolean 'Variable to track whether an error has occurred in a subroutine and stop processing
Dim wsDataSource As Worksheet
Dim strConfirmCompletionMessage As String
Dim wbResultsWorkbook As Workbook
Dim booResultsOpen As Boolean
booCancelMacro = False
booResultsOpen = False
      
Set wsDataSource = CopyWorksheet(strFileName, strDataSourceSheetName, booCancelMacro)
 
If booCancelMacro = True Then
Application.StatusBar = ""
Exit Sub
End If
    
If intSplitColumnNumber > wsDataSource.UsedRange.Columns.Count Then
        MsgBox ("The column you have selected to split on (column " & intSplitColumnNumber & ") is not valid. " _
        & "The maximum number of columns in your data source file is " & wsDataSource.UsedRange.Columns.Count)
        Application.DisplayAlerts = False
        wsDataSource.Parent.Close
        Application.DisplayAlerts = True
        Exit Sub
End If
     
CreateNamedRanges wsDataSource, intSplitColumnNumber, booCancelMacro
    
If booCancelMacro = True Then
        

        Application.DisplayAlerts = False
        Application.StatusBar = ""
        Application.DisplayAlerts = True
        Exit Sub
End If
    
'================
If strFilesOrWorksheets = "S" Then

    CreateNewWorksheetsByRange wsDataSource, strOutputFolder, booApplyBasicFormatting, wbResultsWorkbook, booResultsOpen, booCancelMacro
              
     If booCancelMacro = True Then
    GoTo exitSub
    End If
    
    strConfirmCompletionMessage = "Splitting completed. The new file is " & strOutputFolder & "\split_results.xlsx."""
'================
ElseIf strFilesOrWorksheets = "F" Then

    CreateNewWorkbooksByRange wsDataSource, strOutputFolder, booApplyBasicFormatting, wbResultsWorkbook, booResultsOpen, booCancelMacro
           
    If booCancelMacro = True Then
    GoTo exitSub
    End If
    
    strConfirmCompletionMessage = "Splitting completed. The new files may be found in " & strOutputFolder & "."""
         
Else
    MsgBox ("Error - unexpected error occurred while selecting the output format.")
    Exit Sub
End If

wsDataSource.Parent.Close savechanges:=False

Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox (strConfirmCompletionMessage)
Application.StatusBar = ""
Exit Sub

exitSub:
    
    If Not wsDataSource Is Nothing Then
        wsDataSource.Parent.Close savechanges:=False
    End If

    If booResultsOpen = True Then
        wbResultsWorkbook.Close savechanges:=False
    End If
    
    Application.StatusBar = ""

Exit Sub

errHandler:
    
    Debug.Print "SplitData error in main" & Err.Number & Err.Description
    MsgBox ("How did we end up back at base? " & Err.Number & ": " & Err.Description)
    Select Case Err.Number
        Case 18
            Debug.Print "Macro Cancelled in main" & Err.Number & Err.Description
            MsgBox ("Cancelling macro as requested.")
            Resume exitSub
        Case 1004
        
            If Err.Description = "Method 'Open' of object 'Workbooks' failed" Then  ' Cancelled before data file opened
                MsgBox ("Cancelled as requested.")
                Exit Sub
            ElseIf Err.Description = "Application-defined or object-defined error" Then
                MsgBox ("Cancelled as requested.")
                Debug.Print "Error due to cancellation"
                Exit Sub
            Else
                MsgBox ("Unknow runtime error in main sub " & Err.Description)
                Debug.Print Err.Number & ": " & Err.Description
                End If
                
        Case Else
            Debug.Print "Unexpected error - " & Err.Number & ": " & Err.Description
            MsgBox ("Unexpected error occurred - " & Err.Number & ": " & Err.Description)
            Exit Sub
        End Select


End Sub


'============================
'Name: CopyWorksheet
'Purpose: Takes a copy of the data source file to use in the macro
'in case of any problems.
'============================
Private Function CopyWorksheet( _
                        ByVal strFilePath As String, _
                        ByVal strSheetName As String, _
                        ByRef booCancelMacro As Boolean _
                        ) As Worksheet
                        
On Error GoTo errHandler

Dim wbWorkbookToCopy As Workbook
Dim wsWorksheetToCopy As Worksheet
Dim wbNewWorkbook As Workbook
Dim wsNewSheet As Worksheet

Set wbWorkbookToCopy = Application.Workbooks.Open(strFilePath)
Set wsWorksheetToCopy = wbWorkbookToCopy.Worksheets(strSheetName)

Application.StatusBar = "Copying Data Source Worksheet"

Set wbNewWorkbook = Workbooks.Add(xlWBATWorksheet) 'create a 1 sheet workbook
Set wsNewSheet = wbNewWorkbook.Sheets(1)

wsNewSheet.Range( _
        wsNewSheet.Cells(1, 1), _
        wsNewSheet.Cells(wsWorksheetToCopy.UsedRange.Rows.Count, _
        wsWorksheetToCopy.UsedRange.Columns.Count)).Value = wsWorksheetToCopy.UsedRange.Value
        
If wbNewWorkbook.Worksheets(1).Name <> strSheetName Then   'Don't rename to same name
      wbNewWorkbook.Worksheets(1).Name = strSheetName
End If

Application.CutCopyMode = False

wbWorkbookToCopy.Close
Set CopyWorksheet = wbNewWorkbook.Worksheets(1)
Exit Function

exitFunction:
Debug.Print "exitFunction Called"

If Not wbWorkbookToCopy Is Nothing Then
        wbWorkbookToCopy.Close savechanges:=False
End If
        
 If Not wbNewWorkbook Is Nothing Then
        wbNewWorkbook.Close savechanges:=False
End If
Exit Function
        
 
errHandler:
    Debug.Print "Error in CopyWorksheet " & Err.Number & ": " & Err.Description
    booCancelMacro = True
    
    Select Case Err.Number
        Case 18
        Debug.Print "Case 18"
        Application.StatusBar = "Cancelling macro ..."
        MsgBox ("Macro cancelled as requested.")
        Resume exitFunction
        Case 1004
            Debug.Print "Case 1004"
            Select Case Err.Description
                Case "PasteSpecial method of Range class failed"    'triggered by cancellation
                    Debug.Print "Paste Special"
                    Application.StatusBar = "Cancelling macro ..."
                    MsgBox ("Macro cancelled as requested.")
                    Resume exitFunction
                Case "Application-defined or object-defined error"
                    Application.StatusBar = "Cancelling macro ..."
                    MsgBox ("Macro cancelled as requested.")
                    Resume exitFunction
                Case Else
                    Debug.Print "Case " & Err.Description
                    MsgBox ("Unexpected error occurred while copying the datasource file")
                    Resume exitFunction
            End Select
        
        Case Else
        MsgBox ("Unexpected error occurred while copying the datasource file")
        booCancelMacro = True
        If Not wbNewWorkbook Is Nothing Then
           wbNewWorkbook.Close savechanges:=False
        End If
        
        Exit Function
    End Select

End Function
'===================
'Name: CreateNewWorkbooksByRange
'Purpose: Creates a new workbook for each named range
'(except the header)
'===================
Private Sub CreateNewWorkbooksByRange( _
                wsDataSource As Worksheet, _
                ByVal strFolderPath As String, _
                ByVal booApplyBasicFormatting As Boolean, _
                ByRef wbResultsWorkbook As Workbook, _
                ByRef booWbResultsOpen As Boolean, _
                ByRef booCancelMacro As Boolean _
                )

On Error GoTo errHandler

Dim nNamedRange As Name
Dim wsNewSheet As Worksheet
Dim lonNameCount As Long
Dim strFileName As String

lonNameCount = 1
Application.ScreenUpdating = False

wsDataSource.Activate
For Each nNamedRange In Names

    If nNamedRange.Name <> "Header" Then
        
        Application.StatusBar = "Creating File " & lonNameCount & " of " & Names.Count - 1
        Set wbResultsWorkbook = Application.Workbooks.Add
        booWbResultsOpen = True
        
        Set wsNewSheet = wbResultsWorkbook.Worksheets.Add
        
        wsNewSheet.Range(wsNewSheet.Cells(1, 1), wsNewSheet.Cells(1, wsDataSource.Range("Header").Columns.Count)).Value _
            = wsDataSource.Range("Header").Value
        
        wsNewSheet.Range( _
                wsNewSheet.Cells(2, 1), wsNewSheet.Cells(wsDataSource.Range(nNamedRange).Rows.Count + 1, wsDataSource.Range(nNamedRange).Columns.Count)).Value _
            = wsDataSource.Range(nNamedRange).Value
            
        strFileName = RegexReplace(nNamedRange.Name, "^_+", "", False)
        wsNewSheet.Name = strFileName   'Name the worksheet the same as the file
               
        lonNameCount = lonNameCount + 1
        
        If booApplyBasicFormatting = True Then
            FormatWorksheet wsNewSheet, booCancelMacro
        End If
        
        wbResultsWorkbook.SaveAs Filename:=strFolderPath & "\" & strFileName, FileFormat:=51
        wbResultsWorkbook.Close
        booWbResultsOpen = False
    End If

    If booCancelMacro = True Then
            Exit Sub
    End If

Next nNamedRange

Application.ScreenUpdating = False
Exit Sub
        
exitSub:
    End
    
errHandler:
    Debug.Print "Error in CreateNewWorkbooksByRange: " & Err.Number & Err.Description
    Select Case Err.Number
        Case 18
            
            Application.StatusBar = "Cancelling macro ..."
            booCancelMacro = True
            MsgBox ("Macro cancelled as requested")
            Resume                 'run until this iteration of the loop finishes so everthing is in stable state when it gets cleared up
                                            'intentionally resume not resume next to ensure that everything is done so clean up is predictable
        Case 1004
               Select Case Err.Description
                    Case "Application-defined or object-defined error"
                        MsgBox ("Macro stopped to ensure that file is not overwritten.")    'Triggered when you don't save file that already exists & breakpoint is set!
                    Case "Method 'SaveAs' of object '_Workbook' failed"
                        MsgBox ("Macro stopped to ensure that file is not overwritten.")
                    Case "You cannot save this workbook with the same name as another open workbook or add-in. Choose a different name, or close the other workbook or add-in before saving."
                         MsgBox ("The file you are trying to save the data to (" & strFolderPath & "\" & wbResultsWorkbook.ActiveSheet.Name & ".xlsx) is already open. Please close it and try again.")
                    Case Else
                        MsgBox ("Unknown runtime error occurred while creating new workbooks. " & Err.Number & ": " & Err.Description)
                End Select
            booCancelMacro = True
            Exit Sub
         Case Else
            Debug.Print "Unknown error creating new workbooks " & Err.Number & ": " & Err.Description
            booCancelMacro = True
        End Select


End Sub
'===================
'Name: CreateNewWorksheetsByRange
'Purpose: Creates a new worksheet for each named range
'(except the header)
'===================
Private Sub CreateNewWorksheetsByRange( _
                wsDataSource As Worksheet, _
                ByVal strFolderPath As String, _
                ByVal booApplyBasicFormatting As Boolean, _
                ByRef wbResultsWorkbook As Workbook, _
                ByRef booResultsOpen As Boolean, _
                ByRef booCancelMacro As Boolean _
                )

On Error GoTo errHandler

Dim nNamedRange As Name
Dim wsNewSheet As Worksheet
Dim lonNameCount As Long
Dim strSheetName As String
Application.ScreenUpdating = False
Set wbResultsWorkbook = Application.Workbooks.Add(xlWBATWorksheet)
booResultsOpen = True
wsDataSource.Activate
lonNameCount = 1

For Each nNamedRange In Names
    
    If nNamedRange.Name <> "Header" Then
        Application.StatusBar = "Creating Sheet " & lonNameCount & " of " & Names.Count - 1
        
        Set wsNewSheet = wbResultsWorkbook.Worksheets.Add
        wsNewSheet.Range(wsNewSheet.Cells(1, 1), wsNewSheet.Cells(1, wsDataSource.Range("Header").Columns.Count)).Value _
            = wsDataSource.Range("Header").Value
        
        wsNewSheet.Range( _
                wsNewSheet.Cells(2, 1), wsNewSheet.Cells(wsDataSource.Range(nNamedRange).Rows.Count + 1, wsDataSource.Range(nNamedRange).Columns.Count)).Value _
            = wsDataSource.Range(nNamedRange).Value
        
        strSheetName = RegexReplace(nNamedRange.Name, "^_+", "", False)
        wsNewSheet.Name = strSheetName
        lonNameCount = lonNameCount + 1
        
        If booApplyBasicFormatting = True Then
            FormatWorksheet wsNewSheet, booCancelMacro
        End If
    
    wsDataSource.Activate
    End If

If booCancelMacro = True Then
            Exit Sub
        End If

Next nNamedRange

wbResultsWorkbook.SaveAs Filename:=strFolderPath & "\split_results", FileFormat:=51
wbResultsWorkbook.Close
Application.ScreenUpdating = True

Exit Sub

errHandler:                             'Hand back to SplitData for clear up
    Debug.Print "Error in CreateNewWorksheetsByRange: " & Err.Number & Err.Description
    
    Select Case Err.Number
        Case 18
            Application.StatusBar = "Cancelling macro ..."
            booCancelMacro = True
            MsgBox ("Macro cancelled as requested.")
            Resume                  'run until this iteration of the loop finishes so everthing is in stable state when it gets cleared up
                                            'resume not resume next as want to pick up where cancel occurred and get to predictable point
        Case 1004
                Select Case Err.Description
                    Case "Method 'SaveAs' of object '_Workbook' failed"
                        MsgBox ("Macro stopped to ensure that file is not overwritten.")
                    Case "You cannot save this workbook with the same name as another open workbook or add-in. Choose a different name, or close the other workbook or add-in before saving."
                        MsgBox ("The file you are trying to save the data to (" & strFolderPath & "\" & "split_results.xlsx is already open. Please close it and try again.")
                    Case "Application-defined or object-defined error"
                        MsgBox ("Macro cancelled as requested.")  'Not sure why cancelling sometimes causes this error to fire
                    Case Else
                        MsgBox ("Unknown runtime error occurred while creating new workbooks. " & Err.Number & ": " & Err.Description)
                End Select
            booCancelMacro = True
            Exit Sub
         Case Else
            MsgBox ("Unknown error creating new worksheets " & Err.Number & ": " & Err.Description)
            booCancelMacro = True
        End Select


End Sub
'==================================
'Name: FormatWorksheet
'Purpose: To provide some very basic formatting.
'==================================
Private Sub FormatWorksheet(wsMySheet As Worksheet, booCancelMacro As Boolean)

On Error GoTo errHandler

wsMySheet.Activate

With wsMySheet.Range(Cells(1, 1), Cells(1, wsMySheet.UsedRange.Columns.Count)).Interior
    .Color = RGB(224, 224, 224)
End With

With wsMySheet.Range(Cells(1, 1), Cells(1, wsMySheet.UsedRange.Columns.Count)).Font
    .Name = "Verdana"
    .Bold = True
    .Size = 11
    .Italic = True
End With

With wsMySheet.UsedRange.Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
End With

wsMySheet.UsedRange.Columns.AutoFit
wsMySheet.Cells(1, 1).Select
Exit Sub

errHandler:
    booCancelMacro = True
    Debug.Print "Error in FormatWorksheet: " & Err.Number & Err.Description
    
    Select Case Err.Number
        Case 18
            Application.StatusBar = "Cancelling macro ..."
            MsgBox ("Macro cancelled as requested.")
        Case Else                                   'should only get a runtime error due to cancelling - pass back to caller to be passed back to SplitData
            Application.StatusBar = "Cancelling macro ..."
            MsgBox ("Macro cancelled as requested.")
    End Select

End Sub
'=========================
'Name: CreateNamedRanges
'Purpose: Creates a named range for each group of data
'based upon the split column. These named ranges
'will later be used to create the new worksheets
'or workbooks.
'=========================
Private Sub CreateNamedRanges( _
            wsDataSource As Worksheet, _
            ByVal intColumnToSplitOn As Integer, _
            ByRef booCancelMacro As Boolean)

On Error GoTo errHandler

Dim strRangeName As String
Dim strWarning As String
Dim lngStartRowForRange As Long
Dim x As Long
Dim varDictKey As Variant
Dim dictRanges As Object
Dim intMsgBoxUnsorted As Integer

Set dictRanges = CreateObject("Scripting.Dictionary")
Application.StatusBar = "Creating Named Ranges"
Range(wsDataSource.Cells(1, 1), wsDataSource.Cells(1, wsDataSource.UsedRange.Columns.Count)).Name = "Header"
    
For x = 2 To wsDataSource.UsedRange.Rows.Count
    
        If wsDataSource.Cells(x, intColumnToSplitOn) <> wsDataSource.Cells(x - 1, intColumnToSplitOn) Or x = 2 Then
            strRangeName = wsDataSource.Cells(x, intColumnToSplitOn).Value
            strRangeName = ValidateStringName(strRangeName)
            lngStartRowForRange = x
        End If
        
        If wsDataSource.Cells(x, intColumnToSplitOn) <> wsDataSource.Cells(x + 1, intColumnToSplitOn) Then
            Range(wsDataSource.Cells(lngStartRowForRange, 1), wsDataSource.Cells(x, wsDataSource.UsedRange.Rows.Count)).Select
            Range(wsDataSource.Cells(lngStartRowForRange, 1), wsDataSource.Cells(x, wsDataSource.UsedRange.Rows.Count)).Name = strRangeName
            dictRanges(strRangeName) = dictRanges(strRangeName) + 1
        End If
        
Next x
    
For Each varDictKey In dictRanges.keys
        If dictRanges(varDictKey) > 1 Then
            strWarning = strWarning & "Range " & varDictKey & " occurs " & dictRanges(varDictKey) & " times. "
            
            If Len(strWarning) > 200 Then
                strWarning = Left(strWarning, 200) & "..."
            End If
        
        End If
Next

If strWarning <> "" Then
    strWarning = "The data is not sorted by the split column, and consequently, some of it will be lost in the output file." & vbNewLine & vbNewLine & strWarning
    intMsgBoxUnsorted = MsgBox(strWarning & vbNewLine & vbNewLine & "Do you want to continue?", vbOKCancel, "Warning: Unsorted Data")
                
    If intMsgBoxUnsorted = 2 Then
        booCancelMacro = True
        GoTo exitSub
    End If
    
End If

Exit Sub

exitSub:
    If Not wsDataSource Is Nothing Then
        wsDataSource.Parent.Close savechanges:=False
    End If

Exit Sub
    
errHandler:
    booCancelMacro = True
    Debug.Print "Error in CreateNamedRanges: " & Err.Number & Err.Description
    Select Case Err.Number
        Case 18
            Application.StatusBar = "Cancelling macro ..."
            MsgBox ("Macro cancelled as requested")
            Resume exitSub
        Case Else
            MsgBox ("Unexpected error " & Err.Number & ": " & Err.Description)
            Resume exitSub
        End Select
End Sub

