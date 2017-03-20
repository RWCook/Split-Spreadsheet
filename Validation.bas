Attribute VB_Name = "Validation"
Option Explicit

'==============================
'Name: FileExistenceCheck
'Purpose: Checks that the data source file exists.
'==============================

Public Function FileExistenceCheck(ByVal strFullPathFilename As String) As Boolean

Dim objFSO As Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(strFullPathFilename) Then
    FileExistenceCheck = True
    Else
    FileExistenceCheck = False
End If

End Function

'==============================
'Name: FileIsExcelFileCheck
'Purpose: Checks that the file has an Excel extension.
'==============================

Public Function FileIsExcelFileCheck(ByVal strFileName As String) As Boolean
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")

    With regex
        .Global = False
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "xls[xm]*$"
    End With

If regex.test(strFileName) Then
    FileIsExcelFileCheck = True
Else: FileIsExcelFileCheck = False
End If

End Function

'==============================
'Name: FolderExistenceCheck
'Purpose: Checks that the folder that will be used
'to store the output file or files exists.
'==============================
Public Function FolderExistenceCheck(ByVal strFullPathFolderName As String) As Boolean
Dim objFSO As Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists(strFullPathFolderName) Then
    FolderExistenceCheck = True
Else
    FolderExistenceCheck = False
End If

End Function

'==============================
'Name: IsValidInt
'Purpose: Checks for an integer which is positive
'and greater than 0. Used to check the column
'which is used to split the data is a valid integer.
'==============================

Public Function IsValidInt(strEnteredData As Variant) As Boolean
     
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")

With regex
    .Global = False
    .MultiLine = True
    .IgnoreCase = False
    .Pattern = "^[1-9][0-9]*$"
End With

If regex.test(strEnteredData) Then
        IsValidInt = True
Else: IsValidInt = False
End If

End Function

'==============================
'Name: RegexReplace
'Purpose: To enable the usage or regular expressions
'to replace text. Used in function ValidateStringName.
'==============================

Public Function RegexReplace(ByVal MyString As String, _
                            ByVal strMatch As String, _
                            ByVal strReplace As String, _
                            ByVal booGlobal As Boolean) As String
     
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")

With regex
    .Global = booGlobal
    .MultiLine = True
    .IgnoreCase = False
    .Pattern = strMatch
End With

RegexReplace = regex.Replace(MyString, strReplace)

End Function

'==============================
'Name: ValidateStringName
'Purpose: Replaces any characters that would
'be invalid as part of range's name.
'==============================

Public Function ValidateStringName(strName As String) As String

Const conMaxNameLength As Integer = 30

If strName = "" Then
    strName = "EMPTY"
End If

ValidateStringName = RegexReplace(strName, " ", "_", True)                                          'Remove Spaces
ValidateStringName = RegexReplace(ValidateStringName, "^([0-9])", "_$1", True)      'Prefix initial character with underscore if number
ValidateStringName = RegexReplace(ValidateStringName, "^([A-Za-z][A-Za-z]*[A-Za-z]*[0-9]+)$", "_$1", True)                      'Stop names that could be cell references
ValidateStringName = RegexReplace(ValidateStringName, "^([Rr][0-9]+[Cc][0-9]+)$", "_$1", True)                      'Stop names that could be R1C1 cell references
ValidateStringName = RegexReplace(ValidateStringName, "[\W]", "_", True)                'Any non word characters change to underscore

If Len(ValidateStringName) > conMaxNameLength Then
    ValidateStringName = Mid(ValidateStringName, 1, conMaxNameLength)
End If

End Function

