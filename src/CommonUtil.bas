Attribute VB_Name = "CommonUtil"
' constants
Public Const s_dquote As String = """"
Public Const s_squote As String = "'"
Public Const s_empty As String = s_dquote & s_dquote

' functions
Function enquoteStr(ByVal inputStr As String, Optional ByVal isDoubleQuote As Boolean = True) As String
    If isDoubleQuote Then
        enquoteStr = s_dquote & inputStr & s_dquote
    Else
        enquoteStr = s_squote & inputStr & s_squote
    End If
End Function

Function isRGBColorString(ByVal inputStr As String) As Boolean
    If Len(inputStr) <> 6 Then
        isRGBColorString = False
    Else
        isRGBColorString = Not inputStr Like "*[!0-9A-Fa-f]*"
    End If
End Function

Function isSheetExist(ByVal sheetName As String) As Boolean
    ' ref: https://stackoverflow.com/a/61004981
    isSheetExist = Not IsError(Evaluate(s_squote & sheetName & s_squote & "!a1"))
End Function

Function isTableExists(ByVal tableName As String, targetSheet As Worksheet) As Boolean
    Dim lst_listObj As ListObject
    For Each lst_listObj In targetSheet.ListObjects
        If lst_listObj.Name = tableName Then
            isTableExists = True
            Exit Function
        End If
    Next lst_listObj
    isTableExists = False
End Function

Function getRowDataLen(beginCell As Range) As Integer
    ' note does not update automatically as formula in cell
    While Not IsEmpty(beginCell.Offset(, getRowDataLen).Value)
        getRowDataLen = getRowDataLen + 1
    Wend
End Function

Function IndexMatch(targetRange As Variant, searchVal As Variant, searchRange As Variant, Optional ByVal matchType As Integer = 0) As Variant
    IndexMatch = WorksheetFunction.Index(targetRange, WorksheetFunction.Match(searchVal, searchRange, matchType))
End Function

Function getWorkNameFromSheetName(ByVal sheetName As String) As String
    getWorkNameFromSheetName = Trim(WorksheetFunction.Clean(Worksheets(sheetName).Range("C5").Text))
End Function

Function getTableNameFromWorkName(ByVal workName As String) As String
    getTableNameFromWorkName = "list_" & Replace(workName, " ", "_")
End Function

Function getIndexMatchFormula(ByVal goalArr As String, ByVal searchVal As String, ByVal searchArr As String) As String
    ' getIndexMatchFormula = "INDEX(" & goalArr & ", MATCH(" & searchVal & ", " & searchArr & ", 0))"
    getIndexMatchFormula = "IndexMatch(" & goalArr & ", " & searchVal & ", " & searchArr & ")"
End Function


