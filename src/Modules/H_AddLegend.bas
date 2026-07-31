Attribute VB_Name = "H_AddLegend"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : H_AddLegend - Modul zur Erstellung einer Legende für Newsletter
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit
    Public Sub AddLegend()
12:    Call AddLegendsFromTabel(C_Const.TB_DESCRIPTION)
13: End Sub
    Public Sub AddLegendHotKeys()
15:    Dim sPatpApp    As String
16:    sPatpApp = ThisWorkbook.Path & Application.PathSeparator & FILE_NAME_HOT_KEYS
17:    If FileHave(sPatpApp) Then
18:        Call AddLegendsFromTabel(C_Const.TB_HOT_KEYS)
19:    Else
20:        Debug.Print "This function is not available, no file found:" & vbNewLine & sPatpApp
21:        Debug.Print "You can download it here: " & "https://github.com/vbatools/MacroToolsVBAHotKeys"
22:    End If
23: End Sub

    Private Sub AddLegendsFromTabel(ByVal sTabelName As String)
26:    Dim str_legend As String, str1 As String, str2 As String
27:    Dim objLegend As ListObject
28:    Dim i      As Integer
29:    Dim LenLengh1 As Byte, LenLengh2 As Byte
30:    Set objLegend = SHSNIPPETS.ListObjects(sTabelName)
31:    str_legend = vbNullString
32:    str1 = vbNullString
33:    str2 = vbNullString
34:    LenLengh1 = Len(objLegend.ListColumns(1).Range(1, 1))
35:    LenLengh2 = Len(objLegend.ListColumns(2).Range(1, 1))
36:    For i = 1 To objLegend.ListRows.Count + 1
37:        str1 = addString(objLegend.ListColumns(1).Range(i, 1), LenLengh1)
38:        str2 = addString(objLegend.ListColumns(2).Range(i, 1), LenLengh2)
39:        str_legend = str_legend & str1 & " | " & str2 & " | " & objLegend.ListColumns(1).Range(i, 3) & vbLf
40:    Next i
41:    Debug.Print str_legend
42: End Sub

Private Function addString(ByVal st As String, ByVal maxLen As Byte) As String
45:    addString = st & Space(maxLen - Len(st))
End Function
