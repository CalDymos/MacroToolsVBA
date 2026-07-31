Attribute VB_Name = "ZC_HotKey"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : ZC_HotKey
'* Created    : 23-06-2022 16:21
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

    Public Sub hotKeysStart()
12:
13:    If Not ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then Exit Sub
14:
15:    On Error GoTo errMsg
16:    Dim sPatpApp    As String
17:    sPatpApp = ThisWorkbook.Path & Application.PathSeparator & FILE_NAME_HOT_KEYS
18:    If FileHave(sPatpApp) Then
19:        Call Shell(sPatpApp)
20:    End If
21:    Exit Sub
errMsg:
23:    Call WriteErrorLog("hotKeysStart")
24: End Sub

Public Sub hotKeysStop()
27:    On Error GoTo errMsg
28:
29:    Dim sPatpApp    As String
30:    sPatpApp = ThisWorkbook.Path & Application.PathSeparator & FILE_NAME_HOT_KEYS
31:    If FileHave(sPatpApp) Then
32:        Dim WshShell As Object
33:        Set WshShell = CreateObject("WScript.Shell")
34:        If Not WshShell Is Nothing Then
35:            Dim WshExec As Object
36:            Set WshExec = WshShell.Exec(sPatpApp)
37:            If Not WshShell Is Nothing Then WshExec.Terminate
38:            Set WshExec = Nothing
39:        End If
40:        Set WshShell = Nothing
41:    End If
42:
43:    Exit Sub
errMsg:
45:    Call WriteErrorLog("hotKeysStop")
End Sub
