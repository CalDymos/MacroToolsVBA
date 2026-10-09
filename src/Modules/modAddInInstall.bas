Attribute VB_Name = "modAddInInstall"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : F_AddInInstall - модуль установки надстройки
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Private Module
Option Explicit
' Einrichten der Einstellungen
    Public Sub InstallationAddMacro()
10:    Dim AddFolder As String
    Dim bEvents As Boolean
    Dim bAlerts As Boolean
    Dim bFailed As Boolean
    'Remember the caller's settings; CleanUp restores them on the normal path and after an error
    bEvents = Application.EnableEvents
    bAlerts = Application.DisplayAlerts
11:    On Error GoTo InstallationAdd_Err
12:    ' Prьfung, ob das betreffende Verzeichnis existiert
13:    AddFolder = Replace(Application.UserLibraryPath & "\", "\\", "\")
14:    'Prьfung auf Divergenz
15:    If Dir(AddFolder, vbDirectory) = vbNullString Then
16:        Call MsgBox("Unfortunately, the program cannot install the add-on on this computer." _
                      & vbCrLf & "There is no directory with add-ons." & vbCrLf & _
                      "Contact the developer of the program.", vbCritical, _
                      "Add-on installation failed")
20:        Exit Sub
21:    End If
22:    'Отключаем ранее установленую надстройку
23:    If FileHave(AddFolder & modConst.NAME_ADDIN & ".xlam") Then AddIns(modConst.NAME_ADDIN).Installed = False
24:    ' Проверяем открыта ли надстройка
25:    If WorkbookIsOpen(modConst.NAME_ADDIN & ".xlam") Then
26:        Call MsgBox("The file with the add-in is already open." & vbCrLf & _
                      "It may have already been installed earlier.", vbCritical, _
                      "Program installation failure")
29:        Exit Sub
30:    End If
31:    ' Сохраняем как
    Application.EnableEvents = False
33:    Application.DisplayAlerts = False
34:    If Workbooks.Count = 0 Then Workbooks.Add
35:    ThisWorkbook.SaveAs AddFolder & modConst.NAME_ADDIN & ".xlam", FileFormat:=xlOpenXMLAddIn
36:    AddIns.Add Filename:=AddFolder & modConst.NAME_ADDIN & ".xlam"
37:    AddIns(modConst.NAME_ADDIN).Installed = True
CleanUp:
    'Restore the caller's settings; a failing restore must not re-enter the error handler
    On Error Resume Next
    Application.EnableEvents = bEvents
    Application.DisplayAlerts = bAlerts
    On Error GoTo InstallationAdd_Err
    If bFailed Then Exit Sub
40:    Call MsgBox("The program has been successfully installed!" & vbCrLf & _
                  "Just open or create a new document.", vbInformation, _
                  "Installing the add-in:" & modConst.NAME_ADDIN)
43:    ThisWorkbook.Close False
44:    Exit Sub
InstallationAdd_Err:
    bFailed = True
46:    If Err.Number = 1004 Then
47:        MsgBox "To install the add-on, please close this file and run it again.", _
                      64, "Installation"
49:    Else
50:        MsgBox Err.Description & vbCrLf & "в F_AddInInstall.InstallationAdd " & vbCrLf & "in the line" & Erl, vbExclamation + vbOKOnly, "Mistake:"
51:        Call WriteErrorLog("F_AddInInstall.InstallationAdd")
52:    End If
    Resume CleanUp
53: End Sub

