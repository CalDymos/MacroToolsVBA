Attribute VB_Name = "A_RibbonCallbacks"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : A_RibbonCallbacks - модуль обратных вызовов ленты управления Excel
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 30-09-2024 08:34    CalDymos

Option Private Module
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : MacroToolsLoadRibbon - при активайии проверить наличие обновлений
'* Created    : 08-10-2020 13:45
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : Byte Ranger Software
'* Argument(s):                 Description
'*
'* ByRef ribbon As IRibbonUI :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub MacroToolsLoadRibbon(ByRef ribbon As IRibbonUI)
End Sub
    Private Sub RefrasBtn(ByRef control As IRibbonControl)
1      If VBAIsTrusted Then
2          Call B_CreateMenus.RefreshMenu
3      End If
End Sub
    Private Sub ImportCodeBaseBtn(ByRef control As IRibbonControl)
4      On Error GoTo ErrorHandler
5      If VBAIsTrusted Then
6          Workbooks(C_Const.NAME_ADDIN & ".xlam").Sheets(C_Const.SH_SNIPPETS).Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
7          Call MsgBox("The code base has been unloaded", vbInformation, "Unloading the code base:")
8      End If
9      Exit Sub
ErrorHandler:
10     Select Case Err.Number
        Case 91:
11             Call MsgBox("No open" & Chr(34) & "Excel Files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Mistake:")
12         Case Else:
13             Call MsgBox("Mistake! in ImportCodeBaseBtn" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Mistake:")
14             Call WriteErrorLog("ImportCodeBaseBtn")
15     End Select
16     Err.Clear
End Sub
    Private Sub AddCodeBaseBtn(ByRef control As IRibbonControl)
17     If VBAIsTrusted Then
18         Call AddCodeView.Show
19     End If
End Sub
    Private Sub AddStatBtn(ByRef control As IRibbonControl)
20     If VBAIsTrusted Then
21         Call I_StatisticVBAProj.AddSheetStatistica
22     End If
End Sub
    Sub btnHiddenModule(control As IRibbonControl)
23     Call HiddenModule.Show
End Sub
    Private Sub AddInBtn(ByRef control As IRibbonControl)
24     On Error GoTo ErrorHandler
25     Application.Dialogs(xlDialogAddinManager).Show
26     Exit Sub
ErrorHandler:
27     Err.Clear
28     Call MsgBox("No open" & Chr(34) & "Excel Files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Mistake:")
End Sub
    Private Sub VBABtn(ByRef control As IRibbonControl)
29     Call VBAVBEOpen
End Sub
    Private Sub BtnExportVBA(control As IRibbonControl)
30     Call VBAVBEOpen
31     Call ModuleCommander.Show
End Sub
    Public Sub VBAVBEOpen()
32     If C_PublicFunctions.Num_Not_Stable Then Call SendKeys("%{NUMLOCK}")
33     Call SendKeys("%{F11}")
End Sub
    Private Sub BtnVSC(control As IRibbonControl)
34     Call VersionSistemControls.Show
End Sub
    Private Sub onSwitcherReferenceStyle(ByRef control As IRibbonControl)
35     With Application
36         If .ReferenceStyle = xlR1C1 Then
37             .ReferenceStyle = xlA1
38         Else
39             .ReferenceStyle = xlR1C1
40         End If
41     End With
End Sub
    Private Sub onOpenFileExcel(ByRef control As IRibbonControl)
42     Call O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=True, bBackUp:=False)
End Sub
    Private Sub onCloseFileExcel(ByRef control As IRibbonControl)
43     Call O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=False, bBackUp:=True)
End Sub
     Private Sub onUnProtectVBA(ByRef control As IRibbonControl)
44     Call P_UnProtected.unprotected
End Sub
     Private Sub onUnProtectSheets(ByRef control As IRibbonControl)
45      On Error GoTo ErrorHandler
46      Call ProtectedSheets.Show
47      Exit Sub
ErrorHandler:
48      Select Case Err.Number
        Case 91:
49              Call MsgBox("No open" & Chr(34) & "Excel Files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Mistake:")
50          Case Else:
51              Call MsgBox("Mistake! in onUnUnProtectSheets" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Mistake:")
52              Call WriteErrorLog("onUnUnProtectSheets")
53      End Select
54      Err.Clear
End Sub
     Private Sub onUnProtectVBAUnivable(control As IRibbonControl)
55      Call P_UnProtected.DelPasswordVBAProjectUnivable
End Sub
     Private Sub onAddShapeStatistic(ByRef control As IRibbonControl)
56      Call AddShapeStatistic
End Sub
     Private Sub onOptions(ByRef control As IRibbonControl)
57      Call OptionsCodeFormat.Show
End Sub
     Private Sub onOptionsComments(ByRef control As IRibbonControl)
58      Call SettingsAddCommentsProc.Show
End Sub
     Private Sub BtnHelpMain(control As IRibbonControl)
59      Call HelpMainAddin
End Sub
     Private Sub BtnMainBuilders(control As IRibbonControl)
60      Call URLLinks(C_Const.URL_BILD)
End Sub
     Private Sub BtnHelpControls(control As IRibbonControl)
61      Call URLLinks(C_Const.URL_MOVE_CNTR)
End Sub
     Private Sub BtnHelpSnippets(control As IRibbonControl)
62      Call URLLinks(C_Const.URL_STYLE)
End Sub
     Private Sub BtnHelpPass(control As IRibbonControl)
63      Call URLLinks(C_Const.URL_FILE)
End Sub
     Private Sub HelpMainAddin()
64      Call URLLinks(C_Const.URL_ADDIN)
End Sub
     Private Sub onInToFile(control As IRibbonControl)
65      Call Q_InToFile.InToFile
End Sub
'Version
     Public Sub getVisible(control As IRibbonControl, ByRef visible)
66      visible = C_Const.FlagVisible
End Sub
'btnVersion
     Public Sub onVisible(control As IRibbonControl)
67      Call URLLinks(C_Const.URL_DOWNLOAD)
End Sub
     Private Sub BtnUnProtectSheetsXML(control As IRibbonControl)
68      Call DeletePaswortSheets
End Sub
     Private Sub onProtectVBAUnivable(control As IRibbonControl)
69      Call SetPasswordVBAProjectUnviewable
End Sub
'Themenwechsel
     Private Sub onBlackTheme(control As IRibbonControl)
70      Call V_BlackAndWiteTheme.ChangeColorDarkTheme
End Sub
     Private Sub onWhiteTheme(control As IRibbonControl)
71      Call V_BlackAndWiteTheme.ChangeColorWhiteTheme
End Sub
     Private Sub onToolCharMonitor(control As IRibbonControl)
72      Call CharsMonitor.Show
End Sub
     Private Sub FunctionWizardShowExc()
73        If Application.Dialogs(xlDialogFunctionWizard).Show = False Then
74            ActiveCell.Clear
75        End If
76        Calculate
End Sub
'Obfuscation
Private Sub onCompleteObf(control As IRibbonControl)
    If VBAIsTrusted Then
        Call N_ObfMainNew.StartCompleteObfuscation
    End If
End Sub
     Private Sub onParserVBA(control As IRibbonControl)
77      If VBAIsTrusted Then
78          Call N_ObfParserVBA.StartParser
79      End If
End Sub
     Private Sub onObfuscator(ByRef control As IRibbonControl)
80      If VBAIsTrusted Then
81          Call N_ObfMainNew.StartObfuscation
82      End If
End Sub
     Private Sub onFormatsDel(control As IRibbonControl)
83      If VBAIsTrusted Then
84          Call ObfuscationCode.Show
85      End If
End Sub
'-----------
     Private Sub BtnInfoFile(control As IRibbonControl)
86      Call InfoFile.Show
End Sub
     Private Sub ParserStrings(control As IRibbonControl)
87      Call ZA_ParserString.ParserStringWB
End Sub

     Private Sub ReNameParserString(control As IRibbonControl)
88      Call ZA_ParserString.ReNameStr
End Sub
     Private Sub onDeleteAllLinks(control As IRibbonControl)
89      Call ZB_CleanUpWorkbook.deleteAllLinksInFile
End Sub
     Private Sub onAddListLinks(control As IRibbonControl)
90      Call ZB_CleanUpWorkbook.getListAllLinksInFile
End Sub
Private Sub onDeleteLinksOnList(control As IRibbonControl)
91      Call ZB_CleanUpWorkbook.deleteLinksOnList
End Sub
Private Sub onDelAllShapes(control As IRibbonControl)
        Call ZB_CleanUpWorkbook.DelAllShapes
End Sub
Private Sub onDelAllUnusedFormats(control As IRibbonControl)
        Call ZB_CleanUpWorkbook.DellAllUnusedCustomFormats
End Sub

Private Sub onDelAllCustomFormats(control As IRibbonControl)
        Call ZB_CleanUpWorkbook.DellAllCustomFormats
End Sub
