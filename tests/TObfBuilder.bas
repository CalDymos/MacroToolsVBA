Attribute VB_Name = "TObfBuilder"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : TObfBuilder
'* Purpose    : Builds the T-OBF sample workbooks that test the obfuscation of MACROTools (refactoring
'*              plan, phase A3 and B6). Development tool: it runs in its own helper workbook and is not
'*              part of the add-in.
'* Usage      : Save the helper workbook, then run BuildTObfWorkbooks (Alt+F8). The sample workbooks
'*              are written to the subfolder T-OBF next to the helper workbook; existing ones are
'*              replaced after a confirmation.
'* Workbooks  : 00 Baseline        what the obfuscation handles today (regression for every step)
'*              01 NoUserForm      project without UserForm (A3.3)
'*              02 Quotes          literals with doubled quotes (A3.4)
'*              03 CompilerIf      #If in the target modules of generated procedures (A3.5)
'*              04 DesignerTexts   designer texts with quotes and line breaks (A3.6)
'*              05 EventNames      event handlers missing from the exclusion lists (A3.7)
'*              06 OptionalParams  literal defaults of optional parameters (A3.8)
'* Self-check : Each sample workbook has a sheet Check. A double-click on it runs the listed cases and
'*              writes the actual value and PASS or FAIL next to the expected value. This also works in
'*              the obfuscated copy: event procedures keep their names, and the obfuscation does not
'*              touch cell values.
'* Requires   : Trust access to the VBA project object model (the checks of the sample workbooks read
'*              their own code, too). No library reference: the VBA projects are used late bound.
'* Templates  : In the code templates at the end of this module a backtick stands for a double quote.
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private Const TOBF_FOLDER As String = "T-OBF"
Private Const BUILDER_TITLE As String = "T-OBF builder"
Private Const QUOTE_MARK As String = "`"
Private Const MSFORMS_GUID As String = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
Private Const XL_WBAT_WORKSHEET As Long = -4167     'xlWBATWorksheet: new workbook with one worksheet
Private Const XL_OPENXML_MACRO As Long = 52         'xlOpenXMLWorkbookMacroEnabled (.xlsm)
Private Const CT_STD_MODULE As Long = 1             'vbext_ct_StdModule
Private Const CT_CLASS_MODULE As Long = 2           'vbext_ct_ClassModule
Private Const CT_MS_FORM As Long = 3                'vbext_ct_MSForm
Private Const CT_DOCUMENT As Long = 100             'vbext_ct_Document

Private m_wbBuild As Workbook       'sample workbook under construction; closed unsaved after an error
Private m_sWarnings As String       'remarks for the final message

'Builds all sample workbooks (entry point, Alt+F8)
Public Sub BuildTObfWorkbooks()
    Dim vFiles      As Variant
    Dim sFolder     As String
    Dim sCurrent    As String
    Dim sBuilt      As String
    Dim sError      As String
    Dim lExisting   As Long
    Dim i           As Long
    Dim bScreen     As Boolean
    Dim bEvents     As Boolean
    Dim bAlerts     As Boolean

    If Not ProjectAccessTrusted() Then
        MsgBox "Access to the VBA project object model is not trusted." & vbLf & _
               "Enable it under File > Options > Trust Center > Trust Center Settings > Macro Settings, " & _
               "then run the builder again.", vbExclamation, BUILDER_TITLE
        Exit Sub
    End If
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save this helper workbook first. The sample workbooks are written next to it.", _
               vbExclamation, BUILDER_TITLE
        Exit Sub
    End If

    sFolder = ThisWorkbook.Path & Application.PathSeparator & TOBF_FOLDER
    vFiles = SampleFileNames()
    For i = LBound(vFiles) To UBound(vFiles)
        If IsWorkbookOpen(CStr(vFiles(i))) Then
            MsgBox "Close " & vFiles(i) & " first.", vbExclamation, BUILDER_TITLE
            Exit Sub
        End If
        If Len(Dir$(sFolder & Application.PathSeparator & vFiles(i))) > 0 Then lExisting = lExisting + 1
    Next i
    If lExisting > 0 Then
        If MsgBox(lExisting & " sample workbook(s) in " & sFolder & " will be replaced. Continue?", _
                  vbYesNo + vbQuestion, BUILDER_TITLE) <> vbYes Then Exit Sub
    End If

    bScreen = Application.ScreenUpdating
    bEvents = Application.EnableEvents
    bAlerts = Application.DisplayAlerts
    On Error GoTo ErrHandler
    sCurrent = "the folder " & sFolder
    If Len(Dir$(sFolder, vbDirectory)) = 0 Then MkDir sFolder
    'No events in the new workbooks while their code is written; no prompt when SaveAs replaces a file
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    m_sWarnings = vbNullString
    For i = LBound(vFiles) To UBound(vFiles)
        sCurrent = vFiles(i)
        Application.StatusBar = BUILDER_TITLE & ": " & sCurrent
        BuildSample i, sFolder & Application.PathSeparator & vFiles(i)
        sBuilt = sBuilt & vbLf & "    " & vFiles(i)
    Next i

CleanUp:
    On Error Resume Next
    If Not m_wbBuild Is Nothing Then
        m_wbBuild.Close SaveChanges:=False
        Set m_wbBuild = Nothing
    End If
    Application.StatusBar = False
    Application.DisplayAlerts = bAlerts
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreen
    On Error GoTo 0
    If Len(sError) > 0 Then
        If Len(sBuilt) = 0 Then sBuilt = " none"
        MsgBox sError & vbLf & vbLf & "Built before the error:" & sBuilt, vbCritical, BUILDER_TITLE
    Else
        If Len(m_sWarnings) > 0 Then m_sWarnings = vbLf & vbLf & "Warnings:" & m_sWarnings
        MsgBox "Built in " & sFolder & ":" & sBuilt & m_sWarnings & vbLf & vbLf & _
               "Next: open each workbook, run Debug > Compile VBAProject, then double-click the sheet Check.", _
               vbInformation, BUILDER_TITLE
    End If
    Exit Sub

ErrHandler:
    sError = "Building " & sCurrent & " failed." & vbLf & "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub

'File names of the sample workbooks; the position selects the build procedure in BuildSample
Private Function SampleFileNames() As Variant
    SampleFileNames = Array("T-OBF-00_Baseline.xlsm", "T-OBF-01_NoUserForm.xlsm", "T-OBF-02_Quotes.xlsm", _
                            "T-OBF-03_CompilerIf.xlsm", "T-OBF-04_DesignerTexts.xlsm", _
                            "T-OBF-05_EventNames.xlsm", "T-OBF-06_OptionalParams.xlsm")
End Function

'Creates one sample workbook with its sheet Check and saves it as sPath
Private Sub BuildSample(ByVal lIndex As Long, ByVal sPath As String)
    Dim wsCheck As Worksheet
    Set m_wbBuild = Workbooks.Add(XL_WBAT_WORKSHEET)
    Set wsCheck = m_wbBuild.Worksheets(1)
    wsCheck.Name = "Check"
    PrepareCheckSheet wsCheck
    Select Case lIndex
        Case 0
            BuildBaseline m_wbBuild, wsCheck
        Case 1
            BuildNoUserForm m_wbBuild, wsCheck
        Case 2
            BuildQuotes m_wbBuild, wsCheck
        Case 3
            BuildCompilerIf m_wbBuild, wsCheck
        Case 4
            BuildDesignerTexts m_wbBuild, wsCheck
        Case 5
            BuildEventNames m_wbBuild, wsCheck
        Case 6
            BuildOptionalParams m_wbBuild, wsCheck
    End Select
    wsCheck.Columns("A:F").AutoFit
    wsCheck.Activate
    m_wbBuild.SaveAs Filename:=sPath, FileFormat:=XL_OPENXML_MACRO
    m_wbBuild.Close SaveChanges:=False
    Set m_wbBuild = Nothing
End Sub

'--------------------------------------------------------------------------------------------- samples

'00: what the obfuscation handles today; must pass before and after every step (regression)
Private Sub BuildBaseline(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim vbcForm As Object
    Dim ctl     As Object
    AddComponent wb, CT_STD_MODULE, "modTobfBase", CodeBaseModule()
    AddComponent wb, CT_CLASS_MODULE, "clsTobfBase", CodeBaseClass()
    Set vbcForm = AddComponent(wb, CT_MS_FORM, "frmTobfBase", CodeBaseForm())
    SetFormFrame vbcForm, "Baseline form", 240, 130
    Set ctl = AddFormControl(vbcForm, "Forms.Label.1", "lblTobfInfo", 6, 6, 180, 18)
    ctl.Caption = "Plain caption"
    Set ctl = AddFormControl(vbcForm, "Forms.Label.1", "lblTobfState", 6, 30, 180, 18)
    ctl.Caption = "not ready"
    Set ctl = AddFormControl(vbcForm, "Forms.CommandButton.1", "cmdTobfOk", 6, 60, 72, 24)
    ctl.Caption = "OK"
    SetModuleCode SheetComponent(wb, ws), CodeCheckSheet(CodeBaseActual(), True, False)
    AddCase ws, "B01", "base", "literal in a standard module", "Hello world"
    AddCase ws, "B02", "base", "arithmetic without literals", "5"
    AddCase ws, "B03", "base", "class property", "abc"
    AddCase ws, "B04", "base", "form caption", "Baseline form"
    AddCase ws, "B05", "base", "label caption", "Plain caption"
    AddCase ws, "B06", "base", "button caption", "OK"
    AddCase ws, "B07", "base", "own code in UserForm_Initialize runs after the generated code", "ready"
    AddCase ws, "B08", "base", "Click handler of a renamed control", "connected"
End Sub

'01: project without UserForm (A3.3)
Private Sub BuildNoUserForm(ByVal wb As Workbook, ByVal ws As Worksheet)
    AddComponent wb, CT_STD_MODULE, "modTobfNoForm", CodeNoFormModule()
    SetModuleCode SheetComponent(wb, ws), CodeCheckSheet(CodeNoFormActual(), False, False)
    AddCase ws, "N01", "A3.3", "literal in a project without UserForm", "Good morning"
    AddCase ws, "N02", "A3.3", "literal joined with a number", "abc123"
End Sub

'02: literals with doubled quotes (A3.4)
Private Sub BuildQuotes(ByVal wb As Workbook, ByVal ws As Worksheet)
    AddComponent wb, CT_STD_MODULE, "modTobfQuotes", CodeQuotesModule()
    SetModuleCode SheetComponent(wb, ws), CodeCheckSheet(CodeQuotesActual(), False, False)
    AddCase ws, "Q01", "A3.4", "doubled quotes at the end", "He said `Hi`"
    AddCase ws, "Q02", "A3.4", "literal that starts with doubled quotes", "`Hi` he said"
    AddCase ws, "Q03", "A3.4", "doubled quote in the middle", "a`b"
    AddCase ws, "Q04", "A3.4", "literal of one quote", "`"
    AddCase ws, "Q05", "A3.4", "quote literal between two literals", "x`y"
    AddCase ws, "Q06", "A3.4", "two quoted words", "say `a` and `b`"
    AddCase ws, "Q07", "A3.4", "quoted literal followed by a literal", "`a`b"
    AddCase ws, "Q08", "A3.4", "literal without quotes", "no quotes"
End Sub

'03: #If in the declarations of the modules that receive generated procedures (A3.5)
Private Sub BuildCompilerIf(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim vbcModule As Object
    Dim vbcForm   As Object
    Dim ctl       As Object
    Set vbcModule = AddComponent(wb, CT_STD_MODULE, "modTobfCompIf", CodeCompIfModule())
    Set vbcForm = AddComponent(wb, CT_MS_FORM, "frmTobfCompIf", CodeCompIfForm())
    SetFormFrame vbcForm, "CompIf form", 200, 90
    Set ctl = AddFormControl(vbcForm, "Forms.CommandButton.1", "cmdTobfClose", 6, 6, 72, 24)
    ctl.Caption = "Close"
    CheckLastLine vbcModule
    CheckLastLine vbcForm
    SetModuleCode SheetComponent(wb, ws), CodeCheckSheet(CodeCompIfActual(), True, False)
    AddCase ws, "C01", "A3.5", "literal in the only standard module (with #If)", "compiled"
    AddCase ws, "C02", "A3.5", "constant from an #If branch", "7"
    AddCase ws, "C03", "A3.5", "caption of a form with #If and without UserForm_Initialize", "CompIf form"
    AddCase ws, "C04", "A3.5", "button caption on that form", "Close"
    AddCase ws, "C05", "A3.5", "Click handler on that form", "connected"
End Sub

'04: designer texts with quotes and line breaks (A3.6). The expected values are read back from the
'designer, so they are exactly what the original form shows at run time.
Private Sub BuildDesignerTexts(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim vbcForm As Object
    Dim ctl     As Object
    AddComponent wb, CT_STD_MODULE, "modTobfTexts", CodeTextsModule()
    Set vbcForm = AddComponent(wb, CT_MS_FORM, "frmTobfTexts", CodeTextsForm())
    SetFormFrame vbcForm, "Say `Hi`", 330, 250
    vbcForm.Properties("Tag").Value = Unquote("tag `form`")
    SetModuleCode SheetComponent(wb, ws), CodeCheckSheet(CodeTextsActual(), False, False)
    AddCase ws, "T01", "A3.6", "literal in the standard module", "texts"
    AddCase ws, "T02", "A3.6", "form caption with quotes", CStr(vbcForm.Properties("Caption").Value)
    AddCase ws, "T03", "A3.6", "form tag with quotes", CStr(vbcForm.Properties("Tag").Value)

    Set ctl = AddFormControl(vbcForm, "Forms.Label.1", "lblTobfCrLf", 6, 6, 150, 30)
    ctl.Caption = "Line 1" & vbCrLf & "Line 2"
    AddCase ws, "T04", "A3.6", "label caption with CR LF", ctl.Caption
    Set ctl = AddFormControl(vbcForm, "Forms.Label.1", "lblTobfLf", 6, 42, 150, 30)
    ctl.Caption = "Line A" & vbLf & "Line B"
    AddCase ws, "T05", "A3.6", "label caption with LF", ctl.Caption
    Set ctl = AddFormControl(vbcForm, "Forms.Label.1", "lblTobfQuote", 6, 78, 150, 18)
    ctl.Caption = Unquote("Label `q`")
    AddCase ws, "T06", "A3.6", "label caption with quotes", ctl.Caption
    Set ctl = AddFormControl(vbcForm, "Forms.TextBox.1", "txtTobfMulti", 6, 102, 150, 36)
    ctl.MultiLine = True
    ctl.Text = "first" & vbCrLf & "second"
    AddCase ws, "T07", "A3.6", "multi-line text box text", ctl.Text
    Set ctl = AddFormControl(vbcForm, "Forms.TextBox.1", "txtTobfQuote", 6, 144, 150, 18)
    ctl.Text = Unquote("say `x`")
    AddCase ws, "T08", "A3.6", "text box text with quotes", ctl.Text
    Set ctl = AddFormControl(vbcForm, "Forms.CommandButton.1", "cmdTobfRun", 6, 168, 96, 24)
    ctl.Caption = Unquote("Run `now`")
    ctl.ControlTipText = Unquote("tip `c`")
    ctl.Tag = Unquote("tag `c`")
    AddCase ws, "T09", "A3.6", "button caption with quotes", ctl.Caption
    AddCase ws, "T10", "A3.6", "control tip text with quotes", ctl.ControlTipText
    AddCase ws, "T11", "A3.6", "control tag with quotes", ctl.Tag
    Set ctl = AddFormControl(vbcForm, "Forms.Frame.1", "fraTobf", 168, 6, 140, 60)
    ctl.Caption = Unquote("Frame `f`")
    AddCase ws, "T12", "A3.6", "frame caption with quotes", ctl.Caption
    Set ctl = AddFormControl(vbcForm, "Forms.CheckBox.1", "chkTobf", 168, 72, 140, 18)
    ctl.Caption = Unquote("Check `c`")
    AddCase ws, "T13", "A3.6", "check box caption with quotes", ctl.Caption
    Set ctl = AddFormControl(vbcForm, "Forms.OptionButton.1", "optTobf", 168, 96, 140, 18)
    ctl.Caption = Unquote("Option `o`")
    AddCase ws, "T14", "A3.6", "option button caption with quotes", ctl.Caption
    Set ctl = AddFormControl(vbcForm, "Forms.ToggleButton.1", "tglTobf", 168, 120, 140, 24)
    ctl.Caption = Unquote("Toggle `t`")
    AddCase ws, "T15", "A3.6", "toggle button caption with quotes", ctl.Caption
End Sub

'05: event handlers that the exclusion lists miss (A3.7). The check looks for the handler name that
'binds the procedure to its event, built from the current name of the control or variable.
Private Sub BuildEventNames(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim vbcForm  As Object
    Dim ctl      As Object
    Dim wsEvents As Worksheet
    Dim chEvents As Chart
    Dim i        As Long
    Dim vCases   As Variant
    Set wsEvents = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsEvents.Name = "Events"
    Set chEvents = wb.Charts.Add(After:=wb.Sheets(wb.Sheets.Count))
    chEvents.Name = "EventChart"
    AddComponent wb, CT_STD_MODULE, "modTobfEvents", CodeEventsModule()
    AddComponent wb, CT_CLASS_MODULE, "clsTobfAppEvents", CodeEventsClass()
    AddComponent wb, CT_CLASS_MODULE, "clsTobfChartEvents", CodeEventsChartVar()
    AddComponent wb, CT_CLASS_MODULE, "clsTobfQueryEvents", CodeEventsQueryVar()
    Set vbcForm = AddComponent(wb, CT_MS_FORM, "frmTobfEvents", CodeEventsForm())
    SetFormFrame vbcForm, "Event form", 200, 110
    Set ctl = AddFormControl(vbcForm, "Forms.SpinButton.1", "spnTobf", 6, 6, 18, 36)
    Set ctl = AddFormControl(vbcForm, "Forms.Frame.1", "fraTobf", 36, 6, 120, 60)
    ctl.Caption = "Frame"
    SetModuleCode SheetComponent(wb, wsEvents), CodeEventsSheet()
    SetModuleCode SheetComponent(wb, chEvents), CodeEventsChart()
    SetModuleCode BookComponent(wb), CodeEventsBook()
    SetModuleCode SheetComponent(wb, ws), CodeCheckSheet(CodeEventsActual(), True, True)
    vCases = Array("SpinButton SpinUp", "SpinButton SpinDown", "Frame AddControl", _
                   "UserForm_AddControl (excluded today)", "Worksheet_BeforeDelete", "Worksheet_TableUpdate", _
                   "Worksheet_LensGalleryRenderComplete", "Workbook_SheetBeforeDelete", "Workbook_SheetTableUpdate", _
                   "Workbook_SheetLensGalleryRenderComplete", "Workbook_ModelChange", "Workbook_BeforeRemoteChange", _
                   "Workbook_AfterRemoteChange", "Chart_Select", "Chart_DragOver", "Chart_DragPlot", _
                   "Application AfterCalculate (WithEvents in a class)", "Application ProtectedViewWindowOpen", _
                   "Application ProtectedViewWindowActivate", "SpinButton Change (excluded today)")
    For i = LBound(vCases) To UBound(vCases)
        AddCase ws, "E" & Format$(i + 1, "00"), "A3.7", CStr(vCases(i)), "connected"
    Next i
    AddCase ws, "E21", "A3.7", "literal in the standard module", "events"
    AddCase ws, "E22", "A3.7b", "ActiveX Click on a worksheet", "connected"
    AddCase ws, "E23", "A3.7b", "ActiveX DblClick on a worksheet", "connected"
    AddCase ws, "E24", "A3.7b", "ActiveX GotFocus on a worksheet", "connected"
    AddCase ws, "E25", "A3.7b", "ActiveX LostFocus on a worksheet", "connected"
    AddCase ws, "E26", "A3.7b", "ActiveX KeyDown on a worksheet", "connected"
    AddCase ws, "E27", "A3.7b", "ActiveX KeyPress on a worksheet", "connected"
    AddCase ws, "E28", "A3.7b", "ActiveX KeyUp on a worksheet", "connected"
    AddCase ws, "E29", "A3.7b", "ActiveX Error on a worksheet", "connected"
    AddCase ws, "E30", "A3.7b", "ActiveX DropButtonClick on a worksheet", "connected"
    AddCase ws, "E31", "A3.7b", "ActiveX BeforeDragOver on a worksheet", "connected"
    AddCase ws, "E32", "A3.7b", "ActiveX BeforeDropOrPaste on a worksheet", "connected"
    AddCase ws, "E33", "A3.7b", "ActiveX SpinUp on a worksheet", "connected"
    AddCase ws, "E34", "A3.7b", "ActiveX SpinDown on a worksheet", "connected"
    AddCase ws, "E35", "A3.7b", "ActiveX Scroll on a worksheet", "connected"
    AddCase ws, "E36", "A3.7b", "Application ProtectedViewWindowBeforeClose", "connected"
    AddCase ws, "E37", "A3.7b", "Application ProtectedViewWindowBeforeEdit", "connected"
    AddCase ws, "E38", "A3.7b", "Application ProtectedViewWindowDeactivate", "connected"
    AddCase ws, "E39", "A3.7b", "Application ProtectedViewWindowResize", "connected"
    AddCase ws, "E40", "A3.7b", "Application WorkbookAfterRemoteChange", "connected"
    AddCase ws, "E41", "A3.7b", "Application WorkbookBeforeRemoteChange", "connected"
    AddCase ws, "E42", "A3.7b", "Chart Select via WithEvents variable", "connected"
    AddCase ws, "E43", "A3.7b", "Chart DragOver via WithEvents variable", "connected"
    AddCase ws, "E44", "A3.7b", "Chart DragPlot via WithEvents variable", "connected"
    AddCase ws, "E45", "A3.7b", "QueryTable AfterRefresh via WithEvents variable", "connected"
    AddCase ws, "E46", "A3.7b", "QueryTable BeforeRefresh via WithEvents variable", "connected"
End Sub

'06: literals as default values of optional parameters (A3.8)
Private Sub BuildOptionalParams(ByVal wb As Workbook, ByVal ws As Worksheet)
    AddComponent wb, CT_STD_MODULE, "modTobfOptional", CodeOptionalModule()
    AddComponent wb, CT_CLASS_MODULE, "clsTobfOptional", CodeOptionalClass()
    SetModuleCode SheetComponent(wb, ws), CodeCheckSheet(CodeOptionalActual(), False, False)
    AddCase ws, "P01", "A3.8", "default in a one-line Function header", "single"
    AddCase ws, "P02", "A3.8", "default in a continued header line", "continued"
    AddCase ws, "P03", "A3.8", "default in a Sub header", "sub"
    AddCase ws, "P04", "A3.8", "default in a Property Get header", "prop"
    AddCase ws, "P05", "A3.8", "explicit argument instead of the default", "given"
End Sub

'--------------------------------------------------------------------------------------------- helpers

'True if the VBA project of this workbook can be read (the same setting covers the sample workbooks)
Private Function ProjectAccessTrusted() As Boolean
    Dim lCount As Long
    On Error Resume Next
    lCount = ThisWorkbook.VBProject.VBComponents.Count
    ProjectAccessTrusted = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function IsWorkbookOpen(ByVal sName As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, sName, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

'Header of the sheet Check. Expected and actual values are text, so Excel converts none of them.
Private Sub PrepareCheckSheet(ByVal ws As Worksheet)
    ws.Range("A1:F1").Value = Array("Id", "Step", "Case", "Expected", "Actual", "Result")
    ws.Range("A1:F1").Font.Bold = True
    ws.Columns("D:E").NumberFormat = "@"
    ws.Range("H1").Value = "Double-click any cell of this sheet to run the checks. Line breaks are shown as {CR} and {LF}."
End Sub

'Adds one case to the sheet Check; in sExpected a backtick stands for a double quote
Private Sub AddCase(ByVal ws As Worksheet, ByVal sId As String, ByVal sStep As String, ByVal sCase As String, _
                    ByVal sExpected As String)
    Dim lRow As Long
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lRow, 1).Value = sId
    ws.Cells(lRow, 2).Value = sStep
    ws.Cells(lRow, 3).Value = sCase
    ws.Cells(lRow, 4).Value = Replace(Replace(Unquote(sExpected), vbCr, "{CR}"), vbLf, "{LF}")
End Sub

'Code of the sheet Check: double-click handler and case runner, the function TobfActual of the sample,
'and the helpers it needs
Private Function CodeCheckSheet(ByVal sActual As String, ByVal bHandler As Boolean, ByVal bEventVariable As Boolean) As String
    Dim s As String
    s = CodeCheckCommon() & sActual
    If bHandler Then s = s & CodeCheckHandler()
    If bEventVariable Then s = s & CodeCheckEventVariable()
    CodeCheckSheet = s
End Function

'Adds a component of type lType named sName with the code sCode
Private Function AddComponent(ByVal wb As Workbook, ByVal lType As Long, ByVal sName As String, ByVal sCode As String) As Object
    Dim vbc As Object
    Set vbc = wb.VBProject.VBComponents.Add(lType)
    vbc.Name = sName
    SetModuleCode vbc, sCode
    If lType = CT_MS_FORM Then EnsureMsFormsReference wb
    Set AddComponent = vbc
End Function

'Replaces the whole code of a component. A new module can already hold Option Explicit (VBE option
'Require Variable Declaration); the templates bring their own. No empty line is left at the end.
Private Sub SetModuleCode(ByVal vbc As Object, ByVal sCode As String)
    With vbc.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Right$(sCode, 2) = vbCrLf Then sCode = Left$(sCode, Len(sCode) - 2)
        .InsertLines 1, sCode
    End With
End Sub

'A3.5 needs modules whose last line ends a procedure; an empty last line would hide the defect
Private Sub CheckLastLine(ByVal vbc As Object)
    Dim sLast As String
    With vbc.CodeModule
        sLast = Trim$(.Lines(.CountOfLines, 1))
    End With
    If sLast <> "End Sub" And sLast <> "End Function" Then
        m_sWarnings = m_sWarnings & vbLf & "    " & vbc.Name & ": last line is not End Sub or End Function, A3.5 is not tested"
    End If
End Sub

'Adds a control to the designer of a UserForm, for example sProgId = Forms.Label.1
Private Function AddFormControl(ByVal vbcForm As Object, ByVal sProgId As String, ByVal sName As String, _
                                ByVal dLeft As Double, ByVal dTop As Double, ByVal dWidth As Double, _
                                ByVal dHeight As Double) As Object
    Dim ctl As Object
    Set ctl = vbcForm.Designer.Controls.Add(sProgId)
    ctl.Name = sName
    ctl.Left = dLeft
    ctl.Top = dTop
    ctl.Width = dWidth
    ctl.Height = dHeight
    Set AddFormControl = ctl
End Function

'Caption and size of a UserForm in the designer; in sCaption a backtick stands for a double quote
Private Sub SetFormFrame(ByVal vbcForm As Object, ByVal sCaption As String, ByVal dWidth As Double, ByVal dHeight As Double)
    vbcForm.Properties("Caption").Value = Unquote(sCaption)
    vbcForm.Properties("Width").Value = dWidth
    vbcForm.Properties("Height").Value = dHeight
End Sub

'Microsoft Forms 2.0 normally comes with the first UserForm; the event signatures of the samples need it
Private Sub EnsureMsFormsReference(ByVal wb As Workbook)
    Dim objRef As Object
    For Each objRef In wb.VBProject.References
        If StrComp(objRef.GUID, MSFORMS_GUID, vbTextCompare) = 0 Then Exit Sub
    Next objRef
    wb.VBProject.References.AddFromGuid MSFORMS_GUID, 2, 0
End Sub

'Code module of a worksheet or chart sheet. CodeName can be empty for a sheet added by code, so the
'fallback compares the Name property of the document components with the sheet name.
Private Function SheetComponent(ByVal wb As Workbook, ByVal objSheet As Object) As Object
    Dim vbc As Object
    If Len(objSheet.CodeName) > 0 Then
        Set SheetComponent = wb.VBProject.VBComponents(objSheet.CodeName)
        Exit Function
    End If
    For Each vbc In wb.VBProject.VBComponents
        If vbc.Type = CT_DOCUMENT Then
            If DocumentName(vbc) = objSheet.Name Then
                Set SheetComponent = vbc
                Exit Function
            End If
        End If
    Next vbc
    Err.Raise vbObjectError + 601, "TObfBuilder.SheetComponent", "No code module found for the sheet " & objSheet.Name
End Function

'Code module of the workbook (ThisWorkbook): the document component that belongs to no sheet
Private Function BookComponent(ByVal wb As Workbook) As Object
    Dim vbc      As Object
    Dim objSheet As Object
    Dim sName    As String
    Dim bSheet   As Boolean
    If Len(wb.CodeName) > 0 Then
        Set BookComponent = wb.VBProject.VBComponents(wb.CodeName)
        Exit Function
    End If
    For Each vbc In wb.VBProject.VBComponents
        If vbc.Type = CT_DOCUMENT Then
            sName = DocumentName(vbc)
            bSheet = False
            For Each objSheet In wb.Sheets
                If sName = objSheet.Name Then bSheet = True
            Next objSheet
            If Len(sName) > 0 And Not bSheet Then
                Set BookComponent = vbc
                Exit Function
            End If
        End If
    Next vbc
    Err.Raise vbObjectError + 602, "TObfBuilder.BookComponent", "No code module found for the workbook"
End Function

'Name property of a document component (sheet name or workbook name); empty if it cannot be read
Private Function DocumentName(ByVal vbc As Object) As String
    On Error Resume Next
    DocumentName = CStr(vbc.Properties("Name").Value)
    On Error GoTo 0
End Function

'A backtick in sText becomes a double quote
Private Function Unquote(ByVal sText As String) As String
    Unquote = Replace(sText, QUOTE_MARK, Chr$(34))
End Function

'Appends one template line to sCode
Private Sub AddLine(ByRef sCode As String, ByVal sLine As String)
    sCode = sCode & Unquote(sLine) & vbCrLf
End Sub

'--------------------------------------------------------------------------------------------- templates
'Code of the sample workbooks, line by line; a backtick stands for a double quote.

Private Function CodeCheckCommon() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'T-OBF self-check, written by the T-OBF builder. A double-click on this sheet runs every case in column A"
    AddLine s, "'and writes the actual value (column E) and PASS or FAIL (column F). The expected values (column D) are cell"
    AddLine s, "'values, so the obfuscation cannot change them. Line breaks appear as {CR} and {LF}."
    AddLine s, "Private Sub Worksheet_BeforeDoubleClick(ByVal rngTobfTarget As Range, bTobfCancel As Boolean)"
    AddLine s, "    Dim lTobfRow As Long"
    AddLine s, "    Dim lTobfLast As Long"
    AddLine s, "    Dim lTobfPassed As Long"
    AddLine s, "    Dim sTobfValue As String"
    AddLine s, "    bTobfCancel = True"
    AddLine s, "    lTobfLast = Me.Cells(Me.Rows.Count, 1).End(xlUp).Row"
    AddLine s, "    For lTobfRow = 2 To lTobfLast"
    AddLine s, "        sTobfValue = TobfRunCase(CStr(Me.Cells(lTobfRow, 1).Value))"
    AddLine s, "        Me.Cells(lTobfRow, 5).Value = sTobfValue"
    AddLine s, "        If sTobfValue = CStr(Me.Cells(lTobfRow, 4).Value) Then"
    AddLine s, "            Me.Cells(lTobfRow, 6).Value = `PASS`"
    AddLine s, "            lTobfPassed = lTobfPassed + 1"
    AddLine s, "        Else"
    AddLine s, "            Me.Cells(lTobfRow, 6).Value = `FAIL`"
    AddLine s, "        End If"
    AddLine s, "    Next lTobfRow"
    AddLine s, "    Me.Cells(2, 8).Value = lTobfPassed & ` of ` & (lTobfLast - 1) & ` passed`"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "'Runs one case. A runtime error becomes the actual value, so the remaining cases still run."
    AddLine s, "Private Function TobfRunCase(ByVal sTobfId As String) As String"
    AddLine s, "    Dim sTobfValue As String"
    AddLine s, "    On Error Resume Next"
    AddLine s, "    sTobfValue = TobfActual(sTobfId)"
    AddLine s, "    If Err.Number <> 0 Then sTobfValue = `error ` & Err.Number & `: ` & Err.Description"
    AddLine s, "    On Error GoTo 0"
    AddLine s, "    TobfRunCase = Replace(Replace(sTobfValue, vbCr, `{CR}`), vbLf, `{LF}`)"
    AddLine s, "End Function"
    CodeCheckCommon = s
End Function

Private Function CodeCheckHandler() As String
    Dim s As String
    AddLine s, ""
    AddLine s, "'Returns connected if component sTobfComp has a procedure named sTobfProc (this name binds it to its"
    AddLine s, "'event), otherwise missing. Reads the own VBA project, so access to it must be trusted."
    AddLine s, "Private Function TobfHandler(ByVal sTobfComp As String, ByVal sTobfProc As String) As String"
    AddLine s, "    Dim objTobfModule As Object"
    AddLine s, "    Dim lTobfLine As Long"
    AddLine s, "    Set objTobfModule = ThisWorkbook.VBProject.VBComponents(sTobfComp).CodeModule"
    AddLine s, "    On Error Resume Next"
    AddLine s, "    lTobfLine = objTobfModule.ProcStartLine(sTobfProc, 0)"
    AddLine s, "    On Error GoTo 0"
    AddLine s, "    If lTobfLine > 0 Then"
    AddLine s, "        TobfHandler = `connected`"
    AddLine s, "    Else"
    AddLine s, "        TobfHandler = `missing`"
    AddLine s, "    End If"
    AddLine s, "End Function"
    CodeCheckHandler = s
End Function

Private Function CodeCheckEventVariable() As String
    Dim s As String
    AddLine s, ""
    AddLine s, "'Name of the WithEvents variable declared in component sTobfComp. Comment lines are skipped, because"
    AddLine s, "'step 2 of the obfuscation keeps every original line as a comment."
    AddLine s, "Private Function TobfEventVariable(ByVal sTobfComp As String) As String"
    AddLine s, "    Dim objTobfModule As Object"
    AddLine s, "    Dim lTobfLine As Long"
    AddLine s, "    Dim lTobfPos As Long"
    AddLine s, "    Dim sTobfLine As String"
    AddLine s, "    Set objTobfModule = ThisWorkbook.VBProject.VBComponents(sTobfComp).CodeModule"
    AddLine s, "    For lTobfLine = 1 To objTobfModule.CountOfDeclarationLines"
    AddLine s, "        sTobfLine = Trim$(objTobfModule.Lines(lTobfLine, 1))"
    AddLine s, "        lTobfPos = InStr(1, sTobfLine, `WithEvents `, vbTextCompare)"
    AddLine s, "        If lTobfPos > 0 And Left$(sTobfLine, 1) <> Chr$(39) Then"
    AddLine s, "            sTobfLine = Mid$(sTobfLine, lTobfPos + 11)"
    AddLine s, "            TobfEventVariable = Left$(sTobfLine, InStr(sTobfLine & ` `, ` `) - 1)"
    AddLine s, "            Exit Function"
    AddLine s, "        End If"
    AddLine s, "    Next lTobfLine"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "'Handler check for an Application event caught by the class clsTobfAppEvents"
    AddLine s, "Private Function TobfAppHandler(ByVal sTobfEvent As String) As String"
    AddLine s, "    Dim objTobfApp As clsTobfAppEvents"
    AddLine s, "    Set objTobfApp = New clsTobfAppEvents"
    AddLine s, "    TobfAppHandler = TobfHandler(TypeName(objTobfApp), TobfEventVariable(TypeName(objTobfApp)) & `_` & sTobfEvent)"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "'Handler check for a Chart event caught by a WithEvents variable in clsTobfChartEvents"
    AddLine s, "Private Function TobfChartHandler(ByVal sTobfEvent As String) As String"
    AddLine s, "    Dim objTobfChart As clsTobfChartEvents"
    AddLine s, "    Set objTobfChart = New clsTobfChartEvents"
    AddLine s, "    TobfChartHandler = TobfHandler(TypeName(objTobfChart), TobfEventVariable(TypeName(objTobfChart)) & `_` & sTobfEvent)"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "'Handler check for a QueryTable event caught by a WithEvents variable in clsTobfQueryEvents"
    AddLine s, "Private Function TobfQueryHandler(ByVal sTobfEvent As String) As String"
    AddLine s, "    Dim objTobfQuery As clsTobfQueryEvents"
    AddLine s, "    Set objTobfQuery = New clsTobfQueryEvents"
    AddLine s, "    TobfQueryHandler = TobfHandler(TypeName(objTobfQuery), TobfEventVariable(TypeName(objTobfQuery)) & `_` & sTobfEvent)"
    AddLine s, "End Function"
    CodeCheckEventVariable = s
End Function

Private Function CodeBaseModule() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Baseline: what the obfuscation handles correctly today"
    AddLine s, "Public Function TobfBaseText() As String"
    AddLine s, "    TobfBaseText = `Hello world`"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfBaseSum(ByVal lTobfFirst As Long, ByVal lTobfSecond As Long) As Long"
    AddLine s, "    TobfBaseSum = lTobfFirst + lTobfSecond"
    AddLine s, "End Function"
    CodeBaseModule = s
End Function

Private Function CodeBaseClass() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "Private m_sTobfItem As String"
    AddLine s, ""
    AddLine s, "Public Property Get TobfItemText() As String"
    AddLine s, "    TobfItemText = m_sTobfItem"
    AddLine s, "End Property"
    AddLine s, ""
    AddLine s, "Public Property Let TobfItemText(ByVal sTobfNew As String)"
    AddLine s, "    m_sTobfItem = sTobfNew"
    AddLine s, "End Property"
    CodeBaseClass = s
End Function

Private Function CodeBaseForm() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "Private Sub UserForm_Initialize()"
    AddLine s, "    Me.lblTobfState.Caption = `ready`"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub cmdTobfOk_Click()"
    AddLine s, "    Me.Hide"
    AddLine s, "End Sub"
    CodeBaseForm = s
End Function

Private Function CodeBaseActual() As String
    Dim s As String
    AddLine s, ""
    AddLine s, "'Actual value of one case (Id in column A)"
    AddLine s, "Private Function TobfActual(ByVal sTobfId As String) As String"
    AddLine s, "    Dim objTobfItem As clsTobfBase"
    AddLine s, "    Dim objTobfForm As frmTobfBase"
    AddLine s, "    Select Case sTobfId"
    AddLine s, "        Case `B01`"
    AddLine s, "            TobfActual = TobfBaseText()"
    AddLine s, "        Case `B02`"
    AddLine s, "            TobfActual = CStr(TobfBaseSum(2, 3))"
    AddLine s, "        Case `B03`"
    AddLine s, "            Set objTobfItem = New clsTobfBase"
    AddLine s, "            objTobfItem.TobfItemText = `abc`"
    AddLine s, "            TobfActual = objTobfItem.TobfItemText"
    AddLine s, "        Case Else"
    AddLine s, "            Set objTobfForm = New frmTobfBase"
    AddLine s, "            Select Case sTobfId"
    AddLine s, "                Case `B04`"
    AddLine s, "                    TobfActual = objTobfForm.Caption"
    AddLine s, "                Case `B05`"
    AddLine s, "                    TobfActual = objTobfForm.lblTobfInfo.Caption"
    AddLine s, "                Case `B06`"
    AddLine s, "                    TobfActual = objTobfForm.cmdTobfOk.Caption"
    AddLine s, "                Case `B07`"
    AddLine s, "                    TobfActual = objTobfForm.lblTobfState.Caption"
    AddLine s, "                Case `B08`"
    AddLine s, "                    TobfActual = TobfHandler(TypeName(objTobfForm), objTobfForm.cmdTobfOk.Name & `_Click`)"
    AddLine s, "                Case Else"
    AddLine s, "                    TobfActual = `unknown case`"
    AddLine s, "            End Select"
    AddLine s, "            Unload objTobfForm"
    AddLine s, "    End Select"
    AddLine s, "End Function"
    CodeBaseActual = s
End Function

Private Function CodeNoFormModule() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Project without UserForm (A3.3)"
    AddLine s, "Public Function TobfGreeting() As String"
    AddLine s, "    TobfGreeting = `Good morning`"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfJoined(ByVal lTobfNumber As Long) As String"
    AddLine s, "    TobfJoined = `abc` & CStr(lTobfNumber)"
    AddLine s, "End Function"
    CodeNoFormModule = s
End Function

Private Function CodeNoFormActual() As String
    Dim s As String
    AddLine s, ""
    AddLine s, "'Actual value of one case (Id in column A)"
    AddLine s, "Private Function TobfActual(ByVal sTobfId As String) As String"
    AddLine s, "    Select Case sTobfId"
    AddLine s, "        Case `N01`"
    AddLine s, "            TobfActual = TobfGreeting()"
    AddLine s, "        Case `N02`"
    AddLine s, "            TobfActual = TobfJoined(123)"
    AddLine s, "        Case Else"
    AddLine s, "            TobfActual = `unknown case`"
    AddLine s, "    End Select"
    AddLine s, "End Function"
    CodeNoFormActual = s
End Function

Private Function CodeQuotesModule() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Literals with doubled quotes (A3.4). Each function returns one literal."
    AddLine s, "Public Function TobfQ1() As String"
    AddLine s, "    TobfQ1 = `He said ``Hi```"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfQ2() As String"
    AddLine s, "    TobfQ2 = ```Hi`` he said`"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfQ3() As String"
    AddLine s, "    TobfQ3 = `a``b`"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfQ4() As String"
    AddLine s, "    TobfQ4 = ````"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfQ5() As String"
    AddLine s, "    TobfQ5 = `x` & ```` & `y`"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfQ6() As String"
    AddLine s, "    TobfQ6 = `say ``a`` and ``b```"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfQ7() As String"
    AddLine s, "    TobfQ7 = ```a``` & `b`"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfQ8() As String"
    AddLine s, "    TobfQ8 = `no quotes`"
    AddLine s, "End Function"
    CodeQuotesModule = s
End Function

Private Function CodeQuotesActual() As String
    Dim s As String
    AddLine s, ""
    AddLine s, "'Actual value of one case (Id in column A)"
    AddLine s, "Private Function TobfActual(ByVal sTobfId As String) As String"
    AddLine s, "    Select Case sTobfId"
    AddLine s, "        Case `Q01`"
    AddLine s, "            TobfActual = TobfQ1()"
    AddLine s, "        Case `Q02`"
    AddLine s, "            TobfActual = TobfQ2()"
    AddLine s, "        Case `Q03`"
    AddLine s, "            TobfActual = TobfQ3()"
    AddLine s, "        Case `Q04`"
    AddLine s, "            TobfActual = TobfQ4()"
    AddLine s, "        Case `Q05`"
    AddLine s, "            TobfActual = TobfQ5()"
    AddLine s, "        Case `Q06`"
    AddLine s, "            TobfActual = TobfQ6()"
    AddLine s, "        Case `Q07`"
    AddLine s, "            TobfActual = TobfQ7()"
    AddLine s, "        Case `Q08`"
    AddLine s, "            TobfActual = TobfQ8()"
    AddLine s, "        Case Else"
    AddLine s, "            TobfActual = `unknown case`"
    AddLine s, "    End Select"
    AddLine s, "End Function"
    CodeQuotesActual = s
End Function

Private Function CodeCompIfModule() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, "#If VBA7 Then"
    AddLine s, "Private Const TOBF_LEVEL As Long = 7"
    AddLine s, "#Else"
    AddLine s, "Private Const TOBF_LEVEL As Long = 6"
    AddLine s, "#End If"
    AddLine s, ""
    AddLine s, "'Only standard module of this workbook, with #If in its declarations (A3.5): the decrypt function of the"
    AddLine s, "'string option is inserted here. The last line of the module must end a procedure."
    AddLine s, "Public Function TobfCompIfText() As String"
    AddLine s, "    TobfCompIfText = `compiled`"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfCompIfLevel() As Long"
    AddLine s, "    TobfCompIfLevel = TOBF_LEVEL"
    AddLine s, "End Function"
    CodeCompIfModule = s
End Function

Private Function CodeCompIfForm() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, "#If VBA7 Then"
    AddLine s, "Private m_lTobfClicks As Long"
    AddLine s, "#Else"
    AddLine s, "Private m_lTobfClicks As Integer"
    AddLine s, "#End If"
    AddLine s, ""
    AddLine s, "'Form with #If in its declarations and without UserForm_Initialize (A3.5): the generated"
    AddLine s, "'UserForm_Initialize is inserted here. The last line of the module must end a procedure."
    AddLine s, "Private Sub cmdTobfClose_Click()"
    AddLine s, "    m_lTobfClicks = m_lTobfClicks + 1"
    AddLine s, "    Me.Hide"
    AddLine s, "End Sub"
    CodeCompIfForm = s
End Function

Private Function CodeCompIfActual() As String
    Dim s As String
    AddLine s, ""
    AddLine s, "'Actual value of one case (Id in column A)"
    AddLine s, "Private Function TobfActual(ByVal sTobfId As String) As String"
    AddLine s, "    Dim objTobfForm As frmTobfCompIf"
    AddLine s, "    Select Case sTobfId"
    AddLine s, "        Case `C01`"
    AddLine s, "            TobfActual = TobfCompIfText()"
    AddLine s, "        Case `C02`"
    AddLine s, "            TobfActual = CStr(TobfCompIfLevel())"
    AddLine s, "        Case Else"
    AddLine s, "            Set objTobfForm = New frmTobfCompIf"
    AddLine s, "            Select Case sTobfId"
    AddLine s, "                Case `C03`"
    AddLine s, "                    TobfActual = objTobfForm.Caption"
    AddLine s, "                Case `C04`"
    AddLine s, "                    TobfActual = objTobfForm.cmdTobfClose.Caption"
    AddLine s, "                Case `C05`"
    AddLine s, "                    TobfActual = TobfHandler(TypeName(objTobfForm), objTobfForm.cmdTobfClose.Name & `_Click`)"
    AddLine s, "                Case Else"
    AddLine s, "                    TobfActual = `unknown case`"
    AddLine s, "            End Select"
    AddLine s, "            Unload objTobfForm"
    AddLine s, "    End Select"
    AddLine s, "End Function"
    CodeCompIfActual = s
End Function

Private Function CodeTextsModule() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Standard module, so that the string option has a target module for its decrypt function"
    AddLine s, "Public Function TobfTextsReady() As String"
    AddLine s, "    TobfTextsReady = `texts`"
    AddLine s, "End Function"
    CodeTextsModule = s
End Function

Private Function CodeTextsForm() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Designer texts with quotes and line breaks (A3.6), set by the builder"
    AddLine s, "Private Sub cmdTobfRun_Click()"
    AddLine s, "    Me.Hide"
    AddLine s, "End Sub"
    CodeTextsForm = s
End Function

Private Function CodeTextsActual() As String
    Dim s As String
    AddLine s, ""
    AddLine s, "'Actual value of one case (Id in column A)"
    AddLine s, "Private Function TobfActual(ByVal sTobfId As String) As String"
    AddLine s, "    Dim objTobfForm As frmTobfTexts"
    AddLine s, "    If sTobfId = `T01` Then"
    AddLine s, "        TobfActual = TobfTextsReady()"
    AddLine s, "        Exit Function"
    AddLine s, "    End If"
    AddLine s, "    Set objTobfForm = New frmTobfTexts"
    AddLine s, "    Select Case sTobfId"
    AddLine s, "        Case `T02`"
    AddLine s, "            TobfActual = objTobfForm.Caption"
    AddLine s, "        Case `T03`"
    AddLine s, "            TobfActual = objTobfForm.Tag"
    AddLine s, "        Case `T04`"
    AddLine s, "            TobfActual = objTobfForm.lblTobfCrLf.Caption"
    AddLine s, "        Case `T05`"
    AddLine s, "            TobfActual = objTobfForm.lblTobfLf.Caption"
    AddLine s, "        Case `T06`"
    AddLine s, "            TobfActual = objTobfForm.lblTobfQuote.Caption"
    AddLine s, "        Case `T07`"
    AddLine s, "            TobfActual = objTobfForm.txtTobfMulti.Text"
    AddLine s, "        Case `T08`"
    AddLine s, "            TobfActual = objTobfForm.txtTobfQuote.Text"
    AddLine s, "        Case `T09`"
    AddLine s, "            TobfActual = objTobfForm.cmdTobfRun.Caption"
    AddLine s, "        Case `T10`"
    AddLine s, "            TobfActual = objTobfForm.cmdTobfRun.ControlTipText"
    AddLine s, "        Case `T11`"
    AddLine s, "            TobfActual = objTobfForm.cmdTobfRun.Tag"
    AddLine s, "        Case `T12`"
    AddLine s, "            TobfActual = objTobfForm.fraTobf.Caption"
    AddLine s, "        Case `T13`"
    AddLine s, "            TobfActual = objTobfForm.chkTobf.Caption"
    AddLine s, "        Case `T14`"
    AddLine s, "            TobfActual = objTobfForm.optTobf.Caption"
    AddLine s, "        Case `T15`"
    AddLine s, "            TobfActual = objTobfForm.tglTobf.Caption"
    AddLine s, "        Case Else"
    AddLine s, "            TobfActual = `unknown case`"
    AddLine s, "    End Select"
    AddLine s, "    Unload objTobfForm"
    AddLine s, "End Function"
    CodeTextsActual = s
End Function

Private Function CodeEventsModule() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Standard module, so that the string option has a target module for its decrypt function"
    AddLine s, "Public Function TobfEventsReady() As String"
    AddLine s, "    TobfEventsReady = `events`"
    AddLine s, "End Function"
    CodeEventsModule = s
End Function

Private Function CodeEventsForm() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'MSForms events that the exclusion lists miss (A3.7), plus UserForm_AddControl and Change as controls"
    AddLine s, "Private m_lTobfFormEvents As Long"
    AddLine s, ""
    AddLine s, "Private Sub spnTobf_SpinUp()"
    AddLine s, "    m_lTobfFormEvents = m_lTobfFormEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub spnTobf_SpinDown()"
    AddLine s, "    m_lTobfFormEvents = m_lTobfFormEvents - 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub spnTobf_Change()"
    AddLine s, "    m_lTobfFormEvents = m_lTobfFormEvents + 2"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub fraTobf_AddControl(ByVal ctlTobfAdded As MSForms.Control)"
    AddLine s, "    m_lTobfFormEvents = m_lTobfFormEvents + 10"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub UserForm_AddControl(ByVal ctlTobfAdded As MSForms.Control)"
    AddLine s, "    m_lTobfFormEvents = m_lTobfFormEvents + 100"
    AddLine s, "End Sub"
    CodeEventsForm = s
End Function

Private Function CodeEventsSheet() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Worksheet events that the exclusion lists miss (A3.7)"
    AddLine s, "Private m_lTobfSheetEvents As Long"
    AddLine s, ""
    AddLine s, "Private Sub Worksheet_BeforeDelete()"
    AddLine s, "    m_lTobfSheetEvents = m_lTobfSheetEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub Worksheet_TableUpdate(ByVal tblTobfTarget As TableObject)"
    AddLine s, "    m_lTobfSheetEvents = m_lTobfSheetEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub Worksheet_LensGalleryRenderComplete()"
    AddLine s, "    m_lTobfSheetEvents = m_lTobfSheetEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "'A3.7b: events of ActiveX controls hosted on a worksheet. No control is placed; only the handler"
    AddLine s, "'names count, as for the other cases. Without the A3.7b fix step 1 renames them and the control stops responding."
    AddLine s, "Private m_lTobfAxEvents As Long"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBtn_Click()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBtn_DblClick()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBtn_GotFocus()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBtn_LostFocus()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBtn_KeyDown()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBtn_KeyPress()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBtn_KeyUp()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBtn_Error()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBox_DropButtonClick()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBox_BeforeDragOver()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfBox_BeforeDropOrPaste()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfSpn_SpinUp()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfSpn_SpinDown()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub axTobfScr_Scroll()"
    AddLine s, "    m_lTobfAxEvents = m_lTobfAxEvents + 1"
    AddLine s, "End Sub"
    CodeEventsSheet = s
End Function

Private Function CodeEventsBook() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Workbook events that the exclusion lists miss (A3.7)"
    AddLine s, "Private m_lTobfBookEvents As Long"
    AddLine s, ""
    AddLine s, "Private Sub Workbook_SheetBeforeDelete(ByVal objTobfSheet As Object)"
    AddLine s, "    m_lTobfBookEvents = m_lTobfBookEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub Workbook_SheetTableUpdate(ByVal objTobfSheet As Object, ByVal tblTobfTarget As TableObject)"
    AddLine s, "    m_lTobfBookEvents = m_lTobfBookEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub Workbook_SheetLensGalleryRenderComplete(ByVal objTobfSheet As Object)"
    AddLine s, "    m_lTobfBookEvents = m_lTobfBookEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub Workbook_ModelChange(ByVal mdcTobfChanges As ModelChanges)"
    AddLine s, "    m_lTobfBookEvents = m_lTobfBookEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub Workbook_BeforeRemoteChange()"
    AddLine s, "    m_lTobfBookEvents = m_lTobfBookEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub Workbook_AfterRemoteChange()"
    AddLine s, "    m_lTobfBookEvents = m_lTobfBookEvents + 1"
    AddLine s, "End Sub"
    CodeEventsBook = s
End Function

Private Function CodeEventsChart() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Chart sheet events that the exclusion lists miss (A3.7)"
    AddLine s, "Private m_lTobfChartEvents As Long"
    AddLine s, ""
    AddLine s, "Private Sub Chart_Select(ByVal lTobfElement As Long, ByVal lTobfArgOne As Long, ByVal lTobfArgTwo As Long)"
    AddLine s, "    m_lTobfChartEvents = m_lTobfChartEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub Chart_DragOver()"
    AddLine s, "    m_lTobfChartEvents = m_lTobfChartEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub Chart_DragPlot()"
    AddLine s, "    m_lTobfChartEvents = m_lTobfChartEvents + 1"
    AddLine s, "End Sub"
    CodeEventsChart = s
End Function

Private Function CodeEventsClass() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Application events that the exclusion lists miss (A3.7). The variable is never set; only the names count."
    AddLine s, "Private WithEvents appTobf As Application"
    AddLine s, "Private m_lTobfAppEvents As Long"
    AddLine s, ""
    AddLine s, "Private Sub appTobf_AfterCalculate()"
    AddLine s, "    m_lTobfAppEvents = m_lTobfAppEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub appTobf_ProtectedViewWindowOpen(ByVal pvwTobf As ProtectedViewWindow)"
    AddLine s, "    m_lTobfAppEvents = m_lTobfAppEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub appTobf_ProtectedViewWindowActivate(ByVal pvwTobf As ProtectedViewWindow)"
    AddLine s, "    m_lTobfAppEvents = m_lTobfAppEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub appTobf_ProtectedViewWindowBeforeClose(ByVal pvwTobf As ProtectedViewWindow, ByVal lTobfReason As XlProtectedViewCloseReason, bTobfCancel As Boolean)"
    AddLine s, "    m_lTobfAppEvents = m_lTobfAppEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub appTobf_ProtectedViewWindowBeforeEdit(ByVal pvwTobf As ProtectedViewWindow, bTobfCancel As Boolean)"
    AddLine s, "    m_lTobfAppEvents = m_lTobfAppEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub appTobf_ProtectedViewWindowDeactivate(ByVal pvwTobf As ProtectedViewWindow)"
    AddLine s, "    m_lTobfAppEvents = m_lTobfAppEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub appTobf_ProtectedViewWindowResize(ByVal pvwTobf As ProtectedViewWindow)"
    AddLine s, "    m_lTobfAppEvents = m_lTobfAppEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub appTobf_WorkbookAfterRemoteChange(ByVal wbTobf As Workbook)"
    AddLine s, "    m_lTobfAppEvents = m_lTobfAppEvents + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub appTobf_WorkbookBeforeRemoteChange(ByVal wbTobf As Workbook)"
    AddLine s, "    m_lTobfAppEvents = m_lTobfAppEvents + 1"
    AddLine s, "End Sub"
    CodeEventsClass = s
End Function

Private Function CodeEventsChartVar() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Chart events for a WithEvents variable (A3.7). SERVICEWORDS protects only Chart_... in a chart sheet,"
    AddLine s, "'so the variable path relies on the code lists. The variable is never set; only the names count."
    AddLine s, "Private WithEvents chTobf As Chart"
    AddLine s, "Private m_lTobfChartVar As Long"
    AddLine s, ""
    AddLine s, "Private Sub chTobf_Select(ByVal lTobfElement As Long, ByVal lTobfArgOne As Long, ByVal lTobfArgTwo As Long)"
    AddLine s, "    m_lTobfChartVar = m_lTobfChartVar + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub chTobf_DragOver()"
    AddLine s, "    m_lTobfChartVar = m_lTobfChartVar + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub chTobf_DragPlot()"
    AddLine s, "    m_lTobfChartVar = m_lTobfChartVar + 1"
    AddLine s, "End Sub"
    CodeEventsChartVar = s
End Function

Private Function CodeEventsQueryVar() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'QueryTable events for a WithEvents variable (A3.7). The variable is never set; only the names count."
    AddLine s, "Private WithEvents qtTobf As QueryTable"
    AddLine s, "Private m_lTobfQueryVar As Long"
    AddLine s, ""
    AddLine s, "Private Sub qtTobf_BeforeRefresh(bTobfCancel As Boolean)"
    AddLine s, "    m_lTobfQueryVar = m_lTobfQueryVar + 1"
    AddLine s, "End Sub"
    AddLine s, ""
    AddLine s, "Private Sub qtTobf_AfterRefresh(ByVal bTobfSuccess As Boolean)"
    AddLine s, "    m_lTobfQueryVar = m_lTobfQueryVar + 1"
    AddLine s, "End Sub"
    CodeEventsQueryVar = s
End Function

Private Function CodeEventsActual() As String
    Dim s As String
    AddLine s, ""
    AddLine s, "'Actual value of one case (Id in column A)"
    AddLine s, "Private Function TobfActual(ByVal sTobfId As String) As String"
    AddLine s, "    Dim objTobfForm As frmTobfEvents"
    AddLine s, "    Dim sTobfSheet As String"
    AddLine s, "    Dim sTobfBook As String"
    AddLine s, "    Dim sTobfChart As String"
    AddLine s, "    sTobfSheet = ThisWorkbook.Worksheets(`Events`).CodeName"
    AddLine s, "    sTobfBook = ThisWorkbook.CodeName"
    AddLine s, "    sTobfChart = ThisWorkbook.Charts(`EventChart`).CodeName"
    AddLine s, "    Select Case sTobfId"
    AddLine s, "        Case `E01`, `E02`, `E03`, `E04`, `E20`"
    AddLine s, "            Set objTobfForm = New frmTobfEvents"
    AddLine s, "            Select Case sTobfId"
    AddLine s, "                Case `E01`"
    AddLine s, "                    TobfActual = TobfHandler(TypeName(objTobfForm), objTobfForm.spnTobf.Name & `_SpinUp`)"
    AddLine s, "                Case `E02`"
    AddLine s, "                    TobfActual = TobfHandler(TypeName(objTobfForm), objTobfForm.spnTobf.Name & `_SpinDown`)"
    AddLine s, "                Case `E03`"
    AddLine s, "                    TobfActual = TobfHandler(TypeName(objTobfForm), objTobfForm.fraTobf.Name & `_AddControl`)"
    AddLine s, "                Case `E04`"
    AddLine s, "                    TobfActual = TobfHandler(TypeName(objTobfForm), `UserForm_AddControl`)"
    AddLine s, "                Case `E20`"
    AddLine s, "                    TobfActual = TobfHandler(TypeName(objTobfForm), objTobfForm.spnTobf.Name & `_Change`)"
    AddLine s, "            End Select"
    AddLine s, "            Unload objTobfForm"
    AddLine s, "        Case `E05`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `Worksheet_BeforeDelete`)"
    AddLine s, "        Case `E06`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `Worksheet_TableUpdate`)"
    AddLine s, "        Case `E07`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `Worksheet_LensGalleryRenderComplete`)"
    AddLine s, "        Case `E08`"
    AddLine s, "            TobfActual = TobfHandler(sTobfBook, `Workbook_SheetBeforeDelete`)"
    AddLine s, "        Case `E09`"
    AddLine s, "            TobfActual = TobfHandler(sTobfBook, `Workbook_SheetTableUpdate`)"
    AddLine s, "        Case `E10`"
    AddLine s, "            TobfActual = TobfHandler(sTobfBook, `Workbook_SheetLensGalleryRenderComplete`)"
    AddLine s, "        Case `E11`"
    AddLine s, "            TobfActual = TobfHandler(sTobfBook, `Workbook_ModelChange`)"
    AddLine s, "        Case `E12`"
    AddLine s, "            TobfActual = TobfHandler(sTobfBook, `Workbook_BeforeRemoteChange`)"
    AddLine s, "        Case `E13`"
    AddLine s, "            TobfActual = TobfHandler(sTobfBook, `Workbook_AfterRemoteChange`)"
    AddLine s, "        Case `E14`"
    AddLine s, "            TobfActual = TobfHandler(sTobfChart, `Chart_Select`)"
    AddLine s, "        Case `E15`"
    AddLine s, "            TobfActual = TobfHandler(sTobfChart, `Chart_DragOver`)"
    AddLine s, "        Case `E16`"
    AddLine s, "            TobfActual = TobfHandler(sTobfChart, `Chart_DragPlot`)"
    AddLine s, "        Case `E17`"
    AddLine s, "            TobfActual = TobfAppHandler(`AfterCalculate`)"
    AddLine s, "        Case `E18`"
    AddLine s, "            TobfActual = TobfAppHandler(`ProtectedViewWindowOpen`)"
    AddLine s, "        Case `E19`"
    AddLine s, "            TobfActual = TobfAppHandler(`ProtectedViewWindowActivate`)"
    AddLine s, "        Case `E21`"
    AddLine s, "            TobfActual = TobfEventsReady()"
    AddLine s, "        Case `E22`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBtn_Click`)"
    AddLine s, "        Case `E23`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBtn_DblClick`)"
    AddLine s, "        Case `E24`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBtn_GotFocus`)"
    AddLine s, "        Case `E25`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBtn_LostFocus`)"
    AddLine s, "        Case `E26`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBtn_KeyDown`)"
    AddLine s, "        Case `E27`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBtn_KeyPress`)"
    AddLine s, "        Case `E28`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBtn_KeyUp`)"
    AddLine s, "        Case `E29`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBtn_Error`)"
    AddLine s, "        Case `E30`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBox_DropButtonClick`)"
    AddLine s, "        Case `E31`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBox_BeforeDragOver`)"
    AddLine s, "        Case `E32`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfBox_BeforeDropOrPaste`)"
    AddLine s, "        Case `E33`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfSpn_SpinUp`)"
    AddLine s, "        Case `E34`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfSpn_SpinDown`)"
    AddLine s, "        Case `E35`"
    AddLine s, "            TobfActual = TobfHandler(sTobfSheet, `axTobfScr_Scroll`)"
    AddLine s, "        Case `E36`"
    AddLine s, "            TobfActual = TobfAppHandler(`ProtectedViewWindowBeforeClose`)"
    AddLine s, "        Case `E37`"
    AddLine s, "            TobfActual = TobfAppHandler(`ProtectedViewWindowBeforeEdit`)"
    AddLine s, "        Case `E38`"
    AddLine s, "            TobfActual = TobfAppHandler(`ProtectedViewWindowDeactivate`)"
    AddLine s, "        Case `E39`"
    AddLine s, "            TobfActual = TobfAppHandler(`ProtectedViewWindowResize`)"
    AddLine s, "        Case `E40`"
    AddLine s, "            TobfActual = TobfAppHandler(`WorkbookAfterRemoteChange`)"
    AddLine s, "        Case `E41`"
    AddLine s, "            TobfActual = TobfAppHandler(`WorkbookBeforeRemoteChange`)"
    AddLine s, "        Case `E42`"
    AddLine s, "            TobfActual = TobfChartHandler(`Select`)"
    AddLine s, "        Case `E43`"
    AddLine s, "            TobfActual = TobfChartHandler(`DragOver`)"
    AddLine s, "        Case `E44`"
    AddLine s, "            TobfActual = TobfChartHandler(`DragPlot`)"
    AddLine s, "        Case `E45`"
    AddLine s, "            TobfActual = TobfQueryHandler(`AfterRefresh`)"
    AddLine s, "        Case `E46`"
    AddLine s, "            TobfActual = TobfQueryHandler(`BeforeRefresh`)"
    AddLine s, "        Case Else"
    AddLine s, "            TobfActual = `unknown case`"
    AddLine s, "    End Select"
    AddLine s, "End Function"
    CodeEventsActual = s
End Function

Private Function CodeOptionalModule() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Literals as default values of optional parameters (A3.8). VBA needs constants there."
    AddLine s, "Public Function TobfOptSingle(Optional ByVal sTobfValue As String = `single`) As String"
    AddLine s, "    TobfOptSingle = sTobfValue"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Function TobfOptContinued(ByVal lTobfDummy As Long, _"
    AddLine s, "                                 Optional ByVal sTobfValue As String = `continued`) As String"
    AddLine s, "    TobfOptContinued = sTobfValue"
    AddLine s, "End Function"
    AddLine s, ""
    AddLine s, "Public Sub TobfOptSub(ByRef sTobfOut As String, Optional ByVal sTobfValue As String = `sub`)"
    AddLine s, "    sTobfOut = sTobfValue"
    AddLine s, "End Sub"
    CodeOptionalModule = s
End Function

Private Function CodeOptionalClass() As String
    Dim s As String
    AddLine s, "Option Explicit"
    AddLine s, ""
    AddLine s, "'Property Get with an optional parameter whose default is a literal (A3.8)"
    AddLine s, "Public Property Get TobfOptProp(Optional ByVal sTobfKey As String = `prop`) As String"
    AddLine s, "    TobfOptProp = sTobfKey"
    AddLine s, "End Property"
    CodeOptionalClass = s
End Function

Private Function CodeOptionalActual() As String
    Dim s As String
    AddLine s, ""
    AddLine s, "'Actual value of one case (Id in column A)"
    AddLine s, "Private Function TobfActual(ByVal sTobfId As String) As String"
    AddLine s, "    Dim objTobfOpt As clsTobfOptional"
    AddLine s, "    Dim sTobfOut As String"
    AddLine s, "    Select Case sTobfId"
    AddLine s, "        Case `P01`"
    AddLine s, "            TobfActual = TobfOptSingle()"
    AddLine s, "        Case `P02`"
    AddLine s, "            TobfActual = TobfOptContinued(1)"
    AddLine s, "        Case `P03`"
    AddLine s, "            Call TobfOptSub(sTobfOut)"
    AddLine s, "            TobfActual = sTobfOut"
    AddLine s, "        Case `P04`"
    AddLine s, "            Set objTobfOpt = New clsTobfOptional"
    AddLine s, "            TobfActual = objTobfOpt.TobfOptProp"
    AddLine s, "        Case `P05`"
    AddLine s, "            TobfActual = TobfOptSingle(`given`)"
    AddLine s, "        Case Else"
    AddLine s, "            TobfActual = `unknown case`"
    AddLine s, "    End Select"
    AddLine s, "End Function"
    CodeOptionalActual = s
End Function
