Attribute VB_Name = "T_TransformChainTest"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : T_TransformChainTests - regression tests of the transformation chain
'* Created    : 02-10-2026
'* Usage      : 1. Import this module into the MACROTools add-in project (the project that
'*                 contains N_Obfuscation and K_AddNumbersLine).
'*              2. Trust access to the VBA project object model must be enabled.
'*              3. Run RunTransformChainTests: type RunTransformChainTests in the Immediate
'*                 window, or place the cursor in the procedure and press F5 (the module is
'*                 Option Private Module, so Alt+F8 does not list it).
'*              4. A new workbook opens with the sheet "Results" (one row per test).
'*              5. Optional compile check: in the VBA editor select the project of that
'*                 workbook and run Debug > Compile. All transformed modules kept there
'*                 must compile (fixtures marked "not compilable" are removed).
'*              The test workbook is not saved; close it without saving.
'* Status     : PASS = output, line count and notices as expected
'*              FAIL = deviation (expected and actual text are shown)
'*              INCONCLUSIVE = the VBA editor changed the test input when it was inserted
'*                             (e.g. indentation of Option lines, tabs); adapt the test data
'*              ERROR = runtime error during the test
'*              Diagnostic cases (X35g, X35h, D0 to D10) isolate editor behaviour: they use the text
'*              as the editor stores it, so INCONCLUSIVE there means the precondition is missing.
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

'Keep transformed modules of compilable fixtures in the test workbook for Debug > Compile
Private Const KEEP_COMPILABLE_MODULES As Boolean = True

'Module text of the label-or-call diagnostics (@ = case tag, @LINE = the numbered line, | = line break)
Private Const LC_TEMPLATE As String = "Private cnt@ As Long|Function P@() As Long|    cnt@ = 0|@LINE|End Function|Sub Mark@()|    cnt@ = cnt@ + 1|End Sub"

Private m_wb As Workbook
Private m_ws As Worksheet
Private m_lRow As Long
Private m_nPass As Long
Private m_nFail As Long
Private m_nInconclusive As Long

Public Sub RunTransformChainTests()
    On Error GoTo EH
    Application.ScreenUpdating = False
    Set m_wb = Workbooks.Add(xlWBATWorksheet)
    Set m_ws = m_wb.Worksheets(1)
    m_ws.Name = "Results"
    m_ws.Range("A1:J1").Value = Array("ID", "Method", "Purpose", "Status", "Notices (actual/expected)", _
                                      "Note", "Rationale", "Input", "Expected", "Actual")
    m_ws.Rows(1).Font.Bold = True
    m_lRow = 1
    m_nPass = 0
    m_nFail = 0
    m_nInconclusive = 0

    RunAllCases
    Case_T42_Attributes
    Case_X26_EditorLimit
    Case_X35g_NameAfterNumber
    Case_X35h_NameAfterNumberColon
    Case_X37_InsertBlankPosition
    Case_X38_ReplaceWithContinuation
    RunAttributeDiagnostics

    m_ws.Columns("A:G").AutoFit
    m_ws.Columns("H:J").ColumnWidth = 80
    m_ws.Range("H:J").WrapText = True
    Application.ScreenUpdating = True
    Debug.Print "Transformation chain tests: PASS " & m_nPass & ", FAIL " & m_nFail & ", INCONCLUSIVE " & m_nInconclusive
    MsgBox "PASS: " & m_nPass & vbLf & "FAIL: " & m_nFail & vbLf & "INCONCLUSIVE: " & m_nInconclusive & vbLf & vbLf & _
           "Details: sheet 'Results' of " & m_wb.Name & "." & vbLf & _
           "Optional: Debug > Compile on the project of that workbook.", vbInformation, "Transformation chain tests"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Test run aborted: error " & Err.Number & ": " & Err.Description, vbCritical, "Transformation chain tests"
End Sub

Private Sub AddLine(ByRef s As String, ByRef n As Long, ByVal sLine As String)
    If n > 0 Then s = s & vbCrLf
    s = s & sLine
    n = n + 1
End Sub

'Fills an empty code module with the given text. Excel 2019 appends an empty line when the inserted
'text contains a line continuation; that line is removed again, so the module holds exactly the text.
Private Sub FillModule(ByVal vbc As VBIDE.VBComponent, ByVal sText As String, ByVal nLines As Long)
    With vbc.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If nLines > 0 Then .InsertLines 1, sText
        If .CountOfLines = nLines + 1 Then
            If Len(Trim$(.Lines(.CountOfLines, 1))) = 0 Then .DeleteLines .CountOfLines, 1
        End If
    End With
End Sub

Private Function ReadTextFile(ByVal sPath As String) As String
    Dim f As Integer
    Dim sLine As String
    Dim s As String

    f = FreeFile
    Open sPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, sLine
        s = s & sLine & vbCrLf
    Loop
    Close #f
    ReadTextFile = s
End Function

Private Sub DeleteIfExists(ByVal sPath As String)
    If Len(Dir$(sPath)) > 0 Then Kill sPath
End Sub

'Same order as ObfuscationCode.lbOK_Click with all options selected.
Private Sub RunChain(ByVal vbc As VBIDE.VBComponent)
    K_AddNumbersLine.RemoveLineNumbers vbc, vbLineNumbers_LabelTypes.vbLabelColon
    K_AddNumbersLine.RemoveLineNumbers vbc, vbLineNumbers_LabelTypes.vbLabelTab
    N_Obfuscation.Remove_DebugPrint vbc.CodeModule
    N_Obfuscation.Remove_Comments vbc.CodeModule
    N_Obfuscation.TrimLinesTabAndSpase vbc.CodeModule
    N_Obfuscation.Remove_OptionExplicit vbc.CodeModule
    N_Obfuscation.Remove_EmptyLines vbc.CodeModule
    N_Obfuscation.RemoveBreaksLineInCode vbc.CodeModule
End Sub

Private Sub RunMethod(ByVal vbc As VBIDE.VBComponent, ByVal sMethod As String)
    Select Case sMethod
        Case "RLN"
            K_AddNumbersLine.RemoveLineNumbers vbc, vbLineNumbers_LabelTypes.vbLabelColon
            K_AddNumbersLine.RemoveLineNumbers vbc, vbLineNumbers_LabelTypes.vbLabelTab
        Case "RLN_COLON"
            K_AddNumbersLine.RemoveLineNumbers vbc, vbLineNumbers_LabelTypes.vbLabelColon
        Case "RLN_TAB"
            K_AddNumbersLine.RemoveLineNumbers vbc, vbLineNumbers_LabelTypes.vbLabelTab
        Case "DBG"
            N_Obfuscation.Remove_DebugPrint vbc.CodeModule
        Case "DBG_FALSE"
            N_Obfuscation.Remove_DebugPrint vbc.CodeModule, False
        Case "CMT"
            N_Obfuscation.Remove_Comments vbc.CodeModule
        Case "TRIM"
            N_Obfuscation.TrimLinesTabAndSpase vbc.CodeModule
        Case "OPT"
            N_Obfuscation.Remove_OptionExplicit vbc.CodeModule
        Case "EMPTY"
            N_Obfuscation.Remove_EmptyLines vbc.CodeModule
        Case "BREAK"
            N_Obfuscation.RemoveBreaksLineInCode vbc.CodeModule
        Case "CHAIN", "CHAIN2"
            RunChain vbc
        Case Else
            Err.Raise vbObjectError + 2700, "RunMethod", "Unknown method " & sMethod
    End Select
End Sub

Private Function ModuleText(ByVal cm As VBIDE.CodeModule) As String
    Dim arrLines() As String
    Dim n As Long
    Dim i As Long
    Dim s As String

    n = N_Obfuscation.TrfReadLines(cm, arrLines)
    For i = 1 To n
        If i > 1 Then s = s & vbCrLf
        s = s & arrLines(i)
    Next i
    ModuleText = s
End Function

'Readable cell text: every line prefixed with "| " (keeps leading apostrophes and "=" visible).
Private Function Show(ByVal sText As String) As String
    Show = "| " & Replace(sText, vbCrLf, vbLf & "| ")
End Function

Private Sub WriteResult(ByVal sId As String, ByVal sMethod As String, ByVal sPurpose As String, ByVal sStatus As String, _
                        ByVal sNotices As String, ByVal sNote As String, ByVal sRationale As String, _
                        ByVal sIn As String, ByVal sExp As String, ByVal sAct As String)
    m_lRow = m_lRow + 1
    m_ws.Range(m_ws.Cells(m_lRow, 1), m_ws.Cells(m_lRow, 10)).NumberFormat = "@"
    m_ws.Cells(m_lRow, 1).Value = sId
    m_ws.Cells(m_lRow, 2).Value = sMethod
    m_ws.Cells(m_lRow, 3).Value = sPurpose
    m_ws.Cells(m_lRow, 4).Value = sStatus
    m_ws.Cells(m_lRow, 5).Value = sNotices
    m_ws.Cells(m_lRow, 6).Value = sNote
    m_ws.Cells(m_lRow, 7).Value = sRationale
    m_ws.Cells(m_lRow, 8).Value = Left$(Show(sIn), 32000)
    m_ws.Cells(m_lRow, 9).Value = Left$(Show(sExp), 32000)
    m_ws.Cells(m_lRow, 10).Value = Left$(Show(sAct), 32000)
    Select Case sStatus
        Case "PASS"
            m_nPass = m_nPass + 1
            m_ws.Cells(m_lRow, 4).Interior.Color = RGB(198, 239, 206)
        Case "INCONCLUSIVE"
            m_nInconclusive = m_nInconclusive + 1
            m_ws.Cells(m_lRow, 4).Interior.Color = RGB(255, 235, 156)
        Case Else
            m_nFail = m_nFail + 1
            m_ws.Cells(m_lRow, 4).Interior.Color = RGB(255, 199, 206)
    End Select
End Sub

Private Sub RunCase(ByVal sId As String, ByVal sMethod As String, ByVal sPurpose As String, _
                    ByVal sIn As String, ByVal nIn As Long, ByVal sExp As String, ByVal nExp As Long, _
                    ByVal lNotices As Long, ByVal sRationale As String, ByVal bCompiles As Boolean)
    Dim vbc As VBIDE.VBComponent
    Dim sStored As String
    Dim sAct As String
    Dim nAct As Long
    Dim lNot As Long
    Dim sStatus As String
    Dim sNote As String
    Dim bWritten As Boolean

    On Error GoTo EH
    Set vbc = m_wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    vbc.Name = "T_" & sId
    FillModule vbc, sIn, nIn
    sStored = ModuleText(vbc.CodeModule)
    If sStored <> sIn Or vbc.CodeModule.CountOfLines <> nIn Then
        sStatus = "INCONCLUSIVE"
        sNote = "Editor changed the input when inserting it; adapt the test data."
        sAct = sStored
    Else
        N_Obfuscation.TrfClearNotices
        RunMethod vbc, sMethod
        lNot = N_Obfuscation.TrfNoticeCount
        sAct = ModuleText(vbc.CodeModule)
        nAct = vbc.CodeModule.CountOfLines
        If sMethod = "CHAIN2" Then
            RunChain vbc
            If ModuleText(vbc.CodeModule) <> sAct Then sNote = sNote & "Second run changed the code. "
        End If
        If nAct <> nExp Or sAct <> sExp Then sNote = sNote & "Output differs. "
        If lNot <> lNotices Then sNote = sNote & "Notice count differs. "
        If Len(sNote) = 0 Then sStatus = "PASS" Else sStatus = "FAIL"
    End If
    If Not bCompiles Then sNote = sNote & "Fixture is not compilable; module removed."
    WriteResult sId, sMethod, sPurpose, sStatus, lNot & "/" & lNotices, sNote, sRationale, sIn, sExp, sAct
    bWritten = True
    If Not (KEEP_COMPILABLE_MODULES And bCompiles) Then m_wb.VBProject.VBComponents.Remove vbc
    Exit Sub
EH:
    sNote = "Error " & Err.Number & ": " & Err.Description
    Resume AfterError
AfterError:
    On Error Resume Next
    If Not bWritten Then WriteResult sId, sMethod, sPurpose, "ERROR", lNot & "/" & lNotices, sNote, sRationale, sIn, sExp, sAct
    'an aborted fixture would disturb the optional compile check
    If Not vbc Is Nothing Then m_wb.VBProject.VBComponents.Remove vbc
End Sub

'T42: hidden attributes (module and procedure) survive the chain, including the merge of a
'continued procedure header. Uses export/import because Attribute lines are not visible
'in the code module.
Private Sub Case_T42_Attributes()
    Dim sPath As String
    Dim sOut As String
    Dim f As Integer
    Dim vbc As VBIDE.VBComponent
    Dim sText As String
    Dim sLine As String
    Dim sNote As String
    Dim sStatus As String
    Dim bWritten As Boolean

    On Error GoTo EH
    sPath = Environ$("TEMP") & "\MT_T42Class.cls"
    sOut = Environ$("TEMP") & "\MT_T42Class_out.cls"
    f = FreeFile
    Open sPath For Output As #f
    Print #f, "VERSION 1.0 CLASS"
    Print #f, "BEGIN"
    Print #f, "  MultiUse = -1  'True"
    Print #f, "END"
    Print #f, "Attribute VB_Name = ""T42Class"""
    Print #f, "Attribute VB_GlobalNameSpace = False"
    Print #f, "Attribute VB_Creatable = False"
    Print #f, "Attribute VB_PredeclaredId = True"
    Print #f, "Attribute VB_Exposed = False"
    Print #f, "Option Explicit"
    Print #f, "Private m_amount As Long"
    Print #f, "Public Property Get Total() As Long"
    Print #f, "    Total = 1 + _"
    Print #f, "        2 ' sum _"
    Print #f, "          continued comment"
    Print #f, "End Property"
    Print #f, "' default member"
    Print #f, "Public Property Get Amount( _"
    Print #f, "    ) As Long"
    Print #f, "Attribute Amount.VB_UserMemId = 0"
    Print #f, "    Amount = m_amount ' value"
    Print #f, "End Property"
    Close #f

    Set vbc = m_wb.VBProject.VBComponents.Import(sPath)
    N_Obfuscation.TrfClearNotices
    RunChain vbc
    vbc.Export sOut
    f = FreeFile
    Open sOut For Input As #f
    Do While Not EOF(f)
        Line Input #f, sLine
        sText = sText & sLine & vbCrLf
    Loop
    Close #f

    If InStr(1, sText, "Attribute Amount.VB_UserMemId = 0") = 0 Then sNote = sNote & "VB_UserMemId lost. "
    If InStr(1, sText, "Attribute VB_PredeclaredId = True") = 0 Then sNote = sNote & "VB_PredeclaredId lost. "
    If InStr(1, sText, "Public Property Get Amount() As Long") = 0 Then sNote = sNote & "Header not merged. "
    If InStr(1, sText, "Total = 1 + 2" & vbCrLf & "End Property") = 0 Then sNote = sNote & "Statement before End Property not merged. "
    If InStr(1, sText, "' default member") > 0 Then sNote = sNote & "Comment not removed. "
    If Len(sNote) = 0 Then sStatus = "PASS" Else sStatus = "FAIL"
    WriteResult "T42", "CHAIN", "Attribute-Zeilen in exportierten Modulen", sStatus, N_Obfuscation.TrfNoticeCount & "/0", sNote, _
                "Versteckte Attribute muessen erhalten bleiben, auch wenn die Prozedur davor mit einer fortgesetzten Anweisung endet.", _
                "(import of " & sPath & ")", "Attribute Amount.VB_UserMemId = 0 / VB_PredeclaredId = True / merged header", sText
    bWritten = True
    m_wb.VBProject.VBComponents.Remove vbc
    Kill sPath
    Kill sOut
    Exit Sub
EH:
    sNote = "Error " & Err.Number & ": " & Err.Description
    Resume AfterError
AfterError:
    On Error Resume Next
    If Not bWritten Then WriteResult "T42", "CHAIN", "Attribute-Zeilen in exportierten Modulen", "ERROR", "", sNote, "", "", "", sText
    Close
    If Not vbc Is Nothing Then m_wb.VBProject.VBComponents.Remove vbc
    If Len(Dir$(sPath)) > 0 Then Kill sPath
    If Len(Dir$(sOut)) > 0 Then Kill sOut
End Sub

'X26: behaviour of the real VBA editor when ReplaceLine receives more than 1023 characters.
'TrfApplyLineEdits must restore the original lines and write one notice.
Private Sub Case_X26_EditorLimit()
    Dim vbc As VBIDE.VBComponent
    Dim arrOld() As String
    Dim arrNew() As String
    Dim abKeep() As Boolean
    Dim sOrig As String
    Dim sAct As String
    Dim sStatus As String
    Dim sNote As String
    Dim bWritten As Boolean

    On Error GoTo EH
    Set vbc = m_wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    vbc.Name = "T_X26"
    sOrig = "Sub PX26()" & vbCrLf & "    x = 1 + _" & vbCrLf & "        2" & vbCrLf & "End Sub"
    FillModule vbc, sOrig, 4
    If ModuleText(vbc.CodeModule) <> sOrig Then
        WriteResult "X26", "TrfApplyLineEdits", "Verhalten des Editors bei mehr als 1023 Zeichen", "INCONCLUSIVE", "", _
                    "Editor changed the input when inserting it; adapt the test data.", "", sOrig, sOrig, ModuleText(vbc.CodeModule)
        m_wb.VBProject.VBComponents.Remove vbc
        Exit Sub
    End If
    ReDim arrOld(1 To 4)
    ReDim arrNew(1 To 4)
    ReDim abKeep(1 To 4)
    arrOld(1) = "Sub PX26()"
    arrOld(2) = "    x = 1 + _"
    arrOld(3) = "        2"
    arrOld(4) = "End Sub"
    arrNew(1) = arrOld(1)
    arrNew(2) = "    x = " & String$(1100, "1")
    arrNew(3) = arrOld(3)
    arrNew(4) = arrOld(4)
    abKeep(1) = True
    abKeep(2) = True
    abKeep(3) = False
    abKeep(4) = True
    N_Obfuscation.TrfClearNotices
    N_Obfuscation.TrfApplyLineEdits vbc.CodeModule, arrOld, arrNew, abKeep, 4, "X26"
    sAct = ModuleText(vbc.CodeModule)
    If sAct = sOrig And N_Obfuscation.TrfNoticeCount = 1 Then sStatus = "PASS" Else sStatus = "FAIL"
    WriteResult "X26", "TrfApplyLineEdits", "Verhalten des Editors bei mehr als 1023 Zeichen", sStatus, _
                N_Obfuscation.TrfNoticeCount & "/1", "", _
                "Belegt das tatsaechliche Editorverhalten (Teilung, Fehler oder Kuerzung) und die Wiederherstellung.", _
                sOrig, sOrig, sAct
    bWritten = True
    m_wb.VBProject.VBComponents.Remove vbc
    Exit Sub
EH:
    sNote = "Error " & Err.Number & ": " & Err.Description & " (editor raised an error instead of splitting)"
    Resume AfterError
AfterError:
    On Error Resume Next
    If Not bWritten Then WriteResult "X26", "TrfApplyLineEdits", "Verhalten des Editors bei mehr als 1023 Zeichen", "ERROR", "", _
                                     sNote, "", sOrig, sOrig, ""
    If Not vbc Is Nothing Then m_wb.VBProject.VBComponents.Remove vbc
End Sub

'Diagnostic: does VBA read "Name:" after a line number as a LABEL or as a CALL?
'The function is executed before and after RemoveLineNumbers. The text is used as the editor stores it,
'so the editor's own normalization does not matter. PASS = the number stays (one notice) and the
'result of the run does not change. The note names the reading: result 1 = call, result 0 = label.
Private Sub RunLabelOrCallCase(ByVal sId As String, ByVal sPurpose As String, ByVal sNumberLine As String)
    Dim vbc As VBIDE.VBComponent
    Dim sTag As String
    Dim sIn As String
    Dim sStored As String
    Dim sAct As String
    Dim sNote As String
    Dim sStatus As String
    Dim sRun As String
    Dim vBefore As Variant
    Dim vAfter As Variant
    Dim lNot As Long
    Dim bWritten As Boolean

    On Error GoTo EH
    sTag = Mid$(sId, 2)
    sIn = Replace(Replace(LC_TEMPLATE, "@LINE", sNumberLine), "@", sTag)
    sIn = Replace(sIn, "|", vbCrLf)
    Set vbc = m_wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    vbc.Name = "T_" & sId
    FillModule vbc, sIn, 8
    sStored = ModuleText(vbc.CodeModule)
    sRun = "'" & m_wb.Name & "'!T_" & sId & ".P" & sTag
    vBefore = Application.Run(sRun)
    N_Obfuscation.TrfClearNotices
    K_AddNumbersLine.RemoveLineNumbers vbc, vbLineNumbers_LabelTypes.vbLabelColon
    K_AddNumbersLine.RemoveLineNumbers vbc, vbLineNumbers_LabelTypes.vbLabelTab
    lNot = N_Obfuscation.TrfNoticeCount
    sAct = ModuleText(vbc.CodeModule)
    vAfter = Application.Run(sRun)
    If sStored <> sIn Then sNote = sNote & "Editor stored the line differently (see Input/Actual). "
    If vBefore = 1 Then
        sNote = sNote & "Reading: CALL (Mark was called). "
    ElseIf vBefore = 0 Then
        sNote = sNote & "Reading: LABEL (Mark was not called). "
    Else
        sNote = sNote & "Unexpected result " & vBefore & ". "
    End If
    If vBefore <> vAfter Then sNote = sNote & "Behaviour changed (before " & vBefore & ", after " & vAfter & "). "
    If sAct <> sStored Then sNote = sNote & "Output differs from the stored text. "
    If lNot <> 1 Then sNote = sNote & "Notice count differs. "
    If vBefore = vAfter And sAct = sStored And lNot = 1 Then sStatus = "PASS" Else sStatus = "FAIL"
    WriteResult sId, "RLN + Application.Run", sPurpose, sStatus, lNot & "/1", sNote & "Run before/after: " & vBefore & "/" & vAfter, _
                "Belegt die Lesart von 'Name:' nach einer Zeilennummer durch Ausfuehrung in Excel.", sIn, sStored, sAct
    bWritten = True
    m_wb.VBProject.VBComponents.Remove vbc
    Exit Sub
EH:
    sNote = "Error " & Err.Number & ": " & Err.Description
    Resume AfterError
AfterError:
    On Error Resume Next
    If Not bWritten Then WriteResult sId, "RLN + Application.Run", sPurpose, "ERROR", "", sNote, "", sIn, sStored, sAct
    If Not vbc Is Nothing Then m_wb.VBProject.VBComponents.Remove vbc
End Sub

Private Sub Case_X35g_NameAfterNumber()
    RunLabelOrCallCase "X35g", "Name nach Zeilennummer ohne Doppelpunkt: Aufruf oder Label?", "10 Mark35g:  P35g = cnt35g"
End Sub

Private Sub Case_X35h_NameAfterNumberColon()
    RunLabelOrCallCase "X35h", "Name nach Zeilennummer mit Doppelpunkt: Aufruf oder Label?", "10: Mark35h: P35h = cnt35h"
End Sub

'X37: diagnostic. Where does Excel 2019 put the empty line that it appends when inserted text contains
'a line continuation: directly after the inserted text or at the end of the module?
Private Sub Case_X37_InsertBlankPosition()
    Dim vbc As VBIDE.VBComponent
    Dim sBefore As String
    Dim sAfter As String
    Dim sNote As String
    Dim lBefore As Long
    Dim lAfter As Long
    Dim i As Long
    Dim lBlank As Long
    Dim bWritten As Boolean

    On Error GoTo EH
    Set vbc = m_wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    vbc.Name = "T_X37"
    With vbc.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .InsertLines 1, "Sub A()" & vbCrLf & "End Sub" & vbCrLf & "Sub B()" & vbCrLf & "End Sub"
        sBefore = ModuleText(vbc.CodeModule)
        lBefore = .CountOfLines
        .InsertLines 2, "    x = 1 + _" & vbCrLf & "        2"
        lAfter = .CountOfLines
    End With
    sAfter = ModuleText(vbc.CodeModule)
    For i = 1 To lAfter
        If Len(Trim$(vbc.CodeModule.Lines(i, 1))) = 0 Then lBlank = i
    Next i
    If lAfter = lBefore + 2 Then
        sNote = "No empty line appended."
    ElseIf lBlank = lAfter Then
        sNote = "Empty line appended at the END of the module (line " & lBlank & " of " & lAfter & ")."
    ElseIf lBlank > 0 Then
        sNote = "Empty line appended INSIDE the module at line " & lBlank & " of " & lAfter & " (inserted text: lines 2 to 3)."
    Else
        sNote = "Line count " & lBefore & " -> " & lAfter & ", no empty line found."
    End If
    WriteResult "X37", "InsertLines", "Position der angehaengten Leerzeile (Fortsetzungszeichen im eingefuegten Text)", "PASS", "", sNote, _
                "Diagnose: bestimmt, wie TrfReplaceLine die Leerzeile nach der Wiederherstellung finden muss.", sBefore, "(diagnostic)", sAfter
    bWritten = True
    m_wb.VBProject.VBComponents.Remove vbc
    Exit Sub
EH:
    sNote = "Error " & Err.Number & ": " & Err.Description
    Resume AfterError
AfterError:
    On Error Resume Next
    If Not bWritten Then WriteResult "X37", "InsertLines", "Position der angehaengten Leerzeile", "ERROR", "", sNote, "", sBefore, "", sAfter
    If Not vbc Is Nothing Then m_wb.VBProject.VBComponents.Remove vbc
End Sub

'X38: diagnostic. ReplaceLine with a line that ends with a continuation character: does the editor keep
'the line count and the text? (TrimLinesTabAndSpase and Remove_Comments write such lines.)
Private Sub Case_X38_ReplaceWithContinuation()
    Dim vbc As VBIDE.VBComponent
    Dim sBefore As String
    Dim sAfter As String
    Dim sNote As String
    Dim sStatus As String
    Dim lBefore As Long
    Dim lAfter As Long
    Dim bWritten As Boolean

    On Error GoTo EH
    Set vbc = m_wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    vbc.Name = "T_X38"
    FillModule vbc, "Sub A()" & vbCrLf & "      x = 1 + _" & vbCrLf & "        2" & vbCrLf & "End Sub", 4
    sBefore = ModuleText(vbc.CodeModule)
    lBefore = vbc.CodeModule.CountOfLines
    vbc.CodeModule.ReplaceLine 2, "x = 1 + _"
    lAfter = vbc.CodeModule.CountOfLines
    sAfter = ModuleText(vbc.CodeModule)
    If lAfter = lBefore And vbc.CodeModule.Lines(2, 1) = "x = 1 + _" And vbc.CodeModule.Lines(3, 1) = "        2" Then
        sStatus = "PASS"
        sNote = "ReplaceLine keeps line count and text."
    Else
        sStatus = "FAIL"
        sNote = "Line count " & lBefore & " -> " & lAfter & "; ReplaceLine with a continuation character changes the module."
    End If
    WriteResult "X38", "ReplaceLine", "ReplaceLine mit Fortsetzungszeichen am Zeilenende", sStatus, "", sNote, _
                "Diagnose: belegt, dass Trim und Kommentarentfernung fortgesetzte Zeilen ohne Nebenwirkung ersetzen.", sBefore, _
                "Sub A()" & vbCrLf & "x = 1 + _" & vbCrLf & "        2" & vbCrLf & "End Sub", sAfter
    bWritten = True
    m_wb.VBProject.VBComponents.Remove vbc
    Exit Sub
EH:
    sNote = "Error " & Err.Number & ": " & Err.Description
    Resume AfterError
AfterError:
    On Error Resume Next
    If Not bWritten Then WriteResult "X38", "ReplaceLine", "ReplaceLine mit Fortsetzungszeichen am Zeilenende", "ERROR", "", sNote, "", sBefore, "", sAfter
    If Not vbc Is Nothing Then m_wb.VBProject.VBComponents.Remove vbc
End Sub

'Writes a class file (header attributes plus the given body, lines separated by "|") for the import.
Private Sub WriteClassFile(ByVal sPath As String, ByVal sName As String, ByVal sBody As String)
    Dim f As Integer

    DeleteIfExists sPath
    f = FreeFile
    Open sPath For Output As #f
    Print #f, "VERSION 1.0 CLASS"
    Print #f, "BEGIN"
    Print #f, "  MultiUse = -1  'True"
    Print #f, "END"
    Print #f, "Attribute VB_Name = """ & sName & """"
    Print #f, "Attribute VB_GlobalNameSpace = False"
    Print #f, "Attribute VB_Creatable = False"
    Print #f, "Attribute VB_PredeclaredId = False"
    Print #f, "Attribute VB_Exposed = False"
    Print #f, Replace(sBody, "|", vbCrLf)
    Close #f
End Sub

'Diagnostic: imports a class with a hidden attribute, applies ONE method, exports the class again and
'checks whether the attribute is still there. Which method and which kind of edit loses it?
Private Sub RunAttributeCase(ByVal sId As String, ByVal sPurpose As String, ByVal sMethod As String, _
                             ByVal sBody As String, ByVal sAttr As String)
    Dim vbc As VBIDE.VBComponent
    Dim sName As String
    Dim sPath As String
    Dim sOut As String
    Dim sImported As String
    Dim sText As String
    Dim sNote As String
    Dim sStatus As String
    Dim lNot As Long
    Dim bWritten As Boolean

    On Error GoTo EH
    sName = "T" & sId & "Cls"
    sPath = Environ$("TEMP") & "\MT_" & sName & ".cls"
    sOut = Environ$("TEMP") & "\MT_" & sName & "_out.cls"
    WriteClassFile sPath, sName, sBody
    Set vbc = m_wb.VBProject.VBComponents.Import(sPath)
    DeleteIfExists sOut
    vbc.Export sOut
    sImported = ReadTextFile(sOut)
    If InStr(1, sImported, sAttr, vbBinaryCompare) = 0 Then
        sStatus = "INCONCLUSIVE"
        sNote = "Attribute is missing right after the import; nothing to protect."
        sText = sImported
    Else
        N_Obfuscation.TrfClearNotices
        RunMethod vbc, sMethod
        lNot = N_Obfuscation.TrfNoticeCount
        DeleteIfExists sOut
        vbc.Export sOut
        sText = ReadTextFile(sOut)
        If InStr(1, sText, sAttr, vbBinaryCompare) = 0 Then
            sStatus = "FAIL"
            sNote = "Attribute lost by " & sMethod & "."
        Else
            sStatus = "PASS"
        End If
    End If
    WriteResult sId, sMethod, sPurpose, sStatus, lNot & "/-", sNote, "Diagnose: welche Bearbeitung entfernt das versteckte Attribut?", _
                sImported, sAttr, sText
    bWritten = True
    m_wb.VBProject.VBComponents.Remove vbc
    DeleteIfExists sPath
    DeleteIfExists sOut
    Exit Sub
EH:
    sNote = "Error " & Err.Number & ": " & Err.Description
    Resume AfterError
AfterError:
    On Error Resume Next
    If Not bWritten Then WriteResult sId, sMethod, sPurpose, "ERROR", "", sNote, "", sImported, sAttr, sText
    Close
    If Not vbc Is Nothing Then m_wb.VBProject.VBComponents.Remove vbc
    DeleteIfExists sPath
    DeleteIfExists sOut
End Sub

Private Sub RunAttributeDiagnostics()
    Dim sAttr As String
    Dim sHdr As String
    Dim sEnd As String

    sAttr = "Attribute Amount.VB_UserMemId = 0"
    sHdr = "Public Property Get Amount() As Long|" & sAttr & "|"
    sEnd = "Amount = m_amount|End Property"
    RunAttributeCase "D0", "Kontrolle: Kette ohne jede Aenderung", "CHAIN", "Private m_amount As Long|" & sHdr & sEnd, sAttr
    RunAttributeCase "D1", "TRIM: nur die Kopfzeile (Einrueckung) wird geaendert", "TRIM", _
                     "Private m_amount As Long|    " & sHdr & sEnd, sAttr
    RunAttributeCase "D2", "CMT: nur die Kopfzeile (Kommentar am Ende) wird geaendert", "CMT", _
                     "Private m_amount As Long|Public Property Get Amount() As Long ' header|" & sAttr & "|" & sEnd, sAttr
    RunAttributeCase "D3", "BREAK: fortgesetzte Kopfzeile wird zusammengefuehrt", "BREAK", _
                     "Private m_amount As Long|Public Property Get Amount( _|    ) As Long|" & sAttr & "|" & sEnd, sAttr
    RunAttributeCase "D4", "CMT: nur eine Zeile im Rumpf wird geaendert", "CMT", _
                     "Private m_amount As Long|" & sHdr & "Amount = m_amount ' value|End Property", sAttr
    RunAttributeCase "D5", "TRIM: nur die Fortsetzungszeile der Kopfzeile wird geaendert", "TRIM", _
                     "Private m_amount As Long|Public Property Get Amount( _|    ) As Long|" & sAttr & "|" & sEnd, sAttr
    RunAttributeCase "D6", "EMPTY: Leerzeile ueber der Kopfzeile wird geloescht", "EMPTY", _
                     "Private m_amount As Long||" & sHdr & sEnd, sAttr
    RunAttributeCase "D7", "CMT: Kommentarzeile ueber der Kopfzeile wird geloescht", "CMT", _
                     "Private m_amount As Long|' default member|" & sHdr & sEnd, sAttr
    RunAttributeCase "D8", "TRIM: eingerueckte Variablendeklaration mit Attribut", "TRIM", _
                     "    Public Amount As Long|Attribute Amount.VB_VarDescription = ""Betrag""", "Attribute Amount.VB_VarDescription = ""Betrag"""
    RunAttributeCase "D9", "DBG: Debug.Print im Rumpf wird geloescht", "DBG", _
                     "Private m_amount As Long|" & sHdr & "Debug.Print ""x""|" & sEnd, sAttr
    RunAttributeCase "D10", "EMPTY: Leerzeile im Rumpf unter dem Attribut wird geloescht", "EMPTY", _
                     "Private m_amount As Long|" & sHdr & "Amount = m_amount||End Property", sAttr
End Sub

'===============================================================================
' Generated test cases (single source: cases.py)
'===============================================================================

Private Sub RunAllCases()
    Case_T01
    Case_T02
    Case_T02b
    Case_T03
    Case_T04
    Case_T05
    Case_T06
    Case_T07
    Case_T08
    Case_T09
    Case_T10
    Case_T11
    Case_T11b
    Case_T12
    Case_T13
    Case_T14
    Case_T15
    Case_T16
    Case_T17
    Case_T18
    Case_T19
    Case_T20
    Case_T21
    Case_T22
    Case_T23
    Case_T23b
    Case_T24
    Case_T25
    Case_T26
    Case_T27
    Case_T28
    Case_T29
    Case_T30
    Case_T31
    Case_T32
    Case_T33
    Case_T34
    Case_T35
    Case_T38
    Case_T39
    Case_T40
    Case_T41
    Case_T43
    Case_T44
    Case_X01
    Case_X02
    Case_X05
    Case_X06
    Case_X07
    Case_X09
    Case_X10
    Case_X11
    Case_X13
    Case_X15
    Case_X16
    Case_X17
    Case_X18
    Case_X19
    Case_X21
    Case_X22
    Case_X23
    Case_X28
    Case_X29
    Case_X30
    Case_X31
    Case_X32
    Case_X03
    Case_X34
    Case_X34b
    Case_X35
    Case_X36
    Case_X36b
    Case_X34c
    Case_X35c
    Case_X35d
    Case_X35e
    Case_X35f
End Sub

Private Sub Case_T01()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    RunCase "T01", "CHAIN", "Leeres CodeModule", sIn, nIn, sExp, nExp, 0, "Jede Methode muss bei 0 Zeilen sofort enden.", True
End Sub

Private Sub Case_T02()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Dim x As Long"
    AddLine sExp, nExp, "Dim x As Long"
    RunCase "T02", "CHAIN", "CodeModule mit nur einer Zeile", sIn, nIn, sExp, nExp, 0, "Einzeilige Module ohne Fortsetzung bleiben unveraendert.", True
End Sub

Private Sub Case_T02b()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Option Explicit"
    RunCase "T02b", "OPT", "Einzige Zeile ist Option Explicit", sIn, nIn, sExp, nExp, 0, "Loeschen der einzigen Zeile ergibt ein leeres Modul.", True
End Sub

Private Sub Case_T03()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, ""
    RunCase "T03", "EMPTY", "CodeModule mit nur einer Leerzeile", sIn, nIn, sExp, nExp, 0, "Leerzeile an Position 1 hat keine Vorgaengerzeile.", True
End Sub

Private Sub Case_T04()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P04()"
    AddLine sIn, nIn, "    x = 1"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "    y = 2"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P04()"
    AddLine sExp, nExp, "    x = 1"
    AddLine sExp, nExp, "    y = 2"
    AddLine sExp, nExp, "End Sub"
    RunCase "T04", "EMPTY", "Mehrere aufeinanderfolgende Leerzeilen", sIn, nIn, sExp, nExp, 0, "Zusammenhaengende Leerzeilen werden gemeinsam geloescht, keine Zeile wird uebersprungen.", True
End Sub

Private Sub Case_T05()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "Sub P05()"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P05()"
    AddLine sExp, nExp, "End Sub"
    RunCase "T05", "EMPTY", "Leerzeilen am Modulanfang", sIn, nIn, sExp, nExp, 0, "Zeile 1 und 2 leer.", True
End Sub

Private Sub Case_T06()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P06()"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, ""
    AddLine sExp, nExp, "Sub P06()"
    AddLine sExp, nExp, "End Sub"
    RunCase "T06", "EMPTY", "Leerzeilen am Modulende", sIn, nIn, sExp, nExp, 0, "Letzte Zeilen leer.", True
End Sub

Private Sub Case_T07()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P07()"
    AddLine sIn, nIn, "    Debug.Print ""Text ' kein Kommentar"""
    AddLine sIn, nIn, "    s = ""a'b"" ' echter Kommentar"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P07()"
    AddLine sExp, nExp, "    Debug.Print ""Text ' kein Kommentar"""
    AddLine sExp, nExp, "    s = ""a'b"""
    AddLine sExp, nExp, "End Sub"
    RunCase "T07", "CMT", "Kommentarzeichen innerhalb von Zeichenketten", sIn, nIn, sExp, nExp, 0, "Apostroph im String ist kein Kommentarbeginn.", True
End Sub

Private Sub Case_T08()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P08()"
    AddLine sIn, nIn, "    Debug.Print ""Er sagte """"Hallo"""""""
    AddLine sIn, nIn, "    s = ""x """"' y"""" z"" ' echt"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P08()"
    AddLine sExp, nExp, "    Debug.Print ""Er sagte """"Hallo"""""""
    AddLine sExp, nExp, "    s = ""x """"' y"""" z"""
    AddLine sExp, nExp, "End Sub"
    RunCase "T08", "CMT", "Doppelte Anfuehrungszeichen innerhalb von Zeichenketten", sIn, nIn, sExp, nExp, 0, "Escapte Anfuehrungszeichen beenden den String nicht.", True
End Sub

Private Sub Case_T09()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P09()"
    AddLine sIn, nIn, "Rem tatsaechlicher Kommentar"
    AddLine sIn, nIn, "    x = 1: Rem nach Trenner"
    AddLine sIn, nIn, "    s = ""Rem Test"""
    AddLine sIn, nIn, "    remark = 2"
    AddLine sIn, nIn, "    If x Then Rem nach Then"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P09()"
    AddLine sExp, nExp, "    x = 1"
    AddLine sExp, nExp, "    s = ""Rem Test"""
    AddLine sExp, nExp, "    remark = 2"
    AddLine sExp, nExp, "    If x Then Rem nach Then"
    AddLine sExp, nExp, "End Sub"
    RunCase "T09", "CMT", "Rem-Kommentare", sIn, nIn, sExp, nExp, 1, "Rem am Anweisungsanfang und nach ':' wird entfernt; 'remark' und String bleiben; Rem direkt nach Then bleibt (Notice), da die Wirkung auf die If-Form nicht belegt ist.", False
End Sub

Private Sub Case_T10()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P10()"
    AddLine sIn, nIn, "    ' nur Kommentar"
    AddLine sIn, nIn, "    x = 1 ' Kommentar"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P10()"
    AddLine sExp, nExp, "    x = 1"
    AddLine sExp, nExp, "End Sub"
    RunCase "T10", "CMT", "Kommentare nach ausfuehrbarem Code", sIn, nIn, sExp, nExp, 0, "Reine Kommentarzeile wird geloescht, Code vor dem Kommentar bleibt.", True
End Sub

Private Sub Case_T11()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P11()"
    AddLine sIn, nIn, "    GoTo Weiter"
    AddLine sIn, nIn, "Weiter:"
    AddLine sIn, nIn, "    x = 1"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P11()"
    AddLine sExp, nExp, "    GoTo Weiter"
    AddLine sExp, nExp, "Weiter:"
    AddLine sExp, nExp, "    x = 1"
    AddLine sExp, nExp, "End Sub"
    RunCase "T11", "RLN", "Echte alphanumerische Labels", sIn, nIn, sExp, nExp, 0, "Nur numerische Labels sind Ziel der Methode.", True
End Sub

Private Sub Case_T11b()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P11b()"
    AddLine sIn, nIn, "Weiter: Debug.Print 1"
    AddLine sIn, nIn, "    GoTo Weiter"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P11b()"
    AddLine sExp, nExp, "Weiter:"
    AddLine sExp, nExp, "    GoTo Weiter"
    AddLine sExp, nExp, "End Sub"
    RunCase "T11b", "DBG", "Debug.Print nach alphanumerischem Label", sIn, nIn, sExp, nExp, 0, "Das Label bleibt Sprungziel, nur die Anweisung faellt weg.", True
End Sub

Private Sub Case_T12()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P12()"
    AddLine sIn, nIn, "10:     x = 1"
    AddLine sIn, nIn, "20      y = 2"
    AddLine sIn, nIn, "30"
    AddLine sIn, nIn, "40: q12 = 3"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P12()"
    AddLine sExp, nExp, "     x = 1"
    AddLine sExp, nExp, "        y = 2"
    AddLine sExp, nExp, ""
    AddLine sExp, nExp, "q12 = 3"
    AddLine sExp, nExp, "End Sub"
    RunCase "T12", "RLN", "Numerische Zeilennummern (beide Formate)", sIn, nIn, sExp, nExp, 0, "Doppelpunkt-Format: '10:' weg, genau ein folgendes Leerzeichen weg; Format ohne Doppelpunkt: Ziffern werden durch Leerzeichen ersetzt; reine Nummernzeile wird leer.", True
End Sub

Private Sub Case_T13()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P13()"
    AddLine sIn, nIn, "10  x = 1"
    AddLine sIn, nIn, "    If x = 1 Then GoTo 20"
    AddLine sIn, nIn, "15  x = 2"
    AddLine sIn, nIn, "20  y = x"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P13()"
    AddLine sExp, nExp, "    x = 1"
    AddLine sExp, nExp, "    If x = 1 Then GoTo 20"
    AddLine sExp, nExp, "    x = 2"
    AddLine sExp, nExp, "20  y = x"
    AddLine sExp, nExp, "End Sub"
    RunCase "T13", "RLN", "Numerische Sprungziele", sIn, nIn, sExp, nExp, 0, "F2 a: referenzierte Nummer 20 bleibt, 10 und 15 werden entfernt.", True
End Sub

Private Sub Case_T14()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P14()"
    AddLine sIn, nIn, "    On i GoTo 100, 200"
    AddLine sIn, nIn, "    GoTo 300"
    AddLine sIn, nIn, "100 a = 1"
    AddLine sIn, nIn, "200 a = 2"
    AddLine sIn, nIn, "300 a = 3"
    AddLine sIn, nIn, "400 a = 4"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P14()"
    AddLine sExp, nExp, "    On i GoTo 100, 200"
    AddLine sExp, nExp, "    GoTo 300"
    AddLine sExp, nExp, "100 a = 1"
    AddLine sExp, nExp, "200 a = 2"
    AddLine sExp, nExp, "300 a = 3"
    AddLine sExp, nExp, "    a = 4"
    AddLine sExp, nExp, "End Sub"
    RunCase "T14", "RLN", "GoTo und On ... GoTo mit Liste", sIn, nIn, sExp, nExp, 0, "Alle Ziele der Liste und von GoTo bleiben.", True
End Sub

Private Sub Case_T15()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P15()"
    AddLine sIn, nIn, "    GoSub 500"
    AddLine sIn, nIn, "    Exit Sub"
    AddLine sIn, nIn, "500 x = 1"
    AddLine sIn, nIn, "    Return"
    AddLine sIn, nIn, "600 y = 2"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P15()"
    AddLine sExp, nExp, "    GoSub 500"
    AddLine sExp, nExp, "    Exit Sub"
    AddLine sExp, nExp, "500 x = 1"
    AddLine sExp, nExp, "    Return"
    AddLine sExp, nExp, "    y = 2"
    AddLine sExp, nExp, "End Sub"
    RunCase "T15", "RLN", "GoSub-Verwendung", sIn, nIn, sExp, nExp, 0, "GoSub-Ziel bleibt.", True
End Sub

Private Sub Case_T16()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P16()"
    AddLine sIn, nIn, "    On Error GoTo 70"
    AddLine sIn, nIn, "    x = 1 / 0"
    AddLine sIn, nIn, "60  y = 1"
    AddLine sIn, nIn, "    Exit Sub"
    AddLine sIn, nIn, "70  Resume 60"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P16()"
    AddLine sExp, nExp, "    On Error GoTo 70"
    AddLine sExp, nExp, "    x = 1 / 0"
    AddLine sExp, nExp, "60  y = 1"
    AddLine sExp, nExp, "    Exit Sub"
    AddLine sExp, nExp, "70  Resume 60"
    AddLine sExp, nExp, "End Sub"
    RunCase "T16", "RLN", "Resume-Verwendung", sIn, nIn, sExp, nExp, 0, "Resume- und On-Error-Ziele bleiben, die Fehlerbehandlung ist unveraendert.", True
End Sub

Private Sub Case_T17()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P17()"
    AddLine sIn, nIn, "    On Error GoTo 90"
    AddLine sIn, nIn, "10  x = 1"
    AddLine sIn, nIn, "    On Error GoTo 0"
    AddLine sIn, nIn, "90  y = 2"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P17()"
    AddLine sExp, nExp, "    On Error GoTo 90"
    AddLine sExp, nExp, "    x = 1"
    AddLine sExp, nExp, "    On Error GoTo 0"
    AddLine sExp, nExp, "90  y = 2"
    AddLine sExp, nExp, "End Sub"
    RunCase "T17", "RLN", "On Error GoTo-Verwendung", sIn, nIn, sExp, nExp, 0, "90 ist Ziel, 10 nicht.", True
End Sub

Private Sub Case_T18()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P18()"
    AddLine sIn, nIn, "    On Error GoTo EH"
    AddLine sIn, nIn, "10  x = 1 / 0"
    AddLine sIn, nIn, "    Exit Sub"
    AddLine sIn, nIn, "EH:"
    AddLine sIn, nIn, "    MsgBox ""Fehler in Zeile "" & Erl"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P18()"
    AddLine sExp, nExp, "    On Error GoTo EH"
    AddLine sExp, nExp, "10  x = 1 / 0"
    AddLine sExp, nExp, "    Exit Sub"
    AddLine sExp, nExp, "EH:"
    AddLine sExp, nExp, "    MsgBox ""Fehler in Zeile "" & Erl"
    AddLine sExp, nExp, "End Sub"
    RunCase "T18", "RLN", "Erl-Verwendung (Entscheidung F2 b)", sIn, nIn, sExp, nExp, 1, "F2 b: Prozedur wertet Erl aus, alle Zeilennummern bleiben; eine Notice.", True
End Sub

Private Sub Case_T19()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P19()"
    AddLine sIn, nIn, "    Debug.Print ""a"""
    AddLine sIn, nIn, "    x = 1"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P19()"
    AddLine sExp, nExp, "    x = 1"
    AddLine sExp, nExp, "End Sub"
    RunCase "T19", "DBG", "Einzeilige Debug.Print-Anweisung", sIn, nIn, sExp, nExp, 0, "Ganze Zeile entfaellt.", True
End Sub

Private Sub Case_T20()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P20()"
    AddLine sIn, nIn, "    Debug.Print ""a""; _"
    AddLine sIn, nIn, "        ""b"""
    AddLine sIn, nIn, "    x = 1"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P20()"
    AddLine sExp, nExp, "    x = 1"
    AddLine sExp, nExp, "End Sub"
    RunCase "T20", "DBG", "Mehrzeilige Debug.Print-Anweisung", sIn, nIn, sExp, nExp, 0, "Alle physischen Zeilen der Anweisung entfallen.", True
End Sub

Private Sub Case_T21()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P21()"
    AddLine sIn, nIn, "    s = ""Debug.Print"""
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P21()"
    AddLine sExp, nExp, "    s = ""Debug.Print"""
    AddLine sExp, nExp, "End Sub"
    RunCase "T21", "DBG", "Debug.Print innerhalb einer Zeichenkette", sIn, nIn, sExp, nExp, 0, "String-Inhalt ist kein Code.", True
End Sub

Private Sub Case_T22()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P22()"
    AddLine sIn, nIn, "    x = 1 ' Debug.Print x"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P22()"
    AddLine sExp, nExp, "    x = 1 ' Debug.Print x"
    AddLine sExp, nExp, "End Sub"
    RunCase "T22", "DBG", "Debug.Print innerhalb eines Kommentars", sIn, nIn, sExp, nExp, 0, "Kommentar ist kein Code; die Zeile bleibt vollstaendig.", True
End Sub

Private Sub Case_T23()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P23()"
    AddLine sIn, nIn, "    Call ImportantFunction: Debug.Print ""Test"": Call NextFunction"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub ImportantFunction()"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub NextFunction()"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P23()"
    AddLine sExp, nExp, "    Call ImportantFunction: Call NextFunction"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub ImportantFunction()"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub NextFunction()"
    AddLine sExp, nExp, "End Sub"
    RunCase "T23", "DBG", "Mehrere Anweisungen in einer Zeile", sIn, nIn, sExp, nExp, 0, "Beide Aufrufe bleiben vollstaendig und in derselben Reihenfolge.", True
End Sub

Private Sub Case_T23b()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P23b()"
    AddLine sIn, nIn, "    Call ImportantFunction: Debug.Print ""Test"": Call NextFunction"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub ImportantFunction()"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub NextFunction()"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P23b()"
    AddLine sExp, nExp, "    Call ImportantFunction: Call NextFunction"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub ImportantFunction()"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub NextFunction()"
    AddLine sExp, nExp, "End Sub"
    RunCase "T23b", "DBG_FALSE", "MatchAnywhereInLine:=False liefert dasselbe Ergebnis", sIn, nIn, sExp, nExp, 0, "Parameter ist nach der Korrektur wirkungslos.", True
End Sub

Private Sub Case_T24()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P24()"
    AddLine sIn, nIn, "    Debug.Print FunctionWithSideEffect()"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Function FunctionWithSideEffect()"
    AddLine sIn, nIn, "    n = n + 1"
    AddLine sIn, nIn, "End Function"
    AddLine sExp, nExp, "Sub P24()"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Function FunctionWithSideEffect()"
    AddLine sExp, nExp, "    n = n + 1"
    AddLine sExp, nExp, "End Function"
    RunCase "T24", "DBG", "Debug.Print mit Funktionsaufruf und Seiteneffekt (F3 i a)", sIn, nIn, sExp, nExp, 0, "F3 i a: Auswertung entfaellt mit; der Seiteneffekt geht bewusst verloren.", True
End Sub

Private Sub Case_T25()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P25()"
    AddLine sIn, nIn, "    x = 1 + _"
    AddLine sIn, nIn, "        2"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P25()"
    AddLine sExp, nExp, "    x = 1 + 2"
    AddLine sExp, nExp, "End Sub"
    RunCase "T25", "BREAK", "Einfache Zeilenfortsetzung", sIn, nIn, sExp, nExp, 0, "Ein Leerzeichen ersetzt die Fortsetzung, Token bleiben getrennt.", True
End Sub

Private Sub Case_T26()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P26()"
    AddLine sIn, nIn, "    s = ""a"" & _"
    AddLine sIn, nIn, "        ""b"" & _"
    AddLine sIn, nIn, "        ""c"""
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P26()"
    AddLine sExp, nExp, "    s = ""a"" & ""b"" & ""c"""
    AddLine sExp, nExp, "End Sub"
    RunCase "T26", "BREAK", "Mehrere aufeinanderfolgende Zeilenfortsetzungen", sIn, nIn, sExp, nExp, 0, "Alle Teile in Reihenfolge.", True
End Sub

Private Sub Case_T27()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P27()"
    AddLine sIn, nIn, "    x = Foo(1, _"
    AddLine sIn, nIn, "        2)"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Function Foo(a, b)"
    AddLine sIn, nIn, "    Foo = a + b"
    AddLine sIn, nIn, "End Function"
    AddLine sExp, nExp, "Sub P27()"
    AddLine sExp, nExp, "    x = Foo(1, 2)"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Function Foo(a, b)"
    AddLine sExp, nExp, "    Foo = a + b"
    AddLine sExp, nExp, "End Function"
    RunCase "T27", "BREAK", "Zeilenfortsetzung innerhalb eines Funktionsaufrufs", sIn, nIn, sExp, nExp, 0, "Argumentliste bleibt gueltig.", True
End Sub

Private Sub Case_T28()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P28()"
    AddLine sIn, nIn, "    Call Bar(a:=1, _"
    AddLine sIn, nIn, "        b:=2)"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub Bar(a, b)"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P28()"
    AddLine sExp, nExp, "    Call Bar(a:=1, b:=2)"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub Bar(a, b)"
    AddLine sExp, nExp, "End Sub"
    RunCase "T28", "BREAK", "Zeilenfortsetzung mit benannten Argumenten", sIn, nIn, sExp, nExp, 0, "':=' ist kein Anweisungstrenner.", True
End Sub

Private Sub Case_T29()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "#Const DEBUG_MODE = 1"
    AddLine sIn, nIn, "Sub P29()"
    AddLine sIn, nIn, "#If DEBUG_MODE = 1 Then ' nur Debug"
    AddLine sIn, nIn, "    Debug.Print ""dbg"""
    AddLine sIn, nIn, "    x = 1"
    AddLine sIn, nIn, "#Else"
    AddLine sIn, nIn, "    x = 2"
    AddLine sIn, nIn, "#End If"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "#Const DEBUG_MODE = 1"
    AddLine sExp, nExp, "Sub P29()"
    AddLine sExp, nExp, "#If DEBUG_MODE = 1 Then"
    AddLine sExp, nExp, "x = 1"
    AddLine sExp, nExp, "#Else"
    AddLine sExp, nExp, "x = 2"
    AddLine sExp, nExp, "#End If"
    AddLine sExp, nExp, "End Sub"
    RunCase "T29", "CHAIN", "Bedingte Kompilierung", sIn, nIn, sExp, nExp, 0, "Direktiven bleiben, Kommentar hinter #If wird entfernt, Debug.Print im Block entfaellt.", True
End Sub

Private Sub Case_T30()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P30()"
    AddLine sIn, nIn, "#If Win64 Then"
    AddLine sIn, nIn, "#If VBA7 Then"
    AddLine sIn, nIn, "    x = 1 ' a"
    AddLine sIn, nIn, "#Else"
    AddLine sIn, nIn, "    x = 2"
    AddLine sIn, nIn, "#End If"
    AddLine sIn, nIn, "#Else"
    AddLine sIn, nIn, "    x = 3"
    AddLine sIn, nIn, "#End If"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P30()"
    AddLine sExp, nExp, "#If Win64 Then"
    AddLine sExp, nExp, "#If VBA7 Then"
    AddLine sExp, nExp, "x = 1"
    AddLine sExp, nExp, "#Else"
    AddLine sExp, nExp, "x = 2"
    AddLine sExp, nExp, "#End If"
    AddLine sExp, nExp, "#Else"
    AddLine sExp, nExp, "x = 3"
    AddLine sExp, nExp, "#End If"
    AddLine sExp, nExp, "End Sub"
    RunCase "T30", "CHAIN", "Verschachtelte #If-Bloecke", sIn, nIn, sExp, nExp, 0, "Struktur der Direktiven unveraendert.", True
End Sub

Private Sub Case_T31()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Option Explicit"
    AddLine sIn, nIn, "Option Compare Text"
    AddLine sIn, nIn, "Option Base 1"
    AddLine sIn, nIn, "Option Private Module"
    AddLine sIn, nIn, "Dim x As Long"
    AddLine sExp, nExp, "Option Compare Text"
    AddLine sExp, nExp, "Option Base 1"
    AddLine sExp, nExp, "Option Private Module"
    AddLine sExp, nExp, "Dim x As Long"
    RunCase "T31", "OPT", "Option Explicit im Deklarationsbereich, andere Option-Anweisungen bleiben", sIn, nIn, sExp, nExp, 0, "Nur Option Explicit wird entfernt.", True
End Sub

Private Sub Case_T32()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "    Option Explicit"
    AddLine sIn, nIn, "Dim x As Long"
    AddLine sExp, nExp, "Dim x As Long"
    RunCase "T32", "OPT", "Option Explicit mit fuehrenden Leerzeichen", sIn, nIn, sExp, nExp, 0, "Einrueckung aendert die Anweisung nicht. Der VBE kann die Einrueckung beim Einfuegen entfernen.", True
End Sub

Private Sub Case_T33()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Option Explicit ' Pflicht"
    AddLine sIn, nIn, "Dim x As Long"
    AddLine sExp, nExp, "' Pflicht"
    AddLine sExp, nExp, "Dim x As Long"
    RunCase "T33", "OPT", "Option Explicit mit nachfolgendem Kommentar", sIn, nIn, sExp, nExp, 0, "Kommentar wird nicht veraendert, nur die Anweisung entfaellt.", True
End Sub

Private Sub Case_T34()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "' Option Explicit ist aus"
    AddLine sIn, nIn, "Dim x As Long"
    AddLine sExp, nExp, "' Option Explicit ist aus"
    AddLine sExp, nExp, "Dim x As Long"
    RunCase "T34", "OPT", "'Option Explicit' innerhalb eines Kommentars", sIn, nIn, sExp, nExp, 0, "Kommentar bleibt.", True
End Sub

Private Sub Case_T35()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P35()"
    AddLine sIn, nIn, "    s = ""Option Explicit"""
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P35()"
    AddLine sExp, nExp, "    s = ""Option Explicit"""
    AddLine sExp, nExp, "End Sub"
    RunCase "T35", "OPT", "'Option Explicit' innerhalb einer Zeichenkette", sIn, nIn, sExp, nExp, 0, "Frueher wurde die ganze Zeile geloescht.", True
End Sub

Private Sub Case_T38()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P38()"
    AddLine sIn, nIn, "    s = ""  a  b  """
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P38()"
    AddLine sExp, nExp, "s = ""  a  b  """
    AddLine sExp, nExp, "End Sub"
    RunCase "T38", "TRIM", "Leerzeichen innerhalb einer Zeichenkette", sIn, nIn, sExp, nExp, 0, "Innerer Text unveraendert.", True
End Sub

Private Sub Case_T39()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Dim x As Long ' erste"
    AddLine sIn, nIn, "Sub P39()"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Dim x As Long"
    AddLine sExp, nExp, "Sub P39()"
    AddLine sExp, nExp, "End Sub"
    RunCase "T39", "CHAIN", "Code direkt an der ersten Modulzeile", sIn, nIn, sExp, nExp, 0, "Zeile 1 wird wie jede andere behandelt.", True
End Sub

Private Sub Case_T40()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P40()"
    AddLine sIn, nIn, "    x = 1"
    AddLine sIn, nIn, "End Sub ' Ende"
    AddLine sExp, nExp, "Sub P40()"
    AddLine sExp, nExp, "x = 1"
    AddLine sExp, nExp, "End Sub"
    RunCase "T40", "CHAIN", "Code direkt an der letzten Modulzeile", sIn, nIn, sExp, nExp, 0, "Letzte Zeile wird erreicht.", True
End Sub

Private Sub Case_T41()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P41a()"
    AddLine sIn, nIn, "10  x = 1"
    AddLine sIn, nIn, "    GoTo 10"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub P41b()"
    AddLine sIn, nIn, "10  y = 1"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P41a()"
    AddLine sExp, nExp, "10  x = 1"
    AddLine sExp, nExp, "    GoTo 10"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub P41b()"
    AddLine sExp, nExp, "    y = 1"
    AddLine sExp, nExp, "End Sub"
    RunCase "T41", "RLN", "Zwei direkt aufeinanderfolgende Prozeduren", sIn, nIn, sExp, nExp, 0, "Zeilennummern gelten je Prozedur; die Referenz in P41a schuetzt nur ihre eigene 10.", True
End Sub

Private Sub Case_T43()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Option Explicit"
    AddLine sIn, nIn, "' Modul-Kommentar"
    AddLine sIn, nIn, "Sub P43()"
    AddLine sIn, nIn, "10: x = ""a'b"": Debug.Print ""c ' d"": y = 2 ' Kommentar"
    AddLine sIn, nIn, "    If x = """" Then GoTo 10"
    AddLine sIn, nIn, "20  w43 = Foo43(1, _"
    AddLine sIn, nIn, "        2) ' Ende"
    AddLine sIn, nIn, "    Debug.Print ""a""; _"
    AddLine sIn, nIn, "        ""b"""
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Function Foo43(a, b)"
    AddLine sIn, nIn, "    Foo43 = a + b"
    AddLine sIn, nIn, "End Function"
    AddLine sExp, nExp, "Sub P43()"
    AddLine sExp, nExp, "10: x = ""a'b"": y = 2"
    AddLine sExp, nExp, "If x = """" Then GoTo 10"
    AddLine sExp, nExp, "w43 = Foo43(1, 2)"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Function Foo43(a, b)"
    AddLine sExp, nExp, "Foo43 = a + b"
    AddLine sExp, nExp, "End Function"
    RunCase "T43", "CHAIN2", "Wiederholte Ausfuehrung der gesamten Kette", sIn, nIn, sExp, nExp, 0, "Zweiter Lauf darf nichts mehr aendern (Idempotenz).", True
End Sub

Private Sub Case_T44()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub P44()"
    AddLine sIn, nIn, "10: x = ""a'b"": Debug.Print ""c ' d"": y = 2 ' Kommentar"
    AddLine sIn, nIn, "    If x = """" Then GoTo 10"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub P44()"
    AddLine sExp, nExp, "10: x = ""a'b"": y = 2"
    AddLine sExp, nExp, "If x = """" Then GoTo 10"
    AddLine sExp, nExp, "End Sub"
    RunCase "T44", "CHAIN", "Kombination mehrerer Sonderfaelle in einer Zeile", sIn, nIn, sExp, nExp, 0, "Referenzierte Nummer, Apostroph im String, Debug.Print zwischen Anweisungen und echter Kommentar.", True
End Sub

Private Sub Case_X01()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX01()"
    AddLine sIn, nIn, "    s = """ & String$(1007, "A") & """ & _"
    AddLine sIn, nIn, "        ""B"""
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX01()"
    AddLine sExp, nExp, "    s = """ & String$(1007, "A") & """ & ""B"""
    AddLine sExp, nExp, "End Sub"
    RunCase "X01", "BREAK", "Grenzwert: zusammengefuehrt genau 1023 Zeichen", sIn, nIn, sExp, nExp, 0, "F6 a: 1023 Zeichen passen in eine physische Zeile.", True
End Sub

Private Sub Case_X02()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX02()"
    AddLine sIn, nIn, "    s = """ & String$(1008, "A") & """ & _"
    AddLine sIn, nIn, "        ""B"""
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX02()"
    AddLine sExp, nExp, "    s = """ & String$(1008, "A") & """ & _"
    AddLine sExp, nExp, "        ""B"""
    AddLine sExp, nExp, "End Sub"
    RunCase "X02", "BREAK", "Grenzwert: zusammengefuehrt 1024 Zeichen", sIn, nIn, sExp, nExp, 1, "F6 a: Anweisung bleibt mehrzeilig, Notice.", True
End Sub

Private Sub Case_X05()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX05()"
    AddLine sIn, nIn, "    If x Then Debug.Print x"
    AddLine sIn, nIn, "    If x Then y = 1 Else Debug.Print y"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX05()"
    AddLine sExp, nExp, "    If x Then Debug.Print x"
    AddLine sExp, nExp, "    If x Then y = 1 Else Debug.Print y"
    AddLine sExp, nExp, "End Sub"
    RunCase "X05", "DBG", "Debug.Print in einzeiligem If (F3 ii a)", sIn, nIn, sExp, nExp, 2, "F3 ii a: Zeile unveraendert, je Zeile eine Notice.", True
End Sub

Private Sub Case_X06()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX06()"
    AddLine sIn, nIn, "    x = 1: Debug.Print ""a""; _"
    AddLine sIn, nIn, "        ""b"""
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX06()"
    AddLine sExp, nExp, "    x = 1: Debug.Print ""a""; _"
    AddLine sExp, nExp, "        ""b"""
    AddLine sExp, nExp, "End Sub"
    RunCase "X06", "DBG", "Mehrzeilige Zeile mit Debug.Print und weiterer Anweisung", sIn, nIn, sExp, nExp, 1, "Konservativ: unveraendert und gemeldet (siehe Idempotenz-Hinweis).", True
End Sub

Private Sub Case_X07()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX07()"
    AddLine sIn, nIn, "    x = 1 + _"
    AddLine sIn, nIn, "        2 ' Kommentar"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX07()"
    AddLine sExp, nExp, "    x = 1 + _"
    AddLine sExp, nExp, "        2"
    AddLine sExp, nExp, "End Sub"
    RunCase "X07", "CMT", "Kommentar auf der letzten Zeile einer fortgesetzten Anweisung", sIn, nIn, sExp, nExp, 0, "Frueher uebersprungen; Ursache der Nicht-Idempotenz.", True
End Sub

Private Sub Case_X09()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "' Kommentar _"
    AddLine sIn, nIn, "  Fortsetzung"
    AddLine sIn, nIn, "Sub PX09()"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX09()"
    AddLine sExp, nExp, "End Sub"
    RunCase "X09", "CMT", "Mehrzeiliger Kommentar mit Fortsetzung", sIn, nIn, sExp, nExp, 0, "Fortsetzungszeilen eines Kommentars gehoeren zum Kommentar.", True
End Sub

Private Sub Case_X10()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "' Kommentar _"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "Sub PX10()"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "' Kommentar _"
    AddLine sExp, nExp, ""
    AddLine sExp, nExp, "Sub PX10()"
    AddLine sExp, nExp, "End Sub"
    RunCase "X10", "EMPTY", "Leerzeile nach Kommentar-Fortsetzung bleibt", sIn, nIn, sExp, nExp, 0, "Loeschen wuerde 'Sub PX10()' in den Kommentar ziehen.", True
End Sub

Private Sub Case_X11()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX11()"
    AddLine sIn, nIn, "    t = #10:30:00 AM#: Debug.Print t"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX11()"
    AddLine sExp, nExp, "    t = #10:30:00 AM#"
    AddLine sExp, nExp, "End Sub"
    RunCase "X11", "DBG", "Datumsliteral mit Doppelpunkten", sIn, nIn, sExp, nExp, 0, "':' im Datumsliteral ist kein Trenner.", True
End Sub

Private Sub Case_X13()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX13()"
    AddLine sIn, nIn, "    Dim obj As Object"
    AddLine sIn, nIn, "10  Set obj = Nothing"
    AddLine sIn, nIn, "    obj.GoTo 10"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX13()"
    AddLine sExp, nExp, "    Dim obj As Object"
    AddLine sExp, nExp, "    Set obj = Nothing"
    AddLine sExp, nExp, "    obj.GoTo 10"
    AddLine sExp, nExp, "End Sub"
    RunCase "X13", "RLN", "obj.GoTo ist keine Sprungreferenz", sIn, nIn, sExp, nExp, 0, "Memberaufruf nach '.' referenziert kein Label.", True
End Sub

Private Sub Case_X15()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Option _"
    AddLine sIn, nIn, "    Explicit"
    AddLine sIn, nIn, "Dim x As Long"
    AddLine sExp, nExp, "Dim x As Long"
    RunCase "X15", "OPT", "Option Explicit ueber Zeilenfortsetzung", sIn, nIn, sExp, nExp, 0, "Logische Zeile wird ganz entfernt.", True
End Sub

Private Sub Case_X16()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX16()"
    AddLine sIn, nIn, "10  Debug.Print ""x"""
    AddLine sIn, nIn, "    GoTo 10"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX16()"
    AddLine sExp, nExp, "10"
    AddLine sExp, nExp, "    GoTo 10"
    AddLine sExp, nExp, "End Sub"
    RunCase "X16", "DBG", "Debug.Print nach referenzierter Zeilennummer", sIn, nIn, sExp, nExp, 0, "Label bleibt als Sprungziel erhalten.", True
End Sub

Private Sub Case_X17()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX17()"
    AddLine sIn, nIn, "    Debug.Print x ' Hinweis"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX17()"
    AddLine sExp, nExp, "    ' Hinweis"
    AddLine sExp, nExp, "End Sub"
    RunCase "X17", "DBG", "Debug.Print mit Kommentar: Kommentar bleibt", sIn, nIn, sExp, nExp, 0, "Remove_DebugPrint greift nicht in Kommentare ein.", True
End Sub

Private Sub Case_X18()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX18()"
    AddLine sIn, nIn, "    Open ""x.txt"" For Output As #1: Print #1, ""a"": Debug.Print ""b"": Close #1"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX18()"
    AddLine sExp, nExp, "    Open ""x.txt"" For Output As #1: Print #1, ""a"": Close #1"
    AddLine sExp, nExp, "End Sub"
    RunCase "X18", "DBG", "Dateinummer # ist kein Datumsliteral", sIn, nIn, sExp, nExp, 0, "Trenner zwischen #1 ... #1 werden erkannt.", True
End Sub

Private Sub Case_X19()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX19()"
    AddLine sIn, nIn, "          Dim x As Long"
    AddLine sIn, nIn, "1         x = 1"
    AddLine sIn, nIn, "2         x = 2"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX19()"
    AddLine sExp, nExp, "          Dim x As Long"
    AddLine sExp, nExp, "          x = 1"
    AddLine sExp, nExp, "          x = 2"
    AddLine sExp, nExp, "End Sub"
    RunCase "X19", "RLN", "Zeilennummern im MZ-Tools-Format", sIn, nIn, sExp, nExp, 0, "Spalte des Codes bleibt erhalten.", True
End Sub

Private Sub Case_X21()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX21()"
    AddLine sIn, nIn, "Weiter: Rem c"
    AddLine sIn, nIn, "    GoTo Weiter"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX21()"
    AddLine sExp, nExp, "Weiter:"
    AddLine sExp, nExp, "    GoTo Weiter"
    AddLine sExp, nExp, "End Sub"
    RunCase "X21", "CMT", "Rem nach Label", sIn, nIn, sExp, nExp, 0, "Label bleibt.", True
End Sub

Private Sub Case_X22()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX22()"
    AddLine sIn, nIn, "    x = 1: ' c"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX22()"
    AddLine sExp, nExp, "    x = 1"
    AddLine sExp, nExp, "End Sub"
    RunCase "X22", "CMT", "Trenner vor Kommentar", sIn, nIn, sExp, nExp, 0, "Ueberfluessiger Trenner faellt mit.", True
End Sub

Private Sub Case_X23()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX23()"
    AddLine sIn, nIn, "    Select Case x"
    AddLine sIn, nIn, "    Case Else: Rem c"
    AddLine sIn, nIn, "    End Select"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX23()"
    AddLine sExp, nExp, "    Select Case x"
    AddLine sExp, nExp, "    Case Else"
    AddLine sExp, nExp, "    End Select"
    AddLine sExp, nExp, "End Sub"
    RunCase "X23", "CMT", "Case Else: Rem", sIn, nIn, sExp, nExp, 0, "Case Else ist kein If.", True
End Sub

Private Sub Case_X28()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX28()"
    AddLine sIn, nIn, "    If x Then"
    AddLine sIn, nIn, "    Else: Debug.Print x"
    AddLine sIn, nIn, "    End If"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX28()"
    AddLine sExp, nExp, "    If x Then"
    AddLine sExp, nExp, "    Else:"
    AddLine sExp, nExp, "    End If"
    AddLine sExp, nExp, "End Sub"
    RunCase "X28", "DBG", "Block-Else mit Debug.Print", sIn, nIn, sExp, nExp, 0, "Block-Else bleibt.", True
End Sub

Private Sub Case_X29()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX29()"
    AddLine sIn, nIn, "    x = 1 ' a _"
    AddLine sIn, nIn, "    b"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX29()"
    AddLine sExp, nExp, "    x = 1 ' a b"
    AddLine sExp, nExp, "End Sub"
    RunCase "X29", "BREAK", "Kommentar mit Fortsetzung wird eine Zeile", sIn, nIn, sExp, nExp, 0, "Kommentar bleibt Kommentar.", True
End Sub

Private Sub Case_X30()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX30()"
    AddLine sIn, nIn, "10: x = 1"
    AddLine sIn, nIn, "20  y = 2"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX30()"
    AddLine sExp, nExp, "10: x = 1"
    AddLine sExp, nExp, "    y = 2"
    AddLine sExp, nExp, "End Sub"
    RunCase "X30", "RLN_TAB", "Nur Format ohne Doppelpunkt", sIn, nIn, sExp, nExp, 0, "Doppelpunkt-Format bleibt bei vbLabelTab.", True
End Sub

Private Sub Case_X31()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX31()"
    AddLine sIn, nIn, "10: x = 1"
    AddLine sIn, nIn, "20  y = 2"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX31()"
    AddLine sExp, nExp, "x = 1"
    AddLine sExp, nExp, "20  y = 2"
    AddLine sExp, nExp, "End Sub"
    RunCase "X31", "RLN_COLON", "Nur Doppelpunkt-Format", sIn, nIn, sExp, nExp, 0, "Format ohne Doppelpunkt bleibt bei vbLabelColon.", True
End Sub

Private Sub Case_X32()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX32()"
    AddLine sIn, nIn, "    x = 1 + 2 _"
    AddLine sIn, nIn, "    ' c"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX32()"
    AddLine sExp, nExp, "    x = 1 + 2"
    AddLine sExp, nExp, "End Sub"
    RunCase "X32", "CMT", "Kommentarzeile innerhalb fortgesetzter Anweisung (Kommentar auf eigener Zeile)", sIn, nIn, sExp, nExp, 0, "Fortsetzung der Vorzeile wird entfernt, wenn ihre Folgezeile entfaellt.", False
End Sub

Private Sub Case_X03()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Private Sub LoadProfiles(ByVal table As ListObject, ByVal customerIds As Object)"
    AddLine sIn, nIn, "    Dim data As Variant"
    AddLine sIn, nIn, "    Dim rowIndex As Long"
    AddLine sIn, nIn, "    Dim profile As clsParserProfile"
    AddLine sIn, nIn, "    Dim ParserProfileId As String"
    AddLine sIn, nIn, "    Dim customerId As String"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "    If table.DataBodyRange Is Nothing Then"
    AddLine sIn, nIn, "        Err.Raise ERR_BASE + 30, TypeName(Me), _"
    AddLine sIn, nIn, "            TABLE_PARSER_PROFILES & "" contains no data rows."""
    AddLine sIn, nIn, "    End If"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "    data = table.DataBodyRange.Value2"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "    For rowIndex = 1 To UBound(data, 1)"
    AddLine sIn, nIn, "        ParserProfileId = ReadRequiredText( _"
    AddLine sIn, nIn, "            data(rowIndex, GetColumnIndex(table, ""ParserProfileId"")), _"
    AddLine sIn, nIn, "            TABLE_PARSER_PROFILES, rowIndex, ""ParserProfileId"")"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "        customerId = ReadRequiredText( _"
    AddLine sIn, nIn, "            data(rowIndex, GetColumnIndex(table, ""CustomerId"")), _"
    AddLine sIn, nIn, "            TABLE_PARSER_PROFILES, rowIndex, ""CustomerId"")"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "        If Not customerIds.Exists(NormalizeKey(customerId)) Then"
    AddLine sIn, nIn, "            Err.Raise ERR_BASE + 31, TypeName(Me), _"
    AddLine sIn, nIn, "                ""Parser profile references unknown CustomerId: "" & customerId"
    AddLine sIn, nIn, "        End If"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "        If m_profileIndex.Exists(NormalizeKey(ParserProfileId)) Then"
    AddLine sIn, nIn, "            Err.Raise ERR_BASE + 32, TypeName(Me), _"
    AddLine sIn, nIn, "                ""Duplicate ParserProfileId: "" & ParserProfileId"
    AddLine sIn, nIn, "        End If"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "        Set profile = New clsParserProfile"
    AddLine sIn, nIn, "        profile.Init _"
    AddLine sIn, nIn, "            ParserProfileId, _"
    AddLine sIn, nIn, "            customerId, _"
    AddLine sIn, nIn, "            ReadRequiredText(data(rowIndex, GetColumnIndex(table, ""ProfileName"")), _"
    AddLine sIn, nIn, "                TABLE_PARSER_PROFILES, rowIndex, ""ProfileName""), _"
    AddLine sIn, nIn, "            ReadRequiredText(data(rowIndex, GetColumnIndex(table, ""ProfileVersion"")), _"
    AddLine sIn, nIn, "                TABLE_PARSER_PROFILES, rowIndex, ""ProfileVersion""), _"
    AddLine sIn, nIn, "            ReadBoolean(data(rowIndex, GetColumnIndex(table, ""IsDefault"")), _"
    AddLine sIn, nIn, "                TABLE_PARSER_PROFILES, rowIndex, ""IsDefault""), _"
    AddLine sIn, nIn, "            ReadBoolean(data(rowIndex, GetColumnIndex(table, ""IsActive"")), _"
    AddLine sIn, nIn, "                TABLE_PARSER_PROFILES, rowIndex, ""IsActive""), _"
    AddLine sIn, nIn, "            ReadRequiredText(data(rowIndex, GetColumnIndex(table, ""RegexDialect"")), _"
    AddLine sIn, nIn, "                TABLE_PARSER_PROFILES, rowIndex, ""RegexDialect""), _"
    AddLine sIn, nIn, "            ReadDouble(data(rowIndex, GetColumnIndex(table, ""AcceptThreshold"")), _"
    AddLine sIn, nIn, "                TABLE_PARSER_PROFILES, rowIndex, ""AcceptThreshold""), _"
    AddLine sIn, nIn, "            ReadDouble(data(rowIndex, GetColumnIndex(table, ""ReviewThreshold"")), _"
    AddLine sIn, nIn, "                TABLE_PARSER_PROFILES, rowIndex, ""ReviewThreshold""), _"
    AddLine sIn, nIn, "            ReadOptionalDate(data(rowIndex, GetColumnIndex(table, ""ValidFrom"")), _"
    AddLine sIn, nIn, "                TABLE_PARSER_PROFILES, rowIndex, ""ValidFrom""), _"
    AddLine sIn, nIn, "            ReadOptionalDate(data(rowIndex, GetColumnIndex(table, ""ValidTo"")), _"
    AddLine sIn, nIn, "                TABLE_PARSER_PROFILES, rowIndex, ""ValidTo""), _"
    AddLine sIn, nIn, "            ReadOptionalText(data(rowIndex, GetColumnIndex(table, ""Notes"")))"
    AddLine sIn, nIn, ""
    AddLine sIn, nIn, "        m_profiles.Add profile"
    AddLine sIn, nIn, "        m_profileIndex.Add NormalizeKey(ParserProfileId), profile"
    AddLine sIn, nIn, "        Set profile = Nothing"
    AddLine sIn, nIn, "    Next rowIndex"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Private Sub LoadProfiles(ByVal table As ListObject, ByVal customerIds As Object)"
    AddLine sExp, nExp, "Dim data As Variant"
    AddLine sExp, nExp, "Dim rowIndex As Long"
    AddLine sExp, nExp, "Dim profile As clsParserProfile"
    AddLine sExp, nExp, "Dim ParserProfileId As String"
    AddLine sExp, nExp, "Dim customerId As String"
    AddLine sExp, nExp, "If table.DataBodyRange Is Nothing Then"
    AddLine sExp, nExp, "Err.Raise ERR_BASE + 30, TypeName(Me), TABLE_PARSER_PROFILES & "" contains no data rows."""
    AddLine sExp, nExp, "End If"
    AddLine sExp, nExp, "data = table.DataBodyRange.Value2"
    AddLine sExp, nExp, "For rowIndex = 1 To UBound(data, 1)"
    AddLine sExp, nExp, "ParserProfileId = ReadRequiredText(data(rowIndex, GetColumnIndex(table, ""ParserProfileId"")), TABLE_PARSER_PROFILES, rowIndex, ""ParserProfileId"")"
    AddLine sExp, nExp, "customerId = ReadRequiredText(data(rowIndex, GetColumnIndex(table, ""CustomerId"")), TABLE_PARSER_PROFILES, rowIndex, ""CustomerId"")"
    AddLine sExp, nExp, "If Not customerIds.Exists(NormalizeKey(customerId)) Then"
    AddLine sExp, nExp, "Err.Raise ERR_BASE + 31, TypeName(Me), ""Parser profile references unknown CustomerId: "" & customerId"
    AddLine sExp, nExp, "End If"
    AddLine sExp, nExp, "If m_profileIndex.Exists(NormalizeKey(ParserProfileId)) Then"
    AddLine sExp, nExp, "Err.Raise ERR_BASE + 32, TypeName(Me), ""Duplicate ParserProfileId: "" & ParserProfileId"
    AddLine sExp, nExp, "End If"
    AddLine sExp, nExp, "Set profile = New clsParserProfile"
    AddLine sExp, nExp, "profile.Init _"
    AddLine sExp, nExp, "ParserProfileId, _"
    AddLine sExp, nExp, "customerId, _"
    AddLine sExp, nExp, "ReadRequiredText(data(rowIndex, GetColumnIndex(table, ""ProfileName"")), _"
    AddLine sExp, nExp, "TABLE_PARSER_PROFILES, rowIndex, ""ProfileName""), _"
    AddLine sExp, nExp, "ReadRequiredText(data(rowIndex, GetColumnIndex(table, ""ProfileVersion"")), _"
    AddLine sExp, nExp, "TABLE_PARSER_PROFILES, rowIndex, ""ProfileVersion""), _"
    AddLine sExp, nExp, "ReadBoolean(data(rowIndex, GetColumnIndex(table, ""IsDefault"")), _"
    AddLine sExp, nExp, "TABLE_PARSER_PROFILES, rowIndex, ""IsDefault""), _"
    AddLine sExp, nExp, "ReadBoolean(data(rowIndex, GetColumnIndex(table, ""IsActive"")), _"
    AddLine sExp, nExp, "TABLE_PARSER_PROFILES, rowIndex, ""IsActive""), _"
    AddLine sExp, nExp, "ReadRequiredText(data(rowIndex, GetColumnIndex(table, ""RegexDialect"")), _"
    AddLine sExp, nExp, "TABLE_PARSER_PROFILES, rowIndex, ""RegexDialect""), _"
    AddLine sExp, nExp, "ReadDouble(data(rowIndex, GetColumnIndex(table, ""AcceptThreshold"")), _"
    AddLine sExp, nExp, "TABLE_PARSER_PROFILES, rowIndex, ""AcceptThreshold""), _"
    AddLine sExp, nExp, "ReadDouble(data(rowIndex, GetColumnIndex(table, ""ReviewThreshold"")), _"
    AddLine sExp, nExp, "TABLE_PARSER_PROFILES, rowIndex, ""ReviewThreshold""), _"
    AddLine sExp, nExp, "ReadOptionalDate(data(rowIndex, GetColumnIndex(table, ""ValidFrom"")), _"
    AddLine sExp, nExp, "TABLE_PARSER_PROFILES, rowIndex, ""ValidFrom""), _"
    AddLine sExp, nExp, "ReadOptionalDate(data(rowIndex, GetColumnIndex(table, ""ValidTo"")), _"
    AddLine sExp, nExp, "TABLE_PARSER_PROFILES, rowIndex, ""ValidTo""), _"
    AddLine sExp, nExp, "ReadOptionalText(data(rowIndex, GetColumnIndex(table, ""Notes"")))"
    AddLine sExp, nExp, "m_profiles.Add profile"
    AddLine sExp, nExp, "m_profileIndex.Add NormalizeKey(ParserProfileId), profile"
    AddLine sExp, nExp, "Set profile = Nothing"
    AddLine sExp, nExp, "Next rowIndex"
    AddLine sExp, nExp, "End Sub"
    RunCase "X03", "CHAIN", "Reales Beispiel LoadProfiles (Fehlerbild aus Rueckfrage)", sIn, nIn, sExp, nExp, 1, "profile.Init haette 1166 Zeichen: bleibt mehrzeilig (F6 a); alle anderen Fortsetzungen werden zusammengefuehrt.", False
End Sub

Private Sub Case_X34()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX34()"
    AddLine sIn, nIn, "    Debug.Print ""start"": Init34: Run34"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub Init34()"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub Run34()"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX34()"
    AddLine sExp, nExp, "    Call Init34: Run34"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub Init34()"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub Run34()"
    AddLine sExp, nExp, "End Sub"
    RunCase "X34", "DBG", "Call verhindert, dass ein Aufruf zum Label wird", sIn, nIn, sExp, nExp, 1, "'Init34: Run34' am Zeilenanfang waere eine Labeldefinition; Call ist freigegeben.", True
End Sub

Private Sub Case_X34b()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX34b()"
    AddLine sIn, nIn, "    x = 1: Debug.Print ""a"": Init34b: Run34b"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub Init34b()"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub Run34b()"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX34b()"
    AddLine sExp, nExp, "    x = 1: Init34b: Run34b"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub Init34b()"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub Run34b()"
    AddLine sExp, nExp, "End Sub"
    RunCase "X34b", "DBG", "Aufruf nach Trenner bleibt Aufruf", sIn, nIn, sExp, nExp, 0, "Vor Init34b bleibt eine Anweisung, daher kein Label.", True
End Sub

Private Sub Case_X35()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX35()"
    AddLine sIn, nIn, "10 Init35:  Run35"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub Init35()"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub Run35()"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX35()"
    AddLine sExp, nExp, "10 Init35:  Run35"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub Init35()"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub Run35()"
    AddLine sExp, nExp, "End Sub"
    RunCase "X35", "RLN", "Nummer vor 'Name:' bleibt (ohne Doppelpunkt-Format)", sIn, nIn, sExp, nExp, 1, "Ob 'Init35:' nach der Nummer ein zweites Label oder ein Aufruf ist, ist nicht belegt (Schreibweise des Editors wie bei einem Label): Nummer bleibt, Notice. Die Eingabe steht in der vom Excel-Editor normalisierten Form.", True
End Sub

Private Sub Case_X36()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX36()"
    AddLine sIn, nIn, "    x = Foo36(1, _"
    AddLine sIn, nIn, " _"
    AddLine sIn, nIn, "        2)"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Function Foo36(a, b)"
    AddLine sIn, nIn, "    Foo36 = a + b"
    AddLine sIn, nIn, "End Function"
    AddLine sExp, nExp, "Sub PX36()"
    AddLine sExp, nExp, "x = Foo36(1, _"
    AddLine sExp, nExp, " _"
    AddLine sExp, nExp, "2)"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Function Foo36(a, b)"
    AddLine sExp, nExp, "Foo36 = a + b"
    AddLine sExp, nExp, "End Function"
    RunCase "X36", "TRIM", "Zeile nur aus Fortsetzungszeichen", sIn, nIn, sExp, nExp, 0, "Leerzeichen vor '_' bleibt, sonst endet die Fortsetzung (Review-Fund 4).", False
End Sub

Private Sub Case_X36b()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX36b()"
    AddLine sIn, nIn, "    x = Foo36b(1, _"
    AddLine sIn, nIn, " _"
    AddLine sIn, nIn, "        2)"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Function Foo36b(a, b)"
    AddLine sIn, nIn, "    Foo36b = a + b"
    AddLine sIn, nIn, "End Function"
    AddLine sExp, nExp, "Sub PX36b()"
    AddLine sExp, nExp, "x = Foo36b(1, 2)"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Function Foo36b(a, b)"
    AddLine sExp, nExp, "Foo36b = a + b"
    AddLine sExp, nExp, "End Function"
    RunCase "X36b", "CHAIN2", "Kette mit Zeile nur aus Fortsetzungszeichen", sIn, nIn, sExp, nExp, 0, "Zusammenfuehrung vollstaendig und idempotent.", False
End Sub

Private Sub Case_X34c()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX34c()"
    AddLine sIn, nIn, "    Debug.Print ""a"": Stop: x = 1"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX34c()"
    AddLine sExp, nExp, "    Debug.Print ""a"": Stop: x = 1"
    AddLine sExp, nExp, "End Sub"
    RunCase "X34c", "DBG", "Schluesselwort vor ':' bleibt unveraendert", sIn, nIn, sExp, nExp, 1, "'Call Stop' waere ungueltig; Zeile bleibt, Notice.", True
End Sub

Private Sub Case_X35c()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX35c()"
    AddLine sIn, nIn, "10  Stop: x = 1"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX35c()"
    AddLine sExp, nExp, "10  Stop: x = 1"
    AddLine sExp, nExp, "End Sub"
    RunCase "X35c", "RLN", "Nummer vor Schluesselwort mit ':' bleibt", sIn, nIn, sExp, nExp, 1, "Keine Call-Form fuer Stop; Nummer bleibt, Notice.", True
End Sub

Private Sub Case_X35d()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX35d()"
    AddLine sIn, nIn, "10  On Error GoTo 50"
    AddLine sIn, nIn, "20  x = 1 / 0"
    AddLine sIn, nIn, "30  Exit Sub"
    AddLine sIn, nIn, "50  If Erl = 20 Then Resume 30 Else Resume Next"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX35d()"
    AddLine sExp, nExp, "10  On Error GoTo 50"
    AddLine sExp, nExp, "20  x = 1 / 0"
    AddLine sExp, nExp, "30  Exit Sub"
    AddLine sExp, nExp, "50  If Erl = 20 Then Resume 30 Else Resume Next"
    AddLine sExp, nExp, "End Sub"
    RunCase "X35d", "RLN", "F2 b: Erl steuert den Fehlerpfad", sIn, nIn, sExp, nExp, 1, "Mit F2 a wuerde Resume Next statt Resume 30 ausgefuehrt; F2 b laesst alle Nummern stehen.", True
End Sub

Private Sub Case_X35e()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX35e1()"
    AddLine sIn, nIn, "10  x = 1"
    AddLine sIn, nIn, "    Debug.Print Erl"
    AddLine sIn, nIn, "End Sub"
    AddLine sIn, nIn, "Sub PX35e2()"
    AddLine sIn, nIn, "10  y = 1"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX35e1()"
    AddLine sExp, nExp, "10  x = 1"
    AddLine sExp, nExp, "    Debug.Print Erl"
    AddLine sExp, nExp, "End Sub"
    AddLine sExp, nExp, "Sub PX35e2()"
    AddLine sExp, nExp, "    y = 1"
    AddLine sExp, nExp, "End Sub"
    RunCase "X35e", "RLN", "F2 b gilt nur fuer die Prozedur mit Erl", sIn, nIn, sExp, nExp, 1, "Zeilennummern und Erl sind prozedurlokal.", True
End Sub

Private Sub Case_X35f()
    Dim sIn As String, nIn As Long, sExp As String, nExp As Long
    AddLine sIn, nIn, "Sub PX35f()"
    AddLine sIn, nIn, "    Dim o As Object"
    AddLine sIn, nIn, "10  x = o.Erl"
    AddLine sIn, nIn, "End Sub"
    AddLine sExp, nExp, "Sub PX35f()"
    AddLine sExp, nExp, "    Dim o As Object"
    AddLine sExp, nExp, "    x = o.Erl"
    AddLine sExp, nExp, "End Sub"
    RunCase "X35f", "RLN", "obj.Erl ist kein Erl", sIn, nIn, sExp, nExp, 0, "Membername nach '.' zaehlt nicht als Erl.", True
End Sub

