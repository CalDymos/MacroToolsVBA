Attribute VB_Name = "RemoveUnusedTests"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : RemoveUnusedTests - regression tests of modRemoveUnusedCode
'* Created    : 15-10-2026
'* Usage      : 1. Import this module into the MACROTools add-in project (the project that
'*                 contains modRemoveUnusedCode and modObfuscation).
'*              2. Trust access to the VBA project object model must be enabled.
'*              3. Run RunRemoveUnusedTests: type RunRemoveUnusedTests in the Immediate
'*                 window, or place the cursor in the procedure and press F5 (the module is
'*                 Option Private Module, so Alt+F8 does not list it).
'*              4. A new workbook opens with the sheet "Results" (one row per test).
'*              5. Compile check: in the VBA editor select the project of that workbook and
'*                 run Debug > Compile. The fixtures stay there after the run and must
'*                 compile (the removal must not leave broken code).
'*              The test workbook is not saved; close it without saving.
'* Status     : PASS = remaining code and number of removed items as expected
'*              FAIL = deviation (expected and actual code are shown)
'*              INCONCLUSIVE = the VBA editor changed the test input when it was inserted
'*              ERROR = runtime error during the test
'* Cases      : generated from ruc_cases.py (single source); ruc_model.py checks the expected
'*              results against a Python port of the rules.
'* Fixture format: components separated by "~", each "Kind:Name:line|line|...".
'*              Kinds: M standard module, m standard module that is NOT a target, C class,
'*              F UserForm, A imported class (may contain Attribute lines), X imported class
'*              with VB_Exposed = True, S sheet module of the test workbook.
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

'Keep the fixtures of PASS and FAIL cases in the test workbook for Debug > Compile (names are unique per case)
Private Const KEEP_FIXTURES As Boolean = True

Private m_wb As Workbook
Private m_ws As Worksheet
Private m_lRow As Long
Private m_nPass As Long
Private m_nFail As Long
Private m_nInconclusive As Long

Public Sub RunRemoveUnusedTests()
    On Error GoTo EH
    Application.ScreenUpdating = False
    Set m_wb = Workbooks.Add(xlWBATWorksheet)
    Set m_ws = m_wb.Worksheets(1)
    m_ws.Name = "Results"
    m_ws.Range("A1:J1").Value = Array("ID", "Scope", "Purpose", "Status", "Removed (actual/expected)", _
                                      "Note", "Rationale", "Input", "Expected", "Actual")
    m_ws.Rows(1).Font.Bold = True
    m_lRow = 1
    m_nPass = 0
    m_nFail = 0
    m_nInconclusive = 0

    RunAllCases

    m_ws.Columns("A:G").AutoFit
    m_ws.Columns("H:J").ColumnWidth = 80
    m_ws.Range("H:J").WrapText = True
    Application.ScreenUpdating = True
    Debug.Print "Remove unused code tests: PASS " & m_nPass & ", FAIL " & m_nFail & ", INCONCLUSIVE " & m_nInconclusive
    MsgBox "PASS: " & m_nPass & vbLf & "FAIL: " & m_nFail & vbLf & "INCONCLUSIVE: " & m_nInconclusive & vbLf & vbLf & _
           "Details: sheet 'Results' of " & m_wb.Name & "." & vbLf & _
           "Then: Debug > Compile on the project of that workbook.", vbInformation, "Remove unused code tests"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Test run aborted: error " & Err.Number & ": " & Err.Description, vbCritical, "Remove unused code tests"
End Sub

'Appends one line to a component body (lines are separated by "|"; AddComp drops the first "|").
Private Sub AddLn(ByRef sBody As String, ByVal sLine As String)
    sBody = sBody & "|" & sLine
End Sub

'Appends a component to a fixture: "Kind:Name:body".
Private Sub AddComp(ByRef sSpec As String, ByVal sKind As String, ByVal sName As String, ByVal sBody As String)
    If Len(sSpec) > 0 Then sSpec = sSpec & "~"
    sSpec = sSpec & sKind & ":" & sName & ":" & Mid$(sBody, 2)
End Sub

Private Function ScopeFlags(ByVal sScope As String) As Long
    Select Case sScope
        Case "A": ScopeFlags = enumUnusedScopePrivateOnly
        Case "B": ScopeFlags = enumUnusedScopePublicClassMembers
        Case "C": ScopeFlags = enumUnusedScopePublicClassMembers Or enumUnusedScopePublicModuleProcs
        Case Else: Err.Raise vbObjectError + 2710, "ScopeFlags", "Unknown scope " & sScope
    End Select
End Function

'Splits a fixture into kinds, names and bodies (0-based arrays). Returns the number of components.
Private Function ParseSpec(ByVal sSpec As String, ByRef asKind() As String, ByRef asName() As String, _
                           ByRef asBody() As String) As Long
    Dim v As Variant
    Dim i As Long
    Dim p As Long
    Dim s As String

    v = Split(sSpec, "~")
    ReDim asKind(0 To UBound(v))
    ReDim asName(0 To UBound(v))
    ReDim asBody(0 To UBound(v))
    For i = 0 To UBound(v)
        s = v(i)
        p = InStr(3, s, ":")
        asKind(i) = Left$(s, 1)
        asName(i) = Mid$(s, 3, p - 3)
        asBody(i) = Mid$(s, p + 1)
    Next i
    ParseSpec = UBound(v) + 1
End Function

'Body without Attribute lines (they are hidden in the VBA editor).
Private Function VisibleBody(ByVal sBody As String) As String
    Dim v As Variant
    Dim i As Long
    Dim s As String
    Dim bFirst As Boolean

    If Len(sBody) = 0 Then Exit Function
    v = Split(sBody, "|")
    bFirst = True
    For i = 0 To UBound(v)
        If Left$(v(i), 10) <> "Attribute " Then
            If Not bFirst Then s = s & "|"
            s = s & v(i)
            bFirst = False
        End If
    Next i
    VisibleBody = s
End Function

Private Function VisibleSpec(ByVal sSpec As String) As String
    Dim asKind() As String
    Dim asName() As String
    Dim asBody() As String
    Dim n As Long
    Dim i As Long
    Dim s As String

    n = ParseSpec(sSpec, asKind, asName, asBody)
    For i = 0 To n - 1
        If i > 0 Then s = s & "~"
        s = s & asKind(i) & ":" & asName(i) & ":" & VisibleBody(asBody(i))
    Next i
    VisibleSpec = s
End Function

'Fills an empty code module with the given body. Excel 2019 appends an empty line when the inserted
'text contains a line continuation; that line is removed again (same as TransformChainTests).
Private Sub FillModule(ByVal vbc As VBIDE.VBComponent, ByVal sBody As String)
    Dim nLines As Long

    If Len(sBody) > 0 Then nLines = UBound(Split(sBody, "|")) + 1
    With vbc.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If nLines > 0 Then .InsertLines 1, Replace(sBody, "|", vbCrLf)
        If .CountOfLines = nLines + 1 Then
            If Len(Trim$(.Lines(.CountOfLines, 1))) = 0 Then .DeleteLines .CountOfLines, 1
        End If
    End With
End Sub

Private Sub DeleteIfExists(ByVal sPath As String)
    If Len(Dir$(sPath)) > 0 Then Kill sPath
End Sub

'Writes a class file for the import, followed by the body (may contain Attribute lines).
Private Sub WriteClassFile(ByVal sPath As String, ByVal sName As String, ByVal sBody As String, ByVal bExposed As Boolean)
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
    Print #f, "Attribute VB_Exposed = " & IIf(bExposed, "True", "False")
    Print #f, Replace(sBody, "|", vbCrLf)
    Close #f
End Sub

'Sheet module of the first worksheet of the test workbook.
Private Function SheetComponent() As VBIDE.VBComponent
    Dim vbc As VBIDE.VBComponent

    For Each vbc In m_wb.VBProject.VBComponents
        If vbc.Type = vbext_ct_Document Then
            If vbc.Properties("Name").Value = m_wb.Worksheets(1).Name Then
                Set SheetComponent = vbc
                Exit Function
            End If
        End If
    Next vbc
    Err.Raise vbObjectError + 2711, "SheetComponent", "Sheet module not found."
End Function

Private Function FindComponent(ByVal sName As String) As VBIDE.VBComponent
    On Error Resume Next
    Set FindComponent = m_wb.VBProject.VBComponents(sName)
    On Error GoTo 0
End Function

'Creates one fixture component and returns it.
Private Function CreateComponent(ByVal sKind As String, ByVal sName As String, ByVal sBody As String) As VBIDE.VBComponent
    Dim vbc As VBIDE.VBComponent
    Dim sPath As String

    Select Case sKind
        Case "M", "m"
            Set vbc = m_wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
            vbc.Name = sName
            FillModule vbc, sBody
        Case "C"
            Set vbc = m_wb.VBProject.VBComponents.Add(vbext_ct_ClassModule)
            vbc.Name = sName
            FillModule vbc, sBody
        Case "F"
            Set vbc = m_wb.VBProject.VBComponents.Add(vbext_ct_MSForm)
            vbc.Name = sName
            FillModule vbc, sBody
        Case "A", "X"
            sPath = Environ$("TEMP") & "\MT_Ruc_" & sName & ".cls"
            WriteClassFile sPath, sName, sBody, (sKind = "X")
            Set vbc = m_wb.VBProject.VBComponents.Import(sPath)
            DeleteIfExists sPath
        Case "S"
            Set vbc = SheetComponent()
            FillModule vbc, sBody
        Case Else
            Err.Raise vbObjectError + 2712, "CreateComponent", "Unknown kind " & sKind
    End Select
    Set CreateComponent = vbc
End Function

'Code of a component as one body ("|" between lines).
Private Function ComponentBody(ByVal vbc As VBIDE.VBComponent) As String
    Dim arrLines() As String
    Dim n As Long
    Dim i As Long
    Dim s As String

    n = modObfuscation.TrfReadLines(vbc.CodeModule, arrLines)
    For i = 1 To n
        If i > 1 Then s = s & "|"
        s = s & arrLines(i)
    Next i
    ComponentBody = s
End Function

'Current state of the fixture components in fixture format; removed components are left out.
Private Function Snapshot(ByVal sSpec As String) As String
    Dim asKind() As String
    Dim asName() As String
    Dim asBody() As String
    Dim n As Long
    Dim i As Long
    Dim vbc As VBIDE.VBComponent
    Dim s As String

    n = ParseSpec(sSpec, asKind, asName, asBody)
    For i = 0 To n - 1
        If asKind(i) = "S" Then
            Set vbc = SheetComponent()
        Else
            Set vbc = FindComponent(asName(i))
        End If
        If Not vbc Is Nothing Then
            If Len(s) > 0 Then s = s & "~"
            s = s & asKind(i) & ":" & asName(i) & ":" & ComponentBody(vbc)
        End If
    Next i
    Snapshot = s
End Function

'Readable cell text: one block per component, every code line prefixed with "| ".
Private Function Show(ByVal sSpec As String) As String
    Dim asKind() As String
    Dim asName() As String
    Dim asBody() As String
    Dim n As Long
    Dim i As Long
    Dim s As String

    If Len(sSpec) = 0 Then Exit Function
    n = ParseSpec(sSpec, asKind, asName, asBody)
    For i = 0 To n - 1
        If i > 0 Then s = s & vbLf & vbLf
        s = s & "[" & asKind(i) & "] " & asName(i)
        If Len(asBody(i)) > 0 Then s = s & vbLf & "| " & Replace(asBody(i), "|", vbLf & "| ")
    Next i
    Show = s
End Function

'Removes the fixture components (or keeps them for the compile check) and clears the sheet module.
Private Sub CleanUpCase(ByVal sSpec As String, ByVal bKeep As Boolean)
    Dim asKind() As String
    Dim asName() As String
    Dim asBody() As String
    Dim n As Long
    Dim i As Long
    Dim vbc As VBIDE.VBComponent

    n = ParseSpec(sSpec, asKind, asName, asBody)
    For i = 0 To n - 1
        If asKind(i) = "S" Then
            FillModule SheetComponent(), vbNullString
        ElseIf Not bKeep Then
            Set vbc = FindComponent(asName(i))
            If Not vbc Is Nothing Then m_wb.VBProject.VBComponents.Remove vbc
        End If
    Next i
End Sub

Private Function RemovedList() As String
    Dim i As Long
    Dim s As String

    For i = 1 To modRemoveUnusedCode.RemovedCount
        If i > 1 Then s = s & "; "
        s = s & modRemoveUnusedCode.RemovedItem(i)
    Next i
    RemovedList = s
End Function

Private Sub WriteResult(ByVal sId As String, ByVal sScope As String, ByVal sPurpose As String, ByVal sStatus As String, _
                        ByVal sCount As String, ByVal sNote As String, ByVal sRationale As String, _
                        ByVal sIn As String, ByVal sExp As String, ByVal sAct As String)
    m_lRow = m_lRow + 1
    m_ws.Range(m_ws.Cells(m_lRow, 1), m_ws.Cells(m_lRow, 10)).NumberFormat = "@"
    m_ws.Cells(m_lRow, 1).Value = sId
    m_ws.Cells(m_lRow, 2).Value = sScope
    m_ws.Cells(m_lRow, 3).Value = sPurpose
    m_ws.Cells(m_lRow, 4).Value = sStatus
    m_ws.Cells(m_lRow, 5).Value = sCount
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

'Builds the fixture, runs RemoveUnusedCode on its target components and compares the remaining code.
'bTwice: a second run must remove nothing and leave the code unchanged (idempotence).
Private Sub RunCase(ByVal sId As String, ByVal sScope As String, ByVal sPurpose As String, ByVal sIn As String, _
                    ByVal sExp As String, ByVal lExpCount As Long, ByVal sRationale As String, _
                    Optional ByVal bTwice As Boolean = False)
    Dim asKind() As String
    Dim asName() As String
    Dim asBody() As String
    Dim n As Long
    Dim i As Long
    Dim vbc As VBIDE.VBComponent
    Dim dicTargets As Scripting.Dictionary
    Dim lCount As Long
    Dim lCount2 As Long
    Dim sStored As String
    Dim sAct As String
    Dim sNote As String
    Dim sStatus As String
    Dim bSecondOk As Boolean
    Dim bCreated As Boolean
    Dim bWritten As Boolean

    On Error GoTo EH
    Set dicTargets = New Scripting.Dictionary
    dicTargets.CompareMode = vbTextCompare
    n = ParseSpec(sIn, asKind, asName, asBody)
    bCreated = True
    For i = 0 To n - 1
        Set vbc = CreateComponent(asKind(i), asName(i), asBody(i))
        If asKind(i) <> "S" And vbc.Name <> asName(i) Then
            Err.Raise vbObjectError + 2713, "RunCase", "Component was created as " & vbc.Name & " instead of " & asName(i)
        End If
        If asKind(i) <> "m" Then dicTargets(vbc.Name) = True
    Next i
    Set vbc = Nothing

    sStored = Snapshot(sIn)
    If sStored <> VisibleSpec(sIn) Then
        sStatus = "INCONCLUSIVE"
        sNote = "Editor changed the input when inserting it; adapt the test data."
        sAct = sStored
    Else
        lCount = modRemoveUnusedCode.RemoveUnusedCode(m_wb.VBProject, dicTargets, ScopeFlags(sScope))
        sNote = RemovedList()
        sAct = Snapshot(sIn)
        bSecondOk = True
        If bTwice Then
            lCount2 = modRemoveUnusedCode.RemoveUnusedCode(m_wb.VBProject, dicTargets, ScopeFlags(sScope))
            bSecondOk = (lCount2 = 0 And Snapshot(sIn) = sAct)
            If Not bSecondOk Then sNote = sNote & " | Second run changed the code (" & lCount2 & " removed)."
        End If
        If sAct <> sExp Then sNote = sNote & " | Output differs."
        If lCount <> lExpCount Then sNote = sNote & " | Count differs."
        If sAct = sExp And lCount = lExpCount And bSecondOk Then sStatus = "PASS" Else sStatus = "FAIL"
    End If
    WriteResult sId, sScope, sPurpose, sStatus, lCount & "/" & lExpCount, sNote, sRationale, sIn, sExp, sAct
    bWritten = True
    CleanUpCase sIn, KEEP_FIXTURES And sStatus <> "INCONCLUSIVE"
    Exit Sub
EH:
    sNote = sNote & " | Error " & Err.Number & ": " & Err.Description
    Resume AfterError
AfterError:
    On Error Resume Next
    If Not bWritten Then WriteResult sId, sScope, sPurpose, "ERROR", lCount & "/" & lExpCount, sNote, sRationale, sIn, sExp, sAct
    'an aborted fixture would disturb the compile check
    If bCreated Then CleanUpCase sIn, False
End Sub

'===============================================================================
' Generated test cases (single source: ruc_cases.py)
'===============================================================================

Private Sub RunAllCases()
    Case_R01
    Case_R02
    Case_R03
    Case_R04
    Case_R05
    Case_R06
    Case_R07
    Case_R08
    Case_R09
    Case_R10
    Case_R11
    Case_R12
    Case_R13
    Case_R14a
    Case_R14b
    Case_R14c
    Case_R15b
    Case_R15c
    Case_R16
    Case_R17
    Case_R18
    Case_R19
    Case_R20
    Case_R21
    Case_R22
    Case_R23
    Case_R24
    Case_R25
    Case_R26
    Case_R27
    Case_R28
    Case_R29
    Case_R30
    Case_R31
    Case_R32
    Case_R33
    Case_R34
    Case_R35
End Sub

Private Sub Case_R01()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry01()"
    AddLn b, "    Used01"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Used01()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Dead01()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U01", b
    b = vbNullString
    AddLn b, "Public Sub Entry01()"
    AddLn b, "    Used01"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Used01()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U01", b
    RunCase "R01", "A", "Ungenutzte Private Sub wird entfernt", sIn, sExp, 1, "Grundfall: Dead01 kommt ausserhalb der eigenen Zeilen nicht vor, Used01 schon."
End Sub

Private Sub Case_R02()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry02()"
    AddLn b, "End Sub"
    AddLn b, "Private Function Dead02() As Long"
    AddLn b, "    Dead02 = 1"
    AddLn b, "End Function"
    AddComp sIn, "M", "U02", b
    b = vbNullString
    AddLn b, "Public Sub Entry02()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U02", b
    RunCase "R02", "A", "Private Function, die nur sich selbst zuweist", sIn, sExp, 1, "Die Rueckgabezuweisung Dead02 = 1 liegt in den eigenen Zeilen und zaehlt nicht."
End Sub

Private Sub Case_R03()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry03()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Rec03(ByVal n03 As Long)"
    AddLn b, "    If n03 > 0 Then Rec03 n03 - 1"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U03", b
    b = vbNullString
    AddLn b, "Public Sub Entry03()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U03", b
    RunCase "R03", "A", "Rekursion ohne externen Aufrufer", sIn, sExp, 1, "Der rekursive Aufruf liegt in den eigenen Zeilen."
End Sub

Private Sub Case_R04()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry04()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub DeadA04()"
    AddLn b, "    DeadB04"
    AddLn b, "End Sub"
    AddLn b, "Private Sub DeadB04()"
    AddLn b, "    DeadC04"
    AddLn b, "End Sub"
    AddLn b, "Private Sub DeadC04()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U04", b
    b = vbNullString
    AddLn b, "Public Sub Entry04()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U04", b
    RunCase "R04", "A", "Kaskade ueber mehrere Durchlaeufe", sIn, sExp, 3, "DeadB04 wird erst nach DeadA04 unbenutzt, DeadC04 erst nach DeadB04 (drei Durchlaeufe)."
End Sub

Private Sub Case_R05()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry05()"
    AddLn b, "    Application.Run ""Timer05"""
    AddLn b, "End Sub"
    AddLn b, "Private Sub Timer05()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U05", b
    b = vbNullString
    AddLn b, "Public Sub Entry05()"
    AddLn b, "    Application.Run ""Timer05"""
    AddLn b, "End Sub"
    AddLn b, "Private Sub Timer05()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U05", b
    RunCase "R05", "A", "Name in einer Zeichenkette zaehlt als Verwendung", sIn, sExp, 0, "Application.Run ""Timer05"" haelt die Prozedur (bewusst konservativ)."
End Sub

Private Sub Case_R06()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry06()"
    AddLn b, "    ' Cmt06 is started by hand"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Cmt06()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U06", b
    b = vbNullString
    AddLn b, "Public Sub Entry06()"
    AddLn b, "    ' Cmt06 is started by hand"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Cmt06()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U06", b
    RunCase "R06", "A", "Name nur im Kommentar zaehlt als Verwendung", sIn, sExp, 0, "Konservativ: Kommentare werden mitdurchsucht. Mit Kommentarentfernung vorher faellt Cmt06 weg."
End Sub

Private Sub Case_R07()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry07()"
    AddLn b, "    Dim Init07Count As Long"
    AddLn b, "    Init07Count = 1"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Init07()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U07", b
    b = vbNullString
    AddLn b, "Public Sub Entry07()"
    AddLn b, "    Dim Init07Count As Long"
    AddLn b, "    Init07Count = 1"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U07", b
    RunCase "R07", "A", "Teilwort ist keine Verwendung", sIn, sExp, 1, "Init07Count enthaelt Init07 nur als Teilwort; die Wortgrenze verhindert den Treffer."
End Sub

Private Sub Case_R08()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Private WithEvents m_app08 As Application"
    AddLn b, "Private Sub Class_Initialize()"
    AddLn b, "    Set m_app08 = Application"
    AddLn b, "End Sub"
    AddLn b, "Private Sub m_app08_NewWorkbook(ByVal Wb As Workbook)"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Helper_08()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Dead08()"
    AddLn b, "End Sub"
    AddComp sIn, "C", "C08", b
    b = vbNullString
    AddLn b, "Public Sub Entry08()"
    AddLn b, "    Dim o08 As C08"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U08", b
    b = vbNullString
    AddLn b, "Private WithEvents m_app08 As Application"
    AddLn b, "Private Sub Class_Initialize()"
    AddLn b, "    Set m_app08 = Application"
    AddLn b, "End Sub"
    AddLn b, "Private Sub m_app08_NewWorkbook(ByVal Wb As Workbook)"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Helper_08()"
    AddLn b, "End Sub"
    AddComp sExp, "C", "C08", b
    b = vbNullString
    AddLn b, "Public Sub Entry08()"
    AddLn b, "    Dim o08 As C08"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U08", b
    RunCase "R08", "A", "Ereignisprozeduren und Unterstrich-Namen bleiben", sIn, sExp, 1, "Class_Initialize, WithEvents-Handler und Helper_08 enthalten ""_""; nur Dead08 faellt weg."
End Sub

Private Sub Case_R09()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry09()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub OnLoad09(ribbon As IRibbonUI)"
    AddLn b, "End Sub"
    AddLn b, "Private Sub OnAction09(control As IRibbonControl)"
    AddLn b, "End Sub"
    AddLn b, "Private Sub GetLabel09( _"
    AddLn b, "    control As IRibbonControl, ByRef returnedVal)"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Dead09()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U09", b
    b = vbNullString
    AddLn b, "Public Sub Entry09()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub OnLoad09(ribbon As IRibbonUI)"
    AddLn b, "End Sub"
    AddLn b, "Private Sub OnAction09(control As IRibbonControl)"
    AddLn b, "End Sub"
    AddLn b, "Private Sub GetLabel09( _"
    AddLn b, "    control As IRibbonControl, ByRef returnedVal)"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U09", b
    RunCase "R09", "A", "Ribbon-Callbacks bleiben (IRibbonUI, IRibbonControl, fortgesetzte Deklaration)", sIn, sExp, 1, "Callbacks werden aus customUI.xml aufgerufen; der Typ in der Deklaration schuetzt sie, auch ueber eine Zeilenfortsetzung."
End Sub

Private Sub Case_R10()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry10()"
    AddLn b, "End Sub"
    AddLn b, "#If VBA7 Then"
    AddLn b, "Private Function Ptr10() As LongPtr"
    AddLn b, "#Else"
    AddLn b, "Private Function Ptr10() As Long"
    AddLn b, "#End If"
    AddLn b, "    Ptr10 = 0"
    AddLn b, "End Function"
    AddComp sIn, "M", "U10", b
    b = vbNullString
    AddLn b, "Public Sub Entry10()"
    AddLn b, "End Sub"
    AddLn b, "#If VBA7 Then"
    AddLn b, "Private Function Ptr10() As LongPtr"
    AddLn b, "#Else"
    AddLn b, "Private Function Ptr10() As Long"
    AddLn b, "#End If"
    AddLn b, "    Ptr10 = 0"
    AddLn b, "End Function"
    AddComp sExp, "M", "U10", b
    RunCase "R10", "A", "Doppelte Deklaration in #If-Zweigen bleibt", sIn, sExp, 0, "Ptr10 ist zweimal deklariert; Loeschen wuerde die #If-Struktur zerstoeren."
End Sub

Private Sub Case_R11()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry11()"
    AddLn b, "End Sub"
    AddLn b, "#If DEBUG11 Then"
    AddLn b, "Private Sub Dbg11()"
    AddLn b, "End Sub"
    AddLn b, "#End If"
    AddComp sIn, "M", "U11", b
    b = vbNullString
    AddLn b, "Public Sub Entry11()"
    AddLn b, "End Sub"
    AddLn b, "#If DEBUG11 Then"
    AddLn b, "#End If"
    AddComp sExp, "M", "U11", b
    RunCase "R11", "A", "Prozedur im #If-Block wird ohne die Direktiven entfernt", sIn, sExp, 1, "Geloescht wird von der Deklaration bis End Sub; #If und #End If bleiben stehen."
End Sub

Private Sub Case_R12()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry12()"
    AddLn b, "End Sub"
    AddLn b, "Private Function Bits12() As Long"
    AddLn b, "#If Win64 Then"
    AddLn b, "    Bits12 = 64"
    AddLn b, "#Else"
    AddLn b, "    Bits12 = 32"
    AddLn b, "#End If"
    AddLn b, "End Function"
    AddComp sIn, "M", "U12", b
    b = vbNullString
    AddLn b, "Public Sub Entry12()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U12", b
    RunCase "R12", "A", "Ausgeglichene Direktiven im Rumpf: Prozedur wird entfernt", sIn, sExp, 1, "#If, #Else und #End If liegen vollstaendig innerhalb der Prozedur."
End Sub

Private Sub Case_R13()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Private m_col13 As Collection"
    AddLn b, "Public Function Item13(ByVal i13 As Long) As Variant"
    AddLn b, "Attribute Item13.VB_UserMemId = 0"
    AddLn b, "    Item13 = m_col13(i13)"
    AddLn b, "End Function"
    AddLn b, "Public Function NewEnum13() As IUnknown"
    AddLn b, "Attribute NewEnum13.VB_UserMemId = -4"
    AddLn b, "    Set NewEnum13 = m_col13.[_NewEnum]"
    AddLn b, "End Function"
    AddLn b, "Public Sub Dead13()"
    AddLn b, "End Sub"
    AddComp sIn, "A", "C13", b
    b = vbNullString
    AddLn b, "Public Sub Entry_13()"
    AddLn b, "    Dim o13 As C13"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U13", b
    b = vbNullString
    AddLn b, "Private m_col13 As Collection"
    AddLn b, "Public Function Item13(ByVal i13 As Long) As Variant"
    AddLn b, "    Item13 = m_col13(i13)"
    AddLn b, "End Function"
    AddLn b, "Public Function NewEnum13() As IUnknown"
    AddLn b, "    Set NewEnum13 = m_col13.[_NewEnum]"
    AddLn b, "End Function"
    AddComp sExp, "A", "C13", b
    b = vbNullString
    AddLn b, "Public Sub Entry_13()"
    AddLn b, "    Dim o13 As C13"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U13", b
    RunCase "R13", "C", "Versteckte Attribute schuetzen Mitglieder (Standardmember, NewEnum)", sIn, sExp, 1, "Item13 (VB_UserMemId = 0) und NewEnum13 (-4) werden ohne Namen aufgerufen; nur Dead13 faellt weg."
End Sub

Private Sub Case_R14a()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Used14a()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ClsPub14a()"
    AddLn b, "End Sub"
    AddLn b, "Friend Sub ClsFriend14a()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub ClsPriv14a()"
    AddLn b, "End Sub"
    AddComp sIn, "C", "C14a", b
    b = vbNullString
    AddLn b, "Public Sub Entry_14a()"
    AddLn b, "    Dim o14a As C14a"
    AddLn b, "    Set o14a = New C14a"
    AddLn b, "    o14a.Used14a"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ModPub14a()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub ModPriv14a()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U14a", b
    b = vbNullString
    AddLn b, "Public Sub Used14a()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ClsPub14a()"
    AddLn b, "End Sub"
    AddLn b, "Friend Sub ClsFriend14a()"
    AddLn b, "End Sub"
    AddComp sExp, "C", "C14a", b
    b = vbNullString
    AddLn b, "Public Sub Entry_14a()"
    AddLn b, "    Dim o14a As C14a"
    AddLn b, "    Set o14a = New C14a"
    AddLn b, "    o14a.Used14a"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ModPub14a()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U14a", b
    RunCase "R14a", "A", "Variante A: nur Private", sIn, sExp, 2, "Public und Friend bleiben in Klasse und Modul; ClsPriv14a und ModPriv14a fallen weg."
End Sub

Private Sub Case_R14b()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Used14b()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ClsPub14b()"
    AddLn b, "End Sub"
    AddLn b, "Friend Sub ClsFriend14b()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub ClsPriv14b()"
    AddLn b, "End Sub"
    AddComp sIn, "C", "C14b", b
    b = vbNullString
    AddLn b, "Public Sub Entry_14b()"
    AddLn b, "    Dim o14b As C14b"
    AddLn b, "    Set o14b = New C14b"
    AddLn b, "    o14b.Used14b"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ModPub14b()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub ModPriv14b()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U14b", b
    b = vbNullString
    AddLn b, "Public Sub Used14b()"
    AddLn b, "End Sub"
    AddComp sExp, "C", "C14b", b
    b = vbNullString
    AddLn b, "Public Sub Entry_14b()"
    AddLn b, "    Dim o14b As C14b"
    AddLn b, "    Set o14b = New C14b"
    AddLn b, "    o14b.Used14b"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ModPub14b()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U14b", b
    RunCase "R14b", "B", "Variante B: zusaetzlich Public und Friend in Klassen", sIn, sExp, 4, "ClsPub14b und ClsFriend14b fallen zusaetzlich weg, ModPub14b bleibt."
End Sub

Private Sub Case_R14c()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Used14c()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ClsPub14c()"
    AddLn b, "End Sub"
    AddLn b, "Friend Sub ClsFriend14c()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub ClsPriv14c()"
    AddLn b, "End Sub"
    AddComp sIn, "C", "C14c", b
    b = vbNullString
    AddLn b, "Public Sub Entry_14c()"
    AddLn b, "    Dim o14c As C14c"
    AddLn b, "    Set o14c = New C14c"
    AddLn b, "    o14c.Used14c"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ModPub14c()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub ModPriv14c()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U14c", b
    b = vbNullString
    AddLn b, "Public Sub Used14c()"
    AddLn b, "End Sub"
    AddComp sExp, "C", "C14c", b
    b = vbNullString
    AddLn b, "Public Sub Entry_14c()"
    AddLn b, "    Dim o14c As C14c"
    AddLn b, "    Set o14c = New C14c"
    AddLn b, "    o14c.Used14c"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U14c", b
    RunCase "R14c", "C", "Variante C: zusaetzlich Public in Standardmodulen", sIn, sExp, 5, "ModPub14c faellt zusaetzlich weg; Used14c bleibt wegen o14c.Used14c."
End Sub

Private Sub Case_R15b()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Private Sub Worksheet_Activate()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub SheetPub15b()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub SheetPriv15b()"
    AddLn b, "End Sub"
    AddComp sIn, "S", "Sheet", b
    b = vbNullString
    AddLn b, "Private Sub Worksheet_Activate()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub SheetPub15b()"
    AddLn b, "End Sub"
    AddComp sExp, "S", "Sheet", b
    RunCase "R15b", "B", "Dokumentmodul in Variante B", sIn, sExp, 1, "Worksheet_Activate bleibt (Ereignis), SheetPriv15b faellt weg, SheetPub15b bleibt (Public im Dokumentmodul erst in C)."
End Sub

Private Sub Case_R15c()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Private Sub Worksheet_Activate()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub SheetPub15c()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub SheetPriv15c()"
    AddLn b, "End Sub"
    AddComp sIn, "S", "Sheet", b
    b = vbNullString
    AddLn b, "Private Sub Worksheet_Activate()"
    AddLn b, "End Sub"
    AddComp sExp, "S", "Sheet", b
    RunCase "R15c", "C", "Dokumentmodul in Variante C", sIn, sExp, 2, "SheetPub15c und SheetPriv15c fallen weg, Worksheet_Activate bleibt."
End Sub

Private Sub Case_R16()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry16()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Dead16()"
    AddLn b, "    Dim o16 As C16b"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U16", b
    b = vbNullString
    AddLn b, "Public Sub Run16a()"
    AddLn b, "End Sub"
    AddComp sIn, "C", "C16a", b
    b = vbNullString
    AddLn b, "Public Sub Run16b()"
    AddLn b, "End Sub"
    AddComp sIn, "C", "C16b", b
    b = vbNullString
    AddLn b, "Public Sub Entry16()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U16", b
    RunCase "R16", "A", "Ungenutzte Klasse und Klasse aus entfernter Prozedur", sIn, sExp, 3, "C16a wird nirgends genannt; C16b nur in Dead16, das im ersten Durchlauf faellt."
End Sub

Private Sub Case_R17()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Run17()"
    AddLn b, "End Sub"
    AddComp sIn, "X", "C17", b
    b = vbNullString
    AddLn b, "Public Sub Run17()"
    AddLn b, "End Sub"
    AddComp sExp, "X", "C17", b
    RunCase "R17", "A", "Exponierte Klasse (VB_Exposed = True) bleibt", sIn, sExp, 0, "PublicNotCreatable-Klassen koennen aus anderen Projekten verwendet werden."
End Sub

Private Sub Case_R18()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry18()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Pr" & ChrW(252) & "fen18()"
    AddLn b, "End Sub"
    AddLn b, "Private Function Text18$()"
    AddLn b, "End Function"
    AddLn b, "Private Sub My_Helper18()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U18", b
    b = vbNullString
    AddLn b, "Public Sub Entry18()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Pr" & ChrW(252) & "fen18()"
    AddLn b, "End Sub"
    AddLn b, "Private Function Text18$()"
    AddLn b, "End Function"
    AddLn b, "Private Sub My_Helper18()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U18", b
    RunCase "R18", "A", "Namen mit Umlaut, Typkennzeichen oder Unterstrich bleiben", sIn, sExp, 0, "Die RegExp-Wortgrenze kennt nur ASCII-Buchstaben; solche Namen werden nie entfernt."
End Sub

Private Sub Case_R19()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Private Sub UserForm_Initialize()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub Pub19()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Priv19()"
    AddLn b, "End Sub"
    AddComp sIn, "F", "F19", b
    b = vbNullString
    AddLn b, "Private Sub UserForm_Initialize()"
    AddLn b, "End Sub"
    AddComp sExp, "F", "F19", b
    RunCase "R19", "B", "UserForm: Ereignis bleibt, ungenutzte Mitglieder fallen weg", sIn, sExp, 2, "UserForm_Initialize enthaelt ""_""; Pub19 (Variante B) und Priv19 fallen weg; Formulare werden nie entfernt."
End Sub

Private Sub Case_R20()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry20()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Called20()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U20", b
    b = vbNullString
    AddLn b, "Private Sub Other20()"
    AddLn b, "    Application.Run ""Called20"""
    AddLn b, "End Sub"
    AddComp sIn, "m", "V20", b
    b = vbNullString
    AddLn b, "Public Sub Entry20()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Called20()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U20", b
    b = vbNullString
    AddLn b, "Private Sub Other20()"
    AddLn b, "    Application.Run ""Called20"""
    AddLn b, "End Sub"
    AddComp sExp, "m", "V20", b
    RunCase "R20", "A", "Nicht ausgewaehlte Komponente bleibt unveraendert, ihre Aufrufe zaehlen", sIn, sExp, 0, "V20 ist kein Ziel: Other20 bleibt, und der Aufruf ""Called20"" haelt Called20 in U20."
End Sub

Private Sub Case_R21()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry21()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Ping21()"
    AddLn b, "    Pong21"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Pong21()"
    AddLn b, "    Ping21"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U21", b
    b = vbNullString
    AddLn b, "Public Sub Entry21()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Ping21()"
    AddLn b, "    Pong21"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Pong21()"
    AddLn b, "    Ping21"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U21", b
    RunCase "R21", "A", "Gegenseitige Rekursion bleibt (bekannte Grenze)", sIn, sExp, 0, "Ping21 und Pong21 nennen sich gegenseitig; die Wortsuche sieht beide als benutzt."
End Sub

Private Sub Case_R22()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Run22()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Dead22()"
    AddLn b, "End Sub"
    AddComp sIn, "C", "I22", b
    b = vbNullString
    AddLn b, "Implements I22"
    AddLn b, "Private Sub I22_Run22()"
    AddLn b, "End Sub"
    AddComp sIn, "C", "C22", b
    b = vbNullString
    AddLn b, "Public Sub Entry_22()"
    AddLn b, "    Dim o22 As I22"
    AddLn b, "    Set o22 = New C22"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U22", b
    b = vbNullString
    AddLn b, "Public Sub Run22()"
    AddLn b, "End Sub"
    AddComp sExp, "C", "I22", b
    b = vbNullString
    AddLn b, "Implements I22"
    AddLn b, "Private Sub I22_Run22()"
    AddLn b, "End Sub"
    AddComp sExp, "C", "C22", b
    b = vbNullString
    AddLn b, "Public Sub Entry_22()"
    AddLn b, "    Dim o22 As I22"
    AddLn b, "    Set o22 = New C22"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U22", b
    RunCase "R22", "C", "Implements: Schnittstellenmitglieder und Implementierungen bleiben", sIn, sExp, 1, "Run22 gehoert zur Schnittstelle I22, I22_Run22 enthaelt ""_""; nur das Private Dead22 faellt weg."
End Sub

Private Sub Case_R23()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry23()"
    AddLn b, "End Sub"
    AddLn b, "' note 23 _"
    AddLn b, "Private Sub NotDecl23()"
    AddComp sIn, "M", "U23", b
    b = vbNullString
    AddLn b, "Public Sub Entry23()"
    AddLn b, "End Sub"
    AddLn b, "' note 23 _"
    AddLn b, "Private Sub NotDecl23()"
    AddComp sExp, "M", "U23", b
    RunCase "R23", "A", "Kommentar mit Fortsetzung vor einer Deklaration", sIn, sExp, 0, "Die Zeile nach ""' note 23 _"" gehoert zum Kommentar und ist keine Deklaration."
End Sub

Private Sub Case_R24()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry24()"
    AddLn b, "End Sub"
    AddLn b, "' helper 24"
    AddLn b, "Private Sub Cont24( _"
    AddLn b, "    ByVal a24 As Long)"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U24", b
    b = vbNullString
    AddLn b, "Public Sub Entry24()"
    AddLn b, "End Sub"
    AddLn b, "' helper 24"
    AddComp sExp, "M", "U24", b
    RunCase "R24", "A", "Fortgesetzte Deklaration, Kommentar darueber bleibt", sIn, sExp, 1, "Geloescht wird von der ersten Deklarationszeile bis End Sub; der Kommentar darueber bleibt."
End Sub

Private Sub Case_R25()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry_25()"
    AddLn b, "End Sub"
    AddLn b, "Public Function Udf25() As Long"
    AddLn b, "    Udf25 = 1"
    AddLn b, "End Function"
    AddComp sIn, "M", "U25", b
    b = vbNullString
    AddLn b, "Public Sub Entry_25()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U25", b
    RunCase "R25", "C", "Restrisiko: Public Function aus einer Zellformel faellt in Variante C weg", sIn, sExp, 1, "Belegt das dokumentierte Restrisiko: Aufrufe ausserhalb des Codes (hier eine UDF) werden nicht erkannt."
End Sub

Private Sub Case_R26()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry26()"
    AddLn b, "    U26b.Helper26"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U26a", b
    b = vbNullString
    AddLn b, "Public Sub Helper26()"
    AddLn b, "    Inner26"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Inner26()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Dead26()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U26b", b
    b = vbNullString
    AddLn b, "Public Sub Entry26()"
    AddLn b, "    U26b.Helper26"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U26a", b
    b = vbNullString
    AddLn b, "Public Sub Helper26()"
    AddLn b, "    Inner26"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Inner26()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U26b", b
    RunCase "R26", "A", "Verwendung ueber Modulgrenzen und qualifiziert", sIn, sExp, 1, "U26b.Helper26 und Inner26 sind benutzt; nur Dead26 faellt weg."
End Sub

Private Sub Case_R27()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry27()"
    AddLn b, "    U27b.Helper27"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U27a", b
    b = vbNullString
    AddLn b, "Public Sub Helper27()"
    AddLn b, "    Inner27"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Inner27()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Dead27()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U27b", b
    b = vbNullString
    AddComp sExp, "M", "U27a", b
    b = vbNullString
    AddComp sExp, "M", "U27b", b
    RunCase "R27", "C", "Variante C ohne geschuetzten Einstieg leert beide Module", sIn, sExp, 4, "Entry27 hat keinen Aufrufer im Code; danach fallen Helper27 und Inner27 in Folge. Module bleiben (leer)."
End Sub

Private Sub Case_R28()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Used28()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ClsPub28()"
    AddLn b, "End Sub"
    AddLn b, "Friend Sub ClsFriend28()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub ClsPriv28()"
    AddLn b, "End Sub"
    AddComp sIn, "C", "C28", b
    b = vbNullString
    AddLn b, "Public Sub Entry_28()"
    AddLn b, "    Dim o28 As C28"
    AddLn b, "    Set o28 = New C28"
    AddLn b, "    o28.Used28"
    AddLn b, "End Sub"
    AddLn b, "Public Sub ModPub28()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub ModPriv28()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U28", b
    b = vbNullString
    AddLn b, "Public Sub Used28()"
    AddLn b, "End Sub"
    AddComp sExp, "C", "C28", b
    b = vbNullString
    AddLn b, "Public Sub Entry_28()"
    AddLn b, "    Dim o28 As C28"
    AddLn b, "    Set o28 = New C28"
    AddLn b, "    o28.Used28"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U28", b
    RunCase "R28", "C", "Zweiter Lauf aendert nichts (Idempotenz)", sIn, sExp, 5, "Der zweite Aufruf muss 0 liefern und den Code unveraendert lassen.", True
End Sub

Private Sub Case_R29()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry29()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Num29()"
    AddLn b, "10  x29 = 1"
    AddLn b, "20  x29 = 2"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U29", b
    b = vbNullString
    AddLn b, "Public Sub Entry29()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U29", b
    RunCase "R29", "A", "Zeilennummern im Rumpf", sIn, sExp, 1, "Nummerierte Zeilen stoeren weder die Suche nach End Sub noch das Loeschen."
End Sub

Private Sub Case_R30()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Private m_v30 As Long"
    AddLn b, "Private Property Get Value30() As Long"
    AddLn b, "    Value30 = m_v30"
    AddLn b, "End Property"
    AddLn b, "Private Sub Dead30()"
    AddLn b, "End Sub"
    AddComp sIn, "C", "C30", b
    b = vbNullString
    AddLn b, "Public Sub Entry30()"
    AddLn b, "    Dim o30 As C30"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U30", b
    b = vbNullString
    AddLn b, "Private m_v30 As Long"
    AddLn b, "Private Property Get Value30() As Long"
    AddLn b, "    Value30 = m_v30"
    AddLn b, "End Property"
    AddComp sExp, "C", "C30", b
    b = vbNullString
    AddLn b, "Public Sub Entry30()"
    AddLn b, "    Dim o30 As C30"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U30", b
    RunCase "R30", "A", "Properties werden nie entfernt", sIn, sExp, 1, "Value30 ist eine ungenutzte Private Property und bleibt; Dead30 faellt weg."
End Sub

Private Sub Case_R31()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry31()"
    AddLn b, "End Sub"
    AddLn b, "Private Static Sub Stat31()"
    AddLn b, "End Sub"
    AddLn b, "Static Function Stat31b() As Long"
    AddLn b, "End Function"
    AddComp sIn, "M", "U31", b
    b = vbNullString
    AddLn b, "Public Sub Entry31()"
    AddLn b, "End Sub"
    AddLn b, "Static Function Stat31b() As Long"
    AddLn b, "End Function"
    AddComp sExp, "M", "U31", b
    RunCase "R31", "A", "Static und fehlendes Schluesselwort", sIn, sExp, 1, "Private Static Sub ist Private und faellt weg; Static Function ohne Schluesselwort ist Public und bleibt in A."
End Sub

Private Sub Case_R32()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry32()"
    AddLn b, "End Sub"
    AddLn b, "Private Sub Dead32() ' helper"
    AddLn b, "End Sub ' Dead32"
    AddComp sIn, "M", "U32", b
    b = vbNullString
    AddLn b, "Public Sub Entry32()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U32", b
    RunCase "R32", "A", "Kommentar auf der End-Zeile gehoert zur Prozedur", sIn, sExp, 1, """End Sub ' Dead32"" liegt in den eigenen Zeilen und zaehlt nicht als Verwendung."
End Sub

Private Sub Case_R33()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub Entry33()"
    AddLn b, "    Dim p33 As LongPtr"
    AddLn b, "    p33 = Addr33(AddressOf Callback33)"
    AddLn b, "End Sub"
    AddLn b, "Private Function Addr33(ByVal a33 As LongPtr) As LongPtr"
    AddLn b, "    Addr33 = a33"
    AddLn b, "End Function"
    AddLn b, "Private Function Callback33(ByVal h33 As LongPtr) As Long"
    AddLn b, "End Function"
    AddComp sIn, "M", "U33", b
    b = vbNullString
    AddLn b, "Public Sub Entry33()"
    AddLn b, "    Dim p33 As LongPtr"
    AddLn b, "    p33 = Addr33(AddressOf Callback33)"
    AddLn b, "End Sub"
    AddLn b, "Private Function Addr33(ByVal a33 As LongPtr) As LongPtr"
    AddLn b, "    Addr33 = a33"
    AddLn b, "End Function"
    AddLn b, "Private Function Callback33(ByVal h33 As LongPtr) As Long"
    AddLn b, "End Function"
    AddComp sExp, "M", "U33", b
    RunCase "R33", "A", "AddressOf haelt einen Callback", sIn, sExp, 0, "AddressOf Callback33 nennt den Namen im Code."
End Sub

Private Sub Case_R34()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Enum Color34"
    AddLn b, "    Red34 = 1"
    AddLn b, "End Enum"
    AddComp sIn, "C", "C34a", b
    b = vbNullString
    AddLn b, "Private Enum Size34"
    AddLn b, "    Small34 = 1"
    AddLn b, "End Enum"
    AddComp sIn, "C", "C34b", b
    b = vbNullString
    AddLn b, "Public Sub Entry34()"
    AddLn b, "    Dim c34 As Long"
    AddLn b, "    c34 = Red34"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U34", b
    b = vbNullString
    AddLn b, "Public Enum Color34"
    AddLn b, "    Red34 = 1"
    AddLn b, "End Enum"
    AddComp sExp, "C", "C34a", b
    b = vbNullString
    AddLn b, "Public Sub Entry34()"
    AddLn b, "    Dim c34 As Long"
    AddLn b, "    c34 = Red34"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U34", b
    RunCase "R34", "A", "Klasse mit Public Enum bleibt, Klasse mit Private Enum faellt weg", sIn, sExp, 1, "Public Enum einer Klasse ist projektweit ohne Klassennamen nutzbar (Red34); Private Enum nicht."
End Sub

Private Sub Case_R35()
    Dim sIn As String, sExp As String, b As String
    b = vbNullString
    AddLn b, "Public Sub AutoOpen()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub AutoExec()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub Plain35()"
    AddLn b, "End Sub"
    AddComp sIn, "M", "U35", b
    b = vbNullString
    AddLn b, "Public Sub AutoOpen()"
    AddLn b, "End Sub"
    AddLn b, "Public Sub AutoExec()"
    AddLn b, "End Sub"
    AddComp sExp, "M", "U35", b
    RunCase "R35", "C", "Word-Automakros bleiben auch in Variante C", sIn, sExp, 1, "AutoOpen und AutoExec ruft Word ueber den Namen auf; Plain35 faellt weg."
End Sub
