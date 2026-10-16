Attribute VB_Name = "modRemoveUnusedCode"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : modRemoveUnusedCode - removal of unused Subs, Functions and class modules
'* Created    : 15-10-2026
'* Purpose    : Optional step of "Removing Formatting" (ObfuscationCode). Finds Subs,
'*              Functions and class modules whose name occurs nowhere else in the project
'*              code (RegExp word search) and deletes them.
'* Tests      : RemoveUnusedTests.RunRemoveUnusedTests
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

'===============================================================================
' Rules
'
' Usage is a whole-word, case-insensitive RegExp search over the code of ALL
' components of the project, strings and comments included. A name that occurs
' anywhere outside the procedure's own lines counts as used. This is
' deliberately conservative: Application.Run "Name" or OnAction = "Name" keep
' the procedure.
'
' A Sub or Function is removed when all of these hold:
'   - its component is one of the targets passed by the caller
'   - its scope is covered by the scope flags: Private always; Public and
'     Friend only with the flag for its module type (see enumUnusedScope)
'   - its name occurs nowhere in the project code outside its own lines
' It is always kept when:
'   - its name contains "_" (event handlers, WithEvents, Implements,
'     Auto_Open) or any character other than A-Z, a-z, 0-9 (type characters,
'     umlauts; the RegExp word boundary knows only ASCII letters)
'   - its name is a Word auto macro (AutoExec, AutoNew, AutoOpen, AutoClose,
'     AutoExit)
'   - its declaration names IRibbonControl or IRibbonUI (ribbon callbacks are
'     called from the customUI XML, not from the code)
'   - its name is declared more than once in the module (#If branches)
'   - it is a Public or Friend member of a class that is used in an
'     Implements statement
'   - it carries hidden attributes (VB_UserMemId, VB_ProcData, VB_Description)
'   - its lines contain unbalanced #If directives, or its End line is not found
' Properties are never removed.
'
' A class module is removed when it is a target, its name occurs nowhere
' outside its own code, it declares no Public Enum (its members are used
' without the class name) and it is not exposed to other projects
' (VB_Exposed). UserForms, standard modules and document modules are never
' removed.
'
' Removing a procedure can make others unused, so the passes repeat until
' nothing changes. A procedure is deleted from its declaration to its End line;
' comment lines above it stay. Two unused procedures that call each other are
' both kept (each name occurs in the other one).
'
' Not detected (residual risk, shown in the dialog): callers outside the code
' (ribbon callbacks declared without IRibbonControl, macros assigned to shapes
' or buttons, Alt+F8, worksheet formulas, Word command macros such as
' FileSave, other projects) and names built at runtime
' (Application.Run "P" & i, CallByName with a variable).
'===============================================================================

'Scope flags of RemoveUnusedCode. Private Subs and Functions are always candidates.
'Variant A = enumUnusedScopePrivateOnly
'Variant B = enumUnusedScopePublicClassMembers
'Variant C = enumUnusedScopePublicClassMembers Or enumUnusedScopePublicModuleProcs
Public Enum enumUnusedScope
    enumUnusedScopePrivateOnly = 0          'Private Subs and Functions of every module type
    enumUnusedScopePublicClassMembers = 1   'adds Public and Friend Subs and Functions of class modules and UserForms
    enumUnusedScopePublicModuleProcs = 2    'adds Public Subs and Functions of standard and document modules
End Enum

'Safety limit; every pass except the last one removes at least one item
Private Const RUC_MAX_PASSES As Long = 10000

'Word auto macros: called by name from Word, never from the code (lower case, "|" separated)
Private Const RUC_KEEP_NAMES As String = "|autoexec|autonew|autoopen|autoclose|autoexit|"

Private Type RucProc
    Name As String
    Kind As String              '"Sub", "Function" or "Property"
    Scope As String             '"Private", "Public" or "Friend" (no keyword = Public)
    DeclLine As Long            'first physical line of the declaration
    DeclText As String          'declaration as one line (line continuations joined)
    EndLine As Long             'line of the matching End statement, 0 if not found
End Type

Private m_colRemoved As Collection
Private m_reWord As RegExp          'identifier tokens
Private m_reDecl As RegExp          'Sub, Function or Property declaration
Private m_reEnd As RegExp           'End Sub, End Function or End Property
Private m_reImpl As RegExp          'Implements statement
Private m_reDirective As RegExp     '#If, #ElseIf, #Else, #End If
Private m_reEnum As RegExp          'Public Enum declaration (also without access keyword)

'===============================================================================
' Procedure : RemoveUnusedCode
' Purpose   : Removes unused Subs, Functions and class modules of a VBA project.
'
' Parameters:
'   vbProj     - VBIDE.VBProject, ByVal, project to transform
'   dicTargets - Scripting.Dictionary, ByVal, keys = names of the components
'                that may be changed; all other components are only searched
'   lScope     - Long, ByVal, combination of enumUnusedScope flags
'
' Returns:
'   Number of removed procedures and class modules
'
' Side Effects:
'   Deletes procedures from the target components and removes target class
'   modules (rules: see the top of this module). Exports components to the
'   TEMP folder to read hidden attributes; the files are deleted again.
'   Writes one line per removed item to the Immediate window.
'
' Preconditions:
'   Module code is syntactically valid VBA. Access to the VBA project object
'   model is trusted. The project is not locked.
'
' Postconditions:
'   No target procedure or class module that the rules allow to remove is
'   left unused. RemovedCount and RemovedItem report the removed items.
'
' Errors:
'   No local handler. Errors of the VBIDE object model and of the export are
'   passed to the caller; the project may then be partly transformed.
'
' Notes:
'   Run it after Remove_DebugPrint: a procedure called only inside
'   Debug.Print becomes unused only then (test T24 of the chain). Removing
'   comments first gives fewer false "used" results.
'===============================================================================
Public Function RemoveUnusedCode(ByVal vbProj As VBIDE.VBProject, ByVal dicTargets As Scripting.Dictionary, _
                                 ByVal lScope As Long) As Long
    Dim dicTgt As Scripting.Dictionary
    Dim dicExport As Scripting.Dictionary
    Dim vKey As Variant
    Dim lPass As Long

    InitRegExps
    Set m_colRemoved = New Collection
    Set dicTgt = New Scripting.Dictionary
    dicTgt.CompareMode = vbTextCompare
    For Each vKey In dicTargets.Keys
        dicTgt(CStr(vKey)) = True
    Next vKey
    Set dicExport = New Scripting.Dictionary
    dicExport.CompareMode = vbTextCompare

    For lPass = 1 To RUC_MAX_PASSES
        If RemoveUnusedProcs(vbProj, dicTgt, lScope, dicExport) = 0 Then
            If RemoveUnusedClasses(vbProj, dicTgt, dicExport) = 0 Then Exit For
        End If
    Next lPass
    RemoveUnusedCode = m_colRemoved.Count
End Function

'Number of items removed by the last RemoveUnusedCode run.
Public Function RemovedCount() As Long
    If Not m_colRemoved Is Nothing Then RemovedCount = m_colRemoved.Count
End Function

'Description of removed item lIndex (1-based), for example "Module1.Helper (Private Sub)".
Public Function RemovedItem(ByVal lIndex As Long) As String
    RemovedItem = m_colRemoved.Item(lIndex)
End Function

'===============================================================================
' Passes
'===============================================================================

'Deletes every unused Sub and Function of the target components that the scope allows.
'Returns the number of deleted procedures.
Private Function RemoveUnusedProcs(ByVal vbProj As VBIDE.VBProject, ByVal dicTargets As Scripting.Dictionary, _
                                   ByVal lScope As Long, ByVal dicExport As Scripting.Dictionary) As Long
    Dim dicWords As Scripting.Dictionary
    Dim dicIfaces As Scripting.Dictionary
    Dim vbc As VBIDE.VBComponent
    Dim arrLines() As String
    Dim n As Long
    Dim procs() As RucProc
    Dim nProcs As Long
    Dim abDel() As Boolean
    Dim i As Long
    Dim lCount As Long

    Set dicWords = ProjectWordCounts(vbProj)
    Set dicIfaces = ImplementedInterfaces(vbProj)
    For Each vbc In vbProj.VBComponents
        If dicTargets.Exists(vbc.Name) Then
            n = modObfuscation.TrfReadLines(vbc.CodeModule, arrLines)
            nProcs = ScanProcedures(arrLines, n, procs)
            If nProcs > 0 Then
                ReDim abDel(1 To nProcs)
                For i = 1 To nProcs
                    abDel(i) = IsRemovable(vbc, procs, nProcs, i, arrLines, lScope, dicWords, dicIfaces, dicExport)
                Next i
                'bottom-up, so that the line numbers of pending procedures stay valid
                For i = nProcs To 1 Step -1
                    If abDel(i) Then
                        vbc.CodeModule.DeleteLines procs(i).DeclLine, procs(i).EndLine - procs(i).DeclLine + 1
                        RecordRemoval vbc.Name & "." & procs(i).Name & " (" & procs(i).Scope & " " & procs(i).Kind & ")"
                        lCount = lCount + 1
                    End If
                Next i
            End If
        End If
    Next vbc
    RemoveUnusedProcs = lCount
End Function

'Removes every unused target class module. Returns the number of removed classes.
Private Function RemoveUnusedClasses(ByVal vbProj As VBIDE.VBProject, ByVal dicTargets As Scripting.Dictionary, _
                                     ByVal dicExport As Scripting.Dictionary) As Long
    Dim dicWords As Scripting.Dictionary
    Dim vbc As VBIDE.VBComponent
    Dim colNames As Collection
    Dim vName As Variant
    Dim vInfo As Variant
    Dim sCode As String

    Set dicWords = ProjectWordCounts(vbProj)
    Set colNames = New Collection
    For Each vbc In vbProj.VBComponents
        If vbc.Type = vbext_ct_ClassModule Then
            If dicTargets.Exists(vbc.Name) Then
                If IsPlainName(vbc.Name, True) Then
                    sCode = ComponentText(vbc)
                    If Not m_reEnum.Test(sCode) Then
                        If WordCount(dicWords, vbc.Name) - CountWord(sCode, vbc.Name) = 0 Then
                            vInfo = ExportInfo(vbc, dicExport)
                            If Not CBool(vInfo(1)) Then colNames.Add vbc.Name
                        End If
                    End If
                End If
            End If
        End If
    Next vbc
    For Each vName In colNames
        vbProj.VBComponents.Remove vbProj.VBComponents(CStr(vName))
        'no second attempt, even if the editor removes the component only later (own project)
        dicTargets.Remove CStr(vName)
        RecordRemoval "class module " & CStr(vName)
    Next vName
    RemoveUnusedClasses = colNames.Count
End Function

'True if procedure idx of the component may be deleted (rules: see the top of this module).
Private Function IsRemovable(ByVal vbc As VBIDE.VBComponent, ByRef procs() As RucProc, ByVal nProcs As Long, _
                             ByVal idx As Long, ByRef arrLines() As String, ByVal lScope As Long, _
                             ByVal dicWords As Scripting.Dictionary, ByVal dicIfaces As Scripting.Dictionary, _
                             ByVal dicExport As Scripting.Dictionary) As Boolean
    Dim lOutside As Long
    Dim vInfo As Variant

    With procs(idx)
        If .Kind = "Property" Then Exit Function
        If .EndLine = 0 Then Exit Function
        If Not IsPlainName(.Name, False) Then Exit Function
        If InStr(1, RUC_KEEP_NAMES, "|" & LCase$(.Name) & "|") > 0 Then Exit Function
        If Not ScopeAllowed(vbc.Type, .Scope, lScope) Then Exit Function
        If .Scope <> "Private" And vbc.Type = vbext_ct_ClassModule Then
            If dicIfaces.Exists(vbc.Name) Then Exit Function
        End If
        If InStr(1, .DeclText, "IRibbonControl", vbTextCompare) > 0 Then Exit Function
        If InStr(1, .DeclText, "IRibbonUI", vbTextCompare) > 0 Then Exit Function
        If DeclarationCount(procs, nProcs, .Name) > 1 Then Exit Function
        If Not DirectivesBalanced(arrLines, .DeclLine, .EndLine) Then Exit Function
        lOutside = WordCount(dicWords, .Name) - CountWord(JoinLines(arrLines, .DeclLine, .EndLine), .Name)
        If lOutside > 0 Then Exit Function
        'last, because it exports the component (once per run)
        vInfo = ExportInfo(vbc, dicExport)
        If InStr(1, vInfo(0), "|" & LCase$(.Name) & "|") > 0 Then Exit Function
    End With
    IsRemovable = True
End Function

'True if the scope flags cover a procedure with the given access keyword in a component of the given type.
Private Function ScopeAllowed(ByVal lType As VBIDE.vbext_ComponentType, ByVal sScope As String, ByVal lScope As Long) As Boolean
    Select Case lType
        Case vbext_ct_StdModule, vbext_ct_Document
            ScopeAllowed = (sScope = "Private") Or ((lScope And enumUnusedScopePublicModuleProcs) <> 0)
        Case vbext_ct_ClassModule, vbext_ct_MSForm
            ScopeAllowed = (sScope = "Private") Or ((lScope And enumUnusedScopePublicClassMembers) <> 0)
        Case Else
            ScopeAllowed = False    'ActiveX designers and unknown types are not changed
    End Select
End Function

'Number of Sub and Function declarations named sName in the module (more than one: #If branches).
Private Function DeclarationCount(ByRef procs() As RucProc, ByVal nProcs As Long, ByVal sName As String) As Long
    Dim i As Long

    For i = 1 To nProcs
        If procs(i).Kind <> "Property" Then
            If StrComp(procs(i).Name, sName, vbTextCompare) = 0 Then DeclarationCount = DeclarationCount + 1
        End If
    Next i
End Function

'True if #If, #ElseIf, #Else and #End If within the lines lFirst to lLast form complete blocks.
Private Function DirectivesBalanced(ByRef arrLines() As String, ByVal lFirst As Long, ByVal lLast As Long) As Boolean
    Dim l As Long
    Dim lDepth As Long
    Dim mc As MatchCollection
    Dim sKind As String

    For l = lFirst To lLast
        Set mc = m_reDirective.Execute(arrLines(l))
        If mc.Count > 0 Then
            sKind = LCase$(Replace(Replace(mc(0).SubMatches(0), " ", vbNullString), vbTab, vbNullString))
            Select Case sKind
                Case "if"
                    lDepth = lDepth + 1
                Case "endif"
                    lDepth = lDepth - 1
                    If lDepth < 0 Then Exit Function
                Case Else                       'elseif, else
                    If lDepth = 0 Then Exit Function
            End Select
        End If
    Next l
    DirectivesBalanced = (lDepth = 0)
End Function

'===============================================================================
' Reading the code
'===============================================================================

'Finds all Sub, Function and Property declarations of a module. procs(1 To result) receives them
'in line order. Lines that continue a previous line (also a comment ending with " _") are no
'declaration starts.
Private Function ScanProcedures(ByRef arrLines() As String, ByVal n As Long, ByRef procs() As RucProc) As Long
    Dim l As Long
    Dim lLast As Long
    Dim nProcs As Long
    Dim sLogical As String
    Dim sScope As String
    Dim mc As MatchCollection

    l = 1
    Do While l <= n
        lLast = modObfuscation.TrfLogicalLineEnd(arrLines, l, n)
        sLogical = LogicalText(arrLines, l, lLast)
        Set mc = m_reDecl.Execute(sLogical)
        If mc.Count > 0 Then
            nProcs = nProcs + 1
            If nProcs = 1 Then
                ReDim procs(1 To 16)
            ElseIf nProcs > UBound(procs) Then
                ReDim Preserve procs(1 To 2 * UBound(procs))
            End If
            With procs(nProcs)
                .Name = mc(0).SubMatches(2)
                Select Case LCase$(Left$(mc(0).SubMatches(1), 3))
                    Case "sub": .Kind = "Sub"
                    Case "fun": .Kind = "Function"
                    Case Else: .Kind = "Property"
                End Select
                sScope = mc(0).SubMatches(0) & vbNullString
                If Len(sScope) = 0 Then .Scope = "Public" Else .Scope = StrConv(sScope, vbProperCase)
                .DeclLine = l
                .DeclText = sLogical
                .EndLine = FindEndLine(arrLines, n, lLast + 1, .Kind)
            End With
        End If
        l = lLast + 1
    Loop
    ScanProcedures = nProcs
End Function

'Line of the End statement that closes a declaration of kind sKind, searched from lFrom on.
'0 when another declaration or a different End statement comes first, or the module ends.
Private Function FindEndLine(ByRef arrLines() As String, ByVal n As Long, ByVal lFrom As Long, ByVal sKind As String) As Long
    Dim l As Long
    Dim lLast As Long
    Dim sLogical As String
    Dim mc As MatchCollection

    l = lFrom
    Do While l <= n
        lLast = modObfuscation.TrfLogicalLineEnd(arrLines, l, n)
        sLogical = LogicalText(arrLines, l, lLast)
        Set mc = m_reEnd.Execute(sLogical)
        If mc.Count > 0 Then
            If StrComp(mc(0).SubMatches(0), sKind, vbTextCompare) = 0 And lLast = l Then FindEndLine = l
            Exit Function
        End If
        If m_reDecl.Test(sLogical) Then Exit Function
        l = lLast + 1
    Loop
End Function

'Logical line lFirst to lLast as one line: line continuations removed, parts joined with a space.
Private Function LogicalText(ByRef arrLines() As String, ByVal lFirst As Long, ByVal lLast As Long) As String
    Dim l As Long
    Dim sPart As String
    Dim s As String

    For l = lFirst To lLast
        sPart = arrLines(l)
        If l < lLast Then sPart = Left$(sPart, InStrRev(sPart, "_") - 1)
        If l > lFirst Then s = s & " "
        s = s & sPart
    Next l
    LogicalText = s
End Function

'Lines lFirst to lLast joined with line breaks.
Private Function JoinLines(ByRef arrLines() As String, ByVal lFirst As Long, ByVal lLast As Long) As String
    Dim l As Long
    Dim s As String

    For l = lFirst To lLast
        If l > lFirst Then s = s & vbCrLf
        s = s & arrLines(l)
    Next l
    JoinLines = s
End Function

'Whole code of a component ("" for an empty module).
Private Function ComponentText(ByVal vbc As VBIDE.VBComponent) As String
    With vbc.CodeModule
        If .CountOfLines > 0 Then ComponentText = .Lines(1, .CountOfLines)
    End With
End Function

'Occurrences of every identifier token in the code of all components (keys case-insensitive).
Private Function ProjectWordCounts(ByVal vbProj As VBIDE.VBProject) As Scripting.Dictionary
    Dim dic As Scripting.Dictionary
    Dim vbc As VBIDE.VBComponent
    Dim m As VBScript_RegExp_55.Match
    Dim s As String

    Set dic = New Scripting.Dictionary
    dic.CompareMode = vbTextCompare
    For Each vbc In vbProj.VBComponents
        For Each m In m_reWord.Execute(ComponentText(vbc))
            s = m.Value
            If dic.Exists(s) Then dic(s) = dic(s) + 1 Else dic.Add s, 1&
        Next m
    Next vbc
    Set ProjectWordCounts = dic
End Function

'Count of sWord in the word counts (0 if it does not occur).
Private Function WordCount(ByVal dicWords As Scripting.Dictionary, ByVal sWord As String) As Long
    If dicWords.Exists(sWord) Then WordCount = dicWords(sWord)
End Function

'Whole-word, case-insensitive occurrences of sWord in sText. sWord must pass IsPlainName.
Private Function CountWord(ByVal sText As String, ByVal sWord As String) As Long
    CountWord = NewRegExp("\b" & sWord & "\b", True).Execute(sText).Count
End Function

'Names of the interfaces used in Implements statements of the project (keys case-insensitive).
Private Function ImplementedInterfaces(ByVal vbProj As VBIDE.VBProject) As Scripting.Dictionary
    Dim dic As Scripting.Dictionary
    Dim vbc As VBIDE.VBComponent
    Dim m As VBScript_RegExp_55.Match
    Dim sName As String

    Set dic = New Scripting.Dictionary
    dic.CompareMode = vbTextCompare
    For Each vbc In vbProj.VBComponents
        For Each m In m_reImpl.Execute(ComponentText(vbc))
            sName = m.SubMatches(0)
            sName = Mid$(sName, InStrRev(sName, ".") + 1)
            sName = Replace(Replace(sName, "[", vbNullString), "]", vbNullString)
            dic(sName) = True
        Next m
    Next vbc
    Set ImplementedInterfaces = dic
End Function

'True for a name of the letters A-Z and a-z and the digits 0-9 that starts with a letter.
'bAllowUnderscore also accepts "_" after the first character.
Private Function IsPlainName(ByVal sName As String, ByVal bAllowUnderscore As Boolean) As Boolean
    Dim i As Long

    If Len(sName) = 0 Then Exit Function
    For i = 1 To Len(sName)
        Select Case AscW(Mid$(sName, i, 1))
            Case 65 To 90, 97 To 122                    'A-Z, a-z
            Case 48 To 57                               '0-9
                If i = 1 Then Exit Function
            Case 95                                     '_
                If i = 1 Or Not bAllowUnderscore Then Exit Function
            Case Else
                Exit Function
        End Select
    Next i
    IsPlainName = True
End Function

'===============================================================================
' Hidden attributes (export of the component)
'===============================================================================

'Attribute information of a component, read once per run and cached in dicExport:
'element 0 = "|name|name|" (lower case) of the members with hidden attributes,
'element 1 = True if the module is exposed to other projects (VB_Exposed = True).
Private Function ExportInfo(ByVal vbc As VBIDE.VBComponent, ByVal dicExport As Scripting.Dictionary) As Variant
    If Not dicExport.Exists(vbc.Name) Then dicExport.Add vbc.Name, ReadExportInfo(vbc)
    ExportInfo = dicExport(vbc.Name)
End Function

'Reads the attribute lines of an export ("Attribute Item.VB_UserMemId = 0"). Same rules as
'modObfuscation.TrfAttributedMembers; VB_VarHelpID = -1 (written for every WithEvents variable)
'does not count.
Private Function ReadExportInfo(ByVal vbc As VBIDE.VBComponent) As Variant
    Dim sPath As String
    Dim sFrx As String
    Dim f As Integer
    Dim sLine As String
    Dim sNames As String
    Dim sName As String
    Dim p As Long
    Dim bExposed As Boolean
    Dim lErr As Long
    Dim sErr As String

    On Error GoTo EH
    sNames = "|"
    sPath = TempFolder() & "MT_Ruc_" & vbc.Name & ".txt"
    sFrx = Left$(sPath, Len(sPath) - 4) & ".frx"
    KillFile sPath
    KillFile sFrx
    vbc.Export sPath
    f = FreeFile
    Open sPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, sLine
        If Left$(sLine, 10) = "Attribute " Then
            If StrComp(sLine, "Attribute VB_Exposed = True", vbTextCompare) = 0 Then
                bExposed = True
            ElseIf InStr(1, sLine, ".VB_VarHelpID = -1") = 0 Then
                p = InStr(11, sLine, ".VB_")
                If p > 11 Then
                    sName = Mid$(sLine, 11, p - 11)
                    If Left$(sName, 1) = "[" And Right$(sName, 1) = "]" Then
                        sName = Mid$(sName, 2, Len(sName) - 2)
                    ElseIf InStr(1, sName, " ") > 0 Or InStr(1, sName, "=") > 0 Or InStr(1, sName, Chr$(34)) > 0 Then
                        sName = vbNullString    'not a member attribute
                    End If
                    If Len(sName) > 0 Then
                        If InStr(1, sNames, "|" & LCase$(sName) & "|") = 0 Then sNames = sNames & LCase$(sName) & "|"
                    End If
                End If
            End If
        End If
    Loop
    Close #f
    f = 0
CleanUp:
    On Error Resume Next
    If f <> 0 Then Close #f
    KillFile sPath
    KillFile sFrx
    On Error GoTo 0
    If lErr <> 0 Then
        Err.Raise vbObjectError + 2610, "modRemoveUnusedCode", "Export of " & vbc.Name & _
                  " for the attribute check failed (" & lErr & ": " & sErr & ")."
    End If
    ReadExportInfo = Array(sNames, bExposed)
    Exit Function
EH:
    lErr = Err.Number
    sErr = Err.Description
    Resume CleanUp
End Function

'Folder for temporary files, with a trailing path separator.
Private Function TempFolder() As String
    Dim s As String

    s = Environ$("TEMP")
    If Len(s) = 0 Then s = Environ$("TMP")
    If Len(s) = 0 Then s = CurDir$
    If Right$(s, 1) <> "\" Then s = s & "\"
    TempFolder = s
End Function

Private Sub KillFile(ByVal sPath As String)
    If Len(Dir$(sPath)) > 0 Then Kill sPath
End Sub

'===============================================================================
' Helpers
'===============================================================================

Private Sub RecordRemoval(ByVal sText As String)
    m_colRemoved.Add sText
    Debug.Print "MACROTools removed: " & sText
End Sub

Private Sub InitRegExps()
    If Not m_reWord Is Nothing Then Exit Sub
    Set m_reWord = NewRegExp("[A-Za-z0-9_]+", True)
    Set m_reDecl = NewRegExp("^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s(]+)", False)
    Set m_reEnd = NewRegExp("^\s*(?:\d+:?\s*)?End\s+(Sub|Function|Property)\b", False)
    Set m_reImpl = NewRegExp("^[ \t]*Implements[ \t]+([A-Za-z0-9_.\[\]]+)", True, True)
    Set m_reDirective = NewRegExp("^\s*#\s*(If|ElseIf|Else|End\s*If)\b", False)
    Set m_reEnum = NewRegExp("^[ \t]*(?:Public[ \t]+)?Enum[ \t]", False, True)
End Sub

Private Function NewRegExp(ByVal sPattern As String, ByVal bGlobal As Boolean, Optional ByVal bMultiLine As Boolean = False) As RegExp
    Dim re As RegExp

    Set re = New RegExp
    re.Pattern = sPattern
    re.IgnoreCase = True
    re.Global = bGlobal
    re.MultiLine = bMultiLine
    Set NewRegExp = re
End Function
