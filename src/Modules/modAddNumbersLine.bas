Attribute VB_Name = "modAddNumbersLine"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : K_AddNumbersLine - Creating VBA code line numbering
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* Modified   : Date and Time       Author              Description
'* Updated    : 05-10-2026          CalDymos              RemoveLineNumbers replaced: referenced numbers kept,
'*                                                      all numbers kept in procedures with Erl (F2 b),
'*                                                      '10 Name:' -> label at column 1, '10: Name:' ->
'*                                                      Call Name (both verified in Excel), notices
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit
Public Enum vbLineNumbers_LabelTypes
    vbLabelColon    ' 0
    vbLabelTab    ' 1
End Enum
Private Enum vbLineNumbers_ScopeToAddLineNumbersTo
    vbScopeAllProc    ' 1
    vbScopeThisProc    ' 2
End Enum
    Public Sub AddLineNumbers_()
20:    On Error GoTo ErrorHandler
21:    Dim cmb_txt As String
22:    Dim vbComp As VBIDE.VBComponent
23:    cmb_txt = modCreateMenus.WhatIsTextInComboBoxHave
24:    Select Case cmb_txt
        Case modConst.ALLVBAPROJECT:
26:            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
27:                AddLineNumbers vbCompObj:=vbComp, LabelType:=vbLabelColon, AddLineNumbersToEmptyLines:=True, AddLineNumbersToEndOfProc:=True, Scope:=vbScopeAllProc
28:            Next vbComp
29:        Case modConst.SELECTEDMODULE:
30:            AddLineNumbers vbCompObj:=Application.VBE.ActiveCodePane.CodeModule.Parent, LabelType:=vbLabelColon, AddLineNumbersToEmptyLines:=True, AddLineNumbersToEndOfProc:=True, Scope:=vbScopeAllProc
31:    End Select
32:    Exit Sub
ErrorHandler:
34:    Select Case Err.Number
        Case 91:
36:            Exit Sub
37:        Case Else:
38:            Debug.Print "Mistake! in AddLineNumbers_" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
39:            Call WriteErrorLog("AddLineNumbers_")
40:    End Select
41:    Err.Clear
42: End Sub
    Public Sub RemoveLineNumbers_()
44:    On Error GoTo ErrorHandler
45:    Dim cmb_txt As String
46:    Dim vbComp As VBIDE.VBComponent
47:    cmb_txt = modCreateMenus.WhatIsTextInComboBoxHave
48:    Select Case cmb_txt
        Case modConst.ALLVBAPROJECT:
50:            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
51:                RemoveLineNumbers vbCompObj:=vbComp, LabelType:=vbLabelColon
52:                RemoveLineNumbers vbCompObj:=vbComp, LabelType:=vbLabelTab
53:            Next vbComp
54:        Case modConst.SELECTEDMODULE:
55:            RemoveLineNumbers vbCompObj:=Application.VBE.ActiveCodePane.CodeModule.Parent, LabelType:=vbLabelColon
56:            RemoveLineNumbers vbCompObj:=Application.VBE.ActiveCodePane.CodeModule.Parent, LabelType:=vbLabelTab
57:    End Select
58:    Exit Sub
ErrorHandler:
60:    Select Case Err.Number
        Case 91:
62:            Exit Sub
63:        Case Else:
64:            Debug.Print "Mistake! in RemoveLineNumbers_" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
65:            Call WriteErrorLog("RemoveLineNumbers_")
66:    End Select
67:    Err.Clear
68: End Sub
     Private Sub AddLineNumbers( _
                  ByVal vbCompObj As VBIDE.VBComponent, _
                  ByVal LabelType As vbLineNumbers_LabelTypes, _
                  ByVal AddLineNumbersToEmptyLines As Boolean, _
                  ByVal AddLineNumbersToEndOfProc As Boolean, _
                  ByVal Scope As vbLineNumbers_ScopeToAddLineNumbersTo)
75:    ' USAGE RULES
76:    ' DO NOT MIX LABEL TYPES FOR LINE NUMBERS! IF ADDING LINE NUMBERS AS COLON TYPE, ANY LINE NUMBERS AS VBTAB TYPE MUST BE REMOVE BEFORE, AND RECIPROCALLY ADDING LINE NUMBERS AS VBTAB TYPE
77:    Dim i      As Long
78:    Dim ProcName As String
79:    Dim startOfProcedure As Long
80:    Dim lengthOfProcedure As Long
81:    Dim endOfProcedure As Long
82:    Dim bodyOfProcedure As Long
83:    Dim countOfProcedure As Long
84:    Dim prelinesOfProcedure As Long
85:    Dim PreviousIndentAdded As Long
86:    Dim strLine As String
87:    Dim temp_strLine As String
88:    Dim new_strLine As String
89:    Dim tupe_procedure As vbext_ProcKind
90:    Dim InProcBodyLines As Boolean
91:    Dim FlagSelect As Boolean
92:    With vbCompObj.CodeModule
93:
94:        If Scope = vbScopeAllProc Then
95:            For i = 1 To .CountOfLines - 1
96:                strLine = .Lines(i, 1)
97:                If FlagSelect Then
98:                    FlagSelect = False
99:                    GoTo NextLine
100:                End If
101:                If strLine Like "*Select Case *" Then FlagSelect = True
                ProcName = .ProcOfLine(i, tupe_procedure)    ' Type d'argument ByRef incompatible ~~> Requires VBIDE library as a Reference for the VBA Project
103:                If ProcName <> vbNullString Then
104:                    startOfProcedure = .ProcStartLine(ProcName, tupe_procedure)
105:                    bodyOfProcedure = .ProcBodyLine(ProcName, tupe_procedure)
106:                    countOfProcedure = .ProcCountLines(ProcName, tupe_procedure)
107:                    prelinesOfProcedure = bodyOfProcedure - startOfProcedure
108:                    'postlineOfProcedure = ??? not directly available since endOfProcedure is itself not directly available.
109:                    lengthOfProcedure = countOfProcedure - prelinesOfProcedure    ' includes postlinesOfProcedure !
110:                    'endOfProcedure = ??? not directly available, each line of the proc must be tested until the End statement is reached. See below.
111:                    If endOfProcedure <> 0 And startOfProcedure < endOfProcedure And i > endOfProcedure Then
112:                        GoTo NextLine
113:                    End If
114:                    If i = bodyOfProcedure Then InProcBodyLines = True
115:                    If bodyOfProcedure < i And i < startOfProcedure + countOfProcedure Then
116:                        If Not (.Lines(i - 1, 1) Like "* _") Then
117:                            InProcBodyLines = False
118:                            PreviousIndentAdded = 0
119:                            If Trim$(strLine) = vbNullString And Not AddLineNumbersToEmptyLines Then GoTo NextLine
120:                            If IsProcEndLine(vbCompObj, i) Then
121:                                endOfProcedure = i
122:                                If AddLineNumbersToEndOfProc Then
123:                                    Call IndentProcBodyLinesAsProcEndLine(vbCompObj, LabelType, endOfProcedure, tupe_procedure)
124:                                Else
125:                                    GoTo NextLine
126:                                End If
127:                            End If
128:                            If LabelType = vbLabelColon Then
129:                                If HasLabel(strLine, vbLabelColon) Then strLine = RemoveOneLineNumber(.Lines(i, 1), vbLabelColon)
130:                                If Not HasLabel(strLine, vbLabelColon) Then
131:                                    temp_strLine = strLine
132:                                    On Error Resume Next
133:                                    .ReplaceLine i, CStr(i) & ":" & strLine
134:                                    On Error GoTo 0
135:                                    new_strLine = .Lines(i, 1)
136:                                    If Len(new_strLine) = Len(CStr(i) & ":" & temp_strLine) Then
137:                                        PreviousIndentAdded = Len(CStr(i) & ":")
138:                                    Else
139:                                        PreviousIndentAdded = Len(CStr(i) & ": ")
140:                                    End If
141:                                End If
142:                            ElseIf LabelType = vbLabelTab Then
143:                                If Not HasLabel(strLine, vbLabelTab) Then strLine = RemoveOneLineNumber(.Lines(i, 1), vbLabelTab)
144:                                If Not HasLabel(strLine, vbLabelColon) Then
145:                                    temp_strLine = strLine
146:                                    On Error Resume Next
147:                                    .ReplaceLine i, CStr(i) & vbTab & strLine
148:                                    On Error GoTo 0
149:                                    PreviousIndentAdded = Len(strLine) - Len(temp_strLine)
150:                                End If
151:                            End If
152:                        Else
153:                            If Not InProcBodyLines Then
154:                                If LabelType = vbLabelColon Then
155:                                    On Error Resume Next
156:                                    .ReplaceLine i, Space(PreviousIndentAdded) & strLine
157:                                    On Error GoTo 0
158:                                ElseIf LabelType = vbLabelTab Then
159:                                    On Error Resume Next
160:                                    .ReplaceLine i, Space(4) & strLine
161:                                    On Error GoTo 0
162:                                End If
163:                            Else
164:                            End If
165:                        End If
166:                    End If
167:                End If
NextLine:
169:            Next i
170:        ElseIf AddLineNumbersToEmptyLines And Scope = vbScopeThisProc Then
171:            'TODO selected prosedure
172:        End If
173:
174:    End With
175: End Sub
     Private Function IsProcEndLine( _
                  ByVal vbCompObj As VBIDE.VBComponent, _
                  ByVal lLine As Long) As Boolean
179:    With vbCompObj.CodeModule
180:        If Trim$(.Lines(lLine, 1)) Like "End Sub*" _
                        Or Trim$(.Lines(lLine, 1)) Like "End Function*" _
                        Or Trim$(.Lines(lLine, 1)) Like "End Property*" _
                        Then IsProcEndLine = True
184:    End With
185: End Function
     Private Sub IndentProcBodyLinesAsProcEndLine( _
                  ByVal vbCompObj As VBIDE.VBComponent, _
                  ByVal LabelType As vbLineNumbers_LabelTypes, _
                  ByVal ProcEndLine As Long, _
                  ByVal VBEXT As vbext_ProcKind)
191:    Dim ProcName As String
192:    Dim bodyOfProcedure As Long
193:    Dim j      As Long
194:    Dim endOfProcedure As Long
195:    Dim strEnd As String
196:    Dim strLine As String
197:    With vbCompObj.CodeModule
198:        ProcName = .ProcOfLine(ProcEndLine, VBEXT)
199:        bodyOfProcedure = .ProcBodyLine(ProcName, VBEXT)
200:        endOfProcedure = ProcEndLine
201:        strEnd = .Lines(endOfProcedure, 1)
202:        j = bodyOfProcedure
203:        If j = 1 Then j = 2
204:        Do Until Not .Lines(j - 1, 1) Like "* _" And j <> bodyOfProcedure
205:            strLine = .Lines(j, 1)
206:            If LabelType = vbLabelColon Then
207:                If Mid$(strEnd, Len(CStr(endOfProcedure)) + 1 + 1 + 1, 1) = " " Then
208:                    On Error Resume Next
209:                    .ReplaceLine j, Space(Len(CStr(endOfProcedure)) + 1) & strLine
210:                    On Error GoTo 0
211:                Else
212:                    On Error Resume Next
213:                    .ReplaceLine j, Space(Len(CStr(endOfProcedure)) + 2) & strLine
214:                    On Error GoTo 0
215:                End If
216:            ElseIf LabelType = vbLabelTab Then
217:                If endOfProcedure < 1000 Then
218:                    On Error Resume Next
219:                    .ReplaceLine j, Space(4) & strLine
220:                    On Error GoTo 0
221:                Else
222:                    Debug.Print "This tool is limited to 999 lines of code to work properly."
223:                End If
224:            End If
225:            j = j + 1
226:        Loop
227:    End With
228: End Sub
'===============================================================================
' Procedure : RemoveLineNumbers
' Purpose   : Removes numeric line number labels that no statement refers to.
'
' Parameters:
'   vbCompObj - VBIDE.VBComponent, ByVal, component whose code is changed
'   LabelType - vbLineNumbers_LabelTypes, ByVal:
'               vbLabelColon (0): labels written as "10:" ("10: x = 1",
'                                 also "10 : x = 1")
'               vbLabelTab   (1): labels without colon ("10 x = 1",
'                                 "10<Tab>x = 1", "10        x = 1")
'               other values: no change
'
' Returns:
'   Keiner
'
' Side Effects:
'   Replaces lines of the code module (no line is inserted or deleted):
'   - vbLabelColon: "10:" is removed; if exactly one space follows, it is
'     removed too (inverse of AddLineNumbers).
'   - vbLabelTab: a number followed by a tab is removed with the tab
'     (inverse of AddLineNumbers with vbLabelTab); otherwise the digits are
'     replaced by spaces, so the code keeps its column.
'   - A line that consists only of the removed label becomes empty.
'   - Leading whitespace of procedure header lines is removed (undoes the
'     header indentation of AddLineNumbers; also done when no number exists).
'   - "10 Name: ..." becomes "Name: ..." at column 1; "10: Name: ..." becomes
'     "Call Name: ..." (see Postconditions).
'   Writes notices for procedures with Erl (numbers kept), for inserted Call
'   and for numbers kept before a keyword statement.
'
' Preconditions:
'   Module code is syntactically valid VBA. Access to the VBA project object
'   model is trusted.
'
' Postconditions:
'   A line number stays unchanged if a statement of the same procedure
'   refers to it via GoTo, GoSub, On ... GoTo/GoSub (list), On Error GoTo,
'   Resume, or the implicit GoTo of "If x Then 10" / "Else 10".
'   Decision F2 b: in a procedure that uses Erl, every line number stays,
'   so Erl keeps returning the same values.
'   All other numeric labels of the given type are removed. A lone name
'   followed by ":" right after the number keeps its meaning (verified by
'   execution in Excel 2019):
'   - without colon ("10 Init:  Run"), Init is a second LABEL (test X35g):
'     the number and the blanks after it are removed, so "Init:" stays a
'     label at the start of the line;
'   - with colon ("10: Init: Run"), Init is a CALL (test X35h): without the
'     number it would become a label, so "Call " is put in front of it. If
'     that line would exceed TRF_MAX_LINE_LENGTH, the number stays.
'   Before a keyword statement ("10 Stop: x = 1") the number stays.
'   Alphanumeric labels, numbers in strings, comments, dates and
'   expressions are unchanged.
'
' Errors:
'   No local handler. Errors of the VBIDE object model are passed to the
'   caller (the former On Error Resume Next hid failed replacements).
'
' Notes:
'   Labels are scoped per procedure; procedures are delimited by the lines
'   End Sub / End Function / End Property. References in inactive #If
'   branches also protect a number (conservative). obj.GoTo / obj.Resume are
'   member calls, not references.
'   Erl returns the last numbered line executed before the error; code like
'   "If Erl = 20 Then Resume 30" depends on it. Therefore numbers are kept
'   in the whole procedure as soon as Erl appears in it (F2 b, replaces the
'   former decision F2 a). Erl as a member name (obj.Erl) does not count.
'   Numbers of any length are recognised (formerly 1 to 4 digits only).
'===============================================================================
Public Sub RemoveLineNumbers(ByVal vbCompObj As VBIDE.VBComponent, ByVal LabelType As vbLineNumbers_LabelTypes)
    Dim arrLines() As String
    Dim arrNew() As String
    Dim abKeep() As Boolean
    Dim n As Long

    n = modObfuscation.TrfReadLines(vbCompObj.CodeModule, arrLines)
    If n = 0 Then Exit Sub
    RemoveLineNumberEdits arrLines, n, LabelType, vbCompObj.Name, arrNew, abKeep
    modObfuscation.TrfApplyLineEdits vbCompObj.CodeModule, arrLines, arrNew, abKeep, n, "RemoveLineNumbers"
End Sub

'Computes the edits of RemoveLineNumbers (pure function, no module access).
Private Sub RemoveLineNumberEdits(ByRef arrLines() As String, ByVal n As Long, ByVal LabelType As vbLineNumbers_LabelTypes, _
                                  ByVal sModule As String, ByRef arrNew() As String, ByRef abKeep() As Boolean)
    Dim i As Long
    Dim l As Long
    Dim j As Long
    Dim toks() As TrfToken
    Dim nTok As Long
    Dim info As TrfLineInfo
    Dim sRefs As String
    Dim bErl As Boolean
    Dim bEnd As Boolean
    Dim nLabels As Long
    Dim nKeptErl As Long
    Dim alLine() As Long
    Dim asNum() As String
    Dim alNumPos() As Long
    Dim alNumLen() As Long
    Dim alColonPos() As Long
    Dim abGuard() As Boolean
    Dim asGuardWord() As String
    Dim sNew As String
    Dim t As Long
    Dim lPos As Long

    ReDim arrNew(1 To n)
    ReDim abKeep(1 To n)
    ReDim alLine(1 To n)
    ReDim asNum(1 To n)
    ReDim alNumPos(1 To n)
    ReDim alNumLen(1 To n)
    ReDim alColonPos(1 To n)
    ReDim abGuard(1 To n)
    ReDim asGuardWord(1 To n)
    For i = 1 To n
        arrNew(i) = arrLines(i)
        abKeep(i) = True
    Next i

    sRefs = "|"
    i = 1
    Do While i <= n
        l = modObfuscation.TrfLogicalLineEnd(arrLines, i, n)
        modObfuscation.TrfScanLogicalLine arrLines, i, l, toks, nTok, info
        bEnd = False
        If nTok > 0 And Not info.IsDirective Then
            'numeric label on the first physical line of the logical line
            If toks(0).Kind = TRF_TK_LABEL And IsDigitText(toks(0).Text) Then
                nLabels = nLabels + 1
                alLine(nLabels) = i
                asNum(nLabels) = toks(0).Text
                alNumPos(nLabels) = toks(0).StartPos
                alNumLen(nLabels) = toks(0).Length
                alColonPos(nLabels) = 0
                If nTok > 1 Then
                    If toks(1).Kind = TRF_TK_LABEL And toks(1).Text = ":" Then alColonPos(nLabels) = toks(1).StartPos
                End If
                'a name followed by ":" right after the number would become a label ("10 Init: Run")
                t = 1
                If alColonPos(nLabels) > 0 Then t = 2
                abGuard(nLabels) = False
                asGuardWord(nLabels) = ""
                If t + 1 < nTok Then
                    abGuard(nLabels) = (toks(t).Kind = TRF_TK_WORD And toks(t + 1).Kind = TRF_TK_SEPARATOR)
                    If abGuard(nLabels) Then asGuardWord(nLabels) = toks(t).Text
                End If
            End If
            CollectLineReferences toks, nTok, sRefs, bErl
            If IsProcHeaderTokens(toks, nTok) Then arrNew(i) = modObfuscation.TrfLTrimWS(arrLines(i))
            bEnd = IsProcEndTokens(toks, nTok)
        End If

        If bEnd Or l = n Then
            'end of a procedure (or of the module): labels of this scope are known
            nKeptErl = 0
            For j = 1 To nLabels
                If InStr(1, sRefs, "|" & NormalizeLineNumber(asNum(j)) & "|") = 0 Then
                    sNew = LineWithoutNumber(arrLines(alLine(j)), alNumPos(j), alNumLen(j), alColonPos(j), LabelType)
                    If sNew <> arrLines(alLine(j)) Then
                        If bErl Then
                            'decision F2 b: Erl reads the numbers, all of them stay
                            nKeptErl = nKeptErl + 1
                        ElseIf abGuard(j) Then
                            If modObfuscation.TrfIsStandaloneKeyword(asGuardWord(j)) Then
                                modObfuscation.TrfAddNotice "RemoveLineNumbers", sModule, alLine(j), _
                                    "line number kept: keyword statement " & asGuardWord(j) & " follows"
                            ElseIf alColonPos(j) > 0 Then
                                '"10: Name: ..." calls Name (test X35h); without the number it would be a label
                                lPos = Len(sNew) - Len(modObfuscation.TrfLTrimWS(sNew)) + 1
                                sNew = Left$(sNew, lPos - 1) & "Call " & Mid$(sNew, lPos)
                                If Len(sNew) > TRF_MAX_LINE_LENGTH Then
                                    modObfuscation.TrfAddNotice "RemoveLineNumbers", sModule, alLine(j), _
                                        "line number kept: line with Call would exceed " & TRF_MAX_LINE_LENGTH & " characters"
                                Else
                                    arrNew(alLine(j)) = sNew
                                    modObfuscation.TrfAddNotice "RemoveLineNumbers", sModule, alLine(j), _
                                        "Call inserted before " & asGuardWord(j) & ", otherwise it would become a line label"
                                End If
                            Else
                                '"10 Name: ..." is number plus label Name (test X35g): Name stays a label at column 1
                                arrNew(alLine(j)) = modObfuscation.TrfLTrimWS(Mid$(arrLines(alLine(j)), alNumPos(j) + alNumLen(j)))
                            End If
                        Else
                            arrNew(alLine(j)) = sNew
                        End If
                    End If
                End If
            Next j
            If nKeptErl > 0 Then
                modObfuscation.TrfAddNotice "RemoveLineNumbers", sModule, alLine(1), _
                    "procedure uses Erl; " & nKeptErl & " line number(s) kept (decision F2 b)"
            End If
            nLabels = 0
            sRefs = "|"
            bErl = False
        End If
        i = l + 1
    Loop
End Sub

'Adds the line numbers that statements of one logical line refer to.
Private Sub CollectLineReferences(ByRef toks() As TrfToken, ByVal nTok As Long, ByRef sRefs As String, ByRef bErl As Boolean)
    Dim t As Long
    Dim j As Long

    For t = 0 To nTok - 1
        If toks(t).Kind = TRF_TK_WORD Then
            If Not IsMemberName(toks, t) Then
                Select Case LCase$(toks(t).Text)
                    Case "goto", "gosub"
                        'GoTo 10 / On x GoTo 10, 20, 30
                        j = t + 1
                        Do While j < nTok
                            If toks(j).Kind <> TRF_TK_NUMBER Then Exit Do
                            AddLineReference sRefs, toks(j).Text
                            If j + 2 > nTok - 1 Then Exit Do
                            If toks(j + 1).Text <> "," Then Exit Do
                            j = j + 2
                        Loop
                    Case "resume", "then", "else"
                        'Resume 10 / If x Then 10 / Else 10
                        If t + 1 < nTok Then
                            If toks(t + 1).Kind = TRF_TK_NUMBER Then AddLineReference sRefs, toks(t + 1).Text
                        End If
                    Case "erl"
                        bErl = True
                End Select
            End If
        End If
    Next t
End Sub

'True if token t follows "." or "!" (member access such as Application.Goto).
Private Function IsMemberName(ByRef toks() As TrfToken, ByVal t As Long) As Boolean
    If t = 0 Then Exit Function
    If toks(t - 1).Kind = TRF_TK_PUNCT Then IsMemberName = (toks(t - 1).Text = "." Or toks(t - 1).Text = "!")
End Function

Private Sub AddLineReference(ByRef sRefs As String, ByVal sNumber As String)
    Dim sKey As String

    sKey = "|" & NormalizeLineNumber(sNumber) & "|"
    If InStr(1, sRefs, sKey) = 0 Then sRefs = sRefs & Mid$(sKey, 2)
End Sub

'Line numbers compare as numbers: "010" and "10" are the same label.
Private Function NormalizeLineNumber(ByVal sNumber As String) As String
    Do While Len(sNumber) > 1 And Left$(sNumber, 1) = "0"
        sNumber = Mid$(sNumber, 2)
    Loop
    NormalizeLineNumber = sNumber
End Function

Private Function IsDigitText(ByVal sText As String) As Boolean
    Dim p As Long

    If Len(sText) = 0 Then Exit Function
    For p = 1 To Len(sText)
        If AscW(Mid$(sText, p, 1)) < 48 Or AscW(Mid$(sText, p, 1)) > 57 Then Exit Function
    Next p
    IsDigitText = True
End Function

'True for [Public|Private|Friend] [Static] Sub|Function|Property Get|Let|Set (not Declare).
Private Function IsProcHeaderTokens(ByRef toks() As TrfToken, ByVal nTok As Long) As Boolean
    IsProcHeaderTokens = (Len(modObfuscation.TrfProcHeaderName(toks, nTok)) > 0)
End Function

'True for "End Sub", "End Function", "End Property" (a label before it is allowed).
Private Function IsProcEndTokens(ByRef toks() As TrfToken, ByVal nTok As Long) As Boolean
    Dim t As Long

    Do While t < nTok
        If toks(t).Kind <> TRF_TK_LABEL Then Exit Do
        t = t + 1
    Loop
    If nTok - t <> 2 Then Exit Function
    If toks(t).Kind <> TRF_TK_WORD Or toks(t + 1).Kind <> TRF_TK_WORD Then Exit Function
    If LCase$(toks(t).Text) <> "end" Then Exit Function
    Select Case LCase$(toks(t + 1).Text)
        Case "sub", "function", "property"
            IsProcEndTokens = True
    End Select
End Function

'Returns the line without its line number label of the given type (unchanged if the type does not match).
Private Function LineWithoutNumber(ByVal sLine As String, ByVal lNumPos As Long, ByVal lNumLen As Long, _
                                   ByVal lColonPos As Long, ByVal LabelType As vbLineNumbers_LabelTypes) As String
    Dim sLead As String
    Dim sRest As String

    LineWithoutNumber = sLine
    sLead = Left$(sLine, lNumPos - 1)
    If LabelType = vbLabelColon Then
        If lColonPos = 0 Then Exit Function
        sRest = Mid$(sLine, lColonPos + 1)
        If Len(sRest) >= 2 Then
            If Left$(sRest, 1) = " " And Mid$(sRest, 2, 1) <> " " Then sRest = Mid$(sRest, 2)
        End If
    ElseIf LabelType = vbLabelTab Then
        If lColonPos > 0 Then Exit Function
        sRest = Mid$(sLine, lNumPos + lNumLen)
        If Left$(sRest, 1) = vbTab Then
            sRest = Mid$(sRest, 2)
        Else
            sRest = Space$(lNumLen) & sRest
        End If
    Else
        Exit Function
    End If
    If modObfuscation.TrfIsBlank(sLead & sRest) Then
        LineWithoutNumber = ""
    Else
        LineWithoutNumber = sLead & sRest
    End If
End Function
     Private Function RemoveOneLineNumber(ByVal aString As String, ByVal LabelType As vbLineNumbers_LabelTypes) As Variant
269:    RemoveOneLineNumber = aString
270:    If LabelType = vbLabelColon Then
271:        If aString Like "#:*" Or aString Like "##:*" Or aString Like "###:*" Or aString Like "####:*" Then
272:            RemoveOneLineNumber = Mid$(aString, 1 + InStr(1, aString, ":", vbTextCompare))
273:            If Left$(RemoveOneLineNumber, 2) Like " [! ]*" Then RemoveOneLineNumber = Mid$(RemoveOneLineNumber, 2)
274:        End If
275:    ElseIf LabelType = vbLabelTab Then
276:        If aString Like "#   *" Or aString Like "##  *" Or aString Like "### *" Or aString Like "#### *" Then RemoveOneLineNumber = Mid$(aString, 5)
277:        If aString Like "#" Or aString Like "##" Or aString Like "###" Or aString Like "####" Then RemoveOneLineNumber = vbNullString
278:    End If
211:     If RemoveOneLineNumber Like "*Function *" Or RemoveOneLineNumber Like "*Sub *" _
            Or RemoveOneLineNumber Like "*Property Set *" Or RemoveOneLineNumber Like "*Property Get *" Or RemoveOneLineNumber Like "*Property Let *" Then
281:        RemoveOneLineNumber = RemoveLeadingSpaces(RemoveOneLineNumber)
282:    End If
283: End Function
     Private Function HasLabel(ByVal aString As String, ByVal LabelType As vbLineNumbers_LabelTypes) As Boolean
285:    If LabelType = vbLabelColon Then HasLabel = InStr(1, aString & ":", ":") < InStr(1, aString & " ", " ")
286:    If LabelType = vbLabelTab Then
287:        HasLabel = Mid$(aString, 1, 4) Like "#   " Or Mid$(aString, 1, 4) Like "##  " Or Mid$(aString, 1, 4) Like "### " Or Mid$(aString, 1, 5) Like "#### "
288:    End If
289: End Function
'Removes all spaces at the beginning of the line
Private Function RemoveLeadingSpaces(ByVal aString As String) As String
292:    Do Until Left$(aString, 1) <> " "
293:        aString = Mid$(aString, 2)
294:    Loop
295:    RemoveLeadingSpaces = aString
End Function
