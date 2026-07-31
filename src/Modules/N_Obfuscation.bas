Attribute VB_Name = "N_Obfuscation"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : N_Obfuscation - удаление форматировани€ кода
'* Created    : 15-09-2019 15:48
'* Author     : VBATools / CalDymos
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru / Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

'”дал€ем Option Explicit изо всех модулей
Public Sub Remove_OptionExplicit(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i           As Long
    Dim strLine     As String

    For i = CurCodeModule.CountOfLines To 1 Step -1
        strLine = Trim(CurCodeModule.Lines(i, 1))
        If InStr(1, strLine, "Option Explicit") <> 0 Then
            CurCodeModule.DeleteLines i    'удал€ем всю строку
        End If
    Next i
End Sub
'”дал€ем пустые строки:
Public Sub Remove_EmptyLines(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i&, strLine$

    For i = CurCodeModule.CountOfLines To 1 Step -1
        strLine = Trim(CurCodeModule.Lines(i, 1))
        If strLine = vbCrLf Or strLine = Chr(10) Or strLine = "" Then CurCodeModule.DeleteLines i    'удал€ем пустую строку
    Next i
End Sub
'”дал€ем комментарии:
Public Sub Remove_Comments(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i           As Long
    Dim strLine     As String
    Dim pos         As Long
    Dim iCount      As Long

    Rem (!) ¬спомогательные переменные
    Dim bMultiLine  As Boolean
    Dim s           As String

    With CurCodeModule
        For i = .CountOfLines To 1 Step -1
            Rem (!) обрезаем только справа, чтобы не удал€ть отступ
            strLine = RTrim(.Lines(i, 1))
            pos = 1
try_again:
            pos = InStr(pos, strLine, Chr(39))    'позици€ апострофа
            If pos > 0 Then    'есть апостроф

                Rem (!) ≈сли в строке выше есть перенос, то переходим с обработке этой строки
                If i > 1 Then
                    s = RTrim(.Lines(i - 1, 1))    'строка выше
                    If Right(s, 2) = " _" Then GoTo next_i
                End If

                Rem (!) ≈сли справа строки есть перенос, то запоминаем, что это многострочный комментарий
                If Right(RTrim(strLine), 2) = " _" Then
                    bMultiLine = True
                Else
                    bMultiLine = False
                End If

                'ѕровер€ем не в строке ли апостроф:
                'считаем сколько кавычек слева от апострофа
                iCount = CountChrInString(Left(strLine, pos - 1), """")
                'этот апостроф в строке, значит он не метка комментари€
                If iCount Mod 2 = 1 Then pos = pos + 1: GoTo try_again
                strLine = RTrim(Left(strLine, pos - 1))
                .ReplaceLine i, strLine
                '(!) запоминаем строку
                s = strLine
                Rem (!) ≈сли многострочный коментарий
                If bMultiLine Then
                    Do
                        .DeleteLines i
                        strLine = Trim(.Lines(i, 1))
                    Loop While Right(strLine, 2) = " _"
                    'последнюю строку замен€ем на ту, что запомнили
                    .ReplaceLine i, s
                End If
                If Trim(s) = "" Then .DeleteLines i
            End If
next_i:
        Next i
    End With
End Sub

'Function to count occurrences of Char in string sSTR:
Private Function CountChrInString(sSTR As String, Char As String) As Long
    Dim iResult     As Long
    Dim sParts()    As String

    sParts = Split(sSTR, Char)
    iResult = UBound(sParts, 1): If (iResult = -1) Then iResult = 0
    CountChrInString = iResult
End Function

'Removes leading/trailing spaces and tabs from lines (may break formatting of multi-line strings):
Public Sub TrimLinesTabAndSpase(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i           As Long
    Dim strLine     As String
    Dim strLine2    As String

    For i = CurCodeModule.CountOfLines To 1 Step -1
        strLine = CurCodeModule.Lines(i, 1)
        strLine2 = Trim(strLine)
        If strLine <> strLine2 Then
            On Error Resume Next
            CurCodeModule.ReplaceLine i, strLine2
            On Error GoTo 0
        End If
    Next i
End Sub

'Merges multi-line statements (line continuations) back into a single line
Public Sub RemoveBreaksLineInCode(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim strVar      As String
    With CurCodeModule
        If .CountOfLines = 0 Then Exit Sub
        strVar = .Lines(1, .CountOfLines)
        strVar = Replace(strVar, " _" & vbNewLine, " ")
        .DeleteLines startLine:=1, Count:=.CountOfLines
        .InsertLines Line:=1, String:=strVar
    End With
End Sub

'Removes Debug.Print statements from the code.
'MatchAnywhereInLine:
'   True  - matches "Debug.Print" anywhere in the statement (e.g. "If x Then Debug.Print ...")
'   False - matches only statements that start with "Debug.Print" (after trimming)
'
'Behaviour:
'   - If a physical line contains a single statement, the whole line is removed.
'   - If a physical line contains several statements separated by ":" (colons inside
'     string literals are ignored), only the statement(s) containing Debug.Print are
'     removed; the remaining statements are kept and rejoined with ":".
'   - Multi-line statements (continued with " _") are removed as a whole block if they
'     contain Debug.Print; they are NOT split by ":" (see caveat below).
'
'(!) Caveat: for a single-line "If x Then Debug.Print y: z = 2" construct, "z = 2" is
'    part of the conditional Then-block, not an independent statement. Splitting on ":"
'    would incorrectly make "z = 2" unconditional. This edge case is not detected -
'    review the result if such constructs are expected in the target code.
Public Sub Remove_DebugPrint(ByRef CurCodeModule As VBIDE.CodeModule, Optional ByVal MatchAnywhereInLine As Boolean = True)
    Dim i           As Long
    Dim strLine     As String
    Dim startLine   As Long
    Dim endLine     As Long
    Dim prevLine    As String
    Dim fullBlock   As String
    Dim segments()  As String
    Dim keepParts() As String
    Dim nKeep       As Long
    Dim j           As Long
    Dim bLineHasMatch As Boolean
    Dim newLine     As String

    With CurCodeModule
        For i = .CountOfLines To 1 Step -1
            'skip lines that are themselves a continuation of the previous line;
            'they are handled together with their statement's first line below
            If i > 1 Then
                prevLine = RTrim(.Lines(i - 1, 1))
                If Right(prevLine, 2) = " _" Then GoTo next_i
            End If

            strLine = .Lines(i, 1)

            'find the end of this (possibly multi-line) statement
            endLine = i
            Do While Right(RTrim(.Lines(endLine, 1)), 2) = " _" And endLine < .CountOfLines
                endLine = endLine + 1
            Loop

            If endLine > i Then
                'multi-line statement: remove the whole block if it contains Debug.Print
                fullBlock = .Lines(i, endLine - i + 1)
                If ContainsDebugPrint(fullBlock, MatchAnywhereInLine) Then
                    .DeleteLines startLine:=i, Count:=endLine - i + 1
                End If
            Else
                'single physical line: split into ":"-separated statements
                segments = SplitStatements(strLine)
                ReDim keepParts(0 To UBound(segments))
                nKeep = -1
                bLineHasMatch = False
                For j = 0 To UBound(segments)
                    If ContainsDebugPrint(segments(j), MatchAnywhereInLine) Then
                        bLineHasMatch = True
                    Else
                        nKeep = nKeep + 1
                        keepParts(nKeep) = segments(j)
                    End If
                Next j

                If bLineHasMatch Then
                    If nKeep = -1 Then
                        'no statements left - remove the whole line
                        .DeleteLines i
                    Else
                        newLine = keepParts(0)
                        For j = 1 To nKeep
                            newLine = newLine & ":" & keepParts(j)
                        Next j
                        .ReplaceLine i, Trim(newLine)
                    End If
                End If
            End If
next_i:
        Next i
    End With
End Sub
'Checks whether a single statement (segment) contains a Debug.Print call
Private Function ContainsDebugPrint(ByVal strSegment As String, ByVal MatchAnywhereInLine As Boolean) As Boolean
    Dim s As String
    s = Trim(strSegment)
    If MatchAnywhereInLine Then
        ContainsDebugPrint = (InStr(1, s, "Debug.Print", vbTextCompare) > 0)
    Else
        ContainsDebugPrint = (InStr(1, s, "Debug.Print", vbTextCompare) = 1)
    End If
End Function
'Splits a line into ":"-separated statements, ignoring colons inside string literals
Private Function SplitStatements(ByVal strLine As String) As String()
    Dim result()    As String
    Dim segStart    As Long
    Dim pos         As Long
    Dim inString    As Boolean
    Dim n           As Long
    Dim ch          As String

    ReDim result(0 To 0)
    n = 0
    segStart = 1
    inString = False

    For pos = 1 To Len(strLine)
        ch = Mid(strLine, pos, 1)
        If ch = """" Then
            inString = Not inString
        ElseIf ch = ":" And Not inString Then
            ReDim Preserve result(0 To n)
            result(n) = Mid(strLine, segStart, pos - segStart)
            n = n + 1
            segStart = pos + 1
        End If
    Next pos
    ReDim Preserve result(0 To n)
    result(n) = Mid(strLine, segStart, Len(strLine) - segStart + 1)

    SplitStatements = result
End Function
