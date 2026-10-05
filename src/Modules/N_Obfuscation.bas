Attribute VB_Name = "N_Obfuscation"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : N_Obfuscation - Removal of code formatting
'* Created    : 15-09-2019 15:48
'* Author     : VBATools / CalDymos
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru / Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 02-10-2026          CalDymos            Transformation chain corrected: shared
'*                                                      lexer, line edit helpers, notices, fixes
'*                                                      in all seven transformation methods
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

'===============================================================================
' Shared definitions of the transformation chain
'
' All transformation methods work in three steps:
'   1. TrfReadLines reads the module into an array (1-based, index = line).
'   2. A pure function computes the new text of every line and whether the
'      line is kept. No method inserts lines; lines are only replaced or
'      deleted. All analysis therefore uses the unchanged original indices.
'   3. TrfApplyLineEdits writes only the changed lines back, bottom-up, so
'      that deletions never shift the index of a line that is still pending.
'===============================================================================

'Longest physical line the VBA editor stores. A longer line is split by the
'editor after exactly 1023 characters (verified with a merged statement of
'1166 characters).
Public Const TRF_MAX_LINE_LENGTH As Long = 1023

'Token kinds produced by TrfScanLogicalLine
Public Const TRF_TK_WORD As Long = 1        'identifier or keyword
Public Const TRF_TK_NUMBER As Long = 2      'run of decimal digits
Public Const TRF_TK_STRING As Long = 3      'string literal including its quotes
Public Const TRF_TK_BRACKET As Long = 4     'foreign identifier [ ... ]
Public Const TRF_TK_DATE As Long = 5        'date literal # ... #
Public Const TRF_TK_PUNCT As Long = 6       'any other character, and ":="
Public Const TRF_TK_SEPARATOR As Long = 7   'statement separator ":"
Public Const TRF_TK_LABEL As Long = 8       'label at the start of a logical line (number or name) and its ":"

Public Type TrfToken
    LineIdx As Long             'index of the physical line in the lines array
    StartPos As Long            '1-based position within that physical line
    Length As Long
    Kind As Long
    Text As String
End Type

Public Type TrfLineInfo
    IsDirective As Boolean      'logical line is a compiler directive (#If, #Const, ...)
    HasComment As Boolean
    CommentLineIdx As Long      'physical line on which the comment starts
    CommentPos As Long          'position of the apostrophe or of Rem
    CommentIsRem As Boolean
    CutLineIdx As Long          'cut position that removes the comment (and a separator before it)
    CutPos As Long
    RemUnsafe As Boolean        'Rem follows Then/Else of a single-line If directly
End Type

Private Const TRF_STMT_DEBUGPRINT As Long = 1
Private Const TRF_STMT_OPTIONEXPLICIT As Long = 2

'Results of TrfRebuildLine
Private Const TRF_REBUILD_DONE As Long = 1
Private Const TRF_REBUILD_CALL As Long = 2          'done, "Call " inserted before a lone procedure name
Private Const TRF_REBUILD_KEYWORD As Long = 3       'refused: a keyword statement would stand before ":" at the line start
Private Const TRF_REBUILD_TOO_LONG As Long = 4      'refused: rebuilt line longer than TRF_MAX_LINE_LENGTH

'Words that may occur inside a date literal (#Jan 1, 2020#, #10:30:00 AM#)
Private Const TRF_DATE_WORDS As String = " jan feb mar apr may jun jul aug sep oct nov dec january february" & _
    " march april june july august september october november december am pm a p "

Private m_colNotices As Collection

'===============================================================================
' Procedure : Remove_OptionExplicit
' Purpose   : Removes the module statement "Option Explicit".
'
' Parameters:
'   CurCodeModule - VBIDE.CodeModule, ByRef (only read and edited, never
'                   reassigned), module to transform
'
' Returns:
'   Keiner
'
' Side Effects:
'   Deletes every logical line that consists only of the statement
'   "Option Explicit" (case-insensitive, also when split by a line
'   continuation). A trailing comment of that line is kept on its own line.
'   On a line with several statements only the Option Explicit statement is
'   removed. Writes notices via TrfAddNotice for occurrences it cannot
'   remove safely.
'
' Preconditions:
'   Module code is syntactically valid VBA. Access to the VBA project object
'   model is trusted.
'
' Postconditions:
'   No statement "Option Explicit" remains outside strings and comments.
'   Option Compare, Option Base, Option Private Module, strings, comments,
'   and all other lines are unchanged.
'
' Errors:
'   No local handler. Errors of the VBIDE object model are passed to the
'   caller.
'
' Notes:
'   Removing Option Explicit disables the compile time check for undeclared
'   variables of this module; undeclared variables are no longer reported.
'   This is the requested behaviour and is not prevented.
'===============================================================================
Public Sub Remove_OptionExplicit(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim arrLines() As String
    Dim arrNew() As String
    Dim abKeep() As Boolean
    Dim n As Long

    n = TrfReadLines(CurCodeModule, arrLines)
    If n = 0 Then Exit Sub
    TrfStatementEdits arrLines, n, TRF_STMT_OPTIONEXPLICIT, "Remove_OptionExplicit", CurCodeModule.Parent.Name, arrNew, abKeep
    TrfApplyLineEdits CurCodeModule, arrLines, arrNew, abKeep, n, "Remove_OptionExplicit"
End Sub

'===============================================================================
' Procedure : Remove_EmptyLines
' Purpose   : Deletes lines that are empty or contain only spaces and tabs.
'
' Parameters:
'   CurCodeModule - VBIDE.CodeModule, ByRef (only read and edited, never
'                   reassigned), module to transform
'
' Returns:
'   Keiner
'
' Side Effects:
'   Deletes blank physical lines. Consecutive blank lines are deleted in one
'   DeleteLines call.
'
' Preconditions:
'   Access to the VBA project object model is trusted.
'
' Postconditions:
'   Every remaining line contains a character other than space or tab,
'   except a blank line that directly follows a line continuation.
'
' Errors:
'   No local handler. Errors of the VBIDE object model are passed to the
'   caller.
'
' Notes:
'   A blank line directly after a line ending with " _" is kept: it belongs
'   to the continued statement or comment. Deleting it would join the next
'   line to that statement (for a comment: turn the next line into comment).
'===============================================================================
Public Sub Remove_EmptyLines(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim arrLines() As String
    Dim arrNew() As String
    Dim abKeep() As Boolean
    Dim n As Long
    Dim i As Long

    n = TrfReadLines(CurCodeModule, arrLines)
    If n = 0 Then Exit Sub
    TrfInitEdits arrLines, n, arrNew, abKeep
    For i = 1 To n
        If TrfIsBlank(arrLines(i)) Then
            If i = 1 Then
                abKeep(i) = False
            ElseIf Not TrfIsContinuationLine(arrLines(i - 1)) Then
                abKeep(i) = False
            End If
        End If
    Next i
    TrfApplyLineEdits CurCodeModule, arrLines, arrNew, abKeep, n, "Remove_EmptyLines"
End Sub

'===============================================================================
' Procedure : Remove_Comments
' Purpose   : Removes apostrophe comments and Rem comments.
'
' Parameters:
'   CurCodeModule - VBIDE.CodeModule, ByRef (only read and edited, never
'                   reassigned), module to transform
'
' Returns:
'   Keiner
'
' Side Effects:
'   Cuts every comment off its line, including comments continued with
'   " _" over several physical lines and comments on the last physical line
'   of a continued statement. Lines that become blank are deleted. A
'   statement separator ":" directly before the comment is removed too.
'
' Preconditions:
'   Module code is syntactically valid VBA.
'
' Postconditions:
'   No comment remains, except a Rem that follows Then/Else of a single-line
'   If directly (reported as notice). Strings, bracketed names, date literals
'   and code are unchanged.
'
' Errors:
'   No local handler. Errors of the VBIDE object model are passed to the
'   caller.
'
' Notes:
'   An apostrophe inside a string, a bracketed name or a date literal is not
'   a comment. Rem is a comment only at the start of a statement (line start,
'   after a label, after ":", after Then/Else) and when followed by a space,
'   a tab or the line end. "If x Then Rem c" is left unchanged, because the
'   effect of removing it on the If form is not proven.
'===============================================================================
Public Sub Remove_Comments(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim arrLines() As String
    Dim arrNew() As String
    Dim abKeep() As Boolean
    Dim n As Long

    n = TrfReadLines(CurCodeModule, arrLines)
    If n = 0 Then Exit Sub
    TrfCommentEdits arrLines, n, CurCodeModule.Parent.Name, arrNew, abKeep
    TrfApplyLineEdits CurCodeModule, arrLines, arrNew, abKeep, n, "Remove_Comments"
End Sub

'===============================================================================
' Procedure : TrimLinesTabAndSpase
' Purpose   : Removes leading and trailing spaces and tabs of every line.
'
' Parameters:
'   CurCodeModule - VBIDE.CodeModule, ByRef (only read and edited, never
'                   reassigned), module to transform
'
' Returns:
'   Keiner
'
' Side Effects:
'   Replaces every line whose leading or trailing whitespace changes.
'
' Preconditions:
'   Access to the VBA project object model is trusted.
'
' Postconditions:
'   No line starts or ends with a space or tab, except a line that holds
'   only a line continuation (" _"). Characters between the first and the
'   last non-whitespace character of a line are unchanged.
'
' Errors:
'   No local handler. Errors of the VBIDE object model are passed to the
'   caller (the former On Error Resume Next hid failed replacements).
'
' Notes:
'   The name keeps its historic spelling because it is public. Indentation
'   of comment continuation lines is part of the comment text and is
'   removed as well. A line continuation keeps its whitespace before "_";
'   a line consisting only of a continuation becomes " _".
'===============================================================================
Public Sub TrimLinesTabAndSpase(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim arrLines() As String
    Dim arrNew() As String
    Dim abKeep() As Boolean
    Dim n As Long
    Dim i As Long

    n = TrfReadLines(CurCodeModule, arrLines)
    If n = 0 Then Exit Sub
    TrfInitEdits arrLines, n, arrNew, abKeep
    For i = 1 To n
        arrNew(i) = TrfTrimWS(arrLines(i))
        'a line that holds only a line continuation keeps the whitespace required before "_"
        If arrNew(i) = "_" And TrfIsContinuationLine(arrLines(i)) Then arrNew(i) = " _"
    Next i
    TrfApplyLineEdits CurCodeModule, arrLines, arrNew, abKeep, n, "TrimLinesTabAndSpase"
End Sub

'===============================================================================
' Procedure : RemoveBreaksLineInCode
' Purpose   : Joins statements continued with " _" into one physical line.
'
' Parameters:
'   CurCodeModule - VBIDE.CodeModule, ByRef (only read and edited, never
'                   reassigned), module to transform
'
' Returns:
'   Keiner
'
' Side Effects:
'   Replaces the first line of each continued statement with the joined
'   text and deletes its continuation lines. Only continued statements are
'   touched; the module is no longer deleted and re-inserted as a whole.
'
' Preconditions:
'   Module code is syntactically valid VBA.
'
' Postconditions:
'   Every continued statement whose joined text fits into
'   TRF_MAX_LINE_LENGTH characters is one physical line. Longer statements
'   keep all their line continuations unchanged (decision F6 a) and are
'   reported as notice.
'
' Errors:
'   No local handler. Errors of the VBIDE object model are passed to the
'   caller. If the editor still splits a joined line, TrfApplyLineEdits
'   restores the original lines and writes a notice.
'
' Notes:
'   Parts are joined with one space (no space after "("), which keeps every
'   token boundary: a continuation always has whitespace before "_".
'   Continued comments are joined into one comment line.
'===============================================================================
Public Sub RemoveBreaksLineInCode(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim arrLines() As String
    Dim arrNew() As String
    Dim abKeep() As Boolean
    Dim n As Long

    n = TrfReadLines(CurCodeModule, arrLines)
    If n = 0 Then Exit Sub
    TrfMergeEdits arrLines, n, CurCodeModule.Parent.Name, arrNew, abKeep
    TrfApplyLineEdits CurCodeModule, arrLines, arrNew, abKeep, n, "RemoveBreaksLineInCode"
End Sub

'===============================================================================
' Procedure : Remove_DebugPrint
' Purpose   : Removes Debug.Print statements.
'
' Parameters:
'   CurCodeModule       - VBIDE.CodeModule, ByRef (only read and edited, never
'                         reassigned), module to transform
'   MatchAnywhereInLine - Boolean, ByVal, optional, default True. Kept for
'                         compatibility; it no longer changes the result,
'                         see Notes.
'
' Returns:
'   Keiner
'
' Side Effects:
'   Removes every statement that starts with Debug.Print, together with its
'   arguments (decision F3 i a: the evaluation of the arguments is dropped
'   too). Other statements of the same line are kept in their order. A line
'   that becomes empty is deleted, a comment of that line is kept.
'
' Preconditions:
'   Module code is syntactically valid VBA.
'
' Postconditions:
'   No Debug.Print statement remains, except the cases reported as notice:
'   Debug.Print inside a single-line If (decision F3 ii a), not at the start
'   of a statement, in a continued line that holds other statements too, or
'   when the removal would put a keyword statement before ":" at the line
'   start ("Debug.Print x: Stop: y = 1").
'   If the removal would turn a following procedure call into a line label,
'   "Call " is inserted: "Debug.Print x: Init: Run" becomes "Call Init: Run"
'   (released by the product owner; reported as notice).
'   Strings, comments, labels and all other statements are unchanged.
'
' Errors:
'   No local handler. Errors of the VBIDE object model are passed to the
'   caller.
'
' Notes:
'   Before the correction, True matched the text anywhere (also in strings,
'   comments and in conditions of single-line If) and False matched only
'   lines starting with Debug.Print. Both values now remove exactly the real
'   Debug.Print statements, so the parameter has no effect any more.
'   A removed Debug.Print may have called functions with side effects; that
'   loss is accepted (F3 i a).
'===============================================================================
Public Sub Remove_DebugPrint(ByRef CurCodeModule As VBIDE.CodeModule, Optional ByVal MatchAnywhereInLine As Boolean = True)
    Dim arrLines() As String
    Dim arrNew() As String
    Dim abKeep() As Boolean
    Dim n As Long

    n = TrfReadLines(CurCodeModule, arrLines)
    If n = 0 Then Exit Sub
    TrfStatementEdits arrLines, n, TRF_STMT_DEBUGPRINT, "Remove_DebugPrint", CurCodeModule.Parent.Name, arrNew, abKeep
    TrfApplyLineEdits CurCodeModule, arrLines, arrNew, abKeep, n, "Remove_DebugPrint"
End Sub

'===============================================================================
' Notices: transformations that were intentionally not carried out
'===============================================================================

'===============================================================================
' Procedure : TrfAddNotice
' Purpose   : Records a notice and prints it to the Immediate window.
'
' Parameters:
'   sMethod - String, ByVal, name of the reporting method
'   sModule - String, ByVal, name of the VBA component
'   lLine   - Long, ByVal, line number in the module before the method ran
'   sText   - String, ByVal, reason
'
' Returns:
'   Keiner
'
' Side Effects:
'   Adds an entry to the module level notice list, writes one line to the
'   Immediate window.
'
' Preconditions:
'   Keine
'
' Postconditions:
'   TrfNoticeCount is increased by one.
'
' Errors:
'   Keine
'
' Notes:
'   The list lives until TrfClearNotices is called or the project is reset.
'===============================================================================
Public Sub TrfAddNotice(ByVal sMethod As String, ByVal sModule As String, ByVal lLine As Long, ByVal sText As String)
    Dim sMsg As String

    If m_colNotices Is Nothing Then Set m_colNotices = New Collection
    sMsg = sMethod & " | " & sModule & " | line " & lLine & " | " & sText
    m_colNotices.Add sMsg
    Debug.Print "MACROTools notice: " & sMsg
End Sub

'Number of recorded notices.
Public Function TrfNoticeCount() As Long
    If Not m_colNotices Is Nothing Then TrfNoticeCount = m_colNotices.Count
End Function

'Text of notice lIndex (1-based).
Public Function TrfNoticeText(ByVal lIndex As Long) As String
    TrfNoticeText = m_colNotices.Item(lIndex)
End Function

'Deletes all recorded notices.
Public Sub TrfClearNotices()
    Set m_colNotices = Nothing
End Sub

'===============================================================================
' Reading and writing the module
'===============================================================================

'===============================================================================
' Procedure : TrfReadLines
' Purpose   : Reads all lines of a code module into a 1-based array.
'
' Parameters:
'   cm       - VBIDE.CodeModule, ByVal, module to read
'   arrLines - String(), ByRef, receives the lines; arrLines(i) = line i
'
' Returns:
'   Number of lines (0 for an empty module; arrLines is then not allocated)
'
' Side Effects:
'   Keine
'
' Preconditions:
'   Access to the VBA project object model is trusted.
'
' Postconditions:
'   arrLines(1 To result) holds the module text without line terminators.
'
' Errors:
'   Errors of the VBIDE object model are passed to the caller.
'
' Notes:
'   Reads the module with one call. If splitting that text does not give
'   exactly CountOfLines lines, every line is read separately.
'===============================================================================
Public Function TrfReadLines(ByVal cm As VBIDE.CodeModule, ByRef arrLines() As String) As Long
    Dim n As Long
    Dim i As Long
    Dim v As Variant
    Dim lParts As Long
    Dim bSplitOk As Boolean

    n = cm.CountOfLines
    If n = 0 Then Exit Function
    ReDim arrLines(1 To n)
    v = Split(cm.Lines(1, n), vbCrLf)
    lParts = UBound(v) - LBound(v) + 1
    bSplitOk = (lParts = n)
    If Not bSplitOk And lParts = n + 1 Then bSplitOk = (v(UBound(v)) = "")
    If bSplitOk Then
        For i = 1 To n
            arrLines(i) = v(LBound(v) + i - 1)
        Next i
    Else
        For i = 1 To n
            arrLines(i) = cm.Lines(i, 1)
        Next i
    End If
    TrfReadLines = n
End Function

'===============================================================================
' Procedure : TrfApplyLineEdits
' Purpose   : Writes computed line edits back to the code module.
'
' Parameters:
'   cm      - VBIDE.CodeModule, ByVal, module to edit
'   arrOld  - String(), ByRef, original lines (1 To lCount), not changed
'   arrNew  - String(), ByRef, new text of each kept line, not changed
'   abKeep  - Boolean(), ByRef, False = delete the line, not changed
'   lCount  - Long, ByVal, number of lines; must equal cm.CountOfLines
'   sMethod - String, ByVal, name used in notices
'
' Returns:
'   Keiner
'
' Side Effects:
'   DeleteLines for every run of deleted lines, ReplaceLine for every kept
'   line whose text changed. Unchanged lines are not touched.
'
' Preconditions:
'   The module was not changed since arrOld was read. No line is inserted.
'
' Postconditions:
'   The module consists of arrNew(i) for every i with abKeep(i) = True, in
'   the original order.
'
' Errors:
'   Errors of the VBIDE object model are passed to the caller. Raises an
'   error if the line count shrinks after ReplaceLine (not expected).
'
' Notes:
'   Works from the last line to the first, so a deletion never shifts the
'   index of a line that is still to be processed. A changed line directly
'   above a run of deleted lines is replaced before the run is deleted: the
'   line is then complete (for example a joined procedure header) while its
'   former continuation lines still exist, instead of absorbing the next
'   line (for example End Property) for a moment. If the editor stores a
'   replaced line as several lines (line too long), that line is restored,
'   a run below it that continues it is kept, and a notice is written.
'   The callers never produce lines longer than TRF_MAX_LINE_LENGTH, so this
'   path only protects against unknown editor behaviour.
'===============================================================================
Public Sub TrfApplyLineEdits(ByVal cm As VBIDE.CodeModule, ByRef arrOld() As String, ByRef arrNew() As String, _
                             ByRef abKeep() As Boolean, ByVal lCount As Long, ByVal sMethod As String)
    Dim i As Long
    Dim j As Long
    Dim bDelete As Boolean

    i = lCount
    Do While i >= 1
        If Not abKeep(i) Then
            'lines j..i are deleted
            j = i
            Do While j > 1
                If abKeep(j - 1) Then Exit Do
                j = j - 1
            Loop
            bDelete = True
            If j > 1 Then
                If arrNew(j - 1) <> arrOld(j - 1) Then
                    'after a restore, a run that continues line j - 1 must stay with it
                    If Not TrfReplaceLine(cm, j - 1, arrNew(j - 1), arrOld(j - 1), sMethod) Then
                        bDelete = Not TrfIsContinuationLine(arrOld(j - 1))
                    End If
                End If
            End If
            If bDelete Then cm.DeleteLines j, i - j + 1
            i = j - 2   'line j - 1 is done as well
        Else
            If arrNew(i) <> arrOld(i) Then TrfReplaceLine cm, i, arrNew(i), arrOld(i), sMethod
            i = i - 1
        End If
    Loop
End Sub

'Replaces line lLine; returns False if the editor split the line and the original line was restored.
Private Function TrfReplaceLine(ByVal cm As VBIDE.CodeModule, ByVal lLine As Long, ByVal sNew As String, _
                                ByVal sOld As String, ByVal sMethod As String) As Boolean
    Dim lBefore As Long
    Dim lExtra As Long

    lBefore = cm.CountOfLines
    cm.ReplaceLine lLine, sNew
    lExtra = cm.CountOfLines - lBefore
    If lExtra < 0 Then
        Err.Raise vbObjectError + 2601, "TrfApplyLineEdits", "Line count decreased after ReplaceLine in line " & lLine & "."
    ElseIf lExtra > 0 Then
        cm.DeleteLines lLine, lExtra + 1
        cm.InsertLines lLine, sOld
        TrfAddNotice sMethod, cm.Parent.Name, lLine, "editor split the edited line; original line restored"
    Else
        TrfReplaceLine = True
    End If
End Function

'Copies the lines into arrNew and marks every line as kept.
Private Sub TrfInitEdits(ByRef arrLines() As String, ByVal n As Long, ByRef arrNew() As String, ByRef abKeep() As Boolean)
    Dim i As Long

    ReDim arrNew(1 To n)
    ReDim abKeep(1 To n)
    For i = 1 To n
        arrNew(i) = arrLines(i)
        abKeep(i) = True
    Next i
End Sub

'===============================================================================
' Pure transformations (no access to the code module)
'===============================================================================

'Computes the edits of Remove_DebugPrint (lStmtKind = TRF_STMT_DEBUGPRINT)
'or Remove_OptionExplicit (lStmtKind = TRF_STMT_OPTIONEXPLICIT).
Private Sub TrfStatementEdits(ByRef arrLines() As String, ByVal n As Long, ByVal lStmtKind As Long, _
                              ByVal sMethod As String, ByVal sModule As String, _
                              ByRef arrNew() As String, ByRef abKeep() As Boolean)
    Dim i As Long
    Dim l As Long
    Dim toks() As TrfToken
    Dim nTok As Long
    Dim info As TrfLineInfo

    TrfInitEdits arrLines, n, arrNew, abKeep
    i = 1
    Do While i <= n
        l = TrfLogicalLineEnd(arrLines, i, n)
        TrfScanLogicalLine arrLines, i, l, toks, nTok, info
        TrfRemoveStatements arrLines, i, l, toks, nTok, info, lStmtKind, sMethod, sModule, arrNew, abKeep
        i = l + 1
    Loop
End Sub

'Computes the edits of Remove_Comments.
Private Sub TrfCommentEdits(ByRef arrLines() As String, ByVal n As Long, ByVal sModule As String, _
                            ByRef arrNew() As String, ByRef abKeep() As Boolean)
    Dim i As Long
    Dim l As Long
    Dim j As Long
    Dim toks() As TrfToken
    Dim nTok As Long
    Dim info As TrfLineInfo
    Dim sCode As String

    TrfInitEdits arrLines, n, arrNew, abKeep
    i = 1
    Do While i <= n
        l = TrfLogicalLineEnd(arrLines, i, n)
        TrfScanLogicalLine arrLines, i, l, toks, nTok, info
        If info.HasComment Then
            If info.RemUnsafe Then
                TrfAddNotice "Remove_Comments", sModule, info.CommentLineIdx, _
                    "Rem directly after Then/Else of a single-line If; not removed"
            Else
                'the comment and its continuation lines go
                For j = info.CutLineIdx + 1 To l
                    abKeep(j) = False
                Next j
                sCode = TrfRTrimWS(Left$(arrLines(info.CutLineIdx), info.CutPos - 1))
                If Len(TrfLTrimWS(sCode)) > 0 Then
                    arrNew(info.CutLineIdx) = sCode
                Else
                    abKeep(info.CutLineIdx) = False
                    'the line before continued into the deleted line: end the statement there
                    If info.CutLineIdx > i Then
                        arrNew(info.CutLineIdx - 1) = TrfStripContinuation(arrNew(info.CutLineIdx - 1))
                    End If
                End If
            End If
        End If
        i = l + 1
    Loop
End Sub

'Computes the edits of RemoveBreaksLineInCode.
Private Sub TrfMergeEdits(ByRef arrLines() As String, ByVal n As Long, ByVal sModule As String, _
                          ByRef arrNew() As String, ByRef abKeep() As Boolean)
    Dim i As Long
    Dim l As Long
    Dim j As Long
    Dim sMerged As String
    Dim sPart As String

    TrfInitEdits arrLines, n, arrNew, abKeep
    i = 1
    Do While i <= n
        l = TrfLogicalLineEnd(arrLines, i, n)
        If l > i Then
            sMerged = TrfStripContinuation(arrLines(i))
            For j = i + 1 To l
                If j < l Then
                    sPart = TrfStripContinuation(arrLines(j))
                Else
                    sPart = arrLines(j)
                End If
                sPart = TrfLTrimWS(sPart)
                If Len(sPart) > 0 Then
                    If Len(TrfLTrimWS(sMerged)) = 0 Or Right$(sMerged, 1) = "(" Then
                        sMerged = sMerged & sPart
                    Else
                        sMerged = sMerged & " " & sPart
                    End If
                End If
            Next j
            If Len(sMerged) > TRF_MAX_LINE_LENGTH Then
                TrfAddNotice "RemoveBreaksLineInCode", sModule, i, "joined statement would have " & Len(sMerged) & _
                    " characters (limit " & TRF_MAX_LINE_LENGTH & "); line continuations kept"
            Else
                arrNew(i) = sMerged
                For j = i + 1 To l
                    abKeep(j) = False
                Next j
            End If
        End If
        i = l + 1
    Loop
End Sub

'Removes the matching statements of one logical line (lFirst..lLast).
Private Sub TrfRemoveStatements(ByRef arrLines() As String, ByVal lFirst As Long, ByVal lLast As Long, _
                                ByRef toks() As TrfToken, ByVal nTok As Long, ByRef info As TrfLineInfo, _
                                ByVal lStmtKind As Long, ByVal sMethod As String, ByVal sModule As String, _
                                ByRef arrNew() As String, ByRef abKeep() As Boolean)
    Dim stFirst() As Long
    Dim stLast() As Long
    Dim abRemove() As Boolean
    Dim nSt As Long
    Dim t As Long
    Dim tStart As Long
    Dim i As Long
    Dim j As Long
    Dim lIfStmt As Long
    Dim nMatch As Long
    Dim nRemove As Long
    Dim nOccur As Long
    Dim lLastLabelTok As Long
    Dim sCallWord As String

    If nTok = 0 Or info.IsDirective Then Exit Sub

    'split the tokens into statements (labels and separators are no statements)
    ReDim stFirst(0 To nTok)
    ReDim stLast(0 To nTok)
    tStart = -1
    lLastLabelTok = -1
    For t = 0 To nTok - 1
        Select Case toks(t).Kind
            Case TRF_TK_LABEL
                lLastLabelTok = t
            Case TRF_TK_SEPARATOR
                If tStart >= 0 Then
                    stFirst(nSt) = tStart
                    stLast(nSt) = t - 1
                    nSt = nSt + 1
                    tStart = -1
                End If
            Case Else
                If tStart < 0 Then tStart = t
        End Select
    Next t
    If tStart >= 0 Then
        stFirst(nSt) = tStart
        stLast(nSt) = nTok - 1
        nSt = nSt + 1
    End If

    'a statement starting with If and everything after it belongs to that If
    lIfStmt = nSt
    For i = 0 To nSt - 1
        If toks(stFirst(i)).Kind = TRF_TK_WORD Then
            If LCase$(toks(stFirst(i)).Text) = "if" Then
                lIfStmt = i
                Exit For
            End If
        End If
    Next i

    If nSt > 0 Then ReDim abRemove(0 To nSt - 1)
    For i = 0 To nSt - 1
        If TrfStatementMatches(toks, stFirst(i), stLast(i), lStmtKind) Then
            nMatch = nMatch + 1
            If i < lIfStmt Then
                abRemove(i) = True
                nRemove = nRemove + 1
            End If
        End If
    Next i
    If lStmtKind = TRF_STMT_DEBUGPRINT Then
        nOccur = TrfCountDebugPrint(toks, nTok)
    Else
        nOccur = nMatch
    End If
    If nOccur = 0 Then Exit Sub

    If lFirst < lLast Then
        'continued logical line: only a lone matching statement is removed
        If nSt = 1 And nRemove = 1 And lLastLabelTok < 0 Then
            If info.HasComment Then
                If Len(TrfLeadingWS(arrLines(lFirst))) + Len(arrLines(info.CommentLineIdx)) - info.CommentPos + 1 > TRF_MAX_LINE_LENGTH Then
                    TrfAddNotice sMethod, sModule, lFirst, "rebuilt line would exceed " & TRF_MAX_LINE_LENGTH & " characters; not removed"
                    Exit Sub
                End If
                For j = lFirst To info.CommentLineIdx - 1
                    abKeep(j) = False
                Next j
                arrNew(info.CommentLineIdx) = TrfLeadingWS(arrLines(lFirst)) & Mid$(arrLines(info.CommentLineIdx), info.CommentPos)
            Else
                For j = lFirst To lLast
                    abKeep(j) = False
                Next j
            End If
        Else
            TrfAddNotice sMethod, sModule, lFirst, _
                "statement continued over several lines together with other code; not removed"
        End If
        Exit Sub
    End If

    If nRemove > 0 Then
        Select Case TrfRebuildLine(arrLines(lFirst), toks, info, stFirst, stLast, abRemove, nSt, lLastLabelTok, lFirst, _
                                   arrNew, abKeep, sCallWord)
            Case TRF_REBUILD_CALL
                TrfAddNotice sMethod, sModule, lFirst, _
                    "Call inserted before " & sCallWord & ", otherwise it would become a line label"
            Case TRF_REBUILD_KEYWORD
                TrfAddNotice sMethod, sModule, lFirst, _
                    "keyword statement " & sCallWord & " would stand before "":"" at the line start; not removed"
                Exit Sub
            Case TRF_REBUILD_TOO_LONG
                TrfAddNotice sMethod, sModule, lFirst, _
                    "rebuilt line would exceed " & TRF_MAX_LINE_LENGTH & " characters; not removed"
                Exit Sub
        End Select
    End If
    If nOccur > nRemove Then
        TrfAddNotice sMethod, sModule, lFirst, _
            "statement inside a single-line If or not at the start of a statement; not removed"
    End If
End Sub

'Builds one physical line from its label, the kept statements and its comment.
'If a lone name would now start the statements and be followed by ":", VBA would read
'it as a label definition ("Debug.Print x: Init: Run" -> "Init: Run"). A procedure name
'then gets "Call " in front ("Call Init: Run", released by the product owner); for a
'keyword statement nothing is changed. Returns one of the TRF_REBUILD_* values; sWord
'receives the lone name. arrNew and abKeep are only changed for DONE and CALL.
Private Function TrfRebuildLine(ByVal sLine As String, ByRef toks() As TrfToken, ByRef info As TrfLineInfo, _
                                ByRef stFirst() As Long, ByRef stLast() As Long, ByRef abRemove() As Boolean, _
                                ByVal nSt As Long, ByVal lLastLabelTok As Long, ByVal lLineIdx As Long, _
                                ByRef arrNew() As String, ByRef abKeep() As Boolean, ByRef sWord As String) As Long
    Dim lResult As Long
    Dim lWordPos As Long
    Dim sOut As String
    Dim sCore As String
    Dim sGap As String
    Dim sRest As String
    Dim i As Long
    Dim lFirstKept As Long
    Dim lLastKept As Long
    Dim lEnd As Long
    Dim lPrefixLen As Long

    'indentation, or indentation plus label
    If lLastLabelTok >= 0 Then
        sOut = Left$(sLine, toks(lLastLabelTok).StartPos + toks(lLastLabelTok).Length - 1)
    Else
        sOut = Left$(sLine, toks(0).StartPos - 1)
    End If
    lPrefixLen = Len(sOut)

    lFirstKept = -1
    lLastKept = -1
    For i = 0 To nSt - 1
        If Not abRemove(i) Then
            If Len(sCore) > 0 Then sCore = sCore & ": "
            sCore = sCore & Mid$(sLine, toks(stFirst(i)).StartPos, _
                toks(stLast(i)).StartPos + toks(stLast(i)).Length - toks(stFirst(i)).StartPos)
            If lFirstKept < 0 Then lFirstKept = i
            lLastKept = i
        End If
    Next i
    If Len(sCore) > 0 Then
        If lLastLabelTok >= 0 Then sOut = sOut & " "
        sOut = sOut & sCore
    End If

    If info.HasComment Then
        If lLastKept >= 0 And lLastKept = nSt - 1 Then
            'last statement unchanged: keep the original text up to the comment
            lEnd = toks(stLast(lLastKept)).StartPos + toks(stLast(lLastKept)).Length
            sGap = Mid$(sLine, lEnd, info.CommentPos - lEnd)
        ElseIf Len(sCore) > 0 Then
            If info.CommentIsRem Then sGap = ": " Else sGap = " "
        ElseIf lLastLabelTok >= 0 Then
            sGap = " "
        Else
            sGap = ""
        End If
        sOut = sOut & sGap & Mid$(sLine, info.CommentPos)
    End If

    'a lone name that now starts the statements and is followed by ":" (not ":=")
    lResult = TRF_REBUILD_DONE
    sWord = ""
    If lFirstKept > 0 Then
        If stFirst(lFirstKept) = stLast(lFirstKept) And toks(stFirst(lFirstKept)).Kind = TRF_TK_WORD Then
            sWord = toks(stFirst(lFirstKept)).Text
            sRest = TrfLTrimWS(Mid$(sOut, lPrefixLen + 1))
            lWordPos = Len(sOut) - Len(sRest) + 1
            sRest = TrfLTrimWS(Mid$(sRest, Len(sWord) + 1))
            If Left$(sRest, 1) = ":" And Mid$(sRest, 2, 1) <> "=" Then
                If TrfIsStandaloneKeyword(sWord) Then
                    TrfRebuildLine = TRF_REBUILD_KEYWORD
                    Exit Function
                End If
                sOut = Left$(sOut, lWordPos - 1) & "Call " & Mid$(sOut, lWordPos)
                lResult = TRF_REBUILD_CALL
            End If
        End If
    End If

    sOut = TrfRTrimWS(sOut)
    If Len(sOut) > TRF_MAX_LINE_LENGTH Then
        TrfRebuildLine = TRF_REBUILD_TOO_LONG
        Exit Function
    End If
    If Len(TrfLTrimWS(sOut)) = 0 Then
        abKeep(lLineIdx) = False
    Else
        arrNew(lLineIdx) = sOut
    End If
    TrfRebuildLine = lResult
End Function

'===============================================================================
' Procedure : TrfIsStandaloneKeyword
' Purpose   : True for a keyword that forms a statement on its own and cannot be
'             written with Call (Stop, Loop, Next, ...).
'
' Parameters:
'   sWord - String, ByVal, a single name (any case)
'
' Returns:
'   Boolean
'
' Side Effects:
'   Keine
'
' Preconditions:
'   Keine
'
' Postconditions:
'   Keine
'
' Errors:
'   Keine
'
' Notes:
'   A lone name followed by ":" at the start of a line is a label definition
'   in VBA. Procedure names get "Call " in front; for the words of this list
'   Call would be a compile error, so the line is left unchanged. Beep and
'   Randomize are included conservatively: it is not verified whether VBA
'   accepts Call for them.
'===============================================================================
Public Function TrfIsStandaloneKeyword(ByVal sWord As String) As Boolean
    Select Case LCase$(sWord)
        Case "stop", "end", "loop", "wend", "next", "return", "resume", "else", "do", "close", "reset", "beep", "randomize"
            TrfIsStandaloneKeyword = True
    End Select
End Function

'True if the statement tFirst..tLast is the one searched for.
Private Function TrfStatementMatches(ByRef toks() As TrfToken, ByVal tFirst As Long, ByVal tLast As Long, _
                                     ByVal lStmtKind As Long) As Boolean
    Select Case lStmtKind
        Case TRF_STMT_DEBUGPRINT
            If tLast - tFirst >= 2 Then TrfStatementMatches = TrfIsDebugPrintAt(toks, tFirst)
        Case TRF_STMT_OPTIONEXPLICIT
            If tLast - tFirst = 1 Then
                If toks(tFirst).Kind = TRF_TK_WORD And toks(tLast).Kind = TRF_TK_WORD Then
                    TrfStatementMatches = (LCase$(toks(tFirst).Text) = "option" And LCase$(toks(tLast).Text) = "explicit")
                End If
            End If
    End Select
End Function

'True if the tokens t, t + 1, t + 2 are Debug . Print (t + 2 must exist).
Private Function TrfIsDebugPrintAt(ByRef toks() As TrfToken, ByVal t As Long) As Boolean
    If toks(t).Kind <> TRF_TK_WORD Then Exit Function
    If LCase$(toks(t).Text) <> "debug" Then Exit Function
    If toks(t + 1).Text <> "." Then Exit Function
    If toks(t + 2).Kind <> TRF_TK_WORD Then Exit Function
    TrfIsDebugPrintAt = (LCase$(toks(t + 2).Text) = "print")
End Function

'Counts Debug.Print calls anywhere in the code tokens (not obj.Debug.Print).
Private Function TrfCountDebugPrint(ByRef toks() As TrfToken, ByVal nTok As Long) As Long
    Dim t As Long

    For t = 0 To nTok - 3
        If TrfIsDebugPrintAt(toks, t) Then
            If t = 0 Then
                TrfCountDebugPrint = TrfCountDebugPrint + 1
            ElseIf toks(t - 1).Text <> "." And toks(t - 1).Text <> "!" Then
                TrfCountDebugPrint = TrfCountDebugPrint + 1
            End If
        End If
    Next t
End Function

'===============================================================================
' Lexer
'===============================================================================

'===============================================================================
' Procedure : TrfScanLogicalLine
' Purpose   : Splits one logical line into code tokens and locates its comment.
'
' Parameters:
'   arrLines - String(), ByRef, module lines (not changed)
'   lFirst   - Long, ByVal, first physical line of the logical line
'   lLast    - Long, ByVal, last physical line (see TrfLogicalLineEnd)
'   toks     - TrfToken(), ByRef, receives the tokens (0 To nTok - 1)
'   nTok     - Long, ByRef, receives the number of tokens
'   info     - TrfLineInfo, ByRef, receives directive and comment data
'
' Returns:
'   Keiner
'
' Side Effects:
'   Keine
'
' Preconditions:
'   Every line lFirst..lLast - 1 ends with a line continuation.
'
' Postconditions:
'   Tokens cover the code outside the comment. Strings, bracketed names
'   and date literals are single tokens, so their content is never taken
'   for code, comment or separator.
'
' Errors:
'   Keine
'
' Notes:
'   Statement starts: line start, after a label, after ":" and after the
'   keywords Then and Else. A number as first token is a line number label,
'   a name directly followed by ":" (not ":=") as first token is a label.
'   Keywords taken for labels ("Else:") are harmless: label text is always
'   kept verbatim.
'===============================================================================
Public Sub TrfScanLogicalLine(ByRef arrLines() As String, ByVal lFirst As Long, ByVal lLast As Long, _
                              ByRef toks() As TrfToken, ByRef nTok As Long, ByRef info As TrfLineInfo)
    Dim k As Long
    Dim s As String
    Dim n As Long
    Dim p As Long
    Dim q As Long
    Dim ch As String
    Dim sWord As String
    Dim bStmtStart As Boolean
    Dim bHasIfStmt As Boolean
    Dim lChars As Long

    'every token covers at least one character: the array never has to grow
    For k = lFirst To lLast
        lChars = lChars + Len(arrLines(k))
    Next k
    nTok = 0
    ReDim toks(0 To lChars)
    info.IsDirective = False
    info.HasComment = False
    info.CommentLineIdx = 0
    info.CommentPos = 0
    info.CommentIsRem = False
    info.CutLineIdx = 0
    info.CutPos = 0
    info.RemUnsafe = False
    bStmtStart = True

    For k = lFirst To lLast
        s = arrLines(k)
        If k < lLast Then
            n = TrfContinuationPos(s) - 1
        Else
            n = Len(s)
        End If
        p = 1
        Do While p <= n
            ch = Mid$(s, p, 1)
            If ch = " " Or ch = vbTab Then
                p = p + 1
            ElseIf ch = "'" Then
                TrfSetComment toks, nTok, info, k, p, False, bHasIfStmt
                Exit Sub
            ElseIf ch = """" Then
                q = p + 1
                Do While q <= n
                    If Mid$(s, q, 1) = """" Then
                        If Mid$(s, q + 1, 1) = """" Then
                            q = q + 2
                        Else
                            Exit Do
                        End If
                    Else
                        q = q + 1
                    End If
                Loop
                If q > n Then q = n
                TrfPushToken toks, nTok, k, p, q - p + 1, TRF_TK_STRING, s
                p = q + 1
                bStmtStart = False
            ElseIf ch = "[" Then
                q = InStr(p + 1, s, "]")
                If q = 0 Or q > n Then q = n
                TrfPushToken toks, nTok, k, p, q - p + 1, TRF_TK_BRACKET, s
                p = q + 1
                bStmtStart = False
            ElseIf ch = "#" Then
                If nTok = 0 And TrfIsAsciiLetter(Mid$(s, p + 1, 1)) Then
                    info.IsDirective = True
                    TrfPushToken toks, nTok, k, p, 1, TRF_TK_PUNCT, s
                    p = p + 1
                ElseIf TrfIsDateLiteralAt(s, p, n, q) Then
                    TrfPushToken toks, nTok, k, p, q - p + 1, TRF_TK_DATE, s
                    p = q + 1
                Else
                    TrfPushToken toks, nTok, k, p, 1, TRF_TK_PUNCT, s
                    p = p + 1
                End If
                bStmtStart = False
            ElseIf TrfIsDigit(ch) Then
                q = p
                Do While q <= n
                    If Not TrfIsDigit(Mid$(s, q, 1)) Then Exit Do
                    q = q + 1
                Loop
                If nTok = 0 Then
                    'line number label, optionally followed by ":"
                    TrfPushToken toks, nTok, k, p, q - p, TRF_TK_LABEL, s
                    p = q
                    Do While q <= n
                        If Not TrfIsWS(Mid$(s, q, 1)) Then Exit Do
                        q = q + 1
                    Loop
                    If q <= n Then
                        If Mid$(s, q, 1) = ":" And Mid$(s, q + 1, 1) <> "=" Then
                            TrfPushToken toks, nTok, k, q, 1, TRF_TK_LABEL, s
                            p = q + 1
                        End If
                    End If
                    bStmtStart = True
                Else
                    TrfPushToken toks, nTok, k, p, q - p, TRF_TK_NUMBER, s
                    p = q
                    bStmtStart = False
                End If
            ElseIf TrfIsWordStart(ch) Then
                q = p + 1
                Do While q <= n
                    If Not TrfIsWordChar(Mid$(s, q, 1)) Then Exit Do
                    q = q + 1
                Loop
                sWord = LCase$(Mid$(s, p, q - p))
                If bStmtStart And Not info.IsDirective And sWord = "rem" Then
                    If q > n Or TrfIsWS(Mid$(s, q, 1)) Then
                        TrfSetComment toks, nTok, info, k, p, True, bHasIfStmt
                        Exit Sub
                    End If
                End If
                If nTok = 0 And Not info.IsDirective And Mid$(s, q, 1) = ":" And Mid$(s, q + 1, 1) <> "=" Then
                    'label "Name:"
                    TrfPushToken toks, nTok, k, p, q - p, TRF_TK_LABEL, s
                    TrfPushToken toks, nTok, k, q, 1, TRF_TK_LABEL, s
                    p = q + 1
                    bStmtStart = True
                Else
                    TrfPushToken toks, nTok, k, p, q - p, TRF_TK_WORD, s
                    p = q
                    If bStmtStart And sWord = "if" Then bHasIfStmt = True
                    If (sWord = "then" Or sWord = "else") And Not info.IsDirective Then
                        bStmtStart = True
                    Else
                        bStmtStart = False
                    End If
                End If
            ElseIf ch = ":" Then
                If Mid$(s, p + 1, 1) = "=" Then
                    TrfPushToken toks, nTok, k, p, 2, TRF_TK_PUNCT, s
                    p = p + 2
                    bStmtStart = False
                Else
                    TrfPushToken toks, nTok, k, p, 1, TRF_TK_SEPARATOR, s
                    p = p + 1
                    bStmtStart = True
                End If
            Else
                TrfPushToken toks, nTok, k, p, 1, TRF_TK_PUNCT, s
                p = p + 1
                bStmtStart = False
            End If
        Loop
    Next k
End Sub

'Records the comment that starts at line lLineIdx, position lPos.
Private Sub TrfSetComment(ByRef toks() As TrfToken, ByVal nTok As Long, ByRef info As TrfLineInfo, _
                          ByVal lLineIdx As Long, ByVal lPos As Long, ByVal bRem As Boolean, ByVal bHasIfStmt As Boolean)
    Dim lPrev As Long
    Dim bAfterThenElse As Boolean

    info.HasComment = True
    info.CommentLineIdx = lLineIdx
    info.CommentPos = lPos
    info.CommentIsRem = bRem
    info.CutLineIdx = lLineIdx
    info.CutPos = lPos
    If nTok = 0 Then Exit Sub

    lPrev = nTok - 1
    If toks(lPrev).Kind = TRF_TK_SEPARATOR Then
        If lPrev > 0 Then bAfterThenElse = TrfIsThenElse(toks, lPrev - 1)
        'a separator before the comment goes with it, except after Then/Else ("If x Then:" keeps its form)
        If Not bAfterThenElse Then
            info.CutLineIdx = toks(lPrev).LineIdx
            info.CutPos = toks(lPrev).StartPos
        End If
    Else
        bAfterThenElse = TrfIsThenElse(toks, lPrev)
    End If
    If bRem And bAfterThenElse And bHasIfStmt Then info.RemUnsafe = True
End Sub

'True if token t is the keyword Then or Else, but not the Else of "Case Else".
Private Function TrfIsThenElse(ByRef toks() As TrfToken, ByVal t As Long) As Boolean
    Dim sWord As String

    If toks(t).Kind <> TRF_TK_WORD Then Exit Function
    sWord = LCase$(toks(t).Text)
    If sWord = "then" Then
        TrfIsThenElse = True
    ElseIf sWord = "else" Then
        TrfIsThenElse = True
        If t > 0 Then
            If toks(t - 1).Kind = TRF_TK_WORD Then
                If LCase$(toks(t - 1).Text) = "case" Then TrfIsThenElse = False
            End If
        End If
    End If
End Function

'True if "#" at position p starts a date literal; qEnd receives the position of the closing "#".
Private Function TrfIsDateLiteralAt(ByVal s As String, ByVal p As Long, ByVal n As Long, ByRef qEnd As Long) As Boolean
    Dim q As Long
    Dim i As Long
    Dim ch As String
    Dim sBody As String
    Dim sWord As String
    Dim bDigit As Boolean

    If p > 1 Then
        'type suffix as in x# or 1.5#
        ch = Mid$(s, p - 1, 1)
        If TrfIsWordChar(ch) Or ch = ")" Or ch = "]" Or ch = "." Then Exit Function
    End If
    q = InStr(p + 1, s, "#")
    If q = 0 Or q > n Then Exit Function
    sBody = Mid$(s, p + 1, q - p - 1)
    If Len(sBody) = 0 Then Exit Function
    If Not (TrfIsDigit(Left$(sBody, 1)) Or TrfIsAsciiLetter(Left$(sBody, 1))) Then Exit Function
    sBody = sBody & " "
    For i = 1 To Len(sBody)
        ch = Mid$(sBody, i, 1)
        If TrfIsAsciiLetter(ch) Then
            sWord = sWord & ch
        Else
            If Len(sWord) > 0 Then
                If InStr(1, TRF_DATE_WORDS, " " & LCase$(sWord) & " ") = 0 Then Exit Function
                sWord = ""
            End If
            If TrfIsDigit(ch) Then
                bDigit = True
            ElseIf InStr(1, " /-:.,", ch) = 0 Then
                Exit Function
            End If
        End If
    Next i
    If Not bDigit Then Exit Function
    qEnd = q
    TrfIsDateLiteralAt = True
End Function

'Appends a token; toks is sized by TrfScanLogicalLine to the character count of the logical line.
Private Sub TrfPushToken(ByRef toks() As TrfToken, ByRef nTok As Long, ByVal lLineIdx As Long, _
                         ByVal lPos As Long, ByVal lLen As Long, ByVal lKind As Long, ByRef sLine As String)
    toks(nTok).LineIdx = lLineIdx
    toks(nTok).StartPos = lPos
    toks(nTok).Length = lLen
    toks(nTok).Kind = lKind
    toks(nTok).Text = Mid$(sLine, lPos, lLen)
    nTok = nTok + 1
End Sub

'===============================================================================
' Line helpers
'===============================================================================

'Returns the last line of the logical line that starts at lFirst.
Public Function TrfLogicalLineEnd(ByRef arrLines() As String, ByVal lFirst As Long, ByVal lCount As Long) As Long
    Dim l As Long

    l = lFirst
    Do While l < lCount
        If Not TrfIsContinuationLine(arrLines(l)) Then Exit Do
        l = l + 1
    Loop
    TrfLogicalLineEnd = l
End Function

'True if the line ends with a line continuation (whitespace, "_", optional whitespace).
Public Function TrfIsContinuationLine(ByVal sLine As String) As Boolean
    TrfIsContinuationLine = (TrfContinuationPos(sLine) > 0)
End Function

'Position of the "_" of a line continuation, 0 if the line is not continued.
Private Function TrfContinuationPos(ByVal sLine As String) As Long
    Dim p As Long

    p = Len(sLine)
    Do While p > 0
        If Not TrfIsWS(Mid$(sLine, p, 1)) Then Exit Do
        p = p - 1
    Loop
    If p >= 2 Then
        If Mid$(sLine, p, 1) = "_" And TrfIsWS(Mid$(sLine, p - 1, 1)) Then TrfContinuationPos = p
    End If
End Function

'Returns the line without its line continuation and the whitespace before it.
Private Function TrfStripContinuation(ByVal sLine As String) As String
    Dim p As Long

    p = TrfContinuationPos(sLine)
    If p = 0 Then
        TrfStripContinuation = sLine
    Else
        TrfStripContinuation = TrfRTrimWS(Left$(sLine, p - 1))
    End If
End Function

'True if the text contains nothing but spaces and tabs.
Public Function TrfIsBlank(ByVal sText As String) As Boolean
    TrfIsBlank = (Len(TrfLTrimWS(sText)) = 0)
End Function

'Removes leading spaces and tabs.
Public Function TrfLTrimWS(ByVal sText As String) As String
    Dim p As Long

    p = 1
    Do While p <= Len(sText)
        If Not TrfIsWS(Mid$(sText, p, 1)) Then Exit Do
        p = p + 1
    Loop
    TrfLTrimWS = Mid$(sText, p)
End Function

'Removes trailing spaces and tabs.
Private Function TrfRTrimWS(ByVal sText As String) As String
    Dim p As Long

    p = Len(sText)
    Do While p > 0
        If Not TrfIsWS(Mid$(sText, p, 1)) Then Exit Do
        p = p - 1
    Loop
    TrfRTrimWS = Left$(sText, p)
End Function

'Removes leading and trailing spaces and tabs.
Private Function TrfTrimWS(ByVal sText As String) As String
    TrfTrimWS = TrfRTrimWS(TrfLTrimWS(sText))
End Function

'Leading spaces and tabs of a line.
Private Function TrfLeadingWS(ByVal sText As String) As String
    TrfLeadingWS = Left$(sText, Len(sText) - Len(TrfLTrimWS(sText)))
End Function

Private Function TrfIsWS(ByVal ch As String) As Boolean
    TrfIsWS = (ch = " " Or ch = vbTab)
End Function

Private Function TrfIsDigit(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    TrfIsDigit = (AscW(ch) >= 48 And AscW(ch) <= 57)
End Function

Private Function TrfIsAsciiLetter(ByVal ch As String) As Boolean
    Dim c As Long

    If Len(ch) = 0 Then Exit Function
    c = AscW(ch)
    TrfIsAsciiLetter = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

'Letters (also non-ASCII) and "_" may start a name.
Private Function TrfIsWordStart(ByVal ch As String) As Boolean
    Dim c As Long

    If Len(ch) = 0 Then Exit Function
    c = AscW(ch)
    TrfIsWordStart = TrfIsAsciiLetter(ch) Or c = 95 Or c > 127 Or c < 0
End Function

Private Function TrfIsWordChar(ByVal ch As String) As Boolean
    TrfIsWordChar = TrfIsWordStart(ch) Or TrfIsDigit(ch)
End Function
