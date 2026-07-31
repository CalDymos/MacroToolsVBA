Attribute VB_Name = "M_MoveControl"
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : M_MoveControl - Микро подстройка элементов формы VBA и переименование элементов на форме вместе с кодом
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Public sTagNameConrol As String
Public tpStyle      As ProperControlStyle
Type ProperControlStyle
    sError          As String
    snHeight        As Single
    snWidth         As Single
    bVisible        As Boolean
    bEnabled        As Boolean
    bLocked         As Boolean
    lBackColor      As Long
    lForeColor      As Long
    lBackStyle      As Long
    lBorderColor    As Long
    lBorderStyle    As Long
    bFontBold       As Boolean
    bFontItalic     As Boolean
    bFontStrikethru As Boolean
    bFontUnderline  As Boolean
    sFontName       As String
    cuFontSize      As Currency
End Type

    Public Sub HelpMoveControl()
34:    Call URLLinks(C_Const.URL_MOVE_CNTR)
35: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : MoveControl - Микроподстройка элементов формы
'* Created    : 08-10-2020 14:10
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub MoveControl()
44:    If Application.VBE.ActiveWindow.Type <> vbext_wt_Designer Then Exit Sub
45:    Dim myCommandBar As CommandBar
46:    Dim cntrl       As CommandBarControl
47:    Dim combox      As CommandBarComboBox
48:    Dim sComBoxText As String
49:    Dim cnt         As control
50:
51:    Set myCommandBar = Application.VBE.CommandBars(C_Const.MENUMOVECONTRL)
52:    For Each cntrl In myCommandBar.Controls
53:        If cntrl.Tag = C_Const.MTAGCOM Then
54:            Set combox = myCommandBar.Controls(cntrl.ID)
55:            sComBoxText = combox.Text
56:            Exit For
57:        End If
58:    Next cntrl
59:
60:    Dim objActiveModule As VBComponent
61:    Set objActiveModule = getActiveModule()
62:    For Each cnt In objActiveModule.Designer.Selected
63:        If Not cnt Is Nothing Then
64:            Select Case sTagNameConrol
                Case C_Const.MTAG1:
66:                    Call MoveCnt(cnt, 1, sComBoxText)
67:                Case C_Const.MTAG2:
68:                    Call MoveCnt(cnt, 2, sComBoxText)
69:                Case C_Const.MTAG3:
70:                    Call MoveCnt(cnt, 3, sComBoxText)
71:                Case C_Const.MTAG4:
72:                    Call MoveCnt(cnt, 4, sComBoxText)
73:            End Select
74:        End If
75:    Next cnt
76: End Sub
     Private Sub MoveCnt(ByRef cnt As control, ByVal iVal As Integer, ByVal sComBoxText As String)
78:    Const Shag = 0.4
79:    With cnt
80:        Select Case sComBoxText
            Case C_Const.MOVECONT:
82:                Select Case iVal
                    Case 1:
84:                        .Left = .Left - Shag
85:                    Case 2:
86:                        .Left = .Left + Shag
87:                    Case 3:
88:                        .Top = .Top + Shag
89:                    Case 4:
90:                        .Top = .Top - Shag
91:                End Select
92:            Case C_Const.MOVECONTTOPLEFT:
93:                Select Case iVal
                    Case 1:
95:                        .Left = .Left - Shag
96:                        .Width = .Width + Shag
97:                    Case 2:
98:                        .Left = .Left + Shag
99:                        .Width = .Width - Shag
100:                    Case 3:
101:                        .Top = .Top + Shag
102:                        .Height = .Height - Shag
103:                    Case 4:
104:                        .Top = .Top - Shag
105:                        .Height = .Height + Shag
106:                End Select
107:            Case C_Const.MOVECONTBOTTOMRIGHT:
108:                Select Case iVal
                    Case 1:
110:                        .Width = .Width - Shag
111:                    Case 2:
112:                        .Width = .Width + Shag
113:                    Case 3:
114:                        .Height = .Height + Shag
115:                    Case 4:
116:                        .Height = .Height - Shag
117:                End Select
118:        End Select
119:    End With
120: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RenameControl - переименование конторол на форме вместе скодом
'* Created    : 08-10-2020 14:11
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub RenameControl()
130:    Dim cnt         As control
131:    Dim sNewName    As String
132:    Dim sOldName    As String
133:    Dim NameModeCode As String
134:    Dim strVar      As String
135:    Dim CodeMod     As CodeModule
136:
137:    On Error GoTo ErrorHandler
138:
139:    Set cnt = TakeSelectControl
140:    If cnt Is Nothing Then Exit Sub
tryagin:
142:    sOldName = cnt.Name
143:    sNewName = InputBox("Ведите новое название Control", "Переименование Control:", sOldName)
144:    If sNewName = vbNullString Or sNewName = sOldName Then Exit Sub
145:
146:    cnt.Name = sNewName
147:    Set CodeMod = Application.VBE.SelectedVBComponent.CodeModule
148:    With CodeMod
149:        strVar = .Lines(1, .CountOfLines)
150:        strVar = ReplceCode(strVar, sOldName, sNewName)
151:        .DeleteLines startLine:=1, Count:=.CountOfLines
152:        .InsertLines Line:=1, String:=strVar
153:    End With
154:    Exit Sub
ErrorHandler:
156:    Select Case Err.Number
        Case 40044:
158:            Call MsgBox("Fehler! Ведено не допустимое имя Control [ " & sNewName & " ], задайте дргое имя!", vbCritical, "Ведено не допустимое имя Control:")
159:            Exit Sub
160:        Case -2147319764:
161:            Call MsgBox("Данное Имя Control уже используется [" & sNewName & " ], задайте дргое имя!", vbCritical, "Имя задано неоднозначно:")
162:            Exit Sub
163:        Case Else:
164:            Debug.Print "Fehler! в RenameControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
165:            Call WriteErrorLog("RenameControl")
166:    End Select
167:    Err.Clear
168: End Sub
     Public Sub CopyStyleControl()
170:    Dim cnt         As Object
171:    Set cnt = TakeSelectControl(True)
172:    If cnt Is Nothing Then Exit Sub
173:
174:    'установка по умолчанию значений
175:    tpStyle.lBackStyle = 1
176:    tpStyle.lBorderColor = -2147483642
177:    tpStyle.lBorderStyle = 0
178:    tpStyle.bVisible = True
179:    tpStyle.bLocked = False
180:    tpStyle.bEnabled = True
181:    tpStyle.lBackStyle = 1
182:
183:    On Error Resume Next
184:    With cnt
185:        tpStyle.bEnabled = .Enabled
186:        tpStyle.bFontBold = .Font.Bold
187:        tpStyle.bFontItalic = .Font.Italic
188:        tpStyle.bFontStrikethru = .Font.Strikethrough
189:        tpStyle.bFontUnderline = .Font.Underline
190:        tpStyle.bLocked = .Locked
191:        tpStyle.bVisible = .visible
192:        tpStyle.cuFontSize = .Font.Size
193:        tpStyle.lBackColor = .BackColor
194:        tpStyle.lForeColor = .ForeColor
195:        tpStyle.sFontName = .Font.Name
196:        tpStyle.snHeight = .Height
197:        tpStyle.snWidth = .Width
198:
199:        tpStyle.lBackStyle = .BackStyle
200:        tpStyle.lBorderColor = .BorderColor
201:        tpStyle.lBorderStyle = .BorderStyle
202:    End With
203: End Sub
     Public Sub PasteStyleControl()
205:    If Application.VBE.ActiveWindow.Type <> vbext_wt_Designer Then Exit Sub
206:    Dim objActiveModule As VBComponent
207:    Dim cnt         As control
208:    Set objActiveModule = getActiveModule()
209:    For Each cnt In objActiveModule.Designer.Selected
210:        Call setPropertisControl(cnt)
211:    Next cnt
212: End Sub
     Public Sub PasteStyleForms()
214:    Dim cnt         As Object
215:    Set cnt = TakeSelectControl(True)
216:    If cnt Is Nothing Then Exit Sub
217:    Call setPropertisControl(cnt)
218: End Sub
     Private Sub setPropertisControl(ByVal cnt As Object)
220:    On Error Resume Next
221:    With cnt
222:        .Enabled = tpStyle.bEnabled
223:        .Font.Bold = tpStyle.bFontBold
224:        .Font.Italic = tpStyle.bFontItalic
225:        .Font.Strikethrough = tpStyle.bFontStrikethru
226:        .Font.Underline = tpStyle.bFontUnderline
227:        .Locked = tpStyle.bLocked
228:        .visible = tpStyle.bVisible
229:        .Font.Size = tpStyle.cuFontSize
230:        .BackColor = tpStyle.lBackColor
231:        .ForeColor = tpStyle.lForeColor
232:        .Font.Name = tpStyle.sFontName
233:        .Height = tpStyle.snHeight
234:        .Width = tpStyle.snWidth
235:
236:        .BackStyle = tpStyle.lBackStyle
237:        .BorderColor = tpStyle.lBorderColor
238:        .BorderStyle = tpStyle.lBorderStyle
239:    End With
240:    On Error GoTo 0
241: End Sub



     Public Sub AddIcon()
246:    Dim cnt         As control
247:    Dim objForm     As InsertIconUserForm
248:
249:    On Error GoTo ErrorHandler
250:
251:    Set cnt = TakeSelectControl
252:    If cnt Is Nothing Then Exit Sub
253:
254:    Set objForm = New InsertIconUserForm
255:    With objForm
256:        .Show
257:        cnt.Font.Name = .lbNameFont.Caption
258:        DoEvents
259:        If TypeName(cnt) = "Label" Then
260:            cnt.Caption = VBA.ChrW$(.lbASC.Caption)
261:        Else
262:            cnt.Value = VBA.ChrW$(.lbASC.Caption)
263:        End If
264:    End With
265:
266:    Exit Sub
ErrorHandler:
268:    Select Case Err.Number
        Case Else:
270:            Debug.Print "Fehler! в RenameControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
271:            Call WriteErrorLog("AddIcon")
272:    End Select
273:    Err.Clear
274: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : UperTextInControl
'* Created    : 01-07-2022 11:12
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub UperTextInControl()
283:    Call LowerAndUperTextInControl(True)
284: End Sub
     Public Sub LowerTextInControl()
286:    Call LowerAndUperTextInControl(False)
287: End Sub
     Private Sub LowerAndUperTextInControl(ByVal bUCase As Boolean)
289:    If Application.VBE.ActiveWindow.Type = vbext_wt_Designer Then
290:        Dim objActiveModule As VBComponent
291:        Set objActiveModule = getActiveModule()
292:        If Not objActiveModule Is Nothing Then
293:            If getSelectedControlsCollection.Count > 0 Then
294:                Dim ctl As control
295:                On Error Resume Next
296:                For Each ctl In objActiveModule.Designer.Selected
297:                    If bUCase Then
298:                        Call CallByName(ctl, "Caption", VbLet, VBA.UCase$(CallByName(ctl, "Caption", VbGet)))
299:                    Else
300:                        Call CallByName(ctl, "Caption", VbLet, VBA.LCase$(CallByName(ctl, "Caption", VbGet)))
301:                    End If
302:                Next ctl
303:                On Error GoTo 0
304:            End If
305:        End If
306:    End If
307: End Sub
     Public Sub UperTextInForm()
309:    Call LowerAndUperTextInForm(True)
310: End Sub
     Public Sub LowerTextInForm()
312:    Call LowerAndUperTextInForm(False)
313: End Sub
     Private Sub LowerAndUperTextInForm(ByVal bUCase As Boolean)
315:    Dim oVBComp     As VBIDE.VBComponent
316:    Set oVBComp = Application.VBE.SelectedVBComponent
317:    With oVBComp
318:        If .Type = vbext_ct_MSForm Then
319:            If bUCase Then
320:                .Properties("Caption") = VBA.UCase$(.Properties("Caption"))
321:            Else
322:                .Properties("Caption") = VBA.LCase$(.Properties("Caption"))
323:            End If
324:        End If
325:    End With
326: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : vbaCntAlingHoriz
'* Created    : 04-07-2022 14:39
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub vbaCntAlingHoriz()
335:    If Application.VBE.ActiveWindow.Type <> vbext_wt_Designer Then Exit Sub
336:
337:    Dim lCnt        As Long
338:    Dim dTop        As Double
339:    Dim dLeft       As Double
340:    Dim dHeight     As Double
341:    Dim dWidth      As Double
342:    Dim dSPACE      As Variant
343:    Dim lColCnt     As Variant
344:    Dim dStart      As Double
345:    Dim dMaxWidth   As Double
346:    Dim objActiveModule As VBComponent
347:    Set objActiveModule = getActiveModule()
348:
349:    lColCnt = Application.InputBox("Введите количество строк", "Выровнять по горизонтальной сетке:", Type:=1)
350:    If lColCnt <= 0 Or lColCnt = False Then
351:        Exit Sub
352:    End If
353:    dSPACE = Application.InputBox("Введите расстояние между фигурами", "Выровнять по горизонтальной сетке:", Type:=1)
354:    If TypeName(dSPACE) = "Boolean" Then
355:        Exit Sub
356:    End If
357:    lCnt = 1
358:    Dim cnt         As control
359:
360:    For Each cnt In objActiveModule.Designer.Selected
361:        With cnt
362:            If lCnt = 1 Then
363:                dStart = .Top
364:            Else
365:                If lCnt Mod lColCnt = 1 Or lColCnt = 1 Then
366:                    'New column, move shape right
367:                    .Top = dStart
368:                    .Left = dLeft + dMaxWidth + dSPACE
369:                    dMaxWidth = .Width
370:                Else
371:                    'Same column, move shape down
372:                    .Top = dTop + dHeight + dSPACE
373:                    .Left = dLeft
374:                End If
375:            End If
376:            dTop = .Top
377:            dLeft = .Left
378:            dHeight = .Height
379:            dWidth = .Width
380:            dMaxWidth = WorksheetFunction.Max(dMaxWidth, .Width)
381:        End With
382:        lCnt = lCnt + 1
383:    Next cnt
384: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : vbaCntAlingVert
'* Created    : 04-07-2022 14:39
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub vbaCntAlingVert()
394:    If Application.VBE.ActiveWindow.Type <> vbext_wt_Designer Then Exit Sub
395:
396:    Dim lCnt        As Long
397:    Dim dTop        As Double
398:    Dim dLeft       As Double
399:    Dim dHeight     As Double
400:    Dim dWidth      As Double
401:    Dim dSPACE      As Variant
402:    Dim lColCnt     As Variant
403:    Dim dStart      As Double
404:    Dim dMaxHeight  As Double
405:    Dim objActiveModule As VBComponent
406:    Set objActiveModule = getActiveModule()
407:
408:    lColCnt = Application.InputBox("Введите количество столбцов", "Выровнять по вертикальной сетке:", Type:=1)
409:    If lColCnt <= 0 Or lColCnt = False Then
410:        Exit Sub
411:    End If
412:    dSPACE = Application.InputBox("Введите расстояние между фигурами", "Выровнять по вертикальной сетке:", Type:=1)
413:    If TypeName(dSPACE) = "Boolean" Then
414:        Exit Sub
415:    End If
416:    lCnt = 1
417:    Dim cnt         As control
418:    For Each cnt In objActiveModule.Designer.Selected
419:        With cnt
420:            If lCnt = 1 Then
421:                dStart = .Left
422:            Else
423:                If lCnt Mod lColCnt = 1 Or lColCnt = 1 Then
424:                    .Top = dTop + dMaxHeight + dSPACE
425:                    .Left = dStart
426:                    dMaxHeight = .Height
427:                Else
428:                    .Top = dTop
429:                    .Left = dLeft + dWidth + dSPACE
430:                End If
431:            End If
432:            dTop = .Top
433:            dLeft = .Left
434:            dHeight = .Height
435:            dWidth = .Width
436:            dMaxHeight = WorksheetFunction.Max(dMaxHeight, .Height)
437:        End With
438:        lCnt = lCnt + 1
439:    Next cnt
440: End Sub
'* общие функции**********************************************************
     Private Function TakeSelectControl(Optional bUserForm As Boolean = False) As Object
443:    On Error GoTo ErrorHandler
444:    If Application.VBE.ActiveWindow.Type = vbext_wt_Designer Then
445:        Dim objActiveModule As VBComponent
446:        Set objActiveModule = getActiveModule()
447:        If Not objActiveModule Is Nothing Then
448:            If getSelectedControlsCollection.Count = 1 Then
449:                Dim ctl As control
450:                For Each ctl In objActiveModule.Designer.Selected
451:                    Set TakeSelectControl = ctl
452:                    Exit Function
453:                Next ctl
454:            End If
455:        End If
456:    End If
457:
458:    Dim Form        As UserForm
459:    Set Form = Application.VBE.SelectedVBComponent.Designer
460:    If bUserForm And Not Form Is Nothing Then
461:        Set TakeSelectControl = Form
462:        Exit Function
463:    End If
464:
465:    Exit Function
ErrorHandler:
467:    Select Case Err.Number
        Case 9:
469:            Debug.Print "Для работы инструмента, откройте окно View -> Properties Window"
470:        Case Else:
471:            Debug.Print "Fehler! в TakeSelectControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
472:            Call WriteErrorLog("TakeSelectControl")
473:    End Select
474:    Err.Clear
475: End Function
     Public Function getSelectedControlsCollection() As Collection
477:    Dim ctl         As control
478:    Dim out         As New Collection
479:    Dim Module      As VBComponent
480:    Set Module = getActiveModule
481:    For Each ctl In Module.Designer.Selected
482:        out.Add ctl
483:    Next ctl
484:    Set getSelectedControlsCollection = out
485:    Set out = Nothing
486: End Function
     Public Function getActiveModule() As VBComponent
488:    Set getActiveModule = Application.VBE.SelectedVBComponent
489: End Function



