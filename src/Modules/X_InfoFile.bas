Attribute VB_Name = "X_InfoFile"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : X_InfoFile - модуль редактирования свойств книги excel
'* Created    : 20-07-2020 12:31
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ShowProp функция создания списка свойств файла
'* Created    : 20-07-2020 12:32
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook : ссылка на книгу
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function ShowProp(ByRef wb As Workbook) As String
22:    Dim DP          As DocumentProperty
23:    Dim Txt         As String
24:
25:    With wb
26:        On Error Resume Next
27:        For Each DP In .BuiltinDocumentProperties
28:            Txt = Txt & DP.Name & "||" & " " & DP.Value & vbNewLine
29:        Next
30:        On Error GoTo 0
31:    End With
32:    ShowProp = VBA.Left$(Txt, VBA.Len(Txt) - 2)
33: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DelAllProp - функция удаления всех свойств книги, возращает количество удаленных свойств
'* Created    : 20-07-2020 12:33
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook : ссылка на книгу
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function DelAllProp(ByRef wb As Workbook) As Byte
47:    Dim DP          As DocumentProperty
48:    Dim byItem      As Byte
49:    With wb
50:        On Error Resume Next
51:        For Each DP In .BuiltinDocumentProperties
52:            If DP.Value <> vbNullString Then
53:                DP.Value = vbNullString
54:                byItem = byItem + 1
55:            End If
56:        Next
57:        On Error GoTo 0
58:    End With
59:    DelAllProp = byItem
60: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WriteOneProp - редактирование значения одного выбраного свойства
'* Created    : 20-07-2020 12:34
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByRef WB As Workbook     : ссылка на книгу
'* ByVal NameProp As String : название совйства
'* ByVal Val As String      : новое значение свойства
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function WriteOneProp(ByRef wb As Workbook, ByVal NameProp As String, ByVal Val As String) As String
76:    Dim DP          As DocumentProperty
77:    On Error GoTo errMsg
78:    Set DP = wb.BuiltinDocumentProperties(NameProp)
79:    DP.Value = Val
80:    WriteOneProp = True
81:    Exit Function
errMsg:
83:    WriteOneProp = Err.Description
84:    On Error GoTo 0
85: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ShowCustomProp - создания списка пользовательских свойств файла
'* Created    : 20-07-2020 12:35
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook : ссылка на книгу
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function ShowCustomDocProp(ByRef wb As Workbook) As String
99:    Dim i           As Integer
100:    Dim Txt         As String
101:    With wb
102:        If .CustomDocumentProperties.Count > 0 Then
103:            For i = 1 To .CustomDocumentProperties.Count
104:                Txt = Txt & .CustomDocumentProperties(i).Name & "||" & " " & .CustomDocumentProperties(i).Value & vbNewLine
105:            Next i
106:        End If
107:    End With
108:    ShowCustomDocProp = Txt
109: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ShowCustomSheetProp - creates a list of all worksheets' custom properties
'* Created    : 28-07-2026
'* Author     : Erek
'* Argument(s):             Description
'*
'* ByRef WB As Workbook : reference to the workbook
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ShowCustomSheetProp(ByRef wb As Workbook) As String
    Dim ws          As Worksheet
    Dim i           As Integer
    Dim Txt         As String

    For Each ws In wb.Worksheets
        If ws.CustomProperties.Count > 0 Then
            For i = 1 To ws.CustomProperties.Count
                Txt = Txt & ws.Name & "||" & ws.CustomProperties(i).Name & "||" & " " & ws.CustomProperties(i).Value & vbNewLine
            Next i
        End If
    Next ws

    ShowCustomSheetProp = Txt
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DelAllCustomProp - функция удаления всех пользовательских свойств книги, возращает количество удаленных свойств
'* Created    : 20-07-2020 12:35
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook : ссылка на книгу
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function DelAllCustomDocProp(ByRef wb As Workbook) As Byte
123:    Dim i           As Integer
124:    Dim byItem      As Byte
125:    With wb
126:        If .CustomDocumentProperties.Count > 0 Then
127:            For i = .CustomDocumentProperties.Count To 1 Step -1
128:                .CustomDocumentProperties(i).Delete
129:                byItem = byItem + 1
130:            Next i
131:        End If
132:    End With
133:    DelAllCustomDocProp = byItem
134: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DelOneCustomProp - удаление выпраного пользовательского свойства
'* Created    : 20-07-2020 12:36
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByRef WB As Workbook     : ссылка на книгу
'* ByVal NameProp As String : название совйства
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function DelOneCustomDocProp(ByRef wb As Workbook, ByVal NameProp As String) As Boolean
149:    Dim i           As Integer
150:    With wb
151:        If .CustomDocumentProperties.Count > 0 Then
152:            For i = 1 To .CustomDocumentProperties.Count
153:                If .CustomDocumentProperties(i).Name = NameProp Then
154:                    .CustomDocumentProperties(i).Delete
155:                    DelOneCustomDocProp = True
156:                    Exit Function
157:                End If
158:            Next i
159:        End If
160:    End With
161: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddOneCustomProp - процедура создания пользовательского свойства
'* Created    : 20-07-2020 12:37
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByRef WB As Workbook     : ссылка на книгу
'* ByVal NameProp As String : название совйства
'* ByVal Val As Variant     : значение свойства
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub AddOneCustomDocProp(ByRef wb As Workbook, ByVal NameProp As String, ByVal Val As Variant)
177:    Call DelOneCustomDocProp(wb, NameProp)
178:    Call wb.CustomDocumentProperties.Add(NameProp, False, msoPropertyTypeString, VBA.CStr(Val))
179: End Sub
Public Function GetOneProp(ByRef wb As Workbook, ByVal NameProp As String) As String
181:    GetOneProp = wb.BuiltinDocumentProperties(NameProp).Value
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ShowDefinedNames - creates a list of all defined names (Name Manager)
'* Created    : 28-07-2026
'* Author     : Erek
'* Argument(s):             Description
'*
'* ByRef WB As Workbook : reference to the workbook
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ShowDefinedNames(ByRef wb As Workbook) As String
    Dim nm          As Name
    Dim Txt         As String
    Dim RefersToStr As String
    Dim ScopeStr    As String
    Dim ValueStr    As String
    Dim EvalResult  As Variant
    Dim nameStr As String

    For Each nm In wb.Names
        On Error Resume Next
        RefersToStr = nm.RefersTo
        If Err.Number <> 0 Then
            RefersToStr = "#ERROR"
            Err.Clear
        End If
        On Error GoTo 0

        If TypeName(nm.Parent) = "Worksheet" Then
            ScopeStr = nm.Parent.Name
        Else
            ScopeStr = "Workbook"
        End If

        ' aktuellen Wert ermitteln, wie in der "Value"-Spalte des Namens-Managers
        On Error Resume Next
        Err.Clear
        EvalResult = Application.Evaluate(nm.RefersTo)
        If Err.Number <> 0 Then
            ValueStr = "#ERROR"
        ElseIf IsArray(EvalResult) Then
            ValueStr = "{Array}"
        ElseIf IsError(EvalResult) Then
            ValueStr = "#REF_ERROR"
        Else
            ValueStr = CStr(EvalResult)
        End If
        Err.Clear
        On Error GoTo 0

        nameStr = Replace$(nm.Name, ScopeStr & "!", "")
        Txt = Txt & nameStr & "||" & ValueStr & "||" & RefersToStr & "||" & ScopeStr & "||" & nm.visible & vbNewLine
    Next nm

    ShowDefinedNames = Txt
End Function
