Attribute VB_Name = "modAddCodeViewForm"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : G_AddCodeViewForm - управление формой создания сниппетов
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit
    Public Sub AddCode(ByVal X As Long)
4:    Dim frm    As AddEditCode
5:    Set frm = New AddEditCode
6:    With frm
7:        .Caption = "CREATE SNIPPET:"
8:        .lbOK.Caption = "to create"
9:        .txtRow = X + 1
10:        .Show
11:    End With
12: End Sub
    Public Sub EditCode(ByVal X As Long, ByRef ListB As MSForms.ListBox)
14:    Dim frm    As AddEditCode
15:    Dim snippets As ListObject
16:    Dim st()   As String
18:    If X <= 0 Then
19:        Call MsgBox("Nothing is selected in the table!", vbCritical, "Nothing is selected:")
20:        Exit Sub
21:    End If
    'X is the row in ListRows; the Range of a ListColumn below also counts the header row
    X = X + 1
22:    Set frm = New AddEditCode
23:    Set snippets = SHSNIPPETS.ListObjects(modConst.TB_SNIPPETS)
24:    With frm
25:        .Caption = "CHANGE SNIPPET:"
26:        .lbOK.Caption = "to change"
27:        st = Split(snippets.ListColumns(3).Range(X, 1), ".")
28:        .cmbENUM.Style = fmStyleDropDownCombo
29:        .cmbENUM.Text = st(0)
30:        .txtENUMBack.Text = .cmbENUM.Text
31:        .txtSNIP.Text = snippets.ListColumns(2).Range(X, 1)
32:        .txtSNIPBack.Text = .txtSNIP.Text
33:        .txtCode.Text = snippets.ListColumns(4).Range(X, 1)
34:        .txtCodeBack.Text = .txtCode.Text
35:        .cmbOBJ.Value = snippets.ListColumns(5).Range(X, 1)
36:        .txtRow = X
37:        .Show
38:    End With
39: End Sub
    Public Sub DeletRow(ByVal X As Long, ByRef objList As MSForms.ListBox)
41:    Dim snippets As ListObject
42:    On Error GoTo errMsg
43:    Set snippets = SHSNIPPETS.ListObjects(modConst.TB_SNIPPETS)
44:    If X <= 0 Then
45:        Call MsgBox("Nothing is selected in the table!", vbCritical, "Nothing is selected:")
46:        Exit Sub
47:    End If
48:    If MsgBox("Remove SNIPPET: [" & snippets.DataBodyRange.Cells(X, 2).Value & " ] ?", vbYesNo, "Removing SNIPPET:") = vbYes Then
49:        snippets.ListRows.Item(X).Delete
50:        Call modAddCodeViewForm.TbAdd(objList)
51:    End If
52:    Exit Sub
errMsg:
54:    If Err.Number = 91 Then
55:        Err.Clear
56:    Else
57:        Debug.Print "Error in DeletRow!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
58:        Call WriteErrorLog("DeletRow")
59:    End If
60: End Sub
    Public Sub TbAdd(ByRef objList As MSForms.ListBox)
62:    Dim snippets As ListObject
63:    Dim i      As Long
64:    On Error GoTo errMsg
65:    Set snippets = SHSNIPPETS.ListObjects(modConst.TB_SNIPPETS)
66:    With objList
67:        .Clear
68:        For i = 1 To snippets.DataBodyRange.Rows.Count
69:            .AddItem snippets.ListColumns(5).Range(i + 1, 1).Value
70:            .List(i - 1, 1) = snippets.ListColumns(1).Range(i + 1, 1).Value
71:            .List(i - 1, 2) = snippets.ListColumns(2).Range(i + 1, 1).Value
72:            .List(i - 1, 3) = snippets.ListColumns(3).Range(i + 1, 1).Value
73:        Next i
74:    End With
75:    Exit Sub
errMsg:
77:    If Err.Number = 91 Then
78:        Err.Clear
79:    Else
80:        Debug.Print "Error in TbAdd!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
81:        Call WriteErrorLog("TbAdd")
82:    End If
83: End Sub

'Row of the snippet named sName in the table SNIPPETS as index into ListRows (1 = first data row), 0 if no snippet
'has this name. The name (column 2) is the key of a snippet; it is compared without regard to case, as InsertCode
'finds snippets.
Public Function SnippetRow(ByVal sName As String) As Long
    Dim snippets As ListObject
    Dim vName    As Variant
    Dim i        As Long
    Set snippets = SHSNIPPETS.ListObjects(modConst.TB_SNIPPETS)
    For i = 1 To snippets.ListRows.Count
        vName = snippets.ListColumns(2).DataBodyRange.Cells(i, 1).Value2
        If Not IsError(vName) Then
            If StrComp(CStr(vName), sName, vbTextCompare) = 0 Then
                SnippetRow = i
                Exit Function
            End If
        End If
    Next i
End Function

'Row in the table SNIPPETS (index into ListRows) of the snippet shown as item lIndex of objList, whose column 2 holds
'the snippet name; 0 if lIndex is no item (-1: nothing selected) or the name is not in the table.
Public Function ListItemSnippetRow(ByRef objList As MSForms.ListBox, ByVal lIndex As Long) As Long
    If lIndex < 0 Or lIndex >= objList.ListCount Then Exit Function
    ListItemSnippetRow = SnippetRow(objList.List(lIndex, 2) & vbNullString)
End Function
