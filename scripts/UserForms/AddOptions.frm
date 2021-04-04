VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddOptions 
   Caption         =   "OPTION:"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   OleObjectBlob   =   "AddOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : addOptions - �������� OPTIONs � ������� �������
'* Created    : 17-09-2020 14:06
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

    Private Sub chAll_Change()
11:    Dim bFlag       As Boolean
12:    bFlag = chAll.Value
13:    chOptionExplicit.Value = bFlag
14:    chOptionPrivate.Value = bFlag
15:    chOptionCompare.Value = bFlag
16:    chOptionBase.Value = bFlag
17: End Sub

    Private Sub lbOK_Click()
20:    Unload Me
21: End Sub

    Private Sub lbBase_Click()
24:    Dim stxt        As String
25:    stxt = "������������ �� ������ ������ ��� ���������� ������ ������� ��������, �� ���������." & vbNewLine & vbNewLine
26:    stxt = stxt & "���������" & vbNewLine & "Option Base { 0 | 1 }" & vbNewLine & vbNewLine
27:    stxt = stxt & "��������� Option Base �� ��������� ����� 0, �������� Option Base ������� �� ������������. �������� ������ ���������� � ������ �� ���� ��������." & vbNewLine
28:    stxt = stxt & "�������� Option Base ����� ����������� � ������ ������ ���� ��� � ������ �������������� ����������� ��������, ���������� �����������." & vbNewLine & vbNewLine
29:    stxt = stxt & "����������" & vbNewLine & vbNewLine
30:    stxt = stxt & "����������� To � ����������� Dim, Private, Public, ReDim � Static ������������� ����� ������ ������ ���������� ���������� �������� �������." & vbNewLine
31:    stxt = stxt & "������ ���� ������ ������� �������� �� �������� ���� � ����������� To, ����� ��������������� ����������� Option Base," & vbNewLine
32:    stxt = stxt & "����� ���������� ������������ �� ��������� ������ ������� ��������, ������ 1. ������ ������� �������� �������� ��������," & vbNewLine
33:    stxt = stxt & "����������� � ������� ������� Array, ������ ��������� ����; ��� ����������� �� ���������� Option Base."
34:    stxt = stxt & vbNewLine & vbNewLine & "���������� Option Base ��������� �� ������ ������� �������� �������� ������ ���� ������, � ������� ����������� ���� ��� ����������."
35:    Debug.Print stxt
36: End Sub
    Private Sub lbCompare_Click()
38:    Dim stxt        As String
39:    stxt = "������������ �� ������ ������ ��� ���������� ������ ��������� �� ���������, ������� ����� �������������� ��� ��������� ��������� ������." & vbNewLine & vbNewLine
40:    stxt = stxt & "���������" & vbNewLine & "Option Compare { Binary | Text | Database }" & vbNewLine & vbNewLine
41:    stxt = stxt & "����������" & vbNewLine & vbNewLine
42:    stxt = stxt & "���������� Option Compare ��� �� ������������� ������ ���������� � ������ ����� ����� ����������." & vbNewLine
43:    stxt = stxt & "���������� Option Compare ��������� ������ ��������� ����� (Binary, Text ��� Database) ��� ������." & vbNewLine
44:    stxt = stxt & "���� ������ �� �������� ���������� Option Compare, �� ��������� ������������ ������ ��������� Binary." & vbNewLine
45:    stxt = stxt & "���������� Option Compare Binary ������ ��������� ����� �� ������ ������� ����������, ������������� ���������� �������� �������������� ��������." & vbNewLine
46:    stxt = stxt & "� Microsoft Windows ������� ���������� ������������ ������� ��������� ��������." & vbNewLine
47:    stxt = stxt & "� ��������� ������� ����������� �������� ��������� ��������� ������� ����������:" & vbNewLine & vbNewLine
48:    stxt = stxt & "A < B < E < Z < a < b < e < z < � < � < � < � < � < �" & vbNewLine & vbNewLine
49:    stxt = stxt & "���������� Option Compare Text ������ ��������� ����� ��� ����� �������� �������� �� ������ ��������� ������������ ���������." & vbNewLine
50:    stxt = stxt & "��� �� ��������, ��� � ����, ��� ���������� � ����������� Option Compare Text ������������� ��������� �������: " & vbNewLine & vbNewLine
51:    stxt = stxt & "(A=a) < (B=b) < (E=e) < (Z=z) < (�=�) < (�=�) < (�=�)" & vbNewLine & vbNewLine
52:    stxt = stxt & "���������� Option Compare Database ����� �������������� ������ � Microsoft Access. ��� ���� ������ ��������� ����� �� ������ ������� ����������," & vbNewLine
53:    stxt = stxt & "������������� ������������ ���������� ���� ������, � ������� ������������ ��������� �����. "
54:    Debug.Print stxt
55: End Sub
    Private Sub lbExplicit_Click()
57:    Dim stxt        As String
58:    stxt = "������������ �� ������ ������ ��� ��������������� ������ ���������� ���� ���������� � ���� ������." & vbNewLine & vbNewLine
59:    stxt = stxt & "���������" & vbNewLine & "Option Explicit" & vbNewLine & vbNewLine
60:    stxt = stxt & "����������" & vbNewLine & vbNewLine
61:    stxt = stxt & "���������� Option Explicit ��� �� ������������� ������ ���������� � ������ �� ����� ���������." & vbNewLine
62:    stxt = stxt & "��� ������������� ���������� Option Explicit ���������� ���� ������� ��� ���������� � ������� ���������� Dim, Private, Public, ReDim ��� Static." & vbNewLine
63:    stxt = stxt & "��� ������� ������������ ����������� ��� ���������� ��������� ������ �� ����� ����������." & vbNewLine
64:    stxt = stxt & "����� ���������� Option Explicit �� ������������, ��� ����������� ���������� ����� ��� Variant, ���� ������������ �� ��������� ��� ������ �� �������� � ������� ���������� Def���." & vbNewLine
65:    stxt = stxt & "����������� ���������� Option Explicit, ����� �������� ��������� ����� ����� ��������� ���������� ��� ����� ���������� � ���������, ����� ������� ����������� ���������� �� ������ ����."
66:    Debug.Print stxt
67: End Sub
    Private Sub lbPrivate_Click()
69:    Dim stxt        As String
70:    stxt = "������������ �� ������ ������ ��� ������� ������ �� ������� ������ ����� �������." & vbNewLine & vbNewLine
71:    stxt = stxt & "���������" & vbNewLine & "Option Private Module" & vbNewLine & vbNewLine
72:    stxt = stxt & "����������" & vbNewLine & vbNewLine
73:    stxt = stxt & "����� ������ �������� ���������� Option Private Module, ����� ��������, ��������, ����������, ������� � ������������ ������������� ����, ��������� �� ������ ������," & vbNewLine
74:    stxt = stxt & "�������� ���������� ������ �������, ����������� ���� ������, �� ������������ ��� ������ ���������� ��� ��������." & vbNewLine
75:    stxt = stxt & "Microsoft Excel ������������ �������� ���������� ��������. � ���� ������ ���������� Option Private Module ��������� ���������� �������� ��������� ��������."
76:    Debug.Print stxt
77: End Sub

    Private Sub cmbCancel_Click()
80:    Unload Me
81: End Sub
    Private Sub lbCancel_Click()
83:    Call cmbCancel_Click
84: End Sub

Private Sub UserForm_Activate()
87:    On Error GoTo ErrorHandler
88:
89:    lbExplicit.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
90:    lbPrivate.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
91:    lbCompare.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
92:    lbBase.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
93:
94:    lbModule.Caption = Application.VBE.ActiveCodePane.CodeModule.Parent.Name
95:
96:    Exit Sub
ErrorHandler:
98:    Select Case Err.Number
        Case 91:
100:            Unload Me
101:            Debug.Print "��� ��������� ������, ��������� � ������ ����!"
102:            Exit Sub
103:        Case 76:
104:            Exit Sub
105:        Case Else:
106:            Debug.Print "������! � addOptions" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "� ������ " & Erl
107:            Call WriteErrorLog("addOptions")
108:    End Select
109:    Err.Clear
End Sub
