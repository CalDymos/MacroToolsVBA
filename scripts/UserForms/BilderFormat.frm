VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BilderFormat 
   Caption         =   "����������� ��������:"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10455
   OleObjectBlob   =   "BilderFormat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BilderFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************************
' Author         : VBATools - ����������� �������� �����
' Date           : 15.09.2019
' �������� ����� : info@VBATools.ru
' Copyright      : VBATools.ru
'***********************************************************************************************************
Option Explicit
Private arrFormatDate As Variant
Private arrFormatDateDiscription As Variant
Private arrFormatValue As Variant
Private arrFormatValueDiscription As Variant
Private arrFormatDateCustom As Variant
Private arrFormatDateCustomDiscription As Variant

    Private Function AddCode() As String
10:    Dim sSTR As String, sErr As String
11:
12:    If obtnFormat Then
13:        sSTR = cmbFormat.Value
14:    Else
15:        sSTR = LTrim(txtCustom.Text)
16:    End If
17:
18:    sErr = vbNullString
19:    If lbView.Caption = vbNullString Then sErr = "� ���� ����� �������� �����!" & vbNewLine
20:    If lbView.Caption Like "������: *" Then sErr = sErr & "������, � �������� �������"
21:    If sErr <> vbNullString Then
22:        Call MsgBox(sErr, vbCritical, "������:")
23:        AddCode = vbNullString
24:        Exit Function
25:    End If
26:
27:    sSTR = "VBA.Format$(" & Replace(txtValue.Value, ",", ".") & ", " & Chr(34) & sSTR & Chr(34) & ")"
28:    AddCode = sSTR
29: End Function

    Private Sub lbHelp_Click()
32:    Call URLLinks(C_Const.URL_BILD_FOFMAT)
33: End Sub

    Private Sub lbInsertCode_Click()
36:    Dim iLine  As Integer
37:    Dim txtCode As String, txtLine As String
38:    '��������� ����
39:    txtCode = AddCode()
40:    If txtCode = vbNullString Then Exit Sub
41:    txtLine = C_PublicFunctions.SelectedLineColumnProcedure
42:    If txtLine = vbNullString Then
43:        Me.Hide
44:        Exit Sub
45:    End If
46:    iLine = Split(txtLine, "|")(2)
47:
48:    With Application.VBE.ActiveCodePane
49:        .CodeModule.InsertLines iLine, txtCode
50:    End With
51:    Me.Hide
52: End Sub
    Private Sub btnCopyCode_Click()
54:    Dim sSTR As String, sMsgBoxString As String
55:
56:    sSTR = AddCode()
57:    If sSTR = vbNullString Then Exit Sub
58:    Call C_PublicFunctions.SetTextIntoClipboard(sSTR)
59:
60:    sMsgBoxString = "��� ���������� � ����� ������!" & vbNewLine & "��� ������� ���� ����������� " & Chr(34) & "Ctrl+V" & Chr(34)
61:    Call MsgBox(sMsgBoxString, vbInformation, "����������� ����:")
62:
63:    Me.Hide
64: End Sub

    Private Sub cmbCancel_Click()
67:    Me.Hide
68: End Sub

    Private Sub cmbCustomFormat_Change()
71:    If cmbCustomFormat.ListIndex = -1 Then Exit Sub
72:    If obtnCustomFormat Then
73:        txtDiscription.Text = arrFormatDateCustomDiscription(cmbCustomFormat.ListIndex)
74:    End If
75: End Sub
    Private Sub txtCustom_Change()
77:    Call AddFormat(txtCustom.Text)
78: End Sub
    Private Sub lbClear_Click()
80:    txtCustom.Text = vbNullString
81: End Sub
    Private Sub cmbFormat_Change()
83:    If cmbFormat.ListIndex = -1 Then Exit Sub
84:    If obtnDate And obtnFormat Then
85:        txtDiscription.Text = arrFormatDateDiscription(cmbFormat.ListIndex)
86:    Else
87:        txtDiscription.Text = arrFormatValueDiscription(cmbFormat.ListIndex)
88:    End If
89:
90:    Call AddFormat(cmbFormat.Value)
91: End Sub

    Private Sub lbAddCustom_Click()
94:    If cmbCustomFormat.Value <> vbNullString Then
95:        txtCustom.Text = txtCustom & " " & cmbCustomFormat.Value
96:    End If
97: End Sub

     Private Sub obtnDate_Change()
100:    Call AddList
101: End Sub

     Private Sub AddList()
104:    If obtnFormat Then
105:        cmbCustomFormat.Clear
106:        If obtnDate Then
107:            cmbFormat.List = arrFormatDate
108:        Else
109:            cmbFormat.List = arrFormatValue
110:        End If
111:        obtnValue.visible = True
112:    Else
113:        cmbFormat.Clear
114:        cmbCustomFormat.List = arrFormatDateCustom
115:        obtnValue.visible = False
116:    End If
117:    Call ChengeFlag(obtnFormat)
118: End Sub
     Private Sub ChengeFlag(ByVal Flag As Boolean)
120:    obtnDate.visible = Flag
121:    cmbFormat.visible = Flag
122:    cmbCustomFormat.visible = (Not Flag)
123:    txtCustom.visible = (Not Flag)
124:    lbClear.visible = (Not Flag)
125:    lbAddCustom.visible = (Not Flag)
126:    If Flag Then
127:        Frame2.Left = lbAddCustom.Left - 5
128:    Else
129:        Frame2.Left = cmbCustomFormat.Left + cmbCustomFormat.Width + 5
130:    End If
131: End Sub
     Private Sub obtnFormat_Click()
133:    Call AddList
134: End Sub
     Private Sub obtnCustomFormat_Click()
136:    Call AddList
137: End Sub
     Private Sub txtValue_Change()
139:    If cmbFormat.Value = vbNullString Then
140:        Call AddFormat(cmbCustomFormat.Value)
141:    Else
142:        Call AddFormat(cmbFormat.Value)
143:    End If
144: End Sub
     Private Sub AddFormat(ByVal sSTR As String)
146:    On Error GoTo err_msg
147:    lbView.Caption = Format(CDbl(txtValue.Value), sSTR)
148:    lbView.ForeColor = &H8000000D
149:    Exit Sub
err_msg:
151:    Select Case Err.Number
        Case 6
153:            lbView.Caption = "������: ������ ������� �����!"
154:        Case 13
155:            lbView.Caption = vbNullString
156:        Case Else
157:            lbView.Caption = "������: " & Err.Description & " " & Err.Number
158:    End Select
159:    lbView.ForeColor = &H8080FF
160:    Err.Clear
161: End Sub

     Private Sub txtValue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
164:    Dim Txt    As String
165:    Txt = Me.txtValue    ' ������ ����� �� ���� (��� ����������� ����� ���� � ����� �������)
166:    Select Case KeyAscii
        Case 8:    ' ����� Backspace - ������ �� ������
168:        Case 44: KeyAscii = IIf(InStr(1, Txt, ",") > 0, 0, 44)    ' ���� ������� ��� ���� - �������� ���� �������
169:        Case 46: KeyAscii = IIf(InStr(1, Txt, ",") > 0, 0, 44)    ' �������� ��� ����� ����� �� �������
170:        Case 48 To 57    ' ���� ������� �����  - ������ �� ������
171:        Case Else: KeyAscii = 0    ' ����� �������� ���� �������
172:    End Select
173: End Sub

     Private Sub UserForm_Initialize()
176:    arrFormatDate = Array("General Date", "Long Date", "Medium Date", "Short Date", "Long Time", "Medium Time", "Short Time")
177:    arrFormatDateDiscription = Array("����������� ���� �/��� �������, �������� 4/3/93 05:34 PM. ���� ������� ����� �����������, ������������ ������ ����, �������� 4/3/93. ���� ����������� ����� �����, ������������ ������ �����, �������� 05:34 PM. ����������� ���� ������������ ����������� �������.", _
                "����������� ���� � ������������ � ������� �������� ����, ������������ � �������.", _
                "����������� ���� � �������������� �������� ������� ����, ���������������� �������� ������ �������� ����������.", _
                "����������� ���� � ������������ � ������� �������� ����, ������������ � �������.", _
                "����������� ������� � ������������ � ������� �������� ����, ������������ � �������; �������� ����, ������ � �������.", _
                "����������� ������� � 12-������� ������� � �������������� �����, ����� � ��������� AM/PM.", _
                "����������� ������� � 24-������� �������, �������� 17:45.")
184:    arrFormatValue = Array("General Number", "Currency", "Fixed", "Standard", "Percent", "Scientific", "Yes/No", "True/False", "On/Off")
185:    arrFormatValueDiscription = Array("����������� ����� ��� ����� ����������� ����� ��������.", _
                "����������� ����� � �������������� ����������� ����� ��������, ���� ��� ����������; ������������ ��� ����� ������ �� ����������� ����� � ������� �����. ����� ������������ �� �������� ���������� �������.", _
                "����������� �� ������� ���� ����� ����� ����� � ���� ���� ������ �� ����������� ����� � ������� �����.", _
                "����������� ����� � �������������� ����������� ����� ��������; ������������ �� ������� ���� ���� ����� ����� � ��� ����� ������ �� ����������� ����� � ������� �����.", _
                "����������� �����, ����������� �� 100 �� ������ �������� (%), ������������ ������; ������ ������������ ��� ����� ������ �� ����������� ����� � ������� �����.", _
                "������������ ����������� ���������������� �������������.", _
                "������������ ���, ���� ����� ��������� 0; � ��������� ������ ������������ ��.", _
                "������������ False, ���� ����� ��������� 0; � ��������� ������ ������������ True.", _
                "������������ ���, ���� ����� ��������� 0; � ��������� ������ ������������ ����.")
194:    arrFormatDateCustom = Array("c", "d", "dd", "ddd", "dddd", "ddddd", "dddddd", "w", "ww", "m", "mm", "mmm", "mmmm", "q", "y", "yy", "yyyy", "h", "hh", "n", "nn", "s", "ss", "ttttt")
195:    arrFormatDateCustomDiscription = Array("����������� ����������� ����. � ��������� �������� ���������� ����� �������������� ������ ����� ��� ������������� ����������� ����������� ����. ���� ����������� �������� ����, ����� � ���, ����� �������� ���� �������������. ������, ������������ � �������� ����������� ����������� ���� � ����������������� �������� ������, ������������ ����������� �������.", _
                "����p. ��� � ���� ����� ��� ���� � ������ (1�31).", _
                "�����. ��� � ���� ����� � ����� � ������ (01�31).", _
                "�����. ��� � �������������� ���������� (����). ������������.", _
                "�����. ��� � �������������� ������� ����� (������������������). ������������.", _
                "�����. ���� � �������������� ������� ������� (������� ����, ����� � ���), ���������������� �������� ������� ���� � ���������� �������. ������� �������� ���� �� ��������� �������� m/d/yy.", _
                "�����. �����, ��������������� ����, � �������������� ������� ������� (������� ����, ����� � ���), ���������������� �������� ������� ���� � ���������� �������. ������� �������� ���� �� ��������� �������� mmmm dd, yyyy.", _
                "�����. ��� ������ � ���� ����� (�� 1 ��� ����������� � �� 7 ��� �������).", _
                "�����. ������ ���� � ���� ����� (1�54).", _
                "�����. ������ � ���� ����� ��� ���� � ������ (1�12). ���� m ������� ����� �� ����� h ��� hh, ������������ ����� �� �����, � ������.", _
                "�����. ������ � ���� ����� � ����� � ������ (01�12). ���� m ������� ����� �� ����� h ��� hh, ������������ ����� �� �����, � ������.", _
                "�����. ������������ �������� ������ (������). ������������.", _
                "�����. ������� �������� ������ (��������������). ������������.", _
                "�����. �������� ���� � ���� ����� (1�4).", "����������� ��� ���� � ���� ����� (1�366).", "�����. ���� � ���� 2-�������� ����� (00�99).", _
                "�����. ���� � ���� 4-�������� ����� (100�9999).", "�����. ���� � ���� ����� ��� ���� � ������ (0�23).", _
                "�����. ���� � ���� ����� � ����� � ������ (00�23).", "�����. ������ � ���� ����� ��� ���� � ������ (0�59).", _
                "�����. ������ � ���� ����� � ����� � ������ (00�59).", _
                "�����. ������� � ���� ����� ��� ���� � ������ (0�59).", _
                "�����. ������� � ���� ����� � ����� � ������ (00�59).", _
                "�����. ������� � ������ ������� (������� ���, ������ � �������) � �������������� ����������� ����������� �������, ������������� � ������� �������, ���������� � ���������� �������. ���� � ������ ������������, ���� ������ �������� ���� � ������ � ����� ��������� � ��������� ����� 10:00 A.M. ��� P.M. �������� ������� �� ��������� �������� h:mm:ss.")
215:    cmbFormat.List = arrFormatDate
216:    Call ChengeFlag(obtnFormat)
217:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
218: End Sub
