Attribute VB_Name = "B_CreateMenus"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : B_CreateMenus - Creating menu in VBE
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Updated    : 04-12-2025 8:25    CalDymos            Auto_Open made more robust against “Circular dependencies between modules” errors
Option Private Module
Option Explicit
Public ToolContextEventHandlers As New Collection

#If Win64 Then
Private Declare PtrSafe Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare PtrSafe Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
#Else
Private Declare Function GetKeyboardLayoutName Lib "USER32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare Function LoadKeyboardLayout Lib "USER32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
#End If

Private Const LANG_ENGLISH = 409
Private Const LANG_GERMAN = 407

    Private Sub Auto_Open()
1         On Error GoTo ErrorHandler
2      If Not VBAIsTrusted Then Exit Sub
3      If ThisWorkbook.Name <> C_Const.NAME_ADDIN & ".xlam" Then Exit Sub
4          Call AddContextMenus
5      Exit Sub
ErrorHandler:
          ' Log error but don't crash Excel
6         Debug.Print "MacroTools Auto_Open Error: " & Err.Number & " - " & Err.Description
          Call WriteErrorLog("Auto_Open")
          ' Optional: Retry after delay
7         On Error Resume Next
8         Application.OnTime Now + TimeValue("00:00:02"), "RetryAutoOpen"
9         On Error GoTo 0
10    End Sub

' Retry mechanism if initial load fails
Private Sub RetryAutoOpen()
10        On Error Resume Next
          
11        If Not VBAIsTrusted() Then Exit Sub
12        If ThisWorkbook.Name <> C_Const.NAME_ADDIN & ".xlam" Then Exit Sub
13        Call AddContextMenus
          
14        On Error GoTo 0
End Sub
    Public Sub AddContextMenus()

15     Call AddNewCommandBarMenu(C_Const.MENUMOVECONTRL)
16     Call AddButtom(C_Const.MTAG5, 984, "Hilfe für das Tool", "HelpMoveControl", C_Const.MENUMOVECONTRL, False, True)
17     Call AddButtom(C_Const.MTAG4, 38, "", "MoveControl", C_Const.MENUMOVECONTRL)
18     Call AddButtom(C_Const.MTAG3, 40, "", "MoveControl", C_Const.MENUMOVECONTRL, False, True)
19     Call AddButtom(C_Const.MTAG2, 39, "", "MoveControl", C_Const.MENUMOVECONTRL)
20     Call AddButtom(C_Const.MTAG1, 41, "", "MoveControl", C_Const.MENUMOVECONTRL)
21     Call AddComboBoxMove(C_Const.MENUMOVECONTRL)

22     Call AddNewCommandBarMenu(C_Const.TOOLSMENU)
23     Call AddButtom(C_Const.TAG15, 984, "Hilfe beim Einrichten", "HelpMainAddin", C_Const.TOOLSMENU, False, True)
       'Call AddButtom(C_Const.TAG29, 277, "Die Tastenkombinationen", "AddLegendHotKeys", C_Const.TOOLSMENU, False, True)
24     Call AddButtom(C_Const.TAG12, 0, "FormatBuilder", "subFormatBuilder", C_Const.TOOLSMENU, True, True)
       'Call AddButtom(C_Const.TAG11, 0, "MsgBoxBuilder", "subMsgBoxGenerator", C_Const.TOOLSMENU, True, True)
       'Call AddButtom(C_Const.TAG13, 0, "ProcedureBuilder", "subProcedureBuilder", C_Const.TOOLSMENU, True, True)
       'Call AddButtom(C_Const.TAG23, 107, "Option's Explicit and Private Module", "insertOptionsExplicitAndPrivateModule", C_Const.TOOLSMENU, False, False)
       'Call AddButtom(C_Const.TAG28, 0, "Option's", "subOptionsMenu", C_Const.TOOLSMENU, True, True)

       'Call AddButtom(C_Const.TAG24, 2045, "Copy", "SetInCipBoard", C_Const.TOOLSMENU, True, False)
       'Call AddButtom(C_Const.TAG25, 22, "Paste", "GetFromCipBoard", C_Const.TOOLSMENU, True, True)

25     Call AddButtom(C_Const.TAG20, 1714, "Suche nach unbenutzten Variablen ", "SerchVariableUnUsedInSelectedWorkBook", C_Const.TOOLSMENU, False, False)
       'Call AddButtom(C_Const.TAG19, 3838, "Alle VBE-Fenster schließen ", "CloseAllWindowsVBE", C_Const.TOOLSMENU, False, False)
26     Call AddButtom(C_Const.TAG14, 22, "LogRecorder-Klasse einfügen ", "AddLogRecorderClass", C_Const.TOOLSMENU, False, True)

       'Call AddButtom(C_Const.TAG19, 8, "TODO-Liste ", "ShowTODOList", C_Const.TOOLSMENU, False, False)
27     Call AddButtom(C_Const.TAG18, 1972, "TODO Eintrag einfügen ", "sysAddTODOTop", C_Const.TOOLSMENU, False, False)
28     Call AddButtom(C_Const.TAG17, 456, "Eine Aktualisierungsseite erstellen ", "sysAddModifiedTop", C_Const.TOOLSMENU, False, False)
29     Call AddButtom(C_Const.TAG16, 1546, "Erstellen Sie ein Inventar ", "sysAddHeaderTop", C_Const.TOOLSMENU, False, True)

       'Call AddButtom(C_Const.TAG10, 3917, "Formatierung des CODE löschen", "CutTab", C_Const.TOOLSMENU)
       'Call AddButtom(C_Const.TAG9, 3919, "Formatieren des Codes", "ReBild", C_Const.TOOLSMENU, False, True)
       'Call AddButtom(C_Const.TAG8, 12, "Zeilennummerierung löschen", "RemoveLineNumbers_", C_Const.TOOLSMENU)
       'Call AddButtom(C_Const.TAG7, 11, "Zeilennummerierung erstellen", "AddLineNumbers_", C_Const.TOOLSMENU)
30     Call AddComboBox(C_Const.TOOLSMENU)
       'Call AddButtom(C_Const.TAG27, 210, "Prozeduren Alphabetisch sortieren", "AlphabetizeProcedure", C_Const.TOOLSMENU, False, True)
       'Call AddButtom(C_Const.TAG6, 47, "[Direktbereich] löschen", "ClearImmediateWindow", C_Const.TOOLSMENU, False, True)
31     Call AddButtom(C_Const.TAG5, 2059, "Erstelle eine Legende", "AddLegend", C_Const.TOOLSMENU)
       'Call AddButtom(C_Const.TAG4, 21, "Óäàëèòü ìîäóëü", "DeleteSnippetEnumModule", C_Const.TOOLSMENU)
       'Call AddButtom(C_Const.TAG3, 1753, "Âñòàâèòü ìîäóëü", "AddSnippetEnumModule", C_Const.TOOLSMENU)
       'Call AddButtom(C_Const.TAG2, 22, "Code einfügen", "InsertCode", C_Const.TOOLSMENU, False, False)

       'Call AddButtom(C_Const.TAG26, 9634, "Zuweisung tauschen [=]", "SwapEgual", C_Const.POPMENU, True, False)
32     Call AddButtom(C_Const.TAG21, 0, "UPPER Case", "toUpperCase", C_Const.POPMENU, True, False)
33     Call AddButtom(C_Const.TAG22, 0, "lower Case", "toLowerCase", C_Const.POPMENU, True, False)
       'Call AddButtom(C_Const.TAG1, 22, "Code einfügen", "InsertCode", C_Const.POPMENU, True, False)

34     Call AddButtom(C_Const.RTAG7, 1650, "Aling Horiz", "vbaCntAlingHoriz", C_Const.RENAMEMENU, True)
35     Call AddButtom(C_Const.RTAG8, 1653, "Aling Vert", "vbaCntAlingVert", C_Const.RENAMEMENU, True)
36     Call AddButtom(C_Const.RTAG1, 162, "ReName Control", "RenameControl", C_Const.RENAMEMENU, True)
37     Call AddButtom(C_Const.RTAG2, 22, "Paste Style", "PasteStyleControl", C_Const.RENAMEMENU, True)
38     Call AddButtom(C_Const.RTAG3, 1076, "Copy Style", "CopyStyleControl", C_Const.RENAMEMENU, True)
39     Call AddButtom(C_Const.RTAG4, 704, "Paste Icon", "AddIcon", C_Const.RENAMEMENU, True, True)
40     Call AddButtom(C_Const.RTAG5, 0, "UPPER Case", "UperTextInControl", C_Const.RENAMEMENU, True, False)
41     Call AddButtom(C_Const.RTAG6, 0, "lower Case", "LowerTextInControl", C_Const.RENAMEMENU, True, False)

42     Call AddButtom(C_Const.CTAG1, 2045, "Copy Module", "CopyModyleVBE", C_Const.COPYMODULE, True, False, , 6)

43     Call AddButtom(C_Const.RTAG2, 22, "Paste Style", "PasteStyleForms", C_Const.mMSFORMS, True)
44     Call AddButtom(C_Const.RTAG3, 1076, "Copy Style", "CopyStyleControl", C_Const.mMSFORMS, True)
45     Call AddButtom(C_Const.RTAG5, 0, "UPPER Case", "UperTextInForm", C_Const.mMSFORMS, True, False)
46     Call AddButtom(C_Const.RTAG6, 0, "lower Case", "LowerTextInForm", C_Const.mMSFORMS, True, False)
47    End Sub
     Private Sub AddNewCommandBarMenu(ByVal sNameCommandBar As String)
       Dim myCommandBar As CommandBar
48     On Error GoTo AddNewCommandBar
49     Set myCommandBar = Application.VBE.CommandBars(sNameCommandBar)
50     If myCommandBar Is Nothing Then
AddNewCommandBar:
51         Set myCommandBar = Application.VBE.CommandBars.Add(Name:=sNameCommandBar, Position:=msoBarTop)
52         myCommandBar.visible = True
53          myCommandBar.RowIndex = 3
54      End If
55    End Sub
     Private Sub AddButtom( _
                       ByVal sTag As String, _
                       ByVal Face As Long, _
                       ByVal Capitan As String, _
                       ByVal sOnAction As String, _
                       ByVal sMenu As String, _
                       Optional ByRef VisibleCapiton As Boolean = False, _
                       Optional ByVal Begin_Group As Boolean = False, _
                       Optional ByVal ShortcutText As String = vbNullString, _
                       Optional ByVal Before As Byte = 1)
        Dim btn         As CommandBarButton
        Dim evtContextMenu As VBECommandHandler
56      Set btn = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlButton, Before:=Before)
57      With btn
58          .FaceId = Face
59          If VisibleCapiton Then .Caption = Capitan
60          .TooltipText = Capitan
61          .Tag = sTag
62          .OnAction = "'" & ThisWorkbook.Name & "'!" & sOnAction
63          .Style = msoButtonIconAndCaption
64          .BeginGroup = Begin_Group
65          .ShortcutText = ShortcutText
66      End With
67      Set evtContextMenu = New VBECommandHandler
68      Set evtContextMenu.EvtHandler = btn
69      ToolContextEventHandlers.Add evtContextMenu
70    End Sub
     Private Sub AddComboBox(ByVal sMenu As String)
        Dim combox      As CommandBarComboBox
71      Set combox = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlComboBox, Before:=1)
72      With combox
73          .Tag = C_Const.TAGCOM
74          .AddItem C_Const.SELECTEDMODULE
75          .AddItem C_Const.ALLVBAPROJECT
76          .Text = C_Const.SELECTEDMODULE
77      End With
78    End Sub
     Private Sub AddComboBoxMove(ByVal sMenu As String)
        Dim combox      As CommandBarComboBox
79      Set combox = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlComboBox, Before:=1)
80      With combox
81          .Tag = C_Const.MTAGCOM
82          .AddItem C_Const.MOVECONT
83          .AddItem C_Const.MOVECONTTOPLEFT
84          .AddItem C_Const.MOVECONTBOTTOMRIGHT
85          .Text = C_Const.MOVECONT
86      End With
87    End Sub
     Private Sub Auto_Close()
88      If VBAIsTrusted Then
89          Call DeleteContextMenus
90      End If
91    End Sub
     Public Sub DeleteContextMenus()
        Dim myCommandBar As CommandBar
92      On Error GoTo ErrorHandler

93      Call DeleteButton(C_Const.TAG1, C_Const.POPMENU)
94      Call DeleteButton(C_Const.TAG26, C_Const.POPMENU)
95      Call DeleteButton(C_Const.TAG21, C_Const.POPMENU)
96      Call DeleteButton(C_Const.TAG22, C_Const.POPMENU)

97      Call DeleteButton(C_Const.CTAG1, C_Const.COPYMODULE)

98      Call DeleteButton(C_Const.RTAG1, C_Const.RENAMEMENU)
99      Call DeleteButton(C_Const.RTAG2, C_Const.RENAMEMENU)
100     Call DeleteButton(C_Const.RTAG3, C_Const.RENAMEMENU)
101     Call DeleteButton(C_Const.RTAG4, C_Const.RENAMEMENU)
102     Call DeleteButton(C_Const.RTAG5, C_Const.RENAMEMENU)
103     Call DeleteButton(C_Const.RTAG6, C_Const.RENAMEMENU)
104     Call DeleteButton(C_Const.RTAG7, C_Const.RENAMEMENU)
105     Call DeleteButton(C_Const.RTAG8, C_Const.RENAMEMENU)

106     Call DeleteButton(C_Const.RTAG2, C_Const.mMSFORMS)
107     Call DeleteButton(C_Const.RTAG3, C_Const.mMSFORMS)
108     Call DeleteButton(C_Const.RTAG5, C_Const.mMSFORMS)
109     Call DeleteButton(C_Const.RTAG6, C_Const.mMSFORMS)

110     Set myCommandBar = Application.VBE.CommandBars(C_Const.TOOLSMENU)
111     If Not myCommandBar Is Nothing Then
112         myCommandBar.Delete
113     End If

114     Set myCommandBar = Application.VBE.CommandBars(C_Const.MENUMOVECONTRL)
115     If Not myCommandBar Is Nothing Then
116         myCommandBar.Delete
117     End If

        'î÷èñòêà êîëåêöèè
118     Do Until ToolContextEventHandlers.Count = 0
119         ToolContextEventHandlers.Remove 1
120     Loop

121     Exit Sub
ErrorHandler:

122     Select Case Err
        Case 5:
123             Err.Clear
124         Case Else:
125             Debug.Print "Fehler! â DeleteContextMenus" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "â ñòðîêå " & Erl
126             Call WriteErrorLog("DeleteContextMenus")
127     End Select
128     Err.Clear
129   End Sub
     Private Sub DeleteButton(ByRef sTag As String, ByVal sMenu As String)
        Dim Cbar        As CommandBar
        Dim Ctrl        As CommandBarControl
130     On Error GoTo ErrorHandler
131     Set Cbar = Application.VBE.CommandBars(sMenu)
132     For Each Ctrl In Cbar.Controls
133         If Ctrl.Tag = sTag Then
134             Ctrl.Delete
                'Exit Sub
135         End If
136     Next Ctrl
137     Exit Sub
ErrorHandler:
138     Debug.Print "Fehler! â DeleteButton" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "â ñòðîêå " & Erl
139     Call WriteErrorLog("DeleteButton")
140     Err.Clear
141     Resume Next
142   End Sub
     Public Function VBAIsTrusted() As Boolean
143     On Error GoTo ErrorHandler
        Dim sTxt As String
144     sTxt = Application.VBE.Version
145     VBAIsTrusted = True
146     Exit Function
ErrorHandler:
147     Select Case Err.Number
        Case 1004:
                'If ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then
148             Call MsgBox("WARNUNG!" & C_Const.NAME_ADDIN & vbLf & vbNewLine & _
                                "Deaktiviert: [Zugriff auf das VBA-Projektobjektmodell vertrauen]" & vbLf & _
                                "Zur Aktivierung gehen Sie zu Datei->Optionen->Trust Center->Einstellungen für das Trust Center->Makroeinstellungen" & _
                                vbLf & vbNewLine & "Und starten Sie Excel neu", vbCritical, "ACHTUNG !")
149         Case Else:
150             Debug.Print "Fehler! â VBAIsTrusted" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "â ñòðîêå " & Erl
151             Call WriteErrorLog("VBAIsTrusted")
152     End Select
153     Err.Clear
154     VBAIsTrusted = False
155   End Function
     Public Function WhatIsTextInComboBoxHave() As String
        Dim myCommandBar As CommandBar
        Dim cntrl       As CommandBarControl

156     Set myCommandBar = Application.VBE.CommandBars(C_Const.TOOLSMENU)
157     For Each cntrl In myCommandBar.Controls
158         If cntrl.Tag = C_Const.TAGCOM Then
159             WhatIsTextInComboBoxHave = cntrl.Text
160             Exit Function
161         End If
162     Next cntrl
163   End Function
     Public Sub ClearImmediateWindow()
        Dim KeybLayoutName As String * 8
164     KeybLayoutName = String(8, "0")
165     GetKeyboardLayoutName KeybLayoutName
166     KeybLayoutName = Val(KeybLayoutName)

167     Select Case Val(KeybLayoutName)
        Case LANG_ENGLISH, LANG_GERMAN
168             Call ClearImmediateWindowFunction
169             Call ClearImmediateWindowFunction
170         Case Else
                ' Umschalten auf die englische Version
171             Call LoadKeyboardLayout("00000409", &H1)
172             Call ClearImmediateWindowFunction
173             Call LoadKeyboardLayout("00000419", &H1)
174     End Select
175   End Sub
     Private Sub ClearImmediateWindowFunction()
176     Call SendKeys("^g") ' ctrl + g
177     Call SendKeys("^a") ' ctrl + a
178     Call SendKeys("{DEL}") ' entfernen
179   End Sub
     Public Sub RefreshMenu()
180     Call B_CreateMenus.DeleteContextMenus
181     Call B_CreateMenus.AddContextMenus
182     Call MsgBox("Zurücksetzen der Einstellungen " & C_Const.NAME_ADDIN & " durchgeführt!", vbInformation, "Zurücksetzen der Einstellungen " & C_Const.NAME_ADDIN & ":")
183   End Sub
     Private Sub subMsgBoxGenerator()
184     MsgBoxGenerator.Show
185   End Sub
     Private Sub subFormatBuilder()
186     BilderFormat.Show
187   End Sub
     Private Sub subProcedureBuilder()
188     BilderProcedure.Show
189   End Sub
Private Sub subOptionsMenu()
190     Call Y_Options.subOptions
End Sub
