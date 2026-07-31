Attribute VB_Name = "ZB_CleanUpWorkbook"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : ZB_CleanUpWorkbook - Functions to reduce the size of the workbook / clean up the workbook
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 30-09-2024 08:34    CalDymos            Rename Module
'* Updated    : 30-09-2024 08:34    CalDymos            Add Functions 'DelAllShapes' + 'DellAllUnusedCustomFormats'

Option Explicit
Option Private Module

Private Const sFULL_PATH As String = "Full path to the file:"
Private Const sFILE_LINKS As String = "File with a link"
Private Const sDELETE As String = "DELETE"

    Public Sub deleteAllLinksInFile()
          Dim sFileNameFull As Variant
          Dim bFlag       As Boolean

1         sFileNameFull = SelectedFile(vbNullString, False, "*.xls;*.xlsm;*.xlsx")
2         If TypeName(sFileNameFull) = "Empty" Then Exit Sub

3         If MsgBox("Create backup files ?", vbYesNo + vbQuestion, "Removing passwords:") = vbYes Then
4             bFlag = True
5         End If

6         On Error GoTo errMsg

          Dim sFullNameFile As String
          Dim cEditOpenXML As clsEditOpenXML
          Dim sPathLinks  As String
          Dim bMsg        As Boolean

7         sFullNameFile = sFileNameFull(1)
8         Set cEditOpenXML = New clsEditOpenXML
9         With cEditOpenXML
10            .CreateBackupXML = bFlag
11            .SourceFile = sFullNameFile
12            .UnzipFile
13            sPathLinks = .XLFolder & "externalLinks"
14            If FileHave(sPathLinks, Directory) Then
                  Dim objFso As Object
15                Set objFso = CreateObject("Scripting.FileSystemObject")
16                objFso.DeleteFolder (sPathLinks)
17                Set objFso = Nothing
18                bMsg = True
19            End If
20            .ZipAllFilesInFolder
21        End With
22        Set cEditOpenXML = Nothing
23        If bMsg Then
24            Call MsgBox("The links in the file were completely deleted: [" & sGetBaseName(sFullNameFile) & "]", vbInformation, "Deleting links:")
25        Else
26            Call MsgBox("In the file: [" & sGetBaseName(sFullNameFile) & "] no links to other files!", vbInformation, "Deleting links:")
27        End If

28        Exit Sub
errMsg:
29        Select Case Err.Number
              Case Else
30                Call MsgBox("Error in deleteAllLinksInFile" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbCritical, "Mistake:")
31                Call WriteErrorLog("deleteAllLinksInFile")
32        End Select
33        Set cEditOpenXML = Nothing
34    End Sub

     Public Sub getListAllLinksInFile()
          Dim sFileNameFull As Variant

34        sFileNameFull = SelectedFile(vbNullString, False, "*.xls;*.xlsm;*.xlsx")
35        If TypeName(sFileNameFull) = "Empty" Then Exit Sub

36        On Error GoTo errMsg

          Dim sFullNameFile As String
          Dim cEditOpenXML As clsEditOpenXML
          Dim sPathLinks  As String
          Dim bMsg        As Boolean

37        sFullNameFile = sFileNameFull(1)
38        Set cEditOpenXML = New clsEditOpenXML
39        With cEditOpenXML
40            .CreateBackupXML = False
41            .SourceFile = sFullNameFile
42            .UnzipFile
43            sPathLinks = .XLFolder & "externalLinks\_rels"
44            If FileHave(sPathLinks, Directory) Then
                  Dim objFso As Object
                  Dim objFolder As Object
                  Dim objFile As Object
                  Dim i   As Integer
                  Dim j   As Integer
                  Dim arrFile() As String
                  Dim sXML As String
                  Const sTARGET As String = " Target="

45                Set objFso = CreateObject("Scripting.FileSystemObject")
46                Set objFolder = objFso.GetFolder(sPathLinks)

47                For Each objFile In objFolder.Files
48                    If objFile.Name Like "*.rels" Then
49                        j = j + 1
50                        ReDim Preserve arrFile(1 To 2, 1 To j)
51                        arrFile(1, j) = objFile.Name
52                        sXML = .GetXMLFromFile(arrFile(1, j), sPathLinks & Application.PathSeparator)
53                        If sXML Like "*" & sTARGET & VBA.Chr$(34) & "*" Then
54                            sXML = VBA.Right$(sXML, VBA.Len(sXML) - VBA.InStr(1, sXML, sTARGET) - VBA.Len(sTARGET))
55                            sXML = VBA.Left$(sXML, VBA.InStr(1, sXML, VBA.Chr$(34)) - 1)
56                            arrFile(2, j) = sXML
57                        End If
58                    End If
59                Next
60                Set objFolder = Nothing
61                Set objFso = Nothing
62                bMsg = True
63            End If
64            .ZipAllFilesInFolder
65        End With
66        Set cEditOpenXML = Nothing

67        If bMsg Then
68            ActiveWorkbook.Worksheets.Add
69            With ActiveCell
70                .Value = sFULL_PATH
71                .Offset(0, 1).Value = sFullNameFile
72                .Offset(1, 0).Value = sFILE_LINKS
73                .Offset(1, 1).Value = "The file to which the link goes"
74                .Offset(1, 2).Value = "Action (put down)"
75                .Offset(2, 0).Resize(UBound(arrFile, 2), UBound(arrFile, 1)).Value2 = WorksheetFunction.Transpose(arrFile)
76                .Offset(2, 2).Resize(UBound(arrFile, 2), 1).Value2 = sDELETE
77            End With
78            Call MsgBox("Creating a list of links from a file:[" & sGetBaseName(sFullNameFile) & "]", vbInformation, "Creating a list:")
79        Else
80            Call MsgBox("In the file: [" & sGetBaseName(sFullNameFile) & "] no links to other files!", vbInformation, "Creating a list:")
81        End If

82        Exit Sub
errMsg:
83        Select Case Err.Number
              Case Else
84                Call MsgBox("Error in getListAllLinksInFile" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbCritical, "Mistake:")
85                Call WriteErrorLog("getListAllLinksInFile")
86        End Select
87        Set cEditOpenXML = Nothing

89    End Sub

Public Sub deleteLinksOnList()
          Dim bFlag       As Boolean
          Dim arrVal      As Variant
          Dim errMsg      As String
          Dim sFullNameFile As String

88        On Error GoTo errMsg

89        With ActiveSheet
              Dim lLastRow As Long
90            lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
91            If lLastRow < 3 Then
92                Call MsgBox("There is no data table!", vbCritical, "Mistake:")
93                Exit Sub
94            End If
95            If .Cells(1, 1).Value <> sFULL_PATH Then
96                errMsg = "Field not found [" & sFULL_PATH & "]" & vbNewLine
97            End If
98            If .Cells(2, 1).Value <> sFILE_LINKS Then
99                errMsg = errMsg & "Field not found [" & sFILE_LINKS & "]" & vbNewLine
100           End If

101           sFullNameFile = .Cells(1, 2).Value

102           If sFullNameFile = vbNullString Then
103               errMsg = errMsg & "The file path is not set:" & vbNewLine
104           ElseIf Not FileHave(sFullNameFile) Then
105               errMsg = errMsg & "The path to the file does not exist" & vbNewLine
106           End If

107           If errMsg <> vbNullString Then
108               Call MsgBox("The data table is not recognized:" & vbNewLine & errMsg, vbCritical, "Mistake:")
109               Exit Sub
110           End If

111           arrVal = .Range(.Cells(3, 1), .Cells(lLastRow, 3)).Value2
112       End With

113       If MsgBox("Create backup files ?", vbYesNo + vbQuestion, "Removing passwords:") = vbYes Then
114           bFlag = True
115       End If

          Dim cEditOpenXML As clsEditOpenXML
          Dim sPathLinks  As String
          Dim sPathLinksRels As String
          Dim bMsg        As Boolean
          Dim i           As Integer
          Dim iCount      As Integer
          Dim sfileName   As String

116       Set cEditOpenXML = New clsEditOpenXML
117       With cEditOpenXML
118           .CreateBackupXML = bFlag
119           .SourceFile = sFullNameFile
120           .UnzipFile
121           sPathLinks = .XLFolder & "externalLinks" & Application.PathSeparator
122           sPathLinksRels = sPathLinks & "_rels" & Application.PathSeparator

123           For i = 1 To UBound(arrVal)
124               sfileName = arrVal(i, 1)
125               If arrVal(i, 3) = sDELETE And FileHave(sPathLinksRels & sfileName) Then
126                   Call Kill(sPathLinks & VBA.Replace(sfileName, ".rels", vbNullString))
127                   Call Kill(sPathLinksRels & sfileName)
128                   bMsg = True
129                   iCount = iCount + 1
130               End If
131           Next i
132           .ZipAllFilesInFolder
133       End With
134       Set cEditOpenXML = Nothing

135       If bMsg Then
136           Call MsgBox("The links in the file were deleted: [" & sGetBaseName(sFullNameFile) & "]" & vbNewLine & "Deleted: [" & iCount & "] connections!", vbInformation, "Deleting links:")
137       Else
138           Call MsgBox("In the file: [" & sGetBaseName(sFullNameFile) & "] no links to other files!", vbInformation, "Deleting links:")
139       End If

140       Exit Sub
errMsg:
141       Select Case Err.Number
              Case Else
142               Call MsgBox("Error in deleteLinksOnList" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbCritical, "Mistake:")
143               Call WriteErrorLog("deleteLinksOnList")
144       End Select
145       Set cEditOpenXML = Nothing
End Sub


Public Sub DelAllShapes()
          Dim shp As Shape
          Dim errMsg As String
          Dim Response
          Dim lCount As Long
          Dim lTotal As Long
          Dim sSheet As Worksheet
          
146       On Error GoTo errMsg

147       Response = MsgBox("Delete all shapes in the active worksheet?", vbQuestion + vbYesNo, "Del All Shapes")
148       If Response = vbYes Then
149           Set sSheet = Application.ActiveSheet
150           With sSheet
151               lTotal = .Shapes.Count
152               lCount = 0
153               C_PublicFunctions.ShowProgressBar True
154               For Each shp In .Shapes
155                   shp.Delete
156                   lCount = lCount + 1
157                   C_PublicFunctions.UpdateProgressBar lCount, lTotal
                      
158               Next shp
159               C_PublicFunctions.ShowProgressBar False
160           End With
161           MsgBox CStr(lTotal) & " Shapes have been removed"
162       End If
163       Exit Sub

errMsg:
164       Select Case Err.Number
              Case Else
165               Call MsgBox("Error in DelAllShapes" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbCritical, "Mistake:")
166               Call WriteErrorLog("DelAllShapes")
167       End Select
End Sub

Public Sub DellAllUnusedCustomFormats()
          Dim styleObj As Style
          Dim rngCell As Range
          Dim wb As Workbook
          Dim wsh As Worksheet
          Dim str As String
          Dim iStyleCount As Long
          Dim lTotal As Long
          Dim lCount As Long
          Dim Response
          Dim i As Long
          Dim i2 As Long
          Dim errMsg As String
          Dim dict As New Scripting.Dictionary    ' <- from Tools / References... / "Microsoft Scripting Runtime"

168       On Error GoTo errMsg

169       Response = MsgBox("Delete all unused custom Styles?", vbQuestion + vbYesNo, "Del All unused custom formats")
170       If Response = vbYes Then
171           C_PublicFunctions.ShowProgressBar True, "Collect custom styles..."
172           Set wb = ActiveWorkbook ' the active workbook in excel


              'Debug.Print "BEGINNING # of styles in workbook: " & wb.Styles.Count

              ' dict := list of styles
173           lTotal = wb.Styles.Count
174           For i = 1 To lTotal 'Each styleObj In wb.Styles
175               str = wb.Styles(i).NameLocal
176               iStyleCount = iStyleCount + 1
177               Call dict.Add(str, 0)    ' First time:  adds keys
178                If lCount Mod (lTotal / 100) = 0 Then C_PublicFunctions.UpdateProgressBar iStyleCount, lTotal
179           Next i

180           ResetProgressBar "Determine used styles..."
              'Debug.Print "  dictionary now has " & dict.Count & " entries."
              ' Status, dictionary has styles (key) which are known to workbook


              ' Traverse each visible worksheet and increment count each style occurrence
181           For Each wsh In wb.Worksheets
182               lTotal = wsh.UsedRange.Cells.Count
183               lCount = 0
184               If wsh.visible Then
185                   For Each rngCell In wsh.UsedRange.Cells
186                       str = rngCell.Style
187                       dict.Item(str) = dict.Item(str) + 1     ' This time:  counts occurrences
188                       lCount = lCount + 1
189                        If lCount Mod (lTotal / 100) = 0 Then C_PublicFunctions.UpdateProgressBar lCount, lTotal
190                   Next rngCell
191               End If
192               ResetProgressBar "Determine used styles..."
193           Next wsh
              ' Status, dictionary styles (key) has cell occurrence count (item)
194           Set rngCell = Nothing

              ' Try to delete unused styles
              Dim aKey As Variant
            
195           ResetProgressBar "Delete unused styles..."
196           lTotal = dict.Count
197           lCount = 0
199           For Each aKey In dict.Keys

                  ' display count & stylename
                  '    e.g. "24   Normal"
                  'Debug.Print dict.Item(aKey) & vbTab & aKey

200               If dict.Item(aKey) = 0 And Not wb.Styles(aKey).BuiltIn Then
                      ' Occurrence count (Item) indicates this style is not used
201                   wb.Styles(aKey).Locked = False
202                   wb.Styles(aKey).Delete
203                   dict.Remove aKey
204               End If
205               lCount = lCount + 1
207               If lCount Mod (lTotal / 100) = 0 Then C_PublicFunctions.UpdateProgressBar lCount, lTotal ' Update alle 1%
211           Next aKey
212           C_PublicFunctions.ShowProgressBar False
              'Debug.Print "ENDING # of style in workbook: " & wb.Styles.Count
213           MsgBox CStr(iStyleCount - wb.Styles.Count) & " formats have been removed"
214       End If

215       Set dict = Nothing
          
216       Exit Sub
errMsg:
217       Select Case Err.Number
              Case Else
218               Call MsgBox("DellAllUnusedCustomFormats" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbCritical, "Mistake:")
219               Call WriteErrorLog("DellAllUnusedCustomFormats")
220       End Select
End Sub

Public Sub DellAllCustomFormats()
          Dim N As Long, i As Long
          Dim errMsg As String
          Dim inSc, ninSc As Boolean
          Dim Response
          Dim lCount As Long
          
221       On Error GoTo errMsg
222       Response = MsgBox("Delete all custom Styles?", vbQuestion + vbYesNo, "Del All custom formats")
223       If Response = vbYes Then
224           C_PublicFunctions.ShowProgressBar True
225           With ActiveWorkbook
226               N = .Styles.Count
227               inSc = True

228               For i = N To 1 Step -1
229                   inSc = .Styles(i).BuiltIn
230                   ninSc = Not inSc
231                   If ninSc Then
232                       .Styles(i).Locked = False
233                       .Styles(i).Delete
                          lCount = lCount + 1
                          If lCount Mod (N / 100) = 0 Then C_PublicFunctions.UpdateProgressBar lCount, N ' update alle 1%
234                   End If
235               Next i
236           End With
          C_PublicFunctions.ShowProgressBar False
237       End If
238       Exit Sub
errMsg:
239       Select Case Err.Number
              Case Else
240               Call MsgBox("DellAllCustomFormats" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbCritical, "Mistake:")
241               Call WriteErrorLog("DellAllCustomFormats")
242       End Select
End Sub

