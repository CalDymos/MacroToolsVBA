Attribute VB_Name = "Q_InToFile"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Q_InToFile - ���������� �����
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

    Public Sub InToFile()
13:    Dim strPath As String
14:    On Error GoTo errmsg
15:
16:    strPath = O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=True, bBackUp:=False)
17:    If strPath = vbNullString Then Exit Sub
18:   Q_InToFile.FilenamesCollectionToPath (strPath)
19:
20:    If MsgBox("Delete the folder of the unpacked Excel file" & vbNewLine & "The Excel file is not deleted!", vbYesNo + vbCritical, "Deleting a folder:") = vbYes Then
21:        Call Q_InToFile.RemoveFolderWithContent(strPath)
22:    End If
23:    Exit Sub
errmsg:
25:    Debug.Print "Error in InTo File!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
26:    Call WriteErrorLog("InToFile")
27: End Sub
    Private Sub FilenamesCollectionToPath(ByVal StrPathToFile As String)
29:    ' ���� �� ������� ����� ��� ����� TXT, � ������� �� ���� ������ �� ���.
30:    ' ��������������� ����� � �������� �������� �� ����� ���.
31:    Dim i      As Long
32:    Dim coll   As Collection
33:    On Error GoTo errmsg
34:    ' ��������� � �������� coll ������ ����� ������
35:    Set coll = FilenamesCollection(StrPathToFile, "*.*", 3)
36:
37:    Application.ScreenUpdating = False    ' ��������� ���������� ������
38:    ' ������ ����� �����
39:    Dim SH     As Worksheet: Set SH = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
40:    ' ��������� ��������� �������
41:    With SH.Range("a1").Resize(, 5)
42:        .Value = Array("�", "File name", "Full path", "File size", "File extension")
43:        .Font.Bold = True: .Interior.ColorIndex = 17
44:    End With
45:
46:    ' ������� ���������� �� ����
47:    For i = 1 To coll.Count    ' ���������� ��� �������� ���������, ���������� ���� � ������
48:        SH.Range("a" & SH.Rows.Count).End(xlUp).Offset(1).Resize(, 5).Value = _
                         Array(i, C_PublicFunctions.sGetFileName(coll(i)), coll(i), C_PublicFunctions.FileSize(coll(i)), C_PublicFunctions.sGetExtensionName(coll(i)))    ' ������� �� ���� ��������� ������
50:        DoEvents    ' �������� ������� ���������� ��
51:    Next
52:    SH.Range("a:e").EntireColumn.AutoFit    ' ���������� ������ ��������
53:    [a2].Activate: ActiveWindow.FreezePanes = True    ' ���������� ������ ������ �����
54:    Application.ScreenUpdating = True    ' ��������� ���������� ������
55:    Exit Sub
errmsg:
57:    Debug.Print "Error in FilenamesCollectionToPath!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
58:    Call WriteErrorLog("FilenamesCollectionToPath")
59: End Sub

    Private Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "", _
                    Optional ByVal SearchDeep As Long = 999) As Collection
63:    ' �������� � �������� ��������� ���� � ����� FolderPath,
64:    ' ����� ����� ������� ������ Mask (����� �������� ������ ����� � ����� ������/�����������)
65:    ' � ������� ������ SearchDeep � ��������� (���� SearchDeep=1, �� �������� �� ���������������).
66:    ' ���������� ���������, ���������� ������ ���� ��������� ������
67:    ' (����������� ����������� ����� ��������� GetAllFileNamesUsingFSO)
68:    Dim FSO    As Object
69:    On Error GoTo errmsg
70:    Set FilenamesCollection = New Collection    ' ������ ������ ���������
71:    Set FSO = CreateObject("Scripting.FileSystemObject")    ' ������ ��������� FileSystemObject
72:    Call GetAllFileNamesUsingFSO(FolderPath, Mask, FSO, FilenamesCollection, SearchDeep)  ' �����
73:    Set FSO = Nothing: Application.StatusBar = False    ' ������� ������ ��������� Excel
74:    Exit Function
errmsg:
76:    Debug.Print "Error in FilenamesCollection!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
77:    Call WriteErrorLog("FilenamesCollection")
78: End Function

     Private Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef FSO, _
                     ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
82:    ' ���������� ��� ����� � �������� � ����� FolderPath, ��������� ������ FSO
83:    ' ������� ����� �������������� � ��� ������, ���� SearchDeep > 1
84:    ' ��������� ���� ��������� ������ � ��������� FileNamesColl
85:    Dim curfold As Object, fil As Object, sfol As Object
86:    On Error Resume Next: Set curfold = FSO.GetFolder(FolderPath)
87:    If Not curfold Is Nothing Then    ' ���� ������� �������� ������ � �����
88:
89:        ' ���������������� ��� ������ ��� ������ ���� � ���������������
90:        ' � ������� ������ ����� � ������ ��������� Excel
91:        ' Application.StatusBar = "����� � �����: " & FolderPath
92:
93:        For Each fil In curfold.Files    ' ���������� ��� ����� � ����� FolderPath
94:            If fil.Name Like "*" & Mask Then FileNamesColl.Add fil.Path
95:        Next
96:        SearchDeep = SearchDeep - 1    ' ��������� ������� ������ � ���������
97:        If SearchDeep Then    ' ���� ���� ������ ������
98:            For Each sfol In curfold.SubFolders    ' ���������� ��� �������� � ����� FolderPath
99:                GetAllFileNamesUsingFSO sfol.Path, Mask, FSO, FileNamesColl, SearchDeep
100:            Next
101:        End If
102:        Set fil = Nothing: Set curfold = Nothing: Set sfol = Nothing   ' ������� ����������
103:    End If
104: End Function

     Private Sub RemoveFolderWithContent(ByVal sFolder As String)
107: '    '���� � ����� ����� ������ ��������, ���� �� ������� �������� � �� ����������
108:   Shell "cmd /c rd /S/Q """ & sFolder & """"
109: End Sub

