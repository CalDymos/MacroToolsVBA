VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoFile 
   Caption         =   "File Properties:"
   ClientHeight    =   8196.001
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   13452
   OleObjectBlob   =   "InfoFile.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "InfoFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : InfoFile - ÛÔ‡‚ÎÂÌËÂ Ò‚ÓÈÒÚ‚‡ÏË Ù‡ÈÎ‡
'* Created    : 20-07-2020 15:34
'* Author     : VBATools
'* Contacts   : -
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private Sub cmbMain_Change()
    On Error Resume Next
    Call UpdateList(Me.ListCode, X_InfoFile.ShowProp(Workbooks(cmbMain.Value)))
    Call UpdateList(Me.ListCustomDocProp, X_InfoFile.ShowCustomDocProp(Workbooks(cmbMain.Value)))
    Call UpdateList(Me.ListCustomProp, X_InfoFile.ShowCustomSheetProp(Workbooks(cmbMain.Value)))
    Call UpdateList(Me.ListNames, X_InfoFile.ShowDefinedNames(Workbooks(cmbMain.Value)))
    On Error GoTo 0
End Sub
Private Sub UpdateList(ByRef objList As MSForms.ListBox, ByVal Txt As String)
    Const CHAR_WIDTH_FACTOR As Double = 0.6   ' N‰herung: Punkte pro Zeichen = Fontgrˆﬂe * Faktor
    Const COL_PADDING       As Double = 6     ' zus‰tzlicher Puffer in Punkten pro Spalte
    Const MIN_COL_WIDTH     As Double = 20    ' Mindestbreite pro Spalte in Punkten

    Dim arr         As Variant
    Dim parts       As Variant
    Dim i           As Long
    Dim j           As Integer
    Dim maxCol      As Integer
    Dim maxLen()    As Long
    Dim widths      As String
    Dim ptsPerChar  As Double

    objList.Clear
    If Txt = vbNullString Then Exit Sub

    arr = VBA.Split(Txt, vbNewLine)
    maxCol = objList.ColumnCount - 1
    ReDim maxLen(0 To maxCol)

    With objList
        For i = 0 To UBound(arr)
            If arr(i) <> vbNullString Then
                parts = VBA.Split(arr(i), "||")
                .AddItem
                For j = 0 To UBound(parts)
                    If j > maxCol Then Exit For
                    .List(.ListCount - 1, j) = Trim$(parts(j))
                    If Len(.List(.ListCount - 1, j)) > maxLen(j) Then
                        maxLen(j) = Len(.List(.ListCount - 1, j))
                    End If
                Next j
            End If
        Next i

        ' Spaltenbreiten aus max. Zeicheanzahl je Spalte berechnen
        ptsPerChar = .Font.Size * CHAR_WIDTH_FACTOR
        For j = 0 To maxCol
            widths = widths & IIf(j > 0, ";", "") & _
                     CStr(WorksheetFunction.Max(MIN_COL_WIDTH, maxLen(j) * ptsPerChar + COL_PADDING))
        Next j
        .ColumnWidths = widths
    End With
End Sub

Private Sub Label2_Click()
    Me.Hide
    Call InfoFile2.Show
    Call cmbMain_Change
    Me.Show
End Sub

Private Sub LbDelAllProper_Click()
    If MsgBox("Delete ALL properties ?", vbYesNo + vbQuestion, "Deleting Properties:") = vbYes Then
        Dim iCount  As Byte
        iCount = X_InfoFile.DelAllProp(Workbooks(cmbMain.Value))
        Call cmbMain_Change
        Call MsgBox("Properties removed:" & iCount, vbInformation, "Deleting Properties:")
    End If
End Sub
Private Sub LbEdit_Click()
    Call EditProp
End Sub

Private Sub lbTemplete_Click()
'    Dim tbData As Variant
'    Dim i As Integer
'    tbData = ThisWorkbook.Worksheets(C_Const.SH_SNIPPETS).ListObjects("TB_TEMPLETE").DataBodyRange.Value2
'    tbData = ThisWorkbook.Worksheets(C_Const.SH_SNIPPETS).ListObjects("TB_TEMPLETE").DataBodyRange.Value2
'    For i = 1 To UBound(tbData)
'        Call X_InfoFile.AddOneCustomProp(Workbooks(cmbMain.Value), tbData(i, 1), tbData(i, 2))
'    Next i
'    Call cmbMain_Change
End Sub

Private Sub ListCode_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call EditProp
End Sub
Private Sub EditProp()
    Dim txtNew      As String
    Dim txtOld      As String
    Dim NameProp    As String
    With Me.ListCode
        If IsNumeric(.BoundValue) Then
            txtOld = VBA.Trim$(.List(CInt(.BoundValue) - 1, 2))
            NameProp = .List(CInt(.BoundValue) - 1, 1)
            txtNew = InputBox("Edit the property [" & NameProp & " ] ?", "Editing a property:", txtOld)
            If txtNew <> txtOld Then
                Call X_InfoFile.WriteOneProp(Workbooks(cmbMain.Value), NameProp, txtNew)
                Call cmbMain_Change
            End If
        End If
    End With
End Sub

Private Sub lbAddCustDocProp_Click()
    Call AddCustDocProp(vbNullString, vbNullString)
End Sub

Private Sub lbEditCustDocProp_Click()

    Dim txtOld      As String
    Dim NameProp    As String
    With Me.ListCustomProp
        If IsNumeric(.BoundValue) Then
            txtOld = VBA.Trim$(.List(CInt(.BoundValue) - 1, 2))
            NameProp = .List(CInt(.BoundValue) - 1, 1)
            Call X_InfoFile.DelOneCustomDocProp(Workbooks(cmbMain.Value), NameProp)
            Call AddCustDocProp(NameProp, txtOld)
        End If
    End With
End Sub
Private Sub lbDelOneCustDocProp_Click()
    Dim NameProp    As String
    With Me.ListCustomDocProp
        If IsNumeric(.BoundValue) Then
            NameProp = .List(CInt(.BoundValue) - 1, 1)
            If MsgBox("Delete Property [" & NameProp & " ] ?", vbYesNo + vbQuestion, "Deleting a property:") = vbYes Then
                Call X_InfoFile.DelOneCustomDocProp(Workbooks(cmbMain.Value), NameProp)
                Call cmbMain_Change
            End If
        End If
    End With
End Sub
Private Sub AddCustDocProp(ByVal txtPropName As String, ByVal txtPropValue As String)
    txtPropName = InputBox("Name der Eigenschaft", "Creating a property:", txtPropName)
    If txtPropName <> vbNullString Then
        txtPropValue = InputBox("Wert der Eigenschaft", "Creating a property:", txtPropValue)
        If txtPropValue <> vbNullString Then
            Call X_InfoFile.AddOneCustomDocProp(Workbooks(cmbMain.Value), txtPropName, txtPropValue)
            Call cmbMain_Change
        End If
    End If
End Sub


Private Sub lbDelAllCustomDocProp_Click()
    If MsgBox("Delete ALL properties ?", vbYesNo + vbQuestion, "Deleting Properties:") = vbYes Then
        Dim iCount  As Byte
        iCount = X_InfoFile.DelAllCustomDocProp(Workbooks(cmbMain.Value))
        Call cmbMain_Change
        Call MsgBox("Properties removed:" & iCount, vbInformation, "Deleting Properties:")
    End If
End Sub


Private Sub UserForm_Activate()
    Dim vbProj      As VBIDE.VBProject
    If Workbooks.Count = 0 Then
        Unload Me
        Call MsgBox("No open" & Chr(34) & "Excel Files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Mistake:")
        Exit Sub
    End If
    With Me.cmbMain
        .Clear
        On Error Resume Next
        For Each vbProj In Application.VBE.VBProjects
            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
        Next
        On Error GoTo 0
        .Value = ActiveWorkbook.Name
    End With
End Sub

Private Sub cmbCancel_Click()
    Unload Me
End Sub
Private Sub lbCancel_Click()
    Call cmbCancel_Click
End Sub
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
