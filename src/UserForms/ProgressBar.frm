VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Please wait..."
   ClientHeight    =   1200
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   7500
   OleObjectBlob   =   "ProgressBar.frx":0000
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    Dim sngTop As Single, sngLeft As Single
    Me.StartUpPosition = 0
    sngLeft = Application.Left + Application.Width / 2 - Me.Width / 2
    sngTop = Application.Top + Application.Height / 2 - Me.Height / 2

    Me.Left = sngLeft
    Me.Top = sngTop
End Sub
