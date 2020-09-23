VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pidl As Long
Dim hSHNotify As Long


Private Sub Command1_Click()
'Dim x As New OO7.Spy
'Dim y As New OO7.EventType
Dim x As New Spy
Dim y As New EventType
MsgBox "pidl=" & pidl
k = y.SHCNE_UPDATEDIR
    x.EmploySpy Me.hwnd, y.SHCNE_UPDATEDIR, "C:\Temp", AddressOf WndProc3, pidl, hSHNotify
MsgBox "pidl=" & pidl
Set x = Nothing
Set y = Nothing
End Sub

Private Sub Command2_Click()
 'Dim x As New OO7.Spy
 Dim x As New Spy
 MsgBox "pidl=" & pidl
   MsgBox x.DestroySpy(hSHNotify, pidl, Me.hwnd)
MsgBox "pidl=" & pidl
    Set x = Nothing
End Sub




