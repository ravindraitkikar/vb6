VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3960
      Top             =   540
   End
   Begin VB.CommandButton cmdHello 
Caption               =   "Hello New Word!"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple hello world program.
'Illustrates a command button and a timer.
Option Explicit
Dim ElapsedTime As Integer

Private Sub Form_Load()
    Form1.Caption = "Hello World - Button & Timer example"
    ElapsedTime = 5000
    Timer1.Enabled = False
End Sub


Private Sub cmdHello_Click()
    Timer1.Interval = 1
    Timer1.Enabled = True
    Timer1.Interval = 1000
    cmdHello.Enabled = False
End Sub

Private Sub Timer1_Timer()
     ElapsedTime = ElapsedTime - Timer1.Interval
     cmdHello.Caption = "Goodbye in " & ElapsedTime / 1000 & " seconds)"
     If ElapsedTime <= 0 Then End
End Sub
