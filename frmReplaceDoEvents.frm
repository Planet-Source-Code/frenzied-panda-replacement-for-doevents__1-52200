VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "End Loop"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Loop"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim b As Boolean

Private Sub Command1_Click()
  b = True
  While b
    'It is better if you get the hwnd value beforehand so you
    'do not have to ask the vb dll for it each and every time
    DoMyEvents (Me.hwnd) 'Endless loop
    'If you want other window to be processed
    'then you need to DoMyEvents () with their HWND
    'That is why you cannot pause the ide window
    'when your in this loop.
  Wend
End Sub

Private Sub Command2_Click()
  b = False
End Sub

Private Sub Form_Paint()
  MsgBox "Painting Form"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  b = False
End Sub

