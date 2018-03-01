VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2364
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   2364
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   25
      Left            =   1440
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   600
      Top             =   1920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tugas Pak Rudi Text Bergerak Dan Form Bergerak"
      BeginProperty Font 
         Name            =   "One Stroke Script LET"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6132
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_timer()
Static Kiri As Boolean
Label1.Left = Label1.Left + IIf(Kiri, -50, 100)
If Label1.Left < 0 Then
Kiri = False
ElseIf Label1.Left > Me.Height - Label1.Height - 50 Then
Kiri = True
End If
End Sub
Private Sub timer2_timer()
Form1.Left = Form1.Left - 15
If Form1.Left <= -Form1.Left Then
Form1.Left = Form1.Width
End If
End Sub

