VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5892
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10032
   LinkTopic       =   "Form1"
   ScaleHeight     =   5892
   ScaleWidth      =   10032
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Selanjutnya"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   1680
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   3120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mantap Asooiii Pak De !!"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   324
      Left            =   4320
      TabIndex        =   0
      Top             =   480
      Width           =   3252
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form1.Hide

End Sub

Private Sub Timer1_Timer()
Static atas As Boolean
Label1.Left = Label1.Left + IIf(kiri, -50, 50)
If Label1.Left < 0 Then
kiri = False
ElseIf Label1.Left > Me.Height - Label1.Height - 100 Then
kiri = True
End If
End Sub

Private Sub Timer2_Timer()
Static atas As Boolean
Form1.Left = Form1.Left + IIf(kiri, -50, 50)
If Form1.Left < 0 Then
kiri = False
ElseIf Form1.Left > Me.Height - Form1.Height - 100 Then
kiri = True
End If
End Sub
