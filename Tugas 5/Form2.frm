VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6036
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8172
   LinkTopic       =   "Form2"
   ScaleHeight     =   6036
   ScaleWidth      =   8172
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Kembali"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   2052
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   100
      Top             =   100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mantap Asoii Pak dee !!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   348
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   3036
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form2.Hide
End Sub

Private Sub Timer1_Timer()
Label1.Left = Label1.Left - 15
If Label1.Left <= -Label1.Left Then
Label1.Left = Form1.Width
End If
End Sub
