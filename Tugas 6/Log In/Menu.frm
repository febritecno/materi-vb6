VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H0000FF00&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Penghargaan"
   ClientHeight    =   6084
   ClientLeft      =   120
   ClientTop       =   384
   ClientWidth     =   8628
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6084
   ScaleWidth      =   8628
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Left            =   7320
      Top             =   4680
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   5760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3000
      TabIndex        =   1
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "<= Mungkin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   28.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Left            =   6240
      TabIndex        =   3
      Top             =   4440
      Width           =   2412
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mungkin =>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   28.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   2652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selamat Anda Bisa Menjadi Hacker Kelas Atas"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3972
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   8172
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Red, green, blue As Integer


Private Sub Command1_Click()
Form1.Show
Menu.Hide
End Sub


Private Sub Timer1_Timer()
If blue <= 255 Then blue = blue + 50 Else blue = 0
green = green + 50
If green >= 255 Then green = 0
Red = Red + 50
If Red >= 255 Then
Red = 0
End If
Label1.ForeColor = Int(RGB(Red, green, blue))
Label1.Refresh
End Sub

