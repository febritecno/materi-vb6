VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log In"
   ClientHeight    =   3396
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   7692
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3396
   ScaleWidth      =   7692
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   4452
   End
   Begin VB.TextBox TxtUser 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   4452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Stencil Std"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   5160
      TabIndex        =   1
      Top             =   2280
      Width           =   2292
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "MASUK"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan Username :"
      BeginProperty Font 
         Name            =   "One Stroke Script LET"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   3372
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan Password :"
      BeginProperty Font 
         Name            =   "One Stroke Script LET"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Form_Load()
On Error Resume Next
txtPassword.MaxLength = 7
TxtUser.MaxLength = 5
End Sub




Private Sub Command1_Click()
If TxtUser.Text = "febri" And txtPassword.Text = "blogger" Then
MsgBox "Selamat Anda Bisa Masuk", vbInformation
Unload Me
Menu.Show
Else
MsgBox "Rasain Loo..?? User dan Password Salah", vbCritical
Command1.Enabled = False
TxtUser.Text = ""
txtPassword.Text = ""
TxtUser.SetFocus
SendKeys "{Home}+{End}"
End If
End Sub


Private Sub Command2_Click()
a = MsgBox("Mau Keluar..???", vbOKCancel + vbInformation)
If a = vbOK Then
End
End If
End Sub

Private Sub txtPassword_Change()
If txtPassword.Text = "blogger" Then
Command1.Enabled = True
End If
End Sub

