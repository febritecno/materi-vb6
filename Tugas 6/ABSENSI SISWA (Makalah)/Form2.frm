VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   1308
   ClientLeft      =   3888
   ClientTop       =   2664
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1308
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "MASUK"
      Height          =   852
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   852
   End
   Begin VB.TextBox Text2 
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   1932
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   1932
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1092
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "admin" And Text2.Text = "123" Then
Unload Me
menu.admin.Visible = True
menu.login.Visible = False
menu.lihat.Visible = False
menu.keluar.Visible = True
Else
MsgBox "SALAH KAPRAH", vbCritical, "Info"
End If
End Sub
