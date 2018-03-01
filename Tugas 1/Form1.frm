VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Aplikasi Sederhana"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   756
   ClientWidth     =   7380
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   2  'Cross
   ScaleHeight     =   5280
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Ukuran Font"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   17
      Top             =   4680
      Width           =   3855
      Begin VB.OptionButton Option11 
         Caption         =   "45 px"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2280
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option10 
         Caption         =   "35 px"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option9 
         Caption         =   "25 px"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   840
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option8 
         Caption         =   "12 px"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "65 px"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Font Gaya"
      Height          =   1335
      Left            =   5520
      TabIndex        =   9
      Top             =   3240
      Width           =   1695
      Begin VB.CheckBox Check4 
         Caption         =   "Garis Tengah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Garis Bawah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Tebal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Miring"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   2295
      Begin VB.OptionButton Option7 
         Caption         =   "Putih"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1320
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Hitam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Kuning"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Biru"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Hijau"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Merah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Prestige Elite Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Program :  Febrian Dwi Putra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   2295
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan Kata :"
      BeginProperty Font 
         Name            =   "Stencil Std"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2025
   End
   Begin VB.Menu Cmd_File 
      Caption         =   "&File"
      Begin VB.Menu Cmd_New 
         Caption         =   "&New"
         Checked         =   -1  'True
         Shortcut        =   ^N
      End
      Begin VB.Menu Cmd_About 
         Caption         =   "&About Me"
      End
      Begin VB.Menu Cmd_Exit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Cmd_color 
      Caption         =   "&Warna Form"
      Begin VB.Menu Cmd_Default 
         Caption         =   "&Default"
         Checked         =   -1  'True
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu Cmd_Hijau 
         Caption         =   "&Hijau"
      End
      Begin VB.Menu Cmd_Merah 
         Caption         =   "&Merah"
      End
      Begin VB.Menu Cmd_Kuning 
         Caption         =   "&Kuning"
      End
      Begin VB.Menu Cmd_Putih 
         Caption         =   "&Putih"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_unload(Cancel As Integer)
If MsgBox("Mau Keluar Ya..? Tekan Yes", _
              vbYesNo + vbQuestion, _
              "Keluar") = vbNo Then
        Cancel = 1
End If
End Sub

Private Sub Check1_Click()
Label2.FontItalic = Check1.Value
End Sub

Private Sub Check2_Click()
Label2.FontBold = Check2.Value
End Sub

Private Sub Check3_Click()
Label2.FontUnderline = Check3.Value
End Sub

Private Sub Check4_Click()
Label2.FontStrikethru = Check4.Value
End Sub

Private Sub Cmd_About_Click()
FormAbout.Show
End Sub
Private Sub Cmd_Default_Click()
Form1.BackColor = vbDefault
End Sub
Private Sub Cmd_Exit_Click()
End
End Sub
Private Sub Cmd_Hijau_Click()
Form1.BackColor = vbGreen
End Sub


Private Sub Cmd_Merah_Click()
Form1.BackColor = vbRed
End Sub

Private Sub Cmd_Putih_Click()
Form1.BackColor = vbWhite
End Sub

Private Sub Cmd_Kuning_Click()
Form1.BackColor = vbYellow
End Sub

Private Sub Cmd_New_Click()
If MsgBox("Buat Baru...??", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
Text1.Text = Empty
Label2.Caption = Empty
'Option buttom
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Option9.Value = False
Option10.Value = False
Option11.Value = False
'check box
Check1.Value = Unchecked
Check2.Value = Unchecked
Check3.Value = Unchecked
Check4.Value = Unchecked
Command1.Visible = True
Text1.SetFocus
End If
End Sub


Private Sub Command1_Click()
If Text1.Text = "" Then
   MsgBox ("Tolong Masukan Kata !!")
   Text1.SetFocus
Else
    Call Febri
End If
End Sub
Sub Febri()
Label2.Caption = Text1.Text
Command1.Visible = False
Text1.SetFocus
End Sub
'Warna Kata
Private Sub Option1_Click()
Label2.ForeColor = vbRed
End Sub

Private Sub Option10_Click()
Label2.FontSize = 35
End Sub

Private Sub Option11_Click()
Label2.FontSize = 45
End Sub

Private Sub Option2_Click()
Label2.ForeColor = vbGreen
End Sub
Private Sub Option3_Click()
Label2.ForeColor = vbBlue
End Sub

Private Sub Option4_Click()
Label2.ForeColor = vbYellow
End Sub
Private Sub Option5_Click()
Label2.FontSize = 65
End Sub

Private Sub Option6_Click()
Label2.ForeColor = vbBlack
End Sub

Private Sub Option7_Click()
Label2.ForeColor = vbWhite
End Sub
Private Sub Option8_Click()
Label2.FontSize = 12
End Sub

Private Sub Option9_Click()
Label2.FontSize = 25
End Sub



