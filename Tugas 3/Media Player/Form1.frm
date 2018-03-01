VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Media Player wkwk"
   ClientHeight    =   8376
   ClientLeft      =   192
   ClientTop       =   816
   ClientWidth     =   10428
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8376
   ScaleWidth      =   10428
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4680
      Top             =   6600
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Buka aja"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5520
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Media Player AOE RPL"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   21.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10095
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   10335
      URL             =   "m"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   18225
      _cy             =   12933
   End
   Begin VB.Menu PB 
      Caption         =   "&Pilih Background"
      Begin VB.Menu blank 
         Caption         =   "Default"
         Checked         =   -1  'True
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu merah 
         Caption         =   "Merah"
      End
      Begin VB.Menu biru 
         Caption         =   "Biru"
      End
      Begin VB.Menu kuning 
         Caption         =   "kuning"
      End
      Begin VB.Menu hijau 
         Caption         =   "Hijau"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Red, green, blue As Integer






Private Sub biru_Click()
Form1.BackColor = vbBlue
End Sub

Private Sub blank_Click()
Form1.BackColor = vbWhite
End Sub

Private Sub cmdOpen_Click()
On Error Resume Next
CommonDialog1.ShowOpen
WindowsMediaPlayer1.URL = CommonDialog1.FileName
End Sub

Private Sub cmdExit_Click()
If MsgBox("Bisakah anda keluar,,?", vbQuestion + vbYesNo, "Keluar gak,,?") = vbYes Then
End
End If
End Sub

Private Sub hijau_Click()
Form1.BackColor = vbGreen
End Sub

Private Sub kuning_Click()
Form1.BackColor = vbYellow
End Sub

Private Sub merah_Click()
Form1.BackColor = vbRed
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
