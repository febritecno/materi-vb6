VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2496
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2496
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar Bar 
      Height          =   444
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   6972
      _ExtentX        =   12298
      _ExtentY        =   783
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1920
      Top             =   2520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "PROGRAM APLIKASI ABSENSI SISWA SMKN SUMBERREJO"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7212
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   972
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7212
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Bar.Value = Bar.Value + 2
Screen.MousePointer = vbHourglass
Label4.Caption = Bar.Value & " %"
If Bar.Value < 20 Then
ElseIf Bar.Value < 40 Then
ElseIf Bar.Value < 60 Then
ElseIf Bar.Value < 80 Then
ElseIf Bar.Value < 100 Then
End If
If Bar.Value = 100 Then
If Timer1.Interval >= 1 Then
Unload frmSplash
menu.Show
Screen.MousePointer = vbDefault
End If
End If
Exit Sub
End Sub

