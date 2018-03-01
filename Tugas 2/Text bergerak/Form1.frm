VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6372
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10248
   BeginProperty Font 
      Name            =   "Algerian"
      Size            =   48
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6372
   ScaleWidth      =   10248
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   360
      Top             =   2280
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "smk negeri sumberrejo"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Red, green, blue As Integer




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
