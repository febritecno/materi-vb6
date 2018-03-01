VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormAbout 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "About Program"
   ClientHeight    =   3495
   ClientLeft      =   2520
   ClientTop       =   1125
   ClientWidth     =   5250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame frmConAuthor 
      BackColor       =   &H80000003&
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox PicFrmConAut 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   120
         ScaleHeight     =   141
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   313
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "febrikoplo0@yahoo.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   480
            TabIndex        =   10
            Top             =   1200
            Width           =   2550
         End
         Begin VB.Label lbSite 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "http://tecno-yes.blogspot.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   480
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   360
            Width           =   3090
         End
         Begin VB.Label lbSiteX 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "- Please send Your Comment to E-mail:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   2790
         End
         Begin VB.Label lbSiteX 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "- Programmer's Site :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   60
            Width           =   1500
         End
         Begin VB.Label lbMail2 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "febrikondang0@yahoo.com"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   480
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   1560
            Width           =   1995
         End
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   0
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   1
      Top             =   0
      Width           =   5235
      Begin VB.TextBox TxMain 
         BackColor       =   &H00C0FFC0&
         Height          =   2175
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   840
         Width           =   4695
      End
      Begin ComctlLib.TabStrip TabMain 
         Height          =   2955
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5212
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Tugas Vb"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "About Programmer"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Contact Programmer"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Mongolian Baiti"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label lbMail2 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "febrikondang0@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   840
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2640
      Width           =   1995
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Form1.Show
  FormAbout.Visible = False
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub TabMain_Click()
'On Error Resume Next
    Select Case TabMain.SelectedItem.Index
    Case 1
        frmConAuthor.Visible = False
        TxMain.Visible = True
        TxMain.Text = ReadMeText
    Case 2
        frmConAuthor.Visible = False
        TxMain.Visible = True
        TxMain.Text = AboutText
    Case 3
        frmConAuthor.Visible = True
        TxMain.Visible = False
    End Select
End Sub
Private Function ReadMeText() As String
'On Error Resume Next
ReadMeText = _
    Chr$(34) & "Tugas Vb" & Chr$(34) & vbCrLf & _
    "Tugas ini adalah Tugas yang dikhususkan untuk Pengijahan Kata dan Memunculkan kata." & vbCrLf & _
    Chr$(34) & "Latar Belakang Program" & Chr$(34) & vbCrLf & _
    "Saya Membuat Program Ini Karena saya Ingin Bisa menjadi Programer" & vbCrLf & _
    Chr$(34) & "Please Contact Me" & Chr$(34) & vbCrLf & _
    "Silahkan Bertanya Di http://tecno-yes.blogspot.com" & vbCrLf & _
    Chr$(34) & "Tugas Vb 6.0 is Freeware" & Chr$(34) & vbCrLf & _
    "Anda bebas menggunakan dan menyebarluaskan Tugas selama bukan untuk kepentingan komersial. " & vbCrLf & _
    Chr$(34) & "Tugas Vb 6.0 Bug" & Chr$(34) & vbCrLf & _
    "Segala bentuk kerusakan yang mungkin diakibatkan oleh penggunaan Program ini diluar tanggung jawab programmer."
End Function

Private Function AboutText() As String
AboutText = _
    "Name : Febrian Dwi Putra " & vbCrLf & _
    "Country: Indonesia" & vbCrLf & _
    "Province: Jawa Timur" & vbCrLf & _
    "City: Bojonegoro" & vbCrLf & _
    "Job: " & vbCrLf & _
    "    1. Pelajar Smk" & vbCrLf & _
    "    2. Ingin Punya Keahlian" & vbCrLf & _
    "    3. Mencari Bakat" & vbCrLf & _
    "School: Smkn 1 Sumberrejo" & vbCrLf & _
    "Class: X-Rpl"
End Function
