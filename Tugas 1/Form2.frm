VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Masuk"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form2"
   ScaleHeight     =   2040
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Masuk"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2145
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Hitung As Integer
Dim Jawab As String
Private Sub Form_Load()
  Hitung = 0
              
End Sub

Private Sub cmdOK_Click()
  
  Do While Text1.Text <> "T-Blog"
    Jawab = Text1.Text = "T-Blog"
                                        
      
      If Jawab <> "masino" Then
         Hitung = Hitung + 1
         Tampung (Hitung)
         If Hitung = 3 Then
            
            Print "Password Salah ??"
            Text1.Enabled = False
            
            cmdOK.Enabled = False
            
            cmdCancel.Default = True
           
         End If
         Exit Sub
      Else
         Exit Do
      End If
  Loop
  Print "Password Benar !!"
  
  Form1.Show
  Form2.Visible = False
End Sub

Function Tampung(Hitung)
Dim Hasil As Integer
    Hasil = 0
    Hasil = Hasil + Hitung
    Text1.SetFocus
    SendKeys "{Home}+{End}"
    Print "Kesempatan ke-" & Hasil
    
End Function

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   cmdOK.Default = True
End Sub


