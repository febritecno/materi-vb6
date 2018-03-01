VERSION 5.00
Begin VB.Form menu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Absensi Siswa"
   ClientHeight    =   5256
   ClientLeft      =   2928
   ClientTop       =   1836
   ClientWidth     =   6816
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menu.frx":0000
   ScaleHeight     =   5256
   ScaleWidth      =   6816
   ShowInTaskbar   =   0   'False
   Begin VB.Menu menu 
      Caption         =   "File"
      Begin VB.Menu login 
         Caption         =   "Login"
         Checked         =   -1  'True
      End
      Begin VB.Menu keluar 
         Caption         =   "Admin Keluar"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu admin 
      Caption         =   "Admin"
      Visible         =   0   'False
      Begin VB.Menu mulai 
         Caption         =   "Mulai Absen Siswa"
      End
      Begin VB.Menu dtsis 
         Caption         =   "Data Siswa"
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu lapor 
         Caption         =   "Laporan Absen Siswa"
      End
      Begin VB.Menu lapor2 
         Caption         =   "Laporan Data Siswa"
      End
   End
   Begin VB.Menu lihat 
      Caption         =   "Lihat Daftar Absensi Siswa"
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dtsis_Click()
Form3.Show
End Sub

Private Sub exit_Click()
pesan = MsgBox("Anda Yakin Ingin Keluar Dari Program ini?", vbQuestion + vbYesNo, "Keluar")
If pesan = vbYes Then
Animation
Form1.Hide
End
Else
End If
End Sub


Private Sub keluar_Click()
admin.Visible = False
login.Visible = True
lihat.Visible = True
keluar.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
Animation
End Sub
Public Sub Animation()
Dim I As Long
Dim J As Long
I = Me.ScaleHeight
J = Me.ScaleWidth


While Not I = 0
Me.Height = Me.Height - 25
I = I - 1
Wend


While Not J = 0
Me.Width = Me.Width - 25
J = J - 1
Wend
End Sub
Private Sub lapor_Click()
koneksi
RsAbsen.Open "select * from absen", ConN
If Not RsAbsen.EOF Then
Set DataReport1.DataSource = RsAbsen
DataReport1.Show
Else
MsgBox "Tidak ada Data"
End If
End Sub
Private Sub lapor2_Click()
koneksi
RsAbsen.Open "select * from siswa", ConN
If Not RsAbsen.EOF Then
Set DataReport2.DataSource = RsAbsen
DataReport2.Show
Else
MsgBox "Tidak ada Data"
End If
End Sub

Private Sub lihat_Click()
Form4.Show
End Sub

Private Sub login_Click()
Form2.Show
End Sub

Private Sub mulai_Click()
Form1.Show
End Sub
