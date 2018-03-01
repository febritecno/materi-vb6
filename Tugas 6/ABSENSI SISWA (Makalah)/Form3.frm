VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Siswa"
   ClientHeight    =   5196
   ClientLeft      =   996
   ClientTop       =   2256
   ClientWidth     =   10308
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5196
   ScaleWidth      =   10308
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5172
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10332
      _ExtentX        =   18225
      _ExtentY        =   9123
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   420
      BackColor       =   16744576
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "INPUT SISWA"
      TabPicture(0)   =   "Form3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Shape1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Shape2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "DataGrid1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Adodc1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Combo1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command10"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Command9"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command8"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command7"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text7"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Command6"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "BACKUP DATA"
      TabPicture(1)   =   "Form3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "Command4"
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(4)=   "Text5"
      Tab(1).Control(5)=   "Frame1"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0FF&
         Caption         =   "CARI"
         Height          =   252
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   720
         Width           =   1452
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   960
         TabIndex        =   29
         Top             =   720
         Width           =   3012
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080FFFF&
         Caption         =   "<<"
         Height          =   372
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4320
         Width           =   612
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H000080FF&
         Height          =   372
         Left            =   2040
         Picture         =   "Form3.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4320
         Width           =   372
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H000080FF&
         Height          =   372
         Left            =   2520
         Picture         =   "Form3.frx":014E
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4320
         Width           =   372
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0080FFFF&
         Caption         =   ">>"
         Height          =   372
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4320
         Width           =   612
      End
      Begin VB.Frame Frame1 
         Caption         =   "data"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -73800
         TabIndex        =   20
         Top             =   1320
         Width           =   8532
         Begin VB.DriveListBox Drive1 
            Height          =   288
            Left            =   240
            TabIndex        =   22
            Top             =   480
            Width           =   4575
         End
         Begin VB.DirListBox Dir1 
            Height          =   288
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   4575
         End
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72240
         TabIndex        =   19
         Top             =   3120
         Width           =   5532
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Back up"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71400
         TabIndex        =   18
         Top             =   3840
         Width           =   3495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67680
         TabIndex        =   17
         Top             =   4440
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1200
         TabIndex        =   16
         Top             =   2640
         Width           =   2532
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   372
         Left            =   480
         Top             =   5160
         Width           =   4572
         _ExtentX        =   8065
         _ExtentY        =   656
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=latihan.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=latihan.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "siswa"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form3.frx":0263
         Height          =   3852
         Left            =   6000
         TabIndex        =   15
         Top             =   960
         Width           =   4092
         _ExtentX        =   7218
         _ExtentY        =   6795
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "noinduk"
            Caption         =   "No Induk"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "noabsen"
            Caption         =   "No Absen"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "nama"
            Caption         =   "Nama"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "kelas"
            Caption         =   "Kelas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "alamat"
            Caption         =   "Alamat"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "no"
            Caption         =   "No HP"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1944
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1944
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1944
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1944
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1944
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1944
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Hapus"
         Height          =   732
         Left            =   4080
         TabIndex        =   14
         Top             =   3120
         Width           =   1332
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         Height          =   732
         Left            =   4080
         TabIndex        =   13
         Top             =   2280
         Width           =   1332
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tambah"
         Height          =   672
         Left            =   4080
         TabIndex        =   12
         Top             =   1440
         Width           =   1332
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1200
         TabIndex        =   11
         Top             =   3120
         Width           =   2532
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1200
         TabIndex        =   10
         Top             =   2160
         Width           =   2532
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1200
         TabIndex        =   9
         Top             =   1680
         Width           =   2532
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1200
         TabIndex        =   8
         Top             =   3720
         Width           =   2532
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   2532
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   14  'Copy Pen
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   3732
         Left            =   120
         Top             =   1080
         Width           =   5412
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   -73920
         TabIndex        =   24
         Top             =   4440
         Width           =   6132
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "BACK UP DATA"
         BeginProperty Font 
            Name            =   "Rosewood Std Regular"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   -71040
         TabIndex        =   23
         Top             =   480
         Width           =   2412
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   4092
         Left            =   5880
         Top             =   840
         Width           =   4332
      End
      Begin VB.Label Label6 
         Caption         =   "No HP"
         Height          =   492
         Left            =   240
         TabIndex        =   6
         Top             =   3720
         Width           =   1932
      End
      Begin VB.Label Label5 
         Caption         =   "Alamat"
         Height          =   492
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   2892
      End
      Begin VB.Label Label4 
         Caption         =   "Kelas"
         Height          =   492
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   2892
      End
      Begin VB.Label Label3 
         Caption         =   "Nama"
         Height          =   492
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   2892
      End
      Begin VB.Label Label2 
         Caption         =   "No Absen"
         Height          =   492
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   2892
      End
      Begin VB.Label Label1 
         Caption         =   "No Induk"
         Height          =   492
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   2892
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOP As shfileopstruct) As Long
Private Const FO_copy = &H2
Private Const fof_allowundo = &H40
 
Private Type shfileopstruct
    hwnd As Long
    wfunc As Long
    pfrom As String
    pto As String
    Fflags As Integer
    Faborted As Boolean
    hnamemaps As Long
    sprogress As String
End Type
Private Sub Command1_Click()
If Command1.Caption = "Tambah" Then
Command2.Enabled = False
Command3.Enabled = False
Command1.Caption = "SIMPAN"
hilang
hidup
Text1.SetFocus
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset!noinduk = Text1.Text
Adodc1.Recordset!noabsen = Text2.Text
Adodc1.Recordset!nama = Text3.Text
Adodc1.Recordset!kelas = Combo1.Text
Adodc1.Recordset!alamat = Text4.Text
Adodc1.Recordset!no = Text6.Text
Adodc1.Recordset.Update
Command2.Enabled = True
Command3.Enabled = True
MsgBox "Berhasil di Simpan", , "info"
hilang
mati
Call Form_Load
Command1.Caption = "Tambah"
End If
End Sub

Private Sub hilang()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Combo1.Text = ""
End Sub

Private Sub Command10_Click()
If Not Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveLast
    End If
       Call Form_Load
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Edit" Then
hidup
Text1.SetFocus
Command2.Caption = "SIMPAN"
Else
Adodc1.Recordset!noinduk = Text1.Text
Adodc1.Recordset!noabsen = Text2.Text
Adodc1.Recordset!nama = Text3.Text
Adodc1.Recordset!kelas = Combo1.Text
Adodc1.Recordset!alamat = Text4.Text
Adodc1.Recordset!no = Text6.Text
Adodc1.Recordset.Update
mati
Command2.Caption = "Edit"
End If
End Sub

Private Sub Command3_Click()
If Adodc1.Recordset.BOF Then
MsgBox "Gagal,Database Kosong", vbCritical, "Info"
Else
Adodc1.Recordset.Delete
Adodc1.Recordset.Update
hilang
DataGrid1.Refresh
End If
End Sub
Public Sub copy(ByVal asal As String, ByVal tujuan As String)
Dim X As shfileopstruct
    With X
  .hwnd = 0
        .wfunc = FO_copy
        .pfrom = asal & vbNullChar & vbNullChar
        .pto = tujuan & vbNullChar & vbNullChar
        .Fflags = fof_allowundo
            End With
    SHFileOperation X
End Sub


Private Sub Command5_Click()
On Error Resume Next
If Label5.Caption = "" Then
    MsgBox "Anda belum memilih file yang akan dicopy"
    Exit Sub
ElseIf Text5 = "" Then
    MsgBox "Anda tidak memilih direktori tujuan peng-Copy-an"
    Exit Sub
End If
copy Label8.Caption, Text5.Text
MsgBox "Berhasil di Backup"
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Command4_Click()
menu.Show
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Find "noinduk='" + Text7 + "'", , adSearchForward, 1
If Not Adodc1.Recordset.EOF Then
Text1.Text = Adodc1.Recordset!noinduk
Else
MsgBox "Nomer Tidak Ada ??", vbCritical, "Information"
End If
End Sub

Private Sub Command7_Click()
   If Not Adodc1.Recordset.BOF Then
       Adodc1.Recordset.MoveFirst
    End If
    Call Form_Load
End Sub

Private Sub Command8_Click()
 Adodc1.Recordset.MovePrevious
 If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveNext
 End If
    Call Form_Load
End Sub

Private Sub Command9_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MovePrevious
  End If
 Call Form_Load
End Sub

Private Sub Dir1_Change()
Text5.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub


Private Sub DataGrid1_Click()
Text1.Text = Adodc1.Recordset!noinduk
Text2.Text = Adodc1.Recordset!noabsen
Text3.Text = Adodc1.Recordset!nama
Combo1.Text = Adodc1.Recordset!kelas
Text4.Text = Adodc1.Recordset!alamat
Text6.Text = Adodc1.Recordset!no
End Sub

Private Sub Form_Load()
Combo1.AddItem "X-RPL"
Combo1.AddItem "X-TKJ"
Combo1.AddItem "X-AKT"
Label8.Caption = App.Path & "\latihan.mdb"
Dir1.Path = "C:\"
Text1.Text = Adodc1.Recordset!noinduk
Text2.Text = Adodc1.Recordset!noabsen
Text3.Text = Adodc1.Recordset!nama
Combo1.Text = Adodc1.Recordset!kelas
Text4.Text = Adodc1.Recordset!alamat
Text6.Text = Adodc1.Recordset!no
End Sub

Sub hidup()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text6.Enabled = True
Combo1.Enabled = True
End Sub

Sub mati()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text6.Enabled = False
Combo1.Enabled = False
End Sub

