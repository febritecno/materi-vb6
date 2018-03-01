VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Absensi Siswa"
   ClientHeight    =   7680
   ClientLeft      =   768
   ClientTop       =   600
   ClientWidth     =   11016
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11016
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Batal"
      Height          =   252
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1680
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080FFFF&
      Caption         =   ">>"
      Height          =   372
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5160
      Width           =   612
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000080FF&
      Height          =   372
      Left            =   6360
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5160
      Width           =   372
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000080FF&
      Height          =   372
      Left            =   5880
      Picture         =   "Form1.frx":0115
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5160
      Width           =   372
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FFFF&
      Caption         =   "<<"
      Height          =   372
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5160
      Width           =   612
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FF00&
      Caption         =   "Cari Nomer Induk"
      Height          =   372
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4440
      Width           =   2052
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   4920
      TabIndex        =   32
      Top             =   4440
      Width           =   3732
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":022B
      Height          =   1692
      Left            =   240
      TabIndex        =   31
      Top             =   5880
      Width           =   10572
      _ExtentX        =   18648
      _ExtentY        =   2985
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "TABEL ABSENSI SISWA"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "nrp"
         Caption         =   "Nomer Induk Siswa"
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
         DataField       =   "nama"
         Caption         =   "Nama Siswa"
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
         DataField       =   "jurusan"
         Caption         =   "Jenis Kelamin"
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
         DataField       =   "matkul"
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
         DataField       =   "masuk"
         Caption         =   "Masuk"
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
         DataField       =   "izin"
         Caption         =   "Izin"
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
      BeginProperty Column06 
         DataField       =   "sakit"
         Caption         =   "Sakit"
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
      BeginProperty Column07 
         DataField       =   "alpa"
         Caption         =   "Alpa"
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
      BeginProperty Column08 
         DataField       =   "total"
         Caption         =   "Total Keterangan Tidak Hadir"
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
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   684,284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   11520
      Top             =   8640
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
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
      RecordSource    =   "absen"
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Selesai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5040
      Width           =   2892
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hapus"
      Height          =   492
      Left            =   8760
      Picture         =   "Form1.frx":0240
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3600
      Width           =   2052
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Absen Siswa"
      Height          =   612
      Left            =   8880
      Picture         =   "Form1.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2760
      Width           =   1812
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Edit Data"
      Height          =   612
      Left            =   8880
      Picture         =   "Form1.frx":0471
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2040
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tambah Siswa"
      Height          =   612
      Left            =   8880
      Picture         =   "Form1.frx":0595
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1320
      Width           =   1812
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hitung Absen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2892
      Left            =   4920
      TabIndex        =   9
      Top             =   1200
      Width           =   3735
      Begin VB.Line Line3 
         X1              =   0
         X2              =   3720
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   3720
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3720
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3000
         TabIndex        =   19
         Top             =   2520
         Width           =   612
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Left            =   3000
         TabIndex        =   18
         Top             =   2040
         Width           =   612
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   3000
         TabIndex        =   17
         Top             =   1440
         Width           =   612
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   252
         Left            =   3000
         TabIndex        =   16
         Top             =   960
         Width           =   612
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Tidak Masuk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   1932
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanpa Keterangan"
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Width           =   1452
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Izin"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sakit"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Masuk"
         Height          =   252
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1452
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Alasan Tidak Hadir"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   4092
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Tanpa Keterangan/Alpa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   3132
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Izin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1332
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Sakit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1332
      End
   End
   Begin VB.ComboBox Combo3 
      Height          =   288
      Left            =   1440
      TabIndex        =   4
      Top             =   3480
      Width           =   3015
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      Left            =   1440
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   1440
      TabIndex        =   2
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   612
      Left            =   4920
      Top             =   5040
      Width           =   2772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H008080FF&
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillStyle       =   5  'Downward Diagonal
      Height          =   2292
      Left            =   8760
      Top             =   1200
      Width           =   2052
   End
   Begin VB.Label no 
      BackStyle       =   0  'Transparent
      Caption         =   "No Induk :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   30
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label nama 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   29
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label jenis 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Kelamin :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   360
      TabIndex        =   28
      Top             =   2280
      Width           =   1092
   End
   Begin VB.Label kelas 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   27
      Top             =   3000
      Width           =   972
   End
   Begin VB.Label Labelhadir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kehadiran :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   26
      Top             =   3480
      Width           =   960
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "SILAHKAN ABSEN SISWA "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   696
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   10692
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   4572
      Left            =   240
      Top             =   1080
      Width           =   4332
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo3_Click()
If Combo3.Text = "Hadir" Then
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Label11.Caption = Val(Label11.Caption) + 1
Else
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Tambah Siswa" Then
kosong
Command1.Enabled = True
Command3.Enabled = True
Text1.Enabled = True
            Text2.Enabled = True
            Combo1.Enabled = True
            Combo2.Enabled = True
            Combo3.Enabled = False
            Text1.SetFocus
            Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command11.Visible = True
mati1
   Command1.Caption = "Simpan"
Else
             If Text1 = "" Or Text2 = "" Or Combo1 = "" Or Combo2 = "" Then
        MsgBox "Masih ada data yang kosong..!!!", vbCritical, "Error"
        Text1.SetFocus
        Else
Dim SQLSimpan As String
            SQLSimpan = "Insert Into absen (nrp,nama,jurusan,matkul,masuk,sakit,izin,alpa,total) values ('" & Text1 & "','" & Text2 & "','" & Combo1.Text & "','" & Combo2.Text & "','" & Label11.Caption & "','" & Label12.Caption & "','" & Label13.Caption & "','" & Label14.Caption & "','" & Label15.Caption & "')"
            ConN.Execute SQLSimpan
             Form_Activate
             mati
             hidup1
            kosong
                     Command2.Enabled = True
Command3.Enabled = True
                      Command4.Enabled = True
                      penuh
                      Command11.Visible = False
            Command1.Caption = "Tambah Siswa"
            End If
            End If
End Sub

Private Sub Command10_Click()
If Not Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveLast
    End If
       Call DataGrid1_Click
End Sub

Private Sub Command11_Click()
Command1.Caption = "Tambah Siswa"
hidup1
mati
penuh
Command11.Visible = False
Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Edit Data" Then
Text1.Enabled = False
            Text2.Enabled = True
            Combo1.Enabled = True
            Combo2.Enabled = True
            Combo3.Enabled = False
Command2.Caption = "Simpan"
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
mati1
Text2.SetFocus
Else
Dim SQLAbsen As String
            SQLAbsen = "Update absen Set nama='" & Text2.Text & "'," & " matkul='" & Combo2.Text & "'," & " jurusan='" & Combo1.Text & "' where nrp='" & Text1 & "'"
            ConN.Execute SQLAbsen
            Form_Activate
            Call kosong
            Command1.Enabled = True
            Command3.Enabled = True
            Command4.Enabled = True
            hidup1
            penuh
            Command2.Caption = "Edit Data"
            End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Absen Siswa" Then
Combo3.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
mati1
Command3.Caption = "Simpan"
Else
Dim SQLAbsen As String
            SQLAbsen = "Update absen Set masuk= '" & Label11.Caption & "'," & " sakit='" & Label12.Caption & "'," & " izin='" & Label13.Caption & "'," & " alpa='" & Label14.Caption & "'," & " total='" & Label15.Caption & "' where nrp='" & Text1 & "'"
            ConN.Execute SQLAbsen
            Form_Activate
            Command9_Click
            Call kosong
            hidup1
            Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Combo3.Enabled = False
penuh
            Command3.Caption = "Absen Siswa"
            End If
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
DataGrid1.Refresh
mati
kosong
End Sub

Private Sub Command5_Click()
pesan = MsgBox("Anda Yakin Ingin Keluar Dari Program ini?", vbQuestion + vbYesNo, "Keluar")
If pesan = vbYes Then
Animation
Unload Me
Else
End If
End Sub

Private Sub Command6_Click()
If Text3.Text = "" Then
MsgBox "Tolong Masukan Nomer Induk Siswa !!", vbCritical, "information"
End If
Adodc1.Recordset.Find "nrp='" + Text3 + "'", , adSearchForward, 1
If Not Adodc1.Recordset.EOF Then
Text1.Text = Adodc1.Recordset!nrp
Text2.Text = Adodc1.Recordset!nama
Else
MsgBox "Nomer Tidak Ada ??", vbCritical, "Information"
Text3.Text = ""
Text3.SetFocus
End If
End Sub

Private Sub Command7_Click()
   If Not Adodc1.Recordset.BOF Then
       Adodc1.Recordset.MoveFirst
    End If
    Call DataGrid1_Click
End Sub

Private Sub Command8_Click()
 Adodc1.Recordset.MovePrevious
 If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveNext
 End If
    Call DataGrid1_Click
End Sub

Private Sub Command9_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MovePrevious
  End If
 Call DataGrid1_Click
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
If Adodc1.Recordset.BOF Then
    MsgBox "Tidak ada data!", vbOKOnly, "Informasi!"
Else
Call koneksi
Combo3.Enabled = False
    Text1 = Adodc1.Recordset("nrp")
    Text2 = Adodc1.Recordset("nama")
    Combo1 = Adodc1.Recordset("jurusan")
   Combo2 = Adodc1.Recordset("matkul")
    Label11 = Adodc1.Recordset("masuk")
    Label12 = Adodc1.Recordset("sakit")
        Label13 = Adodc1.Recordset("izin")
            Label14 = Adodc1.Recordset("alpa")
                Label15 = Adodc1.Recordset("total")
                End If
End Sub

Private Sub Form_Activate()
Call koneksi
Adodc1.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\latihan.mdb"
Adodc1.RecordSource = "absen"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Label11.Caption = 0
Label12.Caption = 0
Label13.Caption = 0
Label14.Caption = 0
Label15.Caption = 0
Combo1.AddItem "Laki-Laki"
Combo1.AddItem "Perempuan"
Combo2.AddItem "X-RPL"
Combo2.AddItem "X-TKJ"
Combo2.AddItem "X-AKT"
Combo3.AddItem "Hadir"
Combo3.AddItem "Tidak Hadir"
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
    Text1 = Adodc1.Recordset("nrp")
    Text2 = Adodc1.Recordset("nama")
    Combo1 = Adodc1.Recordset("jurusan")
   Combo2 = Adodc1.Recordset("matkul")
    Label11 = Adodc1.Recordset("masuk")
    Label12 = Adodc1.Recordset("sakit")
        Label13 = Adodc1.Recordset("izin")
            Label14 = Adodc1.Recordset("alpa")
                Label15 = Adodc1.Recordset("total")
End Sub
Private Sub Option1_Click()
If Option1.Value = True Then
Label12.Caption = Val(Label12.Caption) + 1
Else
Label12.Caption = Val(Label12.Caption) + 0
End If
Label15.Caption = Val(Label12.Caption) + Val(Label13.Caption) + Val(Label14.Caption)
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Label13.Caption = Val(Label13.Caption) + 1
Else
Label13.Caption = Val(Label13.Caption) + 0
End If
Label15.Caption = Val(Label12.Caption) + Val(Label13.Caption) + Val(Label14.Caption)
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Label14.Caption = Val(Label14.Caption) + 1
Else
Label14.Caption = Val(Label14.Caption) + 0
End If
Label15.Caption = Val(Label12.Caption) + Val(Label13.Caption) + Val(Label14.Caption)
End Sub

Function CariData()
    Call koneksi
    RsAbsen.Open "Select * From absen where nrp='" & Text1 & "'", ConN
End Function

Private Sub TampilkanData()
Text2 = RsAbsen!nama
Combo1.Text = RsAbsen!jurusan
Combo2.Text = RsAbsen!matkul
Label11.Caption = RsAbsen!masuk
Label12.Caption = RsAbsen!sakit
Label13.Caption = RsAbsen!izin
Label14.Caption = RsAbsen!alpa
Label15.Caption = RsAbsen!total
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text1_LostFocus()
On Error Resume Next
Call CariData
        If Not RsAbsen.EOF Then
            TampilkanData
            Text1.Enabled = False
            Text2.Enabled = False
            Combo1.Enabled = False
            Combo2.Enabled = False
            MsgBox "Nomer Induk Sudah Ada", vbCritical, "Information"
            Command1.Enabled = True
        End If
End Sub

Private Sub kosong()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Label11.Caption = 0
Label12.Caption = 0
Label13.Caption = 0
Label14.Caption = 0
Label15.Caption = 0
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
End Sub

Private Sub mati()
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
End Sub

Private Sub mati1()
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
End Sub

Private Sub hidup1()
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
End Sub

Private Sub penuh()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
    Text1 = Adodc1.Recordset("nrp")
    Text2 = Adodc1.Recordset("nama")
    Combo1 = Adodc1.Recordset("jurusan")
   Combo2 = Adodc1.Recordset("matkul")
    Label11 = Adodc1.Recordset("masuk")
    Label12 = Adodc1.Recordset("sakit")
        Label13 = Adodc1.Recordset("izin")
            Label14 = Adodc1.Recordset("alpa")
                Label15 = Adodc1.Recordset("total")
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

