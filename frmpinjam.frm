VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form frmpinjam 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pinjam Buku"
   ClientHeight    =   6024
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   8304
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6024
   ScaleWidth      =   8304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   6000
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   16.2
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   384
      Left            =   3000
      TabIndex        =   16
      Text            =   "PINJAM"
      Top             =   4560
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmpinjam.frx":0000
      Height          =   732
      Left            =   600
      TabIndex        =   12
      Top             =   6360
      Width           =   3372
      _ExtentX        =   5948
      _ExtentY        =   1291
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "No_buku"
         Caption         =   "No_buku"
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
         DataField       =   "Nama"
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
      BeginProperty Column02 
         DataField       =   "Alamat"
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
      BeginProperty Column03 
         DataField       =   "Tanggal"
         Caption         =   "Tanggal"
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
         DataField       =   "Nominal_hari"
         Caption         =   "Nominal_hari"
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
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1955,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1955,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1955,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1955,906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   3720
      Top             =   6240
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   572
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=perpustakaan.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=perpustakaan.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "pinjam"
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
   Begin VB.CommandButton Command2 
      Caption         =   "BATAL"
      Height          =   492
      Left            =   5160
      TabIndex        =   11
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "PINJAM"
      Height          =   492
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1932
   End
   Begin VB.ComboBox Combo1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   384
      ItemData        =   "frmpinjam.frx":0015
      Left            =   3000
      List            =   "frmpinjam.frx":0017
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   3960
      Width           =   4812
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   3
      Top             =   3240
      Width           =   1932
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   3000
      TabIndex        =   2
      Top             =   2520
      Width           =   4812
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   3000
      TabIndex        =   1
      Top             =   1800
      Width           =   4812
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   408
      Left            =   3000
      TabIndex        =   0
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SILAHKAN ISI DATA DIBAWAH INI"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   0
      TabIndex        =   17
      Top             =   120
      Width           =   8292
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Status Pengunjung :"
      BeginProperty Font 
         Name            =   "RoboKoz"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   324
      Left            =   480
      TabIndex        =   15
      Top             =   4560
      Width           =   1884
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ex : Jl\RT\Des\Kec\Kab\Prop"
      Enabled         =   0   'False
      Height          =   492
      Left            =   5760
      TabIndex        =   14
      Top             =   2280
      Width           =   2292
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ex : 0001 ( Max 4 Karakter )"
      Enabled         =   0   'False
      Height          =   252
      Left            =   4080
      TabIndex        =   13
      Top             =   1080
      Width           =   2052
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Berapa Hari :"
      BeginProperty Font 
         Name            =   "RoboKoz"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   324
      Left            =   480
      TabIndex        =   9
      Top             =   3960
      Width           =   1668
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan Tanggal Pinjam Buku :"
      BeginProperty Font 
         Name            =   "RoboKoz"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   612
      Left            =   480
      TabIndex        =   8
      Top             =   3120
      Width           =   2400
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan Alamat Lengkap :"
      BeginProperty Font 
         Name            =   "RoboKoz"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   2412
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan Kode Buku :"
      BeginProperty Font 
         Name            =   "RoboKoz"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   324
      Left            =   480
      TabIndex        =   6
      Top             =   1080
      Width           =   2016
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan Nama Anda :"
      BeginProperty Font 
         Name            =   "RoboKoz"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   324
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   2088
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   4332
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   8052
   End
End
Attribute VB_Name = "frmpinjam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    On Error Resume Next
    sipp
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Combo1 = "" Then
        MsgBox "Woi Ada Yang Belum Di Isi..!!!", vbOKOnly + vbCritical, "Error"
    Else
    sipp
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields("no_buku") = Text1.Text
        Adodc1.Recordset.Fields("nama") = Text2.Text
        Adodc1.Recordset.Fields("alamat") = Text3.Text
        Adodc1.Recordset.Fields("tanggal") = Text4.Text
        Adodc1.Recordset.Fields("status") = Text5.Text
        Adodc1.Recordset.Fields("nominal_hari") = Combo1.Text
        Adodc1.Recordset.Update
        MsgBox "Data Anda Telah Tersimpan!!, Silahkan Ambil Buku Dan Bilang Petugas", vbOKOnly + vbInformation, "Berhasil!"
    frmCari.Show
    Unload Me
    End If
End Sub

Private Sub Command2_Click()
Menu.Show
Unload Me
mati
Combo1.Text = ""
End Sub

Private Sub Form_Load()
Text1.MaxLength = 4
Combo1.AddItem "2 Hari"
Combo1.AddItem "3 Hari"
Combo1.AddItem "5 Hari"
Combo1.AddItem "1 Minggu"
Combo1.AddItem "1 Bulan"
End Sub

Private Sub Timer1_Timer()
Text4.Text = DateTime.Date
End Sub
 Sub mati()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
 Combo1.Text = ""
 End Sub
