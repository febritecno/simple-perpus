VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmCari 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "     "
   ClientHeight    =   7548
   ClientLeft      =   432
   ClientTop       =   888
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7548
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      Caption         =   ">"
      Height          =   492
      Left            =   1320
      TabIndex        =   21
      Top             =   4080
      Width           =   372
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<"
      Height          =   492
      Left            =   960
      TabIndex        =   20
      Top             =   4080
      Width           =   372
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<<"
      Height          =   492
      Left            =   240
      TabIndex        =   19
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      Height          =   492
      Left            =   1920
      TabIndex        =   18
      Top             =   4080
      Width           =   612
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "PINJAM"
      Height          =   492
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4080
      Width           =   2172
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KEMBALI"
      Height          =   492
      Left            =   5160
      TabIndex        =   16
      Top             =   4080
      Width           =   1812
   End
   Begin MSDataGridLib.DataGrid DataBuku 
      Bindings        =   "frmCari.frx":0000
      Height          =   2412
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   6852
      _ExtentX        =   12086
      _ExtentY        =   4255
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
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
      Caption         =   "DATA BUKU YANG TERSEDIA"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "kode"
         Caption         =   "Kode Buku"
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
         DataField       =   "judul"
         Caption         =   "Judul Buku"
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
         DataField       =   "pengarang"
         Caption         =   "Pengarang"
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
         DataField       =   "penerbit"
         Caption         =   "Penerbit"
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
         DataField       =   "sinopsis"
         Caption         =   "Sinopsis"
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
         DataField       =   "lokasi"
         Caption         =   "Lokasi"
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
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoBuku 
      Height          =   492
      Left            =   15000
      Top             =   2280
      Width           =   1932
      _ExtentX        =   3408
      _ExtentY        =   868
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select * from buku"
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
   Begin VB.TextBox txtPenerbit 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   14
      Top             =   2880
      Width           =   4335
   End
   Begin VB.TextBox txtKode 
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox txtJudul 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox txtPengarang 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtSinopsis 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5700
      Left            =   7320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1680
      Width           =   3612
   End
   Begin VB.TextBox txtLokasi 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   3
      Top             =   3480
      Width           =   1212
   End
   Begin VB.TextBox txtKunci 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   6612
   End
   Begin VB.CommandButton cmdCari 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CARI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2172
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000040C0&
      FillStyle       =   4  'Upward Diagonal
      Height          =   732
      Left            =   120
      Top             =   3960
      Width           =   2532
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   732
      Left            =   2640
      Top             =   3960
      Width           =   4452
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SINOPSIS BUKU :"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   444
      Left            =   7320
      TabIndex        =   9
      Top             =   1080
      Width           =   2604
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H0000C000&
      FillStyle       =   5  'Downward Diagonal
      Height          =   6492
      Left            =   7200
      Top             =   960
      Width           =   3852
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   4  'Dash-Dot
      X1              =   7080
      X2              =   7080
      Y1              =   4920
      Y2              =   7440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   4  'Dash-Dot
      X1              =   120
      X2              =   7080
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label7 
      Caption         =   "KODE BUKU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   2292
   End
   Begin VB.Label Label2 
      Caption         =   "JUDUL BUKU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   2292
   End
   Begin VB.Label Label3 
      Caption         =   "PENGARANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   2292
   End
   Begin VB.Label Label4 
      Caption         =   "PENERBIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   2292
   End
   Begin VB.Label Label6 
      Caption         =   "LOKASI BUKU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   2292
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "JUDUL BUKU :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1776
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H00008000&
      FillStyle       =   2  'Horizontal Line
      Height          =   3012
      Left            =   120
      Top             =   960
      Width           =   6972
   End
End
Attribute VB_Name = "frmCari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCari_Click(Index As Integer)
sipp
On Error Resume Next
Dim kunci As String
kunci = txtKunci.Text
If kunci = "" Then
    MsgBox "Judul Buku harus diisi!", vbOKOnly, "Salah!"
    AdoBuku.Refresh
Else
    AdoBuku.RecordSource = "select * from buku where judul like '%" & kunci & "%'"
    AdoBuku.Refresh
    Set DataBuku.DataSource = AdoBuku
    DataBuku.Refresh
    DataBuku_Click
End If
End Sub

Private Sub Command1_Click()
Menu.Show
Unload Me
End Sub

Private Sub Command2_Click()
frmpinjam.Show
End Sub


Private Sub Command4_Click()
   If Not AdoBuku.Recordset.EOF Then
        AdoBuku.Recordset.MoveLast
       Call Form_Load
    End If
End Sub

Private Sub Command5_Click()
   If Not AdoBuku.Recordset.BOF Then
       AdoBuku.Recordset.MoveFirst
       Call Form_Load
    End If
End Sub

Private Sub Command6_Click()
  AdoBuku.Recordset.MovePrevious
 If AdoBuku.Recordset.BOF Then
    AdoBuku.Recordset.MoveNext
 End If
    Call Form_Load
End Sub

Private Sub Command7_Click()
AdoBuku.Recordset.MoveNext
If AdoBuku.Recordset.EOF Then
    AdoBuku.Recordset.MovePrevious
 End If
 Call Form_Load
End Sub

Private Sub DataBuku_Click()
sipp
On Error Resume Next
If AdoBuku.Recordset.BOF Then
    MsgBox "Tidak ada data!", vbOKOnly, "Informasi!"
Else
    txtKode = AdoBuku.Recordset("kode")
    txtJudul = AdoBuku.Recordset("judul")
    txtPengarang = AdoBuku.Recordset("pengarang")
    txtPenerbit = AdoBuku.Recordset("penerbit")
    txtSinopsis = AdoBuku.Recordset("sinopsis")
    txtLokasi = AdoBuku.Recordset("lokasi")

End If
End Sub
Private Sub Form_Load()
 txtKode.Text = AdoBuku.Recordset("kode")
    txtJudul.Text = AdoBuku.Recordset("judul")
    txtPengarang.Text = AdoBuku.Recordset("pengarang")
    txtPenerbit.Text = AdoBuku.Recordset("penerbit")
    txtSinopsis.Text = AdoBuku.Recordset("sinopsis")
    txtLokasi.Text = AdoBuku.Recordset("lokasi")
End Sub
