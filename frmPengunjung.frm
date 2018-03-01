VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form frmPengunjung 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pengunjung"
   ClientHeight    =   5268
   ClientLeft      =   48
   ClientTop       =   1800
   ClientWidth     =   7488
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5268
   ScaleWidth      =   7488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   5280
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPengunjung.frx":0000
      Height          =   492
      Left            =   3000
      TabIndex        =   10
      Top             =   6360
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   868
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
      EndProperty
   End
   Begin VB.TextBox txtNama 
      Appearance      =   0  'Flat
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
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txtAlamat 
      Appearance      =   0  'Flat
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
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox txtTanggal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2160
      Width           =   1932
   End
   Begin VB.TextBox txtPesan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CommandButton cmdTambah 
      BackColor       =   &H0000FF00&
      Caption         =   "LANJUTKAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1692
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   2040
      Top             =   6240
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
      RecordSource    =   "pengunjung"
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan Data Diri Anda "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   22.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   732
      Left            =   600
      TabIndex        =   11
      Top             =   0
      Width           =   6252
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "ANDA HARUS MENGISI DATA PENGUNJUNG UNTUK MELANJUTKAN DALAM PENCARIAN DATA BUKU"
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
      Height          =   1692
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "NAMA"
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
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   2292
   End
   Begin VB.Label Label2 
      Caption         =   "ALAMAT"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   2292
   End
   Begin VB.Label Label3 
      Caption         =   "TANGGAL"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   2292
   End
   Begin VB.Label Label5 
      Caption         =   "PESAN"
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
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   2292
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   4332
      Left            =   120
      Top             =   840
      Width           =   7212
   End
End
Attribute VB_Name = "frmPengunjung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTambah_Click(Index As Integer)
    On Error Resume Next
    If txtNama = "" Or txtAlamat = "" Or txtPesan = "" Then
        MsgBox "Woi Ada Yang Belum Di Isi..!!!", vbOKOnly + vbCritical, "Error"
        Else
    sipp
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields("nama") = txtNama
        Adodc1.Recordset.Fields("alamat") = txtAlamat
        Adodc1.Recordset.Fields("tanggal") = txtTanggal
        Adodc1.Recordset.Fields("pesan") = txtPesan
        Adodc1.Recordset.Update
        MsgBox "Data Anda Telah Tersimpan!!, Silahkan Mencari Data Buku yang anda Inginkan", vbOKOnly + vbInformation, "Berhasil!"
                        txtNama.Text = ""
txtAlamat.Text = ""
txtTanggal.Text = ""
txtPesan.Text = ""
        frmCari.Show
        Unload Me
    End If
End Sub

Private Sub Timer1_Timer()
txtTanggal.Text = DateTime.Date
End Sub
