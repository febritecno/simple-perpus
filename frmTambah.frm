VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form frmTambah 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tambah Data Buku"
   ClientHeight    =   5736
   ClientLeft      =   2532
   ClientTop       =   396
   ClientWidth     =   7308
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   7308
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      Height          =   192
      Left            =   6600
      TabIndex        =   23
      Top             =   4860
      Width           =   612
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<<"
      Height          =   192
      Left            =   4800
      TabIndex        =   22
      Top             =   4860
      Width           =   612
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<"
      Height          =   192
      Left            =   5640
      TabIndex        =   21
      Top             =   4860
      Width           =   372
   End
   Begin VB.CommandButton Command7 
      Caption         =   ">"
      Height          =   192
      Left            =   6000
      TabIndex        =   20
      Top             =   4860
      Width           =   372
   End
   Begin MSAdodcLib.Adodc AdoBuku 
      Height          =   312
      Left            =   7320
      Top             =   5040
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
      RecordSource    =   "buku"
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
      BackColor       =   &H0080FF80&
      Caption         =   " Lihat V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5160
      Width           =   612
   End
   Begin VB.CommandButton cmdSelesai 
      BackColor       =   &H8000000A&
      Caption         =   "Selesai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5520
      TabIndex        =   18
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdHapus 
      BackColor       =   &H00C0E0FF&
      Caption         =   "HAPUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdCari 
      BackColor       =   &H00FFFF80&
      Caption         =   "CARI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H80000002&
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdSimpan 
      BackColor       =   &H008080FF&
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdTambah 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TAMBAH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataBuku 
      Bindings        =   "frmTambah.frx":0000
      Height          =   2532
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   7092
      _ExtentX        =   12510
      _ExtentY        =   4466
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
      Caption         =   "DATA BUKU PEPUSTAKAAN"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "kode"
         Caption         =   "kode"
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
         Caption         =   "judul"
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
         Caption         =   "pengarang"
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
         Caption         =   "penerbit"
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
         Caption         =   "sinopsis"
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
         Caption         =   "lokasi"
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
         BeginProperty Column05 
            ColumnWidth     =   1955,906
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbPenerbit 
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
      Height          =   396
      Left            =   2760
      TabIndex        =   11
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox txtLokasi 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   420
      Left            =   2760
      MaxLength       =   6
      TabIndex        =   10
      Top             =   4320
      Width           =   1332
   End
   Begin VB.TextBox txtSinopsis 
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
      TabIndex        =   9
      Top             =   2640
      Width           =   4335
   End
   Begin VB.TextBox txtPengarang 
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
      TabIndex        =   8
      Top             =   1440
      Width           =   4332
   End
   Begin VB.TextBox txtJudul 
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
      TabIndex        =   7
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox txtKode 
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
      Height          =   420
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   6
      Top             =   240
      Width           =   1092
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   0
      X2              =   7320
      Y1              =   5760
      Y2              =   5760
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
      Left            =   360
      TabIndex        =   5
      Top             =   4320
      Width           =   2292
   End
   Begin VB.Label Label5 
      Caption         =   "SINOPSIS"
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
      TabIndex        =   4
      Top             =   2640
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
      Left            =   360
      TabIndex        =   3
      Top             =   2040
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
      Left            =   360
      TabIndex        =   2
      Top             =   1440
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
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   2292
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2292
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000040C0&
      FillStyle       =   4  'Upward Diagonal
      Height          =   4692
      Left            =   120
      Top             =   120
      Width           =   7092
   End
End
Attribute VB_Name = "frmTambah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Kosong()
    txtKode = ""
    txtJudul = ""
    txtPengarang = ""
    cmbPenerbit = ""
    txtSinopsis = ""
    txtLokasi = ""
End Sub

Sub mati()
    txtKode.Enabled = False
    txtJudul.Enabled = False
    txtPengarang.Enabled = False
    cmbPenerbit.Enabled = False
    txtSinopsis.Enabled = False
    txtLokasi.Enabled = False
    
    cmdSimpan.Enabled = False
    cmdHapus.Enabled = False
    
    
End Sub
Sub hidup()
    txtJudul.Enabled = True
    txtPengarang.Enabled = True
    cmbPenerbit.Enabled = True
    txtSinopsis.Enabled = True
    txtLokasi.Enabled = True
End Sub

Sub Siap()
    Kosong
    
    txtKode.Enabled = True
    txtJudul.Enabled = True
    txtPengarang.Enabled = True
    cmbPenerbit.Enabled = True
    txtSinopsis.Enabled = True
    txtLokasi.Enabled = True
    txtKode.SetFocus
    
    cmdSimpan.Enabled = True
    cmdUpdate.Enabled = False
    cmdHapus.Enabled = False
End Sub
Private Sub cmdCari_Click()
On Error Resume Next
Dim kunci As String
kunci = InputBox("Masukkan Kode Buku Harus 4 digit, Contoh : 0001", "Cari Data Buku")
If kunci = "" Then
    MsgBox "Data Tidak Ada", vbExclamation, "Mencari Data"
    AdoBuku.Refresh
Else
sipp
    AdoBuku.Recordset.Find "kode='" + kunci + "'", , adSearchForward, 1
If Not AdoBuku.Recordset.EOF Then
txtKode.Text = AdoBuku.Recordset!kode
txtJudul.Text = AdoBuku.Recordset!judul
txtPengarang.Text = AdoBuku.Recordset!pengarang
cmbPenerbit.Text = AdoBuku.Recordset!penerbit
txtSinopsis.Text = AdoBuku.Recordset!sinopsis
txtLokasi.Text = AdoBuku.Recordset!lokasi
Else
MsgBox "gagal"
End If
End If
End Sub

Private Sub cmdHapus_Click()
Dim hapus As String
If txtKode = "" Then
    MsgBox "Cari dulu datanya, baru dihapus!", vbOKOnly, "Salah!"
    cmdCari_Click
Else
sipp
    If AdoBuku.Recordset.RecordCount <> 0 Then
        hapus = MsgBox("Yakin akan dihapus?", vbYesNo, "Peringatan...!")
        If hapus = vbYes Then
            AdoBuku.Recordset.Delete
            AdoBuku.Recordset.MoveNext
            Kosong
        End If
    Else
        MsgBox "Data kosong...", vbInformation, "Informasi!"
    End If
End If
End Sub

Private Sub cmdSelesai_Click()
Menu.Show
Unload Me
End Sub

Private Sub cmdSimpan_Click()
    On Error Resume Next
    If txtKode = "" Or txtJudul = "" Or txtPengarang = "" Or cmbPenerbit = "" Or txtSinopsis = "" Or txtLokasi = "" Then
        MsgBox "Masih ada data yang kosong..!!!", , "Error"
    Else
        sipp
        AdoBuku.Recordset.AddNew
        AdoBuku.Recordset.Fields("kode") = txtKode
        AdoBuku.Recordset.Fields("judul") = txtJudul
        AdoBuku.Recordset.Fields("pengarang") = txtPengarang
        AdoBuku.Recordset.Fields("penerbit") = cmbPenerbit
        AdoBuku.Recordset.Fields("sinopsis") = txtSinopsis
        AdoBuku.Recordset.Fields("lokasi") = txtLokasi
        AdoBuku.Recordset.Update
        MsgBox "Data Buku Telah Disimpan!", vbOKOnly, "Berhasil!"
        Kosong
    End If
End Sub

Private Sub cmdTambah_Click(Index As Integer)
    Siap
    Kosong
End Sub

Private Sub cmdUpdate_Click()
If cmdUpdate.Caption = "EDIT" Then
hidup
cmdUpdate.Caption = "SIMPAN"
txtJudul.SetFocus
        cmdSimpan.Enabled = False
    cmdUpdate.Enabled = True
    cmdHapus.Enabled = True
Else
If txtJudul = "" Or txtPengarang = "" Or cmbPenerbit = "" Or txtSinopsis = "" Or txtLokasi = "" Then
    MsgBox "Masih ada data yang kosong..!!!", vbCritical, "Error!"
        Else
sipp
With AdoBuku.Recordset
    !judul = txtJudul
    !pengarang = txtPengarang
    !penerbit = cmbPenerbit
    !sinopsis = txtSinopsis
    !lokasi = txtLokasi
    .Update
End With
cmdUpdate.Caption = "EDIT"
    End If
End If
End Sub


Private Sub Command2_Click()
If Command2.Caption = " Lihat V" Then
frmTambah.Height = 9096
Command2.Caption = "Tutup"
Else
frmTambah.Height = 6048
Command2.Caption = " Lihat V"
End If
End Sub

Private Sub Command4_Click()
  If Not AdoBuku.Recordset.EOF Then
        AdoBuku.Recordset.MoveLast
        cmdUpdate.Caption = "EDIT"
        mati
       Call muncul
    End If
End Sub

Private Sub Command5_Click()
   If Not AdoBuku.Recordset.BOF Then
       AdoBuku.Recordset.MoveFirst
       cmdUpdate.Caption = "EDIT"
       mati
       Call muncul
    End If
End Sub

Private Sub Command6_Click()
 AdoBuku.Recordset.MovePrevious
     cmdUpdate.Caption = "EDIT"
 mati
 If AdoBuku.Recordset.BOF Then
    AdoBuku.Recordset.MoveNext
 End If
    Call muncul
End Sub

Private Sub Command7_Click()
AdoBuku.Recordset.MoveNext
 mati
cmdUpdate.Caption = "EDIT"
If AdoBuku.Recordset.EOF Then
    AdoBuku.Recordset.MovePrevious
 End If
 Call muncul
End Sub

Private Sub DataBuku_Click()
On Error Resume Next
If AdoBuku.Recordset.BOF Then
    MsgBox "Tidak ada data!", vbOKOnly, "Informasi!"
Else
sipp
    Call Siap
    txtKode.Enabled = False
    txtKode = AdoBuku.Recordset("kode")
    txtJudul = AdoBuku.Recordset("judul")
    txtPengarang = AdoBuku.Recordset("pengarang")
    cmbPenerbit = AdoBuku.Recordset("penerbit")
    txtSinopsis = AdoBuku.Recordset("sinopsis")
    txtLokasi = AdoBuku.Recordset("lokasi")
    
    cmdSimpan.Enabled = False
    cmdUpdate.Enabled = True
    cmdHapus.Enabled = True
End If
txtKode.Enabled = False
    txtJudul.Enabled = False
    txtPengarang.Enabled = False
    cmbPenerbit.Enabled = False
    txtSinopsis.Enabled = False
    txtLokasi.Enabled = False
    
    cmdSimpan.Enabled = False
End Sub

Private Sub Form_Load()
txtKode.Text = AdoBuku.Recordset("kode")
    txtJudul.Text = AdoBuku.Recordset("judul")
    txtPengarang.Text = AdoBuku.Recordset("pengarang")
    cmbPenerbit.Text = AdoBuku.Recordset("penerbit")
    txtSinopsis.Text = AdoBuku.Recordset("sinopsis")
    txtLokasi.Text = AdoBuku.Recordset("lokasi")
    cmbPenerbit.AddItem "Andi Publisher"
    cmbPenerbit.AddItem "Elexmedia Computindo"
    cmbPenerbit.AddItem "Media Kita"
    cmbPenerbit.AddItem "Maxikom"
    cmbPenerbit.AddItem "Gava Media"
    cmbPenerbit.AddItem "Erlangga"
    cmbPenerbit.AddItem "Modula"
    cmbPenerbit.AddItem "Mediakom"
    cmbPenerbit.AddItem "Informatika"
    mati
End Sub


Sub muncul()
 txtKode.Text = AdoBuku.Recordset("kode")
    txtJudul.Text = AdoBuku.Recordset("judul")
    txtPengarang.Text = AdoBuku.Recordset("pengarang")
    cmbPenerbit.Text = AdoBuku.Recordset("penerbit")
    txtSinopsis.Text = AdoBuku.Recordset("sinopsis")
    txtLokasi.Text = AdoBuku.Recordset("lokasi")
    txtKode.Enabled = False
        cmdSimpan.Enabled = False
    cmdUpdate.Enabled = True
    cmdHapus.Enabled = True
End Sub
