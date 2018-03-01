VERSION 5.00
Begin VB.Form Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Perpustakaan"
   ClientHeight    =   3972
   ClientLeft      =   96
   ClientTop       =   480
   ClientWidth     =   7440
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Menu.frx":1601A
   ScaleHeight     =   3972
   ScaleMode       =   0  'User
   ScaleWidth      =   7440
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   -120
      Top             =   3960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "READY"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   3720
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   6480
      TabIndex        =   1
      Top             =   3720
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Untuk Keluar Program Gunakan Fitur Exit DiMenu File Ya.!"
      BeginProperty Font 
         Name            =   "Letter Gothic"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   960
      TabIndex        =   0
      Top             =   3720
      Width           =   5532
   End
   Begin VB.Image Image1 
      Height          =   3720
      Left            =   0
      Picture         =   "Menu.frx":4A560
      Top             =   0
      Width           =   7440
   End
   Begin VB.Menu mApp 
      Caption         =   "File"
      Begin VB.Menu smMasuk 
         Caption         =   "Masuk Akun"
         Checked         =   -1  'True
      End
      Begin VB.Menu smKeluar 
         Caption         =   "Keluar Akun"
         Checked         =   -1  'True
      End
      Begin VB.Menu space0 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mMaster 
      Caption         =   "Operator"
      Begin VB.Menu smTambah 
         Caption         =   "Tambah Buku"
         Checked         =   -1  'True
      End
      Begin VB.Menu petugas 
         Caption         =   "Tambah Petugas"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu Data1 
         Caption         =   "Data Pengunjung"
         Checked         =   -1  'True
      End
      Begin VB.Menu pinjam 
         Caption         =   "Data Peminjaman Buku"
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu Lp_dp 
         Caption         =   "Laporan Data Pengunjung"
         Checked         =   -1  'True
      End
      Begin VB.Menu Lp_dpb 
         Caption         =   "Laporan Data Peminjaman Buku"
      End
      Begin VB.Menu db_buku 
         Caption         =   "Laporan Data Buku"
      End
      Begin VB.Menu space4 
         Caption         =   "-"
      End
      Begin VB.Menu bd 
         Caption         =   "Backup Data"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mPengunjung 
      Caption         =   "Pengunjung"
      Begin VB.Menu smCari 
         Caption         =   "Cari Data Buku"
         Checked         =   -1  'True
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu Data 
         Caption         =   "Pinjam Buku"
      End
   End
   Begin VB.Menu ah 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu alat 
      Caption         =   "Alat"
      Begin VB.Menu kal 
         Caption         =   "Kalkulator"
      End
      Begin VB.Menu cct 
         Caption         =   "Catatan"
      End
   End
   Begin VB.Menu ttn 
      Caption         =   "Tentang"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bd_Click()
Form2.Show
End Sub

Private Sub cct_Click()
Form1.Show
End Sub

Private Sub data_Click()
frmpinjam.Show
End Sub

Private Sub Data1_Click()
frmdata1.Show
End Sub

Private Sub db_buku_Click()
sipp
RS.Open "select * from buku", Conn
If Not RS.EOF Then
Set DataReport3.DataSource = RS
DataReport3.Show
Else
MsgBox "Tidak ada Data"
End If
End Sub

Private Sub exit_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
    mMaster.Visible = False
    smKeluar.Visible = False
End Sub

Private Sub mAbout_Click()
    frmAbout.Show
End Sub


Private Sub kal_Click()
Form9.Show
End Sub

Private Sub Lp_dp_Click()
sipp
RS.Open "select * from pengunjung", Conn
If Not RS.EOF Then
Set DataReport1.DataSource = RS
DataReport1.Show
Else
MsgBox "Tidak ada Data"
End If
End Sub

Private Sub Lp_dpb_Click()
sipp
RS.Open "select * from pinjam", Conn
If Not RS.EOF Then
Set DataReport2.DataSource = RS
DataReport2.Show
Else
MsgBox "Tidak ada Data"
End If
End Sub

Private Sub petugas_Click()
frmlogin2.Show
End Sub

Private Sub pinjam_Click()
frmdata.Show
End Sub

Private Sub smCari_Click()
    frmPengunjung.Show
End Sub
Private Sub smKeluar_Click()
    mMaster.Visible = False
    smMasuk.Visible = True
    smKeluar.Visible = False
End Sub

Private Sub smMasuk_Click()
    frmLogin.Show
End Sub

Private Sub smTambah_Click()
    frmTambah.Show
End Sub

Private Sub Timer1_Timer()
Label2.Caption = DateTime.Time
End Sub



Private Sub ttn_Click()
frmAbout.Show
End Sub
