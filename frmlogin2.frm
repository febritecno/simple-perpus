VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.OCX"
Begin VB.Form frmlogin2 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pengisian Petugas Baru"
   ClientHeight    =   2724
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   4320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2724
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4440
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3120
      TabIndex        =   9
      Top             =   2280
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "TAMBAH"
      Height          =   372
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1692
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
      Caption         =   "Ubah"
      Height          =   252
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Masuk"
      Height          =   252
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "l"
      Height          =   192
      Left            =   0
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   84
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Hapus"
      Height          =   252
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1680
      Width           =   2892
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2892
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   1
      Top             =   720
      Width           =   2892
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   6120
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   550
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
      RecordSource    =   "login"
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
      Bindings        =   "frmlogin2.frx":0000
      Height          =   1956
      Left            =   4440
      TabIndex        =   0
      Top             =   1200
      Width           =   4092
      _ExtentX        =   7218
      _ExtentY        =   3450
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "user_id"
         Caption         =   "User ID"
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
         DataField       =   "password"
         Caption         =   "Password"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "nama_petugas"
         Caption         =   "Nama Petugas"
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
            Object.Visible         =   -1  'True
            ColumnWidth     =   1944
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1944
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "SELAMAT DATANG ADMIN UTAMA"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   4656
      TabIndex        =   15
      Top             =   240
      Width           =   3660
   End
   Begin VB.Line Line1 
      BorderStyle     =   4  'Dash-Dot
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   3360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAMBAH PETUGAS PERPUSTAKAAN"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3852
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Petugas :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1092
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   372
      Left            =   120
      Top             =   120
      Width           =   4092
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      BorderWidth     =   2
      FillStyle       =   4  'Upward Diagonal
      Height          =   492
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   120
      Width           =   4092
   End
End
Attribute VB_Name = "frmlogin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Red, green, blue As Integer

Private Sub Command1_Click()
If Command1.Caption = "TAMBAH" Then
Command1.BackColor = vbRed
Command1.Caption = "SIMPAN"
Text1.SetFocus
Else
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox "Belum diisi woii !!, Isi dulu ya..??", vbCritical, "Error"
Command1.BackColor = vbGreen
Command1.Caption = "TAMBAH"
Else
sipp
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields("user_id") = Text1.Text
        Adodc1.Recordset.Fields("password") = Text2.Text
        Adodc1.Recordset.Fields("nama_petugas") = Text3.Text
        Adodc1.Recordset.Update
        MsgBox "Data Anda Telah Tersimpan!!", vbOKOnly + vbInformation, "Berhasil!"
        Command1.BackColor = vbGreen
        Command1.Caption = "TAMBAH"
End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Delete
DataGrid1.Refresh
End Sub
Private Sub Command4_Click()
If Command4.Caption = "l" Then
Text4.Visible = True
Command5.Visible = True
Command1.Top = 2640
Command2.Top = 2640
Command4.Top = 3000
Me.Height = 3432
Text4.SetFocus
Command4.Caption = "i"
Else
Text4.Visible = False
Command5.Visible = False
Command1.Top = 2280
Command2.Top = 2280
Command4.Top = 2640
Me.Height = 3060
Command4.Caption = "l"
End If
End Sub

Private Sub Command5_Click()
If Text4.Text = "kerangkang" And Command5.Caption = "Masuk" Then
Me.Width = 8736
Text4.Enabled = False
Command5.Caption = "Tutup"
Else
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.Visible = False
Command5.Visible = False
Command1.Top = 2280
Command2.Top = 2280
Command4.Top = 2640
Me.Width = 4392
Me.Height = 3060
Command5.Caption = "Masuk"
Text4.Enabled = True
End If
End Sub

Private Sub Command6_Click()
If Command6.Caption = "Ubah" Then
Command6.BackColor = vbRed
Command6.Caption = "SIMPAN"
Text1.SetFocus
Else
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
        MsgBox "Woi Gak Ada Isinya...!!!", vbOKOnly + vbCritical, "Error"
        Else
sipp
Adodc1.Recordset!user_id = Text1
Adodc1.Recordset!Password = Text2
Adodc1.Recordset!nama_petugas = Text3
Adodc1.Recordset.Update
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.Text = ""
Command6.BackColor = vbGreen
Command6.Caption = "Ubah"
End If
End If
End Sub

Private Sub DataGrid1_Click()
sipp
On Error Resume Next
If Adodc1.Recordset.BOF Then
    MsgBox "Tidak ada data!", vbOKOnly, "Informasi!"
Else
    Text1 = Adodc1.Recordset("user_id")
    Text2 = Adodc1.Recordset("password")
    Text3 = Adodc1.Recordset("nama_petugas")
End If
End Sub
Private Sub Timer1_Timer()
If blue <= 255 Then blue = blue + 50 Else blue = 0
green = green + 50
If green >= 255 Then green = 0
Red = Red + 50
If Red >= 255 Then
Red = 0
End If
Label5.ForeColor = Int(RGB(Red, green, blue))
Label5.Refresh
End Sub
