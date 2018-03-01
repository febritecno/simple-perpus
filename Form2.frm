VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Backup Data Perpustakaan"
   ClientHeight    =   4656
   ClientLeft      =   3228
   ClientTop       =   2388
   ClientWidth     =   6456
   LinkTopic       =   "Form2"
   ScaleHeight     =   4656
   ScaleWidth      =   6456
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
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   5175
      Begin VB.DriveListBox Drive1 
         Height          =   288
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   4575
      End
      Begin VB.DirListBox Dir1 
         Height          =   288
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   4575
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   960
      TabIndex        =   2
      Top             =   2520
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
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
      Height          =   372
      Left            =   4560
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Height          =   372
      Left            =   840
      TabIndex        =   7
      Top             =   3960
      Width           =   3492
   End
   Begin VB.Label Label2 
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
      Left            =   2160
      TabIndex        =   6
      Top             =   0
      Width           =   2412
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H000040C0&
      FillStyle       =   6  'Cross
      Height          =   3972
      Left            =   240
      Top             =   600
      Width           =   6012
   End
End
Attribute VB_Name = "Form2"
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
 
Public Sub copy(ByVal asal As String, ByVal tujuan As String)
Dim x As shfileopstruct
    With x
  .hwnd = 0
        .wfunc = FO_copy
        .pfrom = asal & vbNullChar & vbNullChar
        .pto = tujuan & vbNullChar & vbNullChar
        .Fflags = fof_allowundo
            End With
    SHFileOperation x
End Sub


Private Sub Command1_Click()
On Error Resume Next
If Label1.Caption = "" Then
    MsgBox "Anda belum memilih file yang akan dicopy"
    Exit Sub
ElseIf Text1 = "" Then
    MsgBox "Anda tidak memilih direktori tujuan peng-Copy-an"
    Exit Sub
End If
copy Label1.Caption, Text1.Text
MsgBox "Berhasil di Backup"
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Command2_Click()
Menu.Show
Me.Hide
End Sub

Private Sub Dir1_Change()
Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Label1.Caption = App.Path & "\data.mdb"
Dir1.Path = "C:\"
End Sub




