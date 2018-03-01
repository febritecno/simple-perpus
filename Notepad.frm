VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notepad Vb6"
   ClientHeight    =   2532
   ClientLeft      =   36
   ClientTop       =   660
   ClientWidth     =   5676
   Icon            =   "Notepad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2532
   ScaleWidth      =   5676
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2532
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5652
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Document|*.*|Text Document|*.txt|"
      FilterIndex     =   2
   End
   Begin VB.Menu Cmdfile 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&Baru"
         Checked         =   -1  'True
         Shortcut        =   ^N
      End
      Begin VB.Menu buka 
         Caption         =   "&Buka..."
         Shortcut        =   ^O
      End
      Begin VB.Menu simpan 
         Caption         =   "&Simpan..."
         Shortcut        =   ^S
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu Minimaze 
         Caption         =   "&Minimaze"
         Shortcut        =   ^Q
      End
      Begin VB.Menu print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu Keluar 
         Caption         =   "&Keluar"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub copy_Click()
Clipboard.Clear
   Clipboard.SetText Text1.SelText
End Sub


Private Sub cut_Click()
Clipboard.Clear
Clipboard.SetText Text1.SelText
Text1.SelText = ""
End Sub



Private Sub Keluar_Click()
Me.Hide
End Sub

Private Sub Minimaze_Click()
Form1.WindowState = 1
End Sub

Private Sub new_Click()
Text1.Text = ""
End Sub

Private Sub paste_Click()
Text1.SelText = Clipboard.GetText()
End Sub


Private Sub print_Click()
 On Error GoTo ErrHandler
  Dim BeginPage, EndPage, NumCopies, i
   CommonDialog1.CancelError = True
  CommonDialog1.ShowPrinter
  BeginPage = CommonDialog1.FromPage
  EndPage = CommonDialog1.ToPage
  NumCopies = CommonDialog1.Copies
  For i = 1 To NumCopies
 Printer.Print Text1.Text
  Next i
  Exit Sub
ErrHandler:
   Exit Sub
End Sub

Private Sub simpan_Click()
On Error GoTo ErrorHandler
  CommonDialog1.Filter = "Semua Files |*.*|Text Files |*.txt"
    CommonDialog1.FilterIndex = 2
   CommonDialog1.ShowSave
 CommonDialog1.FileName = CommonDialog1.FileName
Dim iFile As Integer
 Dim SaveFileFromTB As Boolean
 Dim TxtBox As Object
 Dim FilePath As String
Dim Append As Boolean
  iFile = FreeFile
If Append Then
    Open CommonDialog1.FileName For Append As #iFile
Else
    Open CommonDialog1.FileName For Output As #iFile
End If
Print #iFile, Text1.Text
SaveFileFromTB = True
ErrorHandler:
Close #iFile
End Sub

