VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Kalkulator"
   ClientHeight    =   2676
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3144
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2676
   ScaleWidth      =   3144
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAngka 
      Caption         =   ","
      Height          =   375
      Index           =   10
      Left            =   720
      TabIndex        =   17
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "C"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdHitung 
      Caption         =   "="
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdOperator 
      Caption         =   "/"
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Caption         =   "*"
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Caption         =   "-"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdOperator 
      Caption         =   "+"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdAngka 
      Caption         =   "0"
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdAngka 
      Caption         =   "9"
      Height          =   375
      Index           =   8
      Left            =   1320
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdAngka 
      Caption         =   "8"
      Height          =   375
      Index           =   7
      Left            =   720
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdAngka 
      Caption         =   "7"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdAngka 
      Caption         =   "6"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdAngka 
      Caption         =   "5"
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdAngka 
      Caption         =   "4"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdAngka 
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdAngka 
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdAngka 
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim angka(1 To 2) As Single
Dim operator As String

Private Sub cmdAngka_Click(Index As Integer)
    Text1.Text = Text1.Text & cmdAngka(Index).Caption
End Sub

Private Sub cmdHapus_Click()
    Text1.Text = ""
End Sub

Private Sub cmdOperator_Click(Index As Integer)
    If Text1.Text = "" Then Exit Sub
    
    angka(1) = CSng(Text1.Text)
    operator = cmdOperator(Index).Caption
    Text1.Text = ""
End Sub

Private Sub cmdHitung_Click()
    Dim hasil As Single
        
    If Text1.Text = "" Then Exit Sub
    
    angka(2) = CSng(Text1.Text)
    
    Select Case operator
    Case "+"
        hasil = angka(1) + angka(2)
    Case "-"
        hasil = angka(1) - angka(2)
    Case "*"
        hasil = angka(1) * angka(2)
    Case "/"
        hasil = angka(1) / angka(2)
    End Select
    
    Text1.Text = hasil
End Sub
