VERSION 5.00
Begin VB.Form frmRandomShuffle 
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "16"
      Height          =   1095
      Index           =   15
      Left            =   4320
      TabIndex        =   17
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "15"
      Height          =   1095
      Index           =   14
      Left            =   3120
      TabIndex        =   16
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "14"
      Height          =   1095
      Index           =   13
      Left            =   1920
      TabIndex        =   15
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "13"
      Height          =   1095
      Index           =   12
      Left            =   720
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "12"
      Height          =   1095
      Index           =   11
      Left            =   4320
      TabIndex        =   13
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "11"
      Height          =   1095
      Index           =   10
      Left            =   3120
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "10"
      Height          =   1095
      Index           =   9
      Left            =   1920
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "9"
      Height          =   1095
      Index           =   8
      Left            =   720
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "8"
      Height          =   1095
      Index           =   7
      Left            =   4320
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "7"
      Height          =   1095
      Index           =   6
      Left            =   3120
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "6"
      Height          =   1095
      Index           =   5
      Left            =   1920
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "5"
      Height          =   1095
      Index           =   4
      Left            =   720
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "4"
      Height          =   1095
      Index           =   3
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "3"
      Height          =   1095
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "2"
      Height          =   1095
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame FraRandomShuffle 
      Caption         =   "Random Shuffle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdShuffle 
         Caption         =   "Shuffle"
         Height          =   360
         Left            =   2280
         TabIndex        =   2
         Top             =   6000
         Width           =   1470
      End
      Begin VB.CommandButton cmd 
         Caption         =   "1"
         Height          =   1095
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRandomShuffle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blankBtn As Integer
Dim randomNumBtn(1 To 16) As Integer

Private Function shuffle()
    Dim randomNumVisit(1 To 16) As Boolean
    Dim rnum As Integer
    rnum = (Fix(Rnd() * 100) Mod 16) + 1
    For X = 1 To 16
        rnum = (Fix(Rnd() * 100) Mod 16) + 1
        If Not randomNumVisit(rnum) Then
            randomNumBtn(X) = rnum
            randomNumVisit(rnum) = True
        Else
            While randomNumVisit(rnum)
                rnum = (Fix(Rnd() * 100) Mod 16) + 1
            Wend
            randomNumBtn(X) = rnum
            randomNumVisit(rnum) = True
        End If
    Next
    
    For i = 1 To 16
        cmd(i - 1).Caption = randomNumBtn(i)
        If randomNumBtn(i) = 16 Then
            cmd(i - 1).Caption = Empty
            blankBtn = i - 1
        End If
    Next
End Function

Private Sub cmd_Click(Index As Integer)
    If blankBtn <> Index Then
    
        temp = cmd(Index).Caption
        cmd(Index).Caption = cmd(blankBtn).Caption
        cmd(blankBtn).Caption = temp
        
        tempN = randomNumBtn(Index + 1)
        randomNumBtn(Index + 1) = randomNumBtn(blankBtn + 1)
        randomNumBtn(blankBtn + 1) = tempN
        
        blankBtn = Index
        
    End If
    
    Dim sorted As Boolean
    sorted = True
    For i = 1 To 15
        If randomNumBtn(i) > randomNumBtn(i + 1) Then
            sorted = False
        End If
    Next
    If sorted Then
        MsgBox "Puzzle solved"
    End If
End Sub

Private Sub cmdShuffle_Click()
    shuffle
End Sub

Private Sub Form_Load()
    shuffle
End Sub
