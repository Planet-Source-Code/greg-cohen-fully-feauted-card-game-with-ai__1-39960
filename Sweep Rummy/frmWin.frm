VERSION 5.00
Begin VB.Form frmWin 
   Caption         =   "The Game is over... and the scores are in!"
   ClientHeight    =   5415
   ClientLeft      =   2055
   ClientTop       =   2715
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5760
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Thank You!"
      Height          =   495
      Left            =   780
      TabIndex        =   8
      Top             =   4830
      Width           =   4335
   End
   Begin VB.Label lblWinner 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Congradulations to "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   780
      TabIndex        =   9
      Top             =   4440
      Width           =   4395
   End
   Begin VB.Line Line7 
      X1              =   240
      X2              =   5610
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      Height          =   3795
      Index           =   3
      Left            =   4380
      TabIndex        =   7
      Top             =   510
      Width           =   1155
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      Height          =   3795
      Index           =   2
      Left            =   3060
      TabIndex        =   6
      Top             =   510
      Width           =   1155
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      Height          =   3795
      Index           =   1
      Left            =   1710
      TabIndex        =   5
      Top             =   510
      Width           =   1155
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      Height          =   3795
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   540
      Width           =   1155
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   240
      Y1              =   450
      Y2              =   4380
   End
   Begin VB.Line Line5 
      X1              =   5610
      X2              =   5610
      Y1              =   450
      Y2              =   4380
   End
   Begin VB.Line Line4 
      X1              =   4290
      X2              =   4290
      Y1              =   450
      Y2              =   4380
   End
   Begin VB.Line Line3 
      X1              =   2970
      X2              =   2970
      Y1              =   450
      Y2              =   4380
   End
   Begin VB.Line Line2 
      X1              =   1590
      X2              =   1590
      Y1              =   450
      Y2              =   4380
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5610
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label lblPlayername 
      Alignment       =   2  'Center
      Caption         =   "Player (0)"
      Height          =   225
      Index           =   3
      Left            =   4350
      TabIndex        =   3
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblPlayername 
      Alignment       =   2  'Center
      Caption         =   "Player (0)"
      Height          =   225
      Index           =   2
      Left            =   3030
      TabIndex        =   2
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblPlayername 
      Alignment       =   2  'Center
      Caption         =   "Player (0)"
      Height          =   225
      Index           =   1
      Left            =   1710
      TabIndex        =   1
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label lblPlayername 
      Alignment       =   2  'Center
      Caption         =   "Player (0)"
      Height          =   225
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   180
      Width           =   1185
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnd_Click()
    Unload Me
    Unload frmMain
End Sub
