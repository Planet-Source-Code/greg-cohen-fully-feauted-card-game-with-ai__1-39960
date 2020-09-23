VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome to Sweep Rummy!"
   ClientHeight    =   7080
   ClientLeft      =   1650
   ClientTop       =   1800
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   6240
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   5190
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5130
      Top             =   5280
   End
   Begin VB.PictureBox picScreen 
      BackColor       =   &H00FFFFFF&
      Height          =   7005
      Left            =   30
      ScaleHeight     =   6945
      ScaleWidth      =   6105
      TabIndex        =   0
      Top             =   30
      Width           =   6165
      Begin VB.CommandButton Command2 
         Caption         =   "Exit Game"
         Height          =   585
         Left            =   1020
         TabIndex        =   2
         Top             =   6270
         Width           =   3675
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start the Game"
         Height          =   585
         Left            =   1020
         TabIndex        =   1
         Top             =   5580
         Width           =   3675
      End
      Begin VB.Image Image1 
         Height          =   5505
         Left            =   30
         Picture         =   "frmMainMenu.frx":0000
         Top             =   0
         Width           =   6030
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X, Y, Z
Private Sub Command1_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    frmMain.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = False
    Unload Me
End Sub

Private Sub Form_DblClick()
    Unload Me
    frmTest.Show
End Sub

Private Sub Timer1_Timer()
  Static interval As Integer
  Dim H, R
  cdtInit H, R    'MS's call to the cards dll to initialize it. X and Y are variables that cdtInit fills with the card dimentions, I don't use these variables
  Randomize
    Z = Int((53 - 1 + 1) * Rnd + 1)
    'X = X + 1
    'Y = Y + 1
    
    X = Int((370 - 1 + 1) * Rnd + 1)
    Y = Int((370 - 1 + 1) * Rnd + 1)
    
  'picScreen.BackColor    'you must choose a color different from the frmMain.backcolor, try it out and see why, or change the properties, and delete this step
  Call cdtDraw(picScreen.hDC, X, Y, CLng(Z), 0, vbGreen)       'draw the crosshatch pattern on picDestination, using frmMain.BackColor as the fill color
  'picScreen.Refresh
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Timer1.Enabled = True
End Sub
