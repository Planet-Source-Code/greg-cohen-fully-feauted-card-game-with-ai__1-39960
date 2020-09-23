VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test Form"
   ClientHeight    =   8250
   ClientLeft      =   1335
   ClientTop       =   1395
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   11655
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   435
      Left            =   9300
      TabIndex        =   28
      Top             =   5880
      Width           =   2145
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Click here to load the graphical Interface"
      Height          =   4305
      Left            =   8970
      TabIndex        =   24
      Top             =   240
      Width           =   2655
   End
   Begin VB.Frame Frame5 
      Caption         =   "Program Report"
      Height          =   1575
      Left            =   120
      TabIndex        =   22
      Top             =   6360
      Width           =   8775
      Begin VB.TextBox Text1 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   240
         Width           =   8535
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Game"
      Height          =   3255
      Left            =   6240
      TabIndex        =   11
      Top             =   240
      Width           =   2655
      Begin VB.CommandButton Command9 
         Caption         =   "Check Card combination"
         Height          =   435
         Left            =   780
         TabIndex        =   26
         Top             =   810
         Width           =   1635
      End
      Begin VB.ListBox List2 
         Height          =   1815
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Deal!"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Pile"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Players"
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   8775
      Begin VB.ListBox List1 
         Height          =   2010
         Index           =   3
         Left            =   6600
         TabIndex        =   15
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Index           =   2
         Left            =   4440
         TabIndex        =   14
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Index           =   1
         Left            =   2280
         TabIndex        =   13
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Player 4"
         Height          =   255
         Index           =   3
         Left            =   6600
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Player 3"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Player 2"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Player 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Deck Functions"
      Height          =   3255
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   3135
      Begin VB.CommandButton Command5 
         Caption         =   "Deal Next Card"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Shuffle Deck"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   2655
      End
      Begin VB.ListBox lvDeck 
         Height          =   1230
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Initilize Deck"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Card Drawing Functions"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.CommandButton Command10 
         Caption         =   "Draw coloured back"
         Height          =   465
         Left            =   1500
         TabIndex        =   27
         Top             =   1050
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Draw Card Back"
         Height          =   465
         Left            =   1500
         TabIndex        =   25
         Top             =   450
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Draw Card Object"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Draw a Card"
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   1920
         Width           =   1815
      End
      Begin VB.PictureBox PicTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   240
         ScaleHeight     =   1455
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim A As clsDeck
Dim G As clsGame

Private Sub Command1_Click()
  Dim X, Y, Z
  cdtInit X, Y    'MS's call to the cards dll to initialize it. X and Y are variables that cdtInit fills with the card dimentions, I don't use these variables
    Z = InputBox("Please enter a number between 1 and 52")
    
  PicTemp.BackColor = vbWhite   'you must choose a color different from the frmMain.backcolor, try it out and see why, or change the properties, and delete this step
  MsgBox cdtDraw(PicTemp.hDC, 0, 0, CLng(Z), 0, vbGreen)      'draw the crosshatch pattern on picDestination, using frmMain.BackColor as the fill color
End Sub

Private Sub Command10_Click()
  Dim X, Y, Z
  cdtInit X, Y    'MS's call to the cards dll to initialize it. X and Y are variables that cdtInit fills with the card dimentions, I don't use these variables
    Z = InputBox("Please enter a number between 1 and 52")
    
  PicTemp.BackColor = vbBlue   'you must choose a color different from the frmMain.backcolor, try it out and see why, or change the properties, and delete this step
  MsgBox cdtDraw(PicTemp.hDC, 0, 0, CLng(Z), 1, vbYellow)          'draw the crosshatch pattern on picDestination, using frmMain.BackColor as the fill color
End Sub

Private Sub Command11_Click()
    Dim H As New clsCardCombo
    Dim c1(2) As clsCard
    Dim c2 As clsCard
    Dim c3 As clsCard
    Set c1(0) = New clsCard
    Set c1(1) = New clsCard
    Set c1(2) = New clsCard
    Set c2 = New clsCard
    Set c3 = New clsCard
    c1(0).CardValue = Eight
    c1(0).Suit = Hearts
    c1(1).CardValue = Nine
    c1(1).Suit = Hearts
    c1(2).CardValue = Ten
    c1(2).Suit = Hearts
    c2.CardValue = joker
    c2.Suit = Hearts
    c3.CardValue = five
    c3.Suit = Hearts
    Call Command6_Click
    H.ReferenceGame G
    H.FirstCards c1, RunOfthree
    If H.ValidateAdd(c2) = True Then
        H.AddCard c2
        MsgBox "Added!"
    Else
        MsgBox "Failed"
    End If
    
    MsgBox H.SwapJoker(c3)
    
    Set H = Nothing
End Sub

Private Sub Command12_Click()
Dim Temp() As clsCard
Dim Counter As Integer
Temp = G.Player(2).GetHighestPlay(G.Player(2).GetHand)

For Counter = LBound(Temp) To UBound(Temp)
    MsgBox Temp(Counter).GetCardName
Next
End Sub

Private Sub Command2_Click()
    Dim A As clsCard
    
    Set A = New clsCard
        A.CardDrawValue = 5
        A.DrawCard PicTemp.hDC
    Set A = Nothing
        PicTemp.Refresh
End Sub

Private Sub Command3_Click()
    
    Dim B() As clsCard
    Dim C As Integer
    
    Set A = New clsDeck
    
        A.InitilizeDeck
        B = A.GetDeck
   
        lvDeck.Clear
    For C = LBound(B) To UBound(B)
        'lvDeck.AddItem B(C).Suit
        lvDeck.AddItem ConvertCardToName(B(C)) & " - " & B(C).AbsolutePosition & " - " & B(C).CardDrawValue
    Next
End Sub

Private Sub Command4_Click()
    Dim B() As clsCard
    Dim C As Integer
    
            lvDeck.Clear
        A.Shuffle
        B = A.GetDeck
    For C = LBound(B) To UBound(B)
        'lvDeck.AddItem B(C).Suit
        lvDeck.AddItem ConvertCardToName(B(C))
    Next
End Sub

Private Sub Command5_Click()
    MsgBox ConvertCardToName(A.DealNextCard)
End Sub

Private Sub Command6_Click()
    Dim R As Integer
    Dim t As Integer
    Dim f() As clsCard
    Set G = New clsGame
    
    G.Initilize
    G.SetPlayerNames "Robby Ree", "Ben Crupsy", "Anna Thetic", "Emma Roids"
    G.DealCards
    
    For R = 0 To 3
        f = G.Player(R).GetHand
        List1(R).Clear
        For t = LBound(f) To UBound(f)
            List1(R).AddItem ConvertCardToName(f(t))
            'List1(r).ItemData( List1(r).ListCount) = f(t).AbsolutePosition
            List1(R).ItemData(List1(R).ListCount - 1) = f(t).AbsolutePosition
            
            
            Label1(R).Caption = G.Player(R).PlayerName
        Next
    Next
    
    f = G.GetStack
    List2.Clear
    For t = LBound(f) To UBound(f)
        List2.AddItem ConvertCardToName(f(t))
        List2.ItemData(List2.ListCount - 1) = f(t).AbsolutePosition
    Next
    
End Sub

Private Sub Command7_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Command8_Click()
 Dim X, Y, Z
  cdtInit X, Y
    Call cdtDraw(PicTemp.hDC, 0, 0, conPlaid, conBacks, 0)    'Draws card backs
    
    PicTemp.Refresh
End Sub

Private Sub Command9_Click()
    Dim c1(2) As clsCard
    
    Set c1(0) = New clsCard
    Set c1(1) = New clsCard
    Set c1(2) = New clsCard
    c1(0).CardValue = joker
    c1(1).CardValue = joker
    c1(2).CardValue = joker
    
    MsgBox G.CheckCardCombination(c1)
End Sub

Private Sub Form_Load()
    Dim Aa As Byte
    Dim B As String
    
    Aa = FreeFile
    Open App.Path & "\notes.txt" For Input As #Aa
        Do Until EOF(Aa)
            Line Input #Aa, B
            If Text1.Text = "" Then
                Text1.Text = B
            Else
            Text1.Text = Text1.Text & Chr(13) & Chr(10) & B
            End If
        Loop
    Close #Aa
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Aa As Byte
    Dim B As String
    
    Aa = FreeFile
    Open App.Path & "\notes.txt" For Output As #Aa
        Print #Aa, Text1.Text
    Close #Aa
    
    Set G = Nothing
    
 Set A = Nothing
End Sub

Private Sub List1_Click(Index As Integer)
Dim B As clsCard

'MsgBox List1(Index).ItemData(List1(Index).ListIndex)
Set B = G.GetCardFromPosition(List1(Index).ItemData(List1(Index).ListIndex))
MsgBox B.CardOwner.PlayerName
B.DrawCard PicTemp.hDC

PicTemp.Refresh
End Sub

Private Sub List2_Click()
Dim B As clsCard
Set B = G.GetCardFromPosition(List2.ItemData(List2.ListIndex))

B.DrawCard PicTemp.hDC

PicTemp.Refresh
End Sub

Private Sub PicTemp_Click()
    'PicTemp.AutoRedraw
    SavePicture PicTemp.Picture, "c:\temp.bmp"
End Sub

Private Sub Text1_DblClick()
Text1.SelText = Now
End Sub
