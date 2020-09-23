VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0000C000&
   Caption         =   "Sweep Rummy"
   ClientHeight    =   10500
   ClientLeft      =   1125
   ClientTop       =   450
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   12525
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   10350
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   29
      Top             =   8550
      Width           =   2025
      Begin VB.ListBox PlayedCombos 
         Height          =   1035
         Left            =   90
         TabIndex        =   30
         Top             =   390
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Played Quads"
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
         Left            =   90
         TabIndex        =   31
         Top             =   60
         Width           =   2205
      End
   End
   Begin VB.CommandButton cmdTurnDone 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "Done!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10140
      Width           =   1785
   End
   Begin VB.CommandButton cmdExtraOptions 
      Caption         =   "Options"
      Height          =   315
      Left            =   240
      TabIndex        =   26
      Top             =   10260
      Width           =   2205
   End
   Begin VB.PictureBox picCardInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   8040
      ScaleHeight     =   1995
      ScaleWidth      =   2235
      TabIndex        =   17
      Top             =   8550
      Width           =   2265
      Begin VB.CommandButton cmdPlay 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "Play!"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1500
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ListBox lstSelectedCards 
         Height          =   840
         Left            =   60
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblErrorReason 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   60
         TabIndex        =   22
         Top             =   1500
         Width           =   2085
      End
      Begin VB.Label lblPlayableScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1290
         TabIndex        =   21
         Top             =   1230
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Playable score: "
         Height          =   225
         Left            =   60
         TabIndex        =   20
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Cards"
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
         Left            =   90
         TabIndex        =   18
         Top             =   60
         Width           =   2205
      End
   End
   Begin SweepRummy.CardHand PlayerHand 
      Height          =   2010
      Left            =   2550
      TabIndex        =   11
      Top             =   8610
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   3545
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   240
      ScaleHeight     =   1575
      ScaleWidth      =   2205
      TabIndex        =   3
      Top             =   8610
      Width           =   2235
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Height          =   555
         Left            =   180
         TabIndex        =   25
         Top             =   780
         Width           =   1905
      End
      Begin VB.Label lblCardCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         Height          =   225
         Left            =   1230
         TabIndex        =   6
         Top             =   510
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cards remaining:"
         Height          =   225
         Left            =   150
         TabIndex        =   5
         Top             =   510
         Width           =   1275
      End
      Begin VB.Label lblPlayerName 
         BackStyle       =   0  'Transparent
         Caption         =   "Player's Name"
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
         Left            =   150
         TabIndex        =   4
         Top             =   150
         Width           =   2205
      End
   End
   Begin SweepRummy.CardBackDispay Player 
      Height          =   6195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   1740
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   10927
   End
   Begin SweepRummy.CardBackDispay Player 
      Height          =   6195
      Index           =   2
      Left            =   14040
      TabIndex        =   1
      Top             =   1740
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   10927
   End
   Begin SweepRummy.CardCombo CardCombo 
      Height          =   1545
      Index           =   0
      Left            =   8970
      TabIndex        =   28
      Top             =   1920
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   2725
   End
   Begin VB.PictureBox picMessageWindow 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   6765
      Left            =   1410
      ScaleHeight     =   6735
      ScaleWidth      =   12495
      TabIndex        =   7
      Top             =   1740
      Width           =   12525
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   525
         Left            =   3090
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Deal More cards"
         Height          =   465
         Left            =   4500
         TabIndex        =   16
         Top             =   810
         Visible         =   0   'False
         Width           =   1245
      End
      Begin SweepRummy.StackControl StackControl 
         Height          =   1935
         Left            =   1680
         TabIndex        =   12
         Top             =   2040
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3413
      End
      Begin VB.PictureBox PicDeck 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1755
         Left            =   210
         ScaleHeight     =   1755
         ScaleWidth      =   1335
         TabIndex        =   10
         Top             =   2220
         Width           =   1335
      End
      Begin VB.CommandButton cmdTestFunction 
         Caption         =   "The BIG Test Button"
         Height          =   615
         Left            =   4410
         TabIndex        =   9
         Top             =   90
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton cmdDeal 
         Caption         =   "Click here to begin"
         Height          =   615
         Left            =   1500
         TabIndex        =   8
         Top             =   4290
         Width           =   3705
      End
   End
   Begin SweepRummy.CardBackDispay Player 
      Height          =   1830
      Index           =   1
      Left            =   4950
      TabIndex        =   2
      Top             =   30
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3228
   End
   Begin VB.Label lblPlayer3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Player 3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13290
      TabIndex        =   15
      Top             =   1230
      Width           =   1875
   End
   Begin VB.Label lblPlayer2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3600
      TabIndex        =   14
      Top             =   540
      Width           =   1245
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1905
   End
   Begin VB.Menu MnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MnuPopupSort 
         Caption         =   "Sort Hand"
      End
      Begin VB.Menu MnuPopupScore 
         Caption         =   "Calculate Score"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Game As clsGame
Attribute Game.VB_VarHelpID = -1
Private HumanTurn As Boolean
Private PickedUp As Boolean
Private Increment As Integer
Private HIncrement As Integer
Private GameOver As Boolean


Private Sub CardCombo_CardClicked(Index As Integer, CardPos As Integer)
    Dim Temp As clsCard, Returned As clsCard
    Dim Counter As Integer
    Dim Tempory() As clsCard
    Dim CurrentCombo As clsCardCombo
    
    'This returns the clicked upon card. Its actual position in the shuffled is used as an index
    Set Temp = Game.GetCardFromPosition(CardPos)
    'Try to swap jokers
    If Temp.CardValue = joker Then
        'Copy the player's hand into an internal array
        ReDim Tempory(PlayerHand.SelectedCount)
        For Counter = 0 To UBound(Tempory)
            Set Tempory(Counter) = PlayerHand.GetSelectedCard(Counter)
        Next
        'Link the internal class to that of the clicked control
        Set CurrentCombo = CardCombo(Index).GetReference
        
        If Tempory(0) Is Nothing Then
            Exit Sub
        Else
            Set Temp = Tempory(0)
            Set Returned = CurrentCombo.SwapJoker(Temp)
            If Returned Is Nothing Then
                MsgBox "Cannot swap jokers"
            Else
                Game.Player(0).RemoveCard Temp
                Set Temp.CardOwner = Game.Player(0)
                Game.Player(0).AddCard Returned
                
                CurrentCombo.SortCards
            End If
             
        End If
    Else
        'This basically adds the selected cards in the player's
        'hand to an array containing the cards in the cardcombo
        'control and then checks to see if it is valid. If so then
        'the cards can and are added.
        ReDim Tempory(PlayerHand.SelectedCount)
        For Counter = 0 To UBound(Tempory)
            Set Tempory(Counter) = PlayerHand.GetSelectedCard(Counter)
        Next
        
        If Tempory(0) Is Nothing Then Exit Sub
        
        Set CurrentCombo = CardCombo(Index).GetReference
        
        If CurrentCombo.ValidateAddCards(Tempory) = True Then
            For Counter = LBound(Tempory) To UBound(Tempory)
                Set Temp = Tempory(Counter)
                CurrentCombo.AddCard Temp
                Game.Player(0).RemoveCard Temp
            Next
            
            CurrentCombo.SortCards
        Else
            MsgBox "No such luck"
        End If
    End If
    'MsgBox ConvertCardToName(Temp)
End Sub

Private Sub CardCombo_DoubleClick(Index As Integer)
    CardCombo(Index).Drag vbBeginDrag
End Sub



Private Sub CardCombo_HideMe(Index As Integer, sTExt As String)
    CardCombo(Index).Visible = False
    'This just makes things easier to see
    PlayedCombos.AddItem sTExt
End Sub

Private Sub CardCombo_MouseMove(Index As Integer, CardPos As Integer)
    On Error Resume Next
    Dim Temp As clsCard
    Set Temp = Game.GetCardFromPosition(CardPos)
    lblPlayerName.Caption = Temp.GetCardName
    lblInfo.Caption = "Owner: " & Temp.CardOwner.PlayerName
End Sub

Private Sub cmdExtraOptions_Click()
    PopupMenu MnuPopup
End Sub

Private Sub cmdPlay_Click()
    Dim Counter As Integer
    Dim Temp() As clsCard
    Dim CurrentCombo As clsCardCombo
    'MsgBox "Check cards against the table!"
    
    If PickedUp = False Then
        MsgBox "You have not picked up a card yet"
        Exit Sub
    End If
    
    ReDim Temp(PlayerHand.SelectedCount)
    For Counter = 0 To UBound(Temp)
        Set Temp(Counter) = PlayerHand.GetSelectedCard(Counter)
    Next
    
    Call Game.CheckCardCombination(Temp)
    
    Set CurrentCombo = Game.PlayCombination(Temp, Game.Player(0), Game.LastCheckCombinationType)
    CurrentCombo.ReferenceGame Game
    AddNewComboControl CurrentCombo
    
    Game.Player(0).ForceHandUpdate
    'CardCombo1.SetRefence CurrentCombo
    
End Sub


Public Sub AddNewComboControl(clsCombo As clsCardCombo)
    Dim Temp As Integer
    'Horrible interface procedure, used to visually
    'reprisent a cardcombo class...
    Temp = CardCombo.Count
    
    Load CardCombo(Temp)
        
    CardCombo(Temp).SetRefence clsCombo
    CardCombo(Temp).Visible = True
    Set CardCombo(Temp).Container = picMessageWindow
    
    
    'If somebody would like to come up with a better
    'solution, be my guest... I didn't have the time
    'nor the patients to come up with anything better
    Randomize
    CardCombo(Temp).Top = Int(((picMessageWindow.Height - CardCombo(Temp).Height) - CardCombo(Temp).Height + 1) * Rnd + CardCombo(Temp).Height)      '1920
    CardCombo(Temp).Left = 8970 + Int((500 - 0 + 1) * Rnd + 1)
    
    
    CardCombo(Temp).ZOrder 1
End Sub

Private Sub cmdTurnDone_Click()
    Dim Counter As Integer
    Dim Temp As clsCard
    'Remind our faithful user to pick up if he hasn't yet
    If PickedUp = False Then
        MsgBox "You have not picked up yet!"
        Exit Sub
    End If
    'Check to see if the player has selected a card to discard
    Set Temp = PlayerHand.GetSelectedCard(0)
    'Discard the player's card if he has chosen one
    If Temp Is Nothing Then
        MsgBox "Please select a card to discard"
    Else
        If MsgBox("Are you sure you want to discard " & Temp.GetCardName & "?", vbYesNo, "Discard?") = vbYes Then
            Game.Player(0).DiscardCard Temp
            HumanTurn = False
            cmdTurnDone.Enabled = False
            'Let the game continue!
            Game.DetermineNextPlayer
        End If
    End If

End Sub

Private Sub Command1_Click()
    Game.DealCards
    DrawDeck
End Sub

Private Sub cmdDeal_Click()
    cmdDeal.Visible = False
    'Deal the cards
    Game.DealCards
    'Draw the physical deck
    DrawDeck
    'Begin the game
    Game.BeginGame
End Sub

Private Sub cmdTestFunction_Click()
' THIS IS RUBBISH WAS USED FOR DEMONSTRATION PURPOSES
'Game.DealCards

'Dim Counter As Integer
'Dim Temp As Variant
'Dim TempCard As clsCard
'Dim Te() As clsCard

'List1.Clear

'MsgBox Game.CheckCardCombination(Temp)

'For Counter = LBound(Temp) To UBound(Temp)
'    Set TempCard = Game.GetCardFromPosition(Temp(Counter).AbsolutePosition)
'    MsgBox ConvertCardToName(TempCard)
'Next

'ReDim Te(PlayerHand.SelectedCount)'
'
'For Counter = 0 To UBound(Te)
'    Set Te(Counter) = PlayerHand.GetSelectedCard(Counter)
'Next'
'
'Counter = Game.CheckCardCombination(Te)'
'
'If Counter = 0 Then
'    MsgBox Counter & " - " & Game.LastCheckCombinationError
'Else
'    MsgBox Counter
'End If

Game.Player(0).SortCards

End Sub



Private Sub Command2_Click()
' THIS IS RUBBISH WAS USED FOR DEMONSTRATION PURPOSES
'   Game.Player(0).RemoveCard PlayerHand.GetSelectedCard(0)
'   Game.DealFromDeck Game.Player(0)
'   Call Game.Player(1).CheckStack
Game.DealCards
End Sub

Private Sub Form_Load()
    Dim Counter As Integer
    Set Game = New clsGame
    Dim PlayerName As String
    'Set the players names
    PlayerName = InputBox("Please enter your name", "Enter your name")
    If PlayerName = "" Then PlayerName = "Human"
    lblPlayerName.Caption = PlayerName
    'Game.SetPlayerNames playerName, "Bob", "Charlie", "Vinnie"
    Game.SetPlayerNames PlayerName, LoadRandomNameFromFile, LoadRandomNameFromFile, LoadRandomNameFromFile
    'Initilize the game engine
    Game.Initilize
    'Change these to make the bots dumber, go on... impress your friends with your card playing skills
    Game.Player(1).AILevel = NotMensaMaterial   'MildlyIntelligentFrothingMoron
    Game.Player(2).AILevel = NotMensaMaterial   'MildlyIntelligentFrothingMoron
    Game.Player(3).AILevel = NotMensaMaterial   'MildlyIntelligentFrothingMoron
    'Set the alignment of the card holding controls
    Player(0).Alignment = 0
    Player(2).Alignment = 0
    'Link the players to the controls
    PlayerHand.SetPlayer Game.Player(0)
    Player(0).SetPlayer Game.Player(1)
    Player(1).SetPlayer Game.Player(2)
    Player(2).SetPlayer Game.Player(3)
    'Set the labels for the Players
    lblPlayer1.Caption = Game.Player(1).PlayerName
    lblPlayer2.Caption = Game.Player(2).PlayerName
    lblPlayer3.Caption = Game.Player(3).PlayerName
    'Set a reference to the game so that the stack control can receive updates
    StackControl.SetGame Game
End Sub

Public Sub DrawDeck()
    Dim X As Long, Y As Long
    Dim XInc As Byte, Yinc As Byte
    Dim Counter As Integer
    PicDeck.Cls
    cdtInit X, Y
    Call cdtDraw(PicDeck.hDC, 0, 0, 54, conBacks, 0)     'Draws card backs
    
    For Counter = 0 To Int(Game.GetNoRemainingCards / 10)
        Call cdtDraw(PicDeck.hDC, XInc, Yinc, 54, conBacks, 0)      'Draws card backs
        XInc = XInc + 2
        Yinc = Yinc + 2
    Next
    PicDeck.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Game = Nothing
    
    'frmTest.Show
    frmMainMenu.Show
End Sub

Private Sub Game_GameOver()
On Error Resume Next
    Dim Scores(3) As Integer
    Dim Counter As Integer
    Dim Winner As clsPlayer
    Dim WinnerScore As Integer
    Dim Temp As String
    'Display the
    frmWin.Show
    
    GameOver = True
    
    For Counter = 0 To 3
        Scores(Counter) = Game.DeterminePlayerScore(Game.Player(Counter))
        If Scores(Counter) > WinnerScore Then
            Set Winner = Game.Player(Counter)
            WinnerScore = Scores(Counter)
            frmWin.lblScore(0).ForeColor = vbRed
            frmWin.lblScore(1).ForeColor = vbRed
            frmWin.lblScore(2).ForeColor = vbRed
            frmWin.lblScore(3).ForeColor = vbRed
            frmWin.lblScore(Counter).ForeColor = vbBlue
        Else
            frmWin.lblScore(Counter).ForeColor = vbRed
            
        End If
            
        frmWin.lblPlayerName(Counter).Caption = Game.Player(Counter).PlayerName
        Temp = "+ " & Game.DeterminePlayerScoreOnTable(Game.Player(Counter)) & Chr(13) & Chr(10)
        Temp = Temp & " - " & Game.DeterminePlayerScoreInHand(Game.Player(Counter)) & Chr(13) & Chr(10)
        Temp = Temp & "-----" & Chr(13) & Chr(10)
        Temp = Temp & Scores(Counter)
        frmWin.lblScore(Counter).Caption = Temp
        frmWin.lblWinner.Caption = "Congradulations to " & Winner.PlayerName
    Next
    
    
    'MsgBox Winner.PlayerName & " is the winner!"
End Sub

Private Sub Game_HumanTurn()
    cmdPlay.Enabled = True
    cmdTurnDone.Enabled = True
    PickedUp = False
    HumanTurn = True
End Sub

Private Sub MnuPopupScore_Click()
    MsgBox "Your score is currently " & Game.DeterminePlayerScore(Game.Player(0))
End Sub

Private Sub MnuPopupSort_Click()
    Game.Player(0).SortCards
End Sub

Private Sub PicDeck_Click()
    If PickedUp = False Then
        Game.DealFromDeck Game.Player(0)
        PickedUp = True
    Else
        MsgBox "You have already picked up this turn"
    End If
End Sub

Private Sub PicDeck_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblPlayerName.Caption = "Deck"
    lblCardCount.Caption = Game.GetNoRemainingCards
End Sub

Private Sub picMessageWindow_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Left = X
    Source.Top = Y
    Source.ZOrder 1
End Sub

Private Sub Player_AIPlayCardCombo(Index As Integer, Cards() As clsCard)
    'msgbox
    Dim Counter As Integer
    Dim Temp() As clsCard
    Dim CurrentCombo As clsCardCombo
    'MsgBox "Check cards against the table!"
    
    'If PickedUp = False Then
    '    MsgBox "You have not picked up a card yet"
    '    Exit Sub
    'End If
    
    'ReDim Temp(PlayerHand.SelectedCount)
    'For Counter = 0 To UBound(Temp)
    '    Set Temp(Counter) = PlayerHand.GetSelectedCard(Counter)
    'Next
    
    Call Game.CheckCardCombination(Cards)
    
    
    Set CurrentCombo = Game.PlayCombination(Cards, Player(Index).ReferencedPlayer, Game.LastCheckCombinationType)
    CurrentCombo.ReferenceGame Game
    AddNewComboControl CurrentCombo
    
    Player(Index).ReferencedPlayer.ForceHandUpdate
End Sub

Private Sub Player_MouseMove(Index As Integer)
    lblPlayerName.Caption = Player(Index).ReferencedPlayer.PlayerName
    lblCardCount.Caption = Player(Index).ReferencedPlayer.GetNumberOfCards
End Sub

Private Sub PlayerHand_CardCleared()
    lstSelectedCards.Clear
End Sub

Private Sub PlayerHand_CardClicked(Card As clsCard, Selected As Boolean)
Dim Counter As Integer
Dim Temp() As clsCard
    If Selected = True Then
        lstSelectedCards.AddItem Card.GetCardName
    Else
        For Counter = 0 To lstSelectedCards.ListCount - 1
            If lstSelectedCards.List(Counter) = Card.GetCardName Then
                lstSelectedCards.RemoveItem Counter
                Exit For
            End If
        Next
    End If
    
    ReDim Temp(PlayerHand.SelectedCount)
    cmdPlay.Visible = False
    For Counter = 0 To UBound(Temp)
        Set Temp(Counter) = PlayerHand.GetSelectedCard(Counter)
    Next
    lblPlayableScore.Caption = ""
    lblErrorReason.Caption = ""
    Counter = Game.CheckCardCombination(Temp)
    If Counter = 0 Then
        lblPlayableScore.Caption = Counter
        lblErrorReason.Caption = Game.LastCheckCombinationError
    Else
        lblPlayableScore.Caption = Counter
        cmdPlay.Visible = True
    End If
End Sub

Private Sub StackControl_SelectedCardDoubleClick(sSelectedCard As clsCard)
    Dim Temp() As clsCard
    Dim Counter As Integer
    
    If PickedUp = False Then
        Temp = Game.GetAllStackRightOfCard(sSelectedCard)
        
        If MsgBox("Are you sure you want to pick up all the cards from the " & sSelectedCard.GetCardName & "?", vbYesNo) = vbYes Then
              For Counter = LBound(Temp) To UBound(Temp)
               'Add it to the player's hand
               Game.Player(0).AddCard Temp(Counter)
               'Remove it from the stack
               Game.StackRemoveCard Temp(Counter)
              Next
              PickedUp = True
        End If
        
     Else
        MsgBox "You have already picked up this turn"
     End If
     'sSelectedCard.GetCardName
End Sub


Public Function LoadRandomNameFromFile() As String
    Dim FilePath As String
    Dim FileNumber As String
    Dim sName() As String
    Dim Buffer As String
    Dim Counter As Integer
    
    Static CallTimes As Integer
    CallTimes = CallTimes + 1
    On Error GoTo ErrorHandler
    FilePath = App.Path & "\names.txt"
    FileNumber = FreeFile
    ReDim sName(0)
    Open FilePath For Input As #FileNumber
        Do Until EOF(FileNumber)
            Line Input #FileNumber, Buffer
            If sName(0) = "" Then
                sName(0) = Buffer
            Else
                ReDim Preserve sName(UBound(sName) + 1)
                sName(UBound(sName)) = Buffer
            End If
        Loop
    Close #FileNumber
    Randomize
    LoadRandomNameFromFile = sName(Int((UBound(sName) - LBound(sName) + 1) * Rnd + LBound(sName)))
Exit Function
ErrorHandler:
    LoadRandomNameFromFile = Choose(CallTimes, "Bob", "Charlie", "Vinnie")
    CallTimes = 0
End Function

