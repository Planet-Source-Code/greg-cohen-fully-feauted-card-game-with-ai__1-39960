VERSION 5.00
Begin VB.UserControl StackControl 
   BackColor       =   &H00008000&
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
   ScaleHeight     =   2010
   ScaleWidth      =   6240
   Begin VB.PictureBox picCard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   0
      Left            =   90
      ScaleHeight     =   1455
      ScaleWidth      =   1065
      TabIndex        =   1
      Top             =   210
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.HScrollBar CardScroll 
      Enabled         =   0   'False
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   1710
      Width           =   6135
   End
End
Attribute VB_Name = "StackCOntrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents GameReference As clsGame
Attribute GameReference.VB_VarHelpID = -1
Private InitilizedDone As Boolean
Private Cards() As clsCard
Private SelectedCard As clsCard
Private StartCard As Integer
Private MappedID(15) As Integer
Const CardNo = 15

Public Event SelectedCardDoubleClick(sSelectedCard As clsCard)

Private Sub CardScroll_Change()
    StartCard = CardScroll.Value
    
    DrawCards
End Sub


Public Sub SetGame(cGame As clsGame)
    Set GameReference = cGame
    InitilizedDone = True
End Sub

Private Sub GameReference_StackChanged(StackCards() As clsCard)
    ReDim Cards(UBound(StackCards))
    Set SelectedCard = Nothing
    Cards = StackCards
    If Cards(0) Is Nothing Then
        InitCards
    Else
        InitCards
        DrawCards
    End If
End Sub

Public Sub InitCards()
    Dim Counter As Integer
    Dim BreakoutCounter As Integer
    For Counter = 0 To picCard.Count - 1
        picCard(Counter).Visible = False
        picCard(Counter).ZOrder 0
    Next
    
    If UBound(Cards) > CardNo Then
        StartCard = 0
        CardScroll.Max = (UBound(Cards) - CardNo) + 1
        CardScroll.Enabled = True
    Else
        CardScroll.Max = 0
        CardScroll.Enabled = False
    End If
End Sub

Public Sub DrawCards()
    Dim Counter As Integer
    Dim BreakoutCounter As Integer
    On Error Resume Next
    For Counter = 0 To picCard.Count - 1
        picCard(Counter).Top = 210
    Next
'
'    If UBound(Cards) > 15 Then
'        StartCard = 0
'        CardScroll.Max = UBound(Cards) - 15
'    End If
    
    BreakoutCounter = 0
    For Counter = StartCard To UBound(Cards)
        'Breakout when more than the required number of cards have been shown
        If BreakoutCounter = CardNo Then Exit For
        'Show the card
        picCard(BreakoutCounter).Visible = True
        'Draw the card onto the surface and refresh
        Cards(Counter).DrawCard picCard(BreakoutCounter).hDC
        picCard(BreakoutCounter).Refresh
        'Map the card to its direct cards() array ID
        MappedID(BreakoutCounter) = Counter
        'Set the tag equal to the card's absolutevalue property
        picCard(BreakoutCounter).Tag = Cards(Counter).AbsolutePosition
        'Reposition the card if already selected
        If SelectedCard Is Nothing Then
        
        Else
            If Cards(Counter).AbsolutePosition = SelectedCard.AbsolutePosition Then
                picCard(BreakoutCounter).Top = 6
            End If
        End If
        'Increase the counter
        BreakoutCounter = BreakoutCounter + 1
        'MsgBox ConvertCardToName(Cards(Counter))
    Next
End Sub

Private Sub picCard_Click(Index As Integer)
    Dim Counter As Integer
    
    
    For Counter = 0 To picCard.Count - 1
        picCard(Counter).Top = 210
    Next
    
    picCard(Index).Top = 60
    Set SelectedCard = GameReference.GetCardFromPosition(picCard(Index).Tag)
End Sub

Private Sub picCard_DblClick(Index As Integer)
    If SelectedCard Is Nothing Then
    
    Else
        If GetCardFromPosition(picCard(Index).Tag).AbsolutePosition = SelectedCard.AbsolutePosition Then
            'MsgBox "You double clicked on " & SelectedCard.GetCardName
            RaiseEvent SelectedCardDoubleClick(SelectedCard)
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim Counter As Integer
    Dim LeftInc As Integer
    'picCard(0).BackColor = UserControl.BackColor
    LeftInc = 300
    For Counter = 1 To 20
        Load picCard(Counter)
        picCard(Counter).Left = LeftInc
        LeftInc = LeftInc + 300
    Next
    StartCard = 0
End Sub

Private Sub UserControl_Resize()
    CardScroll.Width = UserControl.Width - CardScroll.Left
End Sub


Private Function GetCardFromPosition(sID As Integer) As clsCard
    Dim Counter As Long
    Set GetCardFromPosition = Nothing
    For Counter = LBound(Cards) To UBound(Cards)
        If Cards(Counter).AbsolutePosition = sID Then
            Set GetCardFromPosition = Cards(Counter)
            Exit Function
        End If
    Next
End Function

Public Function GetSelectedCard() As clsCard
    Set GetSelectedCard = SelectedCard
End Function
