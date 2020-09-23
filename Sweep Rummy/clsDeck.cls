VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Deck(107) As clsCard
Private m_DeckPosition As Integer


Public Event OutOfCards()

Public Sub InitilizeDeck()
'This procedure adds two sets of decks to the cards
Dim OuterCounter As Integer
Dim CardCounter As Integer
Dim SuitCounter As Integer
Dim CurrentPos As Integer
    'Create the objects
    For OuterCounter = LBound(m_Deck) To UBound(m_Deck)
        Set m_Deck(OuterCounter) = New clsCard
    Next
    'Create the cards
    For OuterCounter = 1 To 2
        For CardCounter = 1 To 13
                For SuitCounter = 1 To 4
                m_Deck(CurrentPos).CardValue = CardCounter
                m_Deck(CurrentPos).Suit = SuitCounter
                m_Deck(CurrentPos).CardDrawValue = CurrentPos - (52 * (OuterCounter - 1))
                m_Deck(CurrentPos).AbsolutePosition = CurrentPos
                CurrentPos = CurrentPos + 1
            Next
        Next
    Next
    'Add the four jokers
    For OuterCounter = 1 To 4
        m_Deck(CurrentPos).CardValue = joker
        m_Deck(CurrentPos).AbsolutePosition = CurrentPos
        m_Deck(CurrentPos).CardDrawValue = -1
        CurrentPos = CurrentPos + 1
    Next
End Sub

Public Sub Shuffle()
    Dim Temp As clsCard
    Dim ShuffleCounter As Integer
    Dim SwapCounter As Integer
    Dim OldCard As Integer
    Dim NewCard As Integer
    'Randomize the pseudo-random generator
    Randomize
    'Begin the shuffling loop
    For ShuffleCounter = 1 To Int((10 - 2 + 1) * Rnd + 2)
        For SwapCounter = 1 To Int((30000 - 10000 + 1) * Rnd + 10000)
            Randomize Timer + ShuffleCounter + SwapCounter
            'Choose a card
            OldCard = Int((UBound(m_Deck) - LBound(m_Deck) + 1) * Rnd + LBound(m_Deck))
            'Choose a new location
            NewCard = Int((UBound(m_Deck) - LBound(m_Deck) + 1) * Rnd + LBound(m_Deck))
            'Swap the cards around
            Set Temp = m_Deck(OldCard)
            Set m_Deck(OldCard) = m_Deck(NewCard)
            Set m_Deck(NewCard) = Temp
            Set Temp = Nothing
        Next
    Next
End Sub

Public Sub CutDeck()
    'No Idea
End Sub

Public Function GetDeck() As clsCard()
    GetDeck = m_Deck
End Function

Public Function DealNextCard() As clsCard
    If m_DeckPosition > UBound(m_Deck) Then
        'MsgBox "Out of Cards!"
        RaiseEvent OutOfCards
        Exit Function
    End If
    Set DealNextCard = m_Deck(m_DeckPosition)
    m_DeckPosition = m_DeckPosition + 1
End Function


Public Function GetCardFromPosition(sID As Integer) As clsCard
    Dim Counter As Long
    Set GetCardFromPosition = Nothing
    For Counter = LBound(m_Deck) To UBound(m_Deck)
        If m_Deck(Counter).AbsolutePosition = sID Then
            Set GetCardFromPosition = m_Deck(Counter)
            Exit Function
        End If
    Next
End Function

Public Function GetRemainingCardNo() As Integer
    
    GetRemainingCardNo = UBound(m_Deck) - m_DeckPosition + 1 ' Zero based array!
End Function
