VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class is used to hold various played
'combinations of cards. Thus is has to be
'able to hold an array of clsCards. It is
'also aware of what type of cards it holds
'and who owns each card.

Private Cards() As clsCard          'Internal Collection of cards
Private GameReference As clsGame    'A refence to the active game
Attribute GameReference.VB_VarHelpID = -1
Private InitialPlayer As clsPlayer  'A reference to the initial player of the card
Private Initilized As Boolean       'Value tracking whether game has been initilized
Private PlayType As PlayTypes       'Allows the object to know what type of play it is holding
'Variables used to store position of the graphical reprisentation
Private LocationTop As Integer
Private LocationLeft As Integer

Public Event CardsChanged(Cards() As clsCard)
Public Event HideMe()

Public Sub ReferenceGame(Game As clsGame)
    Set GameReference = Game
End Sub

Public Sub FirstCards(vCards As Variant, sPlayType As PlayTypes)
'This sub is used to add the first cards to the collection
Dim Counter As Integer
    'Resize the Internal Array and copy across the information
    ReDim Cards(UBound(vCards))
    For Counter = LBound(vCards) To UBound(vCards)
        Set Cards(Counter) = vCards(Counter)
    Next
    'Set the play type
    PlayType = sPlayType
    'Sort the cards
    SortCards
    
    'Raise the event
    RaiseEvent CardsChanged(Cards)
End Sub

Public Function ValidateAdd(sCard As clsCard) As Boolean
    Dim Temp() As clsCard
    Dim Counter As Integer
    ReDim Temp(UBound(Cards) + 1)
    'Fill the tempory array with the current cards
    For Counter = LBound(Cards) To UBound(Cards)
        Set Temp(Counter) = Cards(Counter)
    Next
    'Add the new card
    Set Temp(UBound(Temp)) = sCard
    'Use the Game class to check this new combination
    ValidateAdd = GameReference.CheckCardCombination(Temp)
End Function

Public Function ValidateAddCards(sCard As Variant) As Boolean
    Dim Temp() As clsCard
    Dim Counter As Integer
    Dim TempCard As clsCard
    ReDim Temp(UBound(Cards) + UBound(sCard) + 1)
    'Fill the tempory array with the current cards
    For Counter = LBound(Cards) To UBound(Cards)
        Set Temp(Counter) = Cards(Counter)
    Next
    'Add the new card
    For Counter = LBound(sCard) To UBound(sCard)
        Set TempCard = sCard(Counter)
        Set Temp(UBound(Cards) + Counter + 1) = TempCard
    Next
    
    'Use the Game class to check this new combination
    ValidateAddCards = GameReference.CheckCardCombination(Temp)
End Function


Public Sub AddCard(sCard As clsCard)
    'Resize the array to handle one more card
    ReDim Preserve Cards(UBound(Cards) + 1)
    'Add the new card
    Set Cards(UBound(Cards)) = sCard
    'Change the status of the play type
    If PlayType = Trip Then
        PlayType = Quad
        RaiseEvent HideMe
    ElseIf PlayType = RunOfthree Then
        PlayType = RunOfMore
    End If
    
    RaiseEvent CardsChanged(Cards)
End Sub

Public Function GetCombo() As clsCard()
    GetCombo = Cards
End Function


Public Function SwapJoker(NewCard As clsCard) As clsCard
    Dim Temp() As clsCard
    Dim Counter As Integer
    Dim JokerFound As Boolean
    
    ReDim Temp(UBound(Cards))
    For Counter = LBound(Cards) To UBound(Cards)
        If Cards(Counter).CardValue = joker Then
            JokerFound = True
        End If
        Set Temp(Counter) = Cards(Counter)
    Next
    
    If JokerFound = False Then
        Set SwapJoker = Nothing
        Exit Function
    Else
        'Find the first joker and replace it with the new one
        For Counter = LBound(Temp) To UBound(Temp)
            If Temp(Counter).CardValue = joker Then
                Set SwapJoker = Temp(Counter)
                Set Temp(Counter) = NewCard
                Exit For
            End If
        Next
        'Test this new combination
        If GameReference.CheckCardCombination(Temp) <> 0 Then
            For Counter = LBound(Temp) To UBound(Temp)
                Set Cards(Counter) = Temp(Counter)
                
            Next
            'set SwapJoker = Tru
            RaiseEvent CardsChanged(Cards)
            Exit Function
        Else
            Set SwapJoker = Nothing
            
            Exit Function
        End If
    End If
End Function
'Functality needed
'DONE 1. Add Card
'DONE 2. Add Array of cards (Initial cards)
'     3. Swap Joker
'DONE 4. Check if card can be added
'     5. Sort cards appropriately (including ACES and jokers)

'Bubble Sort Sort
'-----------------------------------------------
'    Counter = 1
'    Do
'        Swapped = False
'        For InnerCounter = LBound(FirstResponses) To (UBound(FirstResponses) - Counter)
'            If FirstResponses(InnerCounter).PercentageSentence < FirstResponses(InnerCounter + 1).PercentageSentence Then
'                Temp = FirstResponses(InnerCounter)
'                FirstResponses(InnerCounter) = FirstResponses(InnerCounter + 1)
'                FirstResponses(InnerCounter + 1) = Temp
'                Swapped = True
'            End If
'        Next
'        Counter = Counter + 1
'    Loop Until Swapped = False

Public Function ValidateJokerAdd(NewCard As clsCard) As Boolean
    Dim Temp() As clsCard
    Dim Counter As Integer
    Dim JokerFound As Boolean
    
    ReDim Temp(UBound(Cards))
    For Counter = LBound(Cards) To UBound(Cards)
        If Cards(Counter).CardValue = joker Then
            JokerFound = True
        End If
        Set Temp(Counter) = Cards(Counter)
    Next
    
    If JokerFound = False Then
        ValidateJokerAdd = False
        Exit Function
    Else
        'Find the first joker and replace it with the new one
        For Counter = LBound(Temp) To UBound(Temp)
            If Temp(Counter).CardValue = joker Then
                'Set SwapJoker = Temp(Counter)
                Set Temp(Counter) = NewCard
                Exit For
            End If
        Next
        'Test this new combination
        If GameReference.CheckCardCombination(Temp) <> 0 Then
            ValidateJokerAdd = True
        Else
            ValidateJokerAdd = False
        End If
    End If
End Function

Public Sub SortCards()
    Dim Counter As Integer
    Dim Innercounter As Integer
    Dim Swapped As Boolean
    Dim Temp As clsCard
    Dim Resort() As clsCard
    'Preliminary Sort
    Counter = 1
    Do
        Swapped = False
        For Innercounter = LBound(Cards) To (UBound(Cards) - Counter)
            If Cards(Innercounter).CardValue > Cards(Innercounter + 1).CardValue Then
                Set Temp = Cards(Innercounter)
                Set Cards(Innercounter) = Cards(Innercounter + 1)
                Set Cards(Innercounter + 1) = Temp
                Swapped = True
            End If
        Next
        Counter = Counter + 1
    Loop Until Swapped = False
    
    If PlayType = RunOfMore Or PlayType = RunOfthree Then
        'We need to take care of jokers and Aces
        
    End If
    'Tell any controls subscribing to this class that the hand has changed
    RaiseEvent CardsChanged(Cards)
End Sub
