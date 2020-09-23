VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PlayerName As String
Public PlayerComputer As Boolean
Public AILevel As AILevels
Private GameReference As clsGame
Private mPlayerHand() As clsCard
Private mScore As Integer
Public PlayerID As Integer
Private Opened As Boolean

Public Event HandChange()
Public Event AIPlayCardCombo(Cards() As clsCard)
Public Event Win()

Public Sub SetReference(Game As clsGame)
    Set GameReference = Game
End Sub


Public Sub AddCard(ObjCard As clsCard)
    Dim Temp As Integer
    If mPlayerHand(0) Is Nothing Then
        Set ObjCard.CardOwner = Me
        Set mPlayerHand(0) = ObjCard
    Else
        Set ObjCard.CardOwner = Me
        Temp = UBound(mPlayerHand) + 1
        ReDim Preserve mPlayerHand(Temp)
        Set mPlayerHand(Temp) = ObjCard
    End If
    
    RaiseEvent HandChange
End Sub

Public Sub RemoveCard(ObjCard As clsCard)
    On Error Resume Next
    Dim Counter As Integer
    Dim Flag As Boolean
    Dim Temp() As clsCard
    Dim Position As Integer
    For Counter = LBound(mPlayerHand) To UBound(mPlayerHand)
        If mPlayerHand(Counter).AbsolutePosition = ObjCard.AbsolutePosition Then
            Flag = True
        End If
    Next
    
    If Flag = True Then
        If UBound(mPlayerHand) <> 0 Then
            ReDim Temp(UBound(mPlayerHand) - 1)
            Position = 0
            For Counter = LBound(mPlayerHand) To UBound(mPlayerHand)
                If mPlayerHand(Counter).AbsolutePosition <> ObjCard.AbsolutePosition Then
                    Set Temp(Position) = mPlayerHand(Counter)
                    Position = Position + 1
                Else
                    'We have found the card we are removing
                    
                End If
            Next
            
            ReDim mPlayerHand(UBound(Temp))
            
            For Counter = LBound(Temp) To UBound(Temp)
                Set mPlayerHand(Counter) = Temp(Counter)
            Next
            
            RaiseEvent HandChange
        Else
            MsgBox "Player has no more cards"
            
            Set mPlayerHand(0) = Nothing
            
            GameReference.PlayerWin Me
            
            RaiseEvent Win
'            Stop
        End If
    Else
        MsgBox "No card found to remove!!"
    End If
    
End Sub

Public Function GetHand() As clsCard()
    GetHand = mPlayerHand
End Function

Private Sub Class_Initialize()
    ReDim mPlayerHand(0)
End Sub


Public Function GetNumberOfCards() As Integer
    If mPlayerHand(0) Is Nothing Then
        GetNumberOfCards = 0
    Else
        GetNumberOfCards = UBound(mPlayerHand) + 1
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
        For Innercounter = LBound(mPlayerHand) To (UBound(mPlayerHand) - Counter)
            If mPlayerHand(Innercounter).CardValue > mPlayerHand(Innercounter + 1).CardValue Then
                Set Temp = mPlayerHand(Innercounter)
                Set mPlayerHand(Innercounter) = mPlayerHand(Innercounter + 1)
                Set mPlayerHand(Innercounter + 1) = Temp
                Swapped = True
            End If
        Next
        Counter = Counter + 1
    Loop Until Swapped = False
    
    RaiseEvent HandChange
End Sub

Public Sub Play()
Dim Counter As Integer
Dim Outercounter As Integer
Dim Innercounter As Integer
Dim TempCard As clsCard
Dim Temp() As clsCard
Dim OtherCard As clsCard
Dim Flag As Boolean
Dim Combo As clsCardCombo
Dim Combos() As clsCardCombo
    MsgBox PlayerName & " is going to play..."
    
    Select Case AILevel
    
    Case UtterlyUselessMindlessDrone  'Utterly Useless Mindless Drone - Name says it all
        'Pick up from the deck
        GameReference.DealFromDeck Me
        'All this does is discard a card at random
        Randomize
        Counter = Int((UBound(mPlayerHand) - LBound(mPlayerHand) + 1) * Rnd + LBound(mPlayerHand))
        Set TempCard = mPlayerHand(Counter)
        DiscardCard TempCard
    Case SemiConsciousFumblingIdiot
        'This player is much more intelligent than the mindless drone but that isnt saying much
        'It does not pick up from the stack but will try and play.
        'Pick up from the deck
        GameReference.DealFromDeck Me
        'Try and get the top card
        ReDim Temp(2)
        Temp = GetHighestPlay(mPlayerHand)
        If Temp(0) Is Nothing Then
        Else
            RaiseEvent AIPlayCardCombo(Temp)
        End If
        'All this does is discard a card at random
        Randomize
        Counter = Int((UBound(mPlayerHand) - LBound(mPlayerHand) + 1) * Rnd + LBound(mPlayerHand))
        Set TempCard = mPlayerHand(Counter)
        DiscardCard TempCard
    Case MildlyIntelligentFrothingMoron
        'This player is almost identicle to the semiconciouse idiot but
        'is capable of playing more than one trick at a time and is
        'also not stupid enough to chuck out a joker.
        GameReference.DealFromDeck Me
PlayRunPhase:
        'Try and get the top card
        Flag = True
        Do Until Flag = False
            ReDim Temp(2)
            Temp = GetHighestPlay(mPlayerHand)
            If Temp(0) Is Nothing Then
                Flag = False
            Else
                RaiseEvent AIPlayCardCombo(Temp)
            End If
        Loop
        'Check if we can swap jokers...
  
        'Check if we can add
        On Error GoTo DiscardPhase
        Combos = GameReference.GetCardCombos
        Flag = False
            Do Until Flag = True
                If Combos(LBound(Combos)) Is Nothing Then
                
                Else
                    Flag = True
                    For Outercounter = LBound(Combos) To UBound(Combos)
                        For Counter = LBound(mPlayerHand) To UBound(mPlayerHand)
                            If Combos(Outercounter).ValidateAdd(mPlayerHand(Counter)) = True Then
                                Set TempCard = mPlayerHand(Counter)
                                Combos(Outercounter).AddCard TempCard
                                Combos(Outercounter).SortCards
                                RemoveCard TempCard
                                Flag = False
                                Exit For
                            End If
                        Next
                    Next
                End If
            Loop
        
        'Discard Phase
DiscardPhase:
        'All this does is discard a card at random
        Randomize
        Flag = True
        
        Do Until Flag = False
            Counter = Int((UBound(mPlayerHand) - LBound(mPlayerHand) + 1) * Rnd + LBound(mPlayerHand))
            Set TempCard = mPlayerHand(Counter)
            If TempCard Is Nothing Then
                Flag = False
            Else
                If TempCard.CardValue <> joker Then
                    Flag = False
                End If
            End If
            
        Loop
        
        
        DiscardCard TempCard
    Case NotMensaMaterial
        'This player is almost identicle to the semiconciouse idiot but
        'is capable of playing more than one trick at a time and is
        'also not stupid enough to chuck out a joker.
        'GameReference.DealFromDeck Me
        If CheckStack(mPlayerHand) > 0 Then
            Temp = GameReference.GetStack
            For Counter = LBound(Temp) To UBound(Temp)
                AddCard Temp(Counter)
                GameReference.StackRemoveCard Temp(Counter)
            Next
        Else
            GameReference.DealFromDeck Me
        End If
        
'SwapJokers:
'        Combos = GameReference.GetCardCombos
'        'Attempt to swap for a joker
'        For Counter = LBound(Combos) To UBound(Combos)
'            For Innercounter = LBound(mPlayerHand) To UBound(mPlayerHand)
'                If Combos(Counter).ValidateJokerAdd(mPlayerHand(Innercounter)) = True Then
'                    Set TempCard = Combos(Counter).SwapJoker(mPlayerHand(Innercounter))
'                    Set OtherCard = mPlayerHand(Innercounter)
'                    RemoveCard OtherCard
'                    Set OtherCard.CardOwner = TempCard.CardOwner
'                    Set TempCard.CardOwner = Me
'                    AddCard TempCard
'
'                End If
'            Next
'        Next
PlayRunPhase2:
        'Try and get the top card
        Flag = True
        Do Until Flag = False
            ReDim Temp(2)
            Temp = GetHighestPlay(mPlayerHand)
            If Temp(0) Is Nothing Then
                Flag = False
            Else
                RaiseEvent AIPlayCardCombo(Temp)
            End If
        Loop
        
        
        'Check if we can add
        On Error GoTo DiscardPhase
        Combos = GameReference.GetCardCombos
        Flag = False
            Do Until Flag = True
                If Combos(LBound(Combos)) Is Nothing Then
                
                Else
                    Flag = True
                    For Outercounter = LBound(Combos) To UBound(Combos)
                        For Counter = LBound(mPlayerHand) To UBound(mPlayerHand)
                            If Combos(Outercounter).ValidateAdd(mPlayerHand(Counter)) = True Then
                                Set TempCard = mPlayerHand(Counter)
                                Combos(Outercounter).AddCard TempCard
                                Combos(Outercounter).SortCards
                                RemoveCard TempCard
                                Flag = False
                                Exit For
                            End If
                        Next
                    Next
                End If
            Loop
        
        'Discard Phase
DiscardPhase2:
        'All this does is discard a card at random
        Randomize
        Flag = True
        
        Do Until Flag = False
            Counter = Int((UBound(mPlayerHand) - LBound(mPlayerHand) + 1) * Rnd + LBound(mPlayerHand))
            Set TempCard = mPlayerHand(Counter)
            If TempCard.CardValue <> joker Then
                Flag = False
            End If
        Loop
        
        
        DiscardCard TempCard
    End Select


    
    GameReference.DetermineNextPlayer
End Sub

Public Function GetHighestPlay(Cards As Variant) As clsCard()
'This function, given an array of cards will produce the highest
'three card combination
Dim Counter As Integer
Dim Innercounter As Integer
Dim Outercounter As Integer
Dim Score As Integer
Dim TopScore As Integer
Dim Temp() As clsCard
Dim Hand() As clsCard
Dim TopHand() As clsCard
    'Resize the tempory internal array and store the values in it
    ReDim Hand(UBound(Cards))
    For Counter = LBound(Cards) To UBound(Cards)
        Set Hand(Counter) = Cards(Counter)
    Next
    'Try every three-card combination
    ReDim Temp(2)
    ReDim TopHand(2)
    For Counter = LBound(Hand) To UBound(Hand)
        For Outercounter = LBound(Hand) To UBound(Hand)
            For Innercounter = LBound(Hand) To UBound(Hand)
                If Counter <> Outercounter And Counter <> Innercounter And Innercounter <> Outercounter Then
                    Set Temp(0) = Hand(Counter)
                    Set Temp(1) = Hand(Outercounter)
                    Set Temp(2) = Hand(Innercounter)
                    Score = GameReference.CheckCardCombination(Temp)
                    If Score > TopScore Then
                        TopScore = Score
                        Set TopHand(0) = Temp(0)
                        Set TopHand(1) = Temp(1)
                        Set TopHand(2) = Temp(2)
                        'MsgBox TopHand(0).GetCardName & Chr(13) & Chr(10) & TopHand(1).GetCardName & Chr(13) & Chr(10) & TopHand(2).GetCardName
                    End If
                End If
            Next
        Next
    Next
    '
    If TopScore = 0 Then
        Set Temp(0) = Nothing
        Set Temp(1) = Nothing
        Set Temp(2) = Nothing
    End If
    'Return the top combination
    GetHighestPlay = TopHand
End Function

Public Function CheckStack(Hand As Variant) As Integer
    Dim Counter As Integer
    Dim Innercounter As Integer
    Dim Outercounter As Integer
    Dim Score As Integer
    Dim Values As Integer
    Dim Tempo As Integer
    Dim TempCard As clsCard
    Dim Temp() As clsCard
    Dim InitialScore As Integer
    Dim Combinations(2) As clsCard
    
    'Temp = GameReference.GetStack
    ReDim Temp(UBound(Hand))
    For Counter = LBound(Hand) To UBound(Hand)
        Set TempCard = Hand(Counter)
        Set Temp(Counter) = TempCard
    Next
    'Get the original Size of the array
    For Counter = LBound(Temp) To UBound(Temp)
        For Outercounter = LBound(Temp) To UBound(Temp)
            For Innercounter = LBound(Temp) To UBound(Temp)
              If Counter <> Outercounter And Counter <> Innercounter And Innercounter <> Outercounter Then
                If Temp(Counter) Is Nothing Then
                Else
                    If Temp(Innercounter) Is Nothing Then
                    Else
                        If Temp(Outercounter) Is Nothing Then
                        Else
                            Set Combinations(0) = Temp(Counter)
                            Set Combinations(1) = Temp(Outercounter)
                            Set Combinations(2) = Temp(Innercounter)
                            Tempo = GameReference.CheckCardCombination(Combinations)
                            If Tempo <> 0 Then
                                Set Temp(Counter) = Nothing
                                Set Temp(Outercounter) = Nothing
                                Set Temp(Innercounter) = Nothing
                            End If
                            InitialScore = InitialScore + Tempo
                        End If
                    End If
                End If
              End If
            Next
        Next
        Values = Values + ConvertCardtoScore(Temp(Counter))
    Next
    
    Temp = GameReference.GetStack
    Tempo = UBound(Temp)
    ReDim Preserve Temp(Tempo + UBound(Hand))
    
    For Counter = LBound(Hand) To UBound(Hand)
        Set TempCard = Hand(Counter)
        Set Temp(Tempo + Counter) = TempCard
    Next
    
    For Counter = LBound(Temp) To UBound(Temp)
        For Outercounter = LBound(Temp) To UBound(Temp)
            For Innercounter = LBound(Temp) To UBound(Temp)
              If Counter <> Outercounter And Counter <> Innercounter And Innercounter <> Outercounter Then
                If Temp(Counter) Is Nothing Then
                Else
                    If Temp(Innercounter) Is Nothing Then
                    Else
                        If Temp(Outercounter) Is Nothing Then
                        Else
                            Set Combinations(0) = Temp(Counter)
                            Set Combinations(1) = Temp(Outercounter)
                            Set Combinations(2) = Temp(Innercounter)
                            Tempo = GameReference.CheckCardCombination(Combinations)
                            If Tempo <> 0 Then
                                Set Temp(Counter) = Nothing
                                Set Temp(Outercounter) = Nothing
                                Set Temp(Innercounter) = Nothing
                            End If
                            Score = Score + Tempo
                        End If
                    End If
                End If
              End If
            Next
        Next
        
    
    Next
    
    If ((Score - InitialScore) - Values) > -20 And Score <> 0 Then
        CheckStack = Score
    End If
End Function

Public Sub DiscardCard(Card As clsCard)
    GameReference.AddToStack Card
    Me.RemoveCard Card
End Sub

Public Sub ForceHandUpdate()
    RaiseEvent HandChange
End Sub

