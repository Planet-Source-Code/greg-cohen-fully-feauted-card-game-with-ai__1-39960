VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variables relating to the last combination checked using the checkcombination function
Private LastCombinationError As String          'Stores a reason why the check failed (if it did)
Private LastCombinationChecked As PlayTypes     'Stores the type of play detected by the function
'Game Variables
Private WithEvents InternalDeck As clsDeck      'The class holding the game's internal deck of cards
Attribute InternalDeck.VB_VarHelpID = -1
Private Players(3) As clsPlayer                 'An array holding the values for the four players
Private Stack() As clsCard                      'The stack/pile of cards played
Private mGameOpened As Boolean                  'Variable storing whether or not the game has been officially opened
Private mPlayersTurn As Byte                    'Variable keeping track of whose turn it is
Private mTurnCounter As Integer                 'Variable counting the number of turns played
Private CardsPlayed() As clsCardCombo           'Array of classes containing arrays of cards
Private WinnerAbortGame As Boolean
'Events
Public Event StackChanged(StackCards() As clsCard) 'This event is called whenever a card is removed from or added to the stack
Public Event HumanTurn()
Public Event GameOver()
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                       -+-+-+ Class Functions +-+-+-
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Class_Initialize()
    Dim Counter As Integer
    'Create a new deck
    Set InternalDeck = New clsDeck
    'Initilize the players
    For Counter = LBound(Players) To UBound(Players)
        Set Players(Counter) = New clsPlayer
        Players(Counter).PlayerID = Counter
        Players(Counter).SetReference Me
    Next
    ReDim CardsPlayed(0)
End Sub

Private Sub Class_Terminate()
    Dim Counter As Integer
    'Clean up code
    Set InternalDeck = Nothing
    For Counter = LBound(Players) To UBound(Players)
        Set Players(Counter) = Nothing
    Next
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                       -+-+-+ Player Functions +-+-+-
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Public Property Get Player(sPlayer As Integer) As clsPlayer
    Set Player = Players(sPlayer)
End Property

'This function is just an easy way of setting the player names, and should only really be used in the testing phase (but we all know the truth...)
Public Sub SetPlayerNames(Player1 As String, Player2 As String, Player3 As String, Player4 As String)
    Players(0).PlayerName = Player1
    Players(1).PlayerName = Player2
    Players(2).PlayerName = Player3
    Players(3).PlayerName = Player4
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                   -+-+-+ Deck/Card Functions +-+-+-
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Public Sub Initilize()
Dim Temp As Integer
    'Initilize and shuffle the deck
    InternalDeck.InitilizeDeck
    InternalDeck.Shuffle
    'Set player names
    
    'Determine a random starting player
    Randomize
    Temp = Int((3 - 0 + 1) * Rnd)
    mPlayersTurn = Temp
    'Initilize the Card combination array
    ReDim CardsPlayed(0)
End Sub



Public Sub DealCards()
    Dim CardCounter As Integer
    Dim PlayerCounter As Integer
    Dim Temp As clsCard
    
    
    'Deal eleven times, one to each player each time
    For CardCounter = 1 To 11
        For PlayerCounter = LBound(Players) To UBound(Players)
            Set Temp = InternalDeck.DealNextCard
            If WinnerAbortGame = True Then Exit Sub
            Set Temp.CardOwner = Players(PlayerCounter)
            Players(PlayerCounter).AddCard Temp
        Next
    Next
    'Add the first card to the stack
    ReDim Stack(0)
    Set Stack(0) = InternalDeck.DealNextCard
    RaiseEvent StackChanged(Stack)
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                        -+-+-+ Table Functions +-+-+-
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Public Function PlayCombination(CardCombination As Variant, Player As clsPlayer, sPlayType As PlayTypes) As clsCardCombo
    Dim NewCombo As clsCardCombo
    Dim Counter As Integer
    Dim TempCard As clsCard
    'Create a new object
    Set NewCombo = New clsCardCombo
    'Add the initial cards
    NewCombo.FirstCards CardCombination, sPlayType
    'Add it to the array
    If CardsPlayed(0) Is Nothing Then
        Set CardsPlayed(0) = NewCombo
    Else
        ReDim Preserve CardsPlayed(UBound(CardsPlayed) + 1)
        Set CardsPlayed(UBound(CardsPlayed)) = NewCombo
    End If
    'Remove the played cards from the player
    For Counter = LBound(CardCombination) To UBound(CardCombination)
        Set TempCard = CardCombination(Counter)
        Player.RemoveCard TempCard
    Next
    Set PlayCombination = NewCombo
    'Destroy the tempory object
    Set NewCombo = Nothing
End Function

Public Function GetCombos() As clsCardCombo()
        GetCombos = CardsPlayed
End Function

Public Sub CheckCardsAgainstTable(CardCombination As Variant)

End Sub

Public Function GetCardCombos() As clsCardCombo()
    If CardsPlayed(0) Is Nothing Then
        'GetCardCombos = Nothing
    Else
        GetCardCombos = CardsPlayed
    End If
End Function
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                        -+-+-+ Stack Functions +-+-+-
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'
Public Sub AddToStack(Card As clsCard)
    If Stack(0) Is Nothing Then
        Set Stack(0) = Card
    Else
        ReDim Preserve Stack(UBound(Stack) + 1)
        Set Stack(UBound(Stack)) = Card
    End If
    
    RaiseEvent StackChanged(Stack)
End Sub
'This function is used to retrieve the cards on the stack
Public Function GetStack() As clsCard()
    GetStack = Stack
End Function
'This function is used to retrieve all the cards on the right
'of the selected card in the stack. Used when picking up the
'stack
Public Function GetAllStackRightOfCard(Card As clsCard) As clsCard()
    Dim Counter As Integer
    Dim Position As Integer
    Dim Temp() As clsCard
    'Find the location of the card
    Position = -1
    For Counter = LBound(Stack) To UBound(Stack)
        If Stack(Counter).AbsolutePosition = Card.AbsolutePosition Then
            Position = Counter
            Exit For
        End If
    Next
    'Check that the card was found
    If Position <> -1 Then
        ReDim Temp(0)
        'Set Temp(0) = Card
        For Counter = Position To UBound(Stack)
            If Temp(0) Is Nothing Then
                Set Temp(0) = Stack(Counter)
            Else
                ReDim Preserve Temp(UBound(Temp) + 1)
                Set Temp(UBound(Temp)) = Stack(Counter)
            End If
        Next
        
        GetAllStackRightOfCard = Temp
    Else
        MsgBox "The card passed to GetAllStackRightOfCard was not found in the stack"
    End If
    
End Function

Public Sub StackRemoveCard(Card As clsCard)
    Dim Counter As Integer
    Dim Flag As Boolean
    Dim Temp() As clsCard
    Dim Position As Integer
    For Counter = LBound(Stack) To UBound(Stack)
        If Stack(Counter).AbsolutePosition = Card.AbsolutePosition Then
            Flag = True
        End If
    Next
    
    If Flag = True Then
        If UBound(Stack) <> 0 Then
            ReDim Temp(UBound(Stack) - 1)
            Position = 0
            For Counter = LBound(Stack) To UBound(Stack)
                If Stack(Counter).AbsolutePosition <> Card.AbsolutePosition Then
                    Set Temp(Position) = Stack(Counter)
                    Position = Position + 1
                Else
                    'We have found the card we are removing
                    
                End If
            Next
            
            ReDim Stack(UBound(Temp))
            
            For Counter = LBound(Temp) To UBound(Temp)
                Set Stack(Counter) = Temp(Counter)
            Next
            
            RaiseEvent StackChanged(Stack)
        Else
            ReDim Stack(0)
            
            RaiseEvent StackChanged(Stack)
        End If
    Else
        MsgBox "No card found to remove!!"
    End If
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                   -+-+-+ Check Combination Functions +-+-+-
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function CheckCardCombination(ParamArray vCards() As Variant) As Integer
    On Error Resume Next
    Dim Cards() As clsCard              'Internal array of the selected cards
    Dim Used() As Boolean               'Internal array used when checking runs
    Dim Counter As Integer              'General counter used to iterate through the cards() array
    Dim Outercounter As Integer         'Second counter often used to reiterate during an iteration
    Dim JokerCount As Integer           'Holds the number of jokers found in the selected cards
    Dim Temp As Integer                 'Tempory storage for numbers commonly used with the flag variables
    Dim Flag As Boolean                 'Used when checking for cards to indicate results
    Dim CurrentCard As clsCard          'Tempory storage for current card
    Dim CurrentCardValue As Integer     'Stores the array index of the current card
    Dim Breakout As Boolean             'Used to break out of loops before conditions are met
    Dim KingPresent As Boolean
    'Clear any state variables
    LastCombinationError = ""
    LastCombinationChecked = Error
    'Resize the Cards array
    ReDim Cards(UBound(vCards(0)))
    'Add the old items to the internalCardArray
    For Counter = LBound(vCards(0)) To UBound(vCards(0))
        Set Cards(Counter) = vCards(0)(Counter)
    Next
    'Firstly we must check to see if there are more than 3 cards
    If UBound(Cards) < 2 Then
        CheckCardCombination = 0
        LastCombinationError = "Too Few Cards Selected"
        Exit Function
    End If
    'Secondly check to see if there are two duplicates in the selected cards
    For Counter = LBound(Cards) To UBound(Cards)
        'Check to see that the card is not a joker
        If Cards(Counter).CardValue <> joker Then
            'Compare this card to the others
            For Outercounter = LBound(Cards) To UBound(Cards)
                'Make sure we are not comparing this card to itself
                If Outercounter <> Counter Then
                    If Cards(Counter).CardValue = Cards(Outercounter).CardValue Then
                        If Cards(Counter).Suit = Cards(Outercounter).Suit Then
                            'Check for the unlikely possibility of a 14 card run
                            If UBound(Cards) <> 13 Then
                                CheckCardCombination = 0
                                LastCombinationError = "Duplicate " & Cards(Counter).GetCardName & " selected"
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        Else
            'Just so we can remeber how many jokers are in the selected cards
            JokerCount = JokerCount + 1
        End If
    Next
    'Thirdly check that the cards are not all jokers (note: Jokercount is 1 based where as ubound(cards) is zero-based)
    If (JokerCount - 1) = UBound(Cards) Then
        CheckCardCombination = ConvertCardtoScore(Cards(0)) * JokerCount
        Exit Function
    End If
    'Now examine the suits to see if it is a rung or a trip/quad
    Temp = -1
    Flag = False
    'Compare the other card's suits to this one
    For Counter = LBound(Cards) To UBound(Cards)
        If Cards(Counter).CardValue <> joker Then
            If Temp = -1 Then
                Temp = Cards(Counter).Suit
            Else
                If Cards(Counter).Suit <> Temp Then
                    Flag = True
                    Exit For
                End If
            End If
        End If
    Next
    'Decide based on the outcome of this
    If Flag = True Then
        'We are dealing with a trip/quad
        'Properties of a trip/quad
        ' 1. Either 3 or 4 cards
        ' 2. No duplicates (checked above)
        ' 3. Jokers are allowed
        'We start by ensuring no more than four cards selected
        If UBound(Cards) <= 4 Then
            'Now we must check to see that they are all the same cards
            Temp = -1
            Flag = False
            'Check that the cards are all the same value
            For Counter = LBound(Cards) To UBound(Cards)
                If Cards(Counter).CardValue <> joker Then
                    If Temp = -1 Then
                        Temp = Cards(Counter).CardValue
                    Else
                        If Cards(Counter).CardValue <> Temp Then
                            Flag = True
                            Exit For
                        End If
                    End If
                End If
            Next
            'Decide based on the results above
            If Flag = False Then
                'Seems to be a valid trip/quad
                If UBound(Cards) > 3 Then
                    CheckCardCombination = 0
                    LastCombinationError = "Maximum of four cards can be played like this"
                    Exit Function
                End If
                CheckCardCombination = 0
                'Total up the score for the cards
                For Counter = LBound(Cards) To UBound(Cards)
                    CheckCardCombination = CheckCardCombination + ConvertCardtoScore(Cards(Counter))
                Next
                If UBound(Cards) = 2 Then
                    LastCombinationChecked = Trip
                Else
                    LastCombinationChecked = Quad
                End If
                Exit Function
            Else
                'Pack up and leave
                CheckCardCombination = 0
                LastCombinationError = "Cards are not of the same type"
                Exit Function
            End If
        Else
            'Pack up and leave
            CheckCardCombination = 0
            LastCombinationError = "Too many cards for a trip or quad"
            Exit Function
        End If
    Else
        'We are dealing with a rung
        'Properties of a rung
        ' 1. Any number of cards between 3 and 14
        ' 2. All cards must be in the same suit (checked above)
        ' 3. Technically possible to have doubles (although extremely unlikely)
        ' 4. Aces can be found at the beginning or the end
        ' 5. Jokers can take the place of any card
        
        'To start off we must look for the lowest card
        Set CurrentCard = Nothing
        KingPresent = False
        'Initilize the Used card array
        ReDim Used(UBound(Cards))
        'Set all the values to false
        For Counter = LBound(Used) To UBound(Used)
            Used(Counter) = False
            'Use this opportunity to check if there is a king present
            If Cards(Counter).CardValue = King Or Cards(Counter).CardValue = Queen Then KingPresent = True
        Next
        'Loop through the cards and look for the lowest card
        For Counter = LBound(Cards) To UBound(Cards)
            If Cards(Counter).CardValue <> joker Then
                If Cards(Counter).CardValue = Ace And KingPresent = True Then
                    'We must ignore this situation as the ace is not the lowest
                    'card
                Else
                    If CurrentCard Is Nothing Then
                        Set CurrentCard = Cards(Counter)
                    Else
                        If Cards(Counter).CardValue < CurrentCard.CardValue Then
                            Set CurrentCard = Cards(Counter)
                            CurrentCardValue = Counter
                        End If
                    End If
                End If
            End If
        Next
        'We now have the lowest card so we must mark it off as checked
        Used(CurrentCardValue) = True
        'Prepare for the checking loop
        Breakout = False
        Temp = 0
        Flag = False
        'Begin the loop
        For Counter = LBound(Cards) To UBound(Cards) - 1 'considering we already have the first card
            'Determine the next card we are looking for
            If Breakout = True Then
                Temp = intDetermineNextCard(CurrentCard.CardValue + 1)
                Breakout = False
            Else
                Temp = DetermineNextCard(CurrentCard)
            End If
            'Find the card
            Flag = False
            For Outercounter = LBound(Cards) To UBound(Cards)
                If Cards(Outercounter).CardValue = Temp Then
                    'If Used(OuterCounter) = True Then MsgBox "The next card determined has already been used? What is going on? Send for the men in white coats..."
                    Flag = True
                    Set CurrentCard = Cards(Outercounter)
                    CurrentCardValue = Outercounter
                    Exit For
                End If
            Next
            'See if we have found the card
            If Flag = False Then
                'Since we have not found the next card, check if there are any jokers left
                If JokerCount > 0 Then
                    'If there is, use it and continue
                    JokerCount = JokerCount - 1
                    'Increment the cardcounter (temp) so it looks for the next card
                    Breakout = True
                Else
                    'Otherwise we do not have a valid combination
                    CheckCardCombination = 0
                    LastCombinationError = "Could not find " & ConvertValueToName(Temp) & " of " & ConvertSuitToName(CurrentCard.Suit)
                    Exit Function
                End If
            End If
        Next
        
        'Seems to be a valid run
        If UBound(Cards) > 13 Then
            CheckCardCombination = 0
            LastCombinationError = "This run exceeds the lenght of a 14 cards (Ace to Ace)"
            Exit Function
        End If
        CheckCardCombination = 0
        If UBound(Cards) = 2 Then
            LastCombinationChecked = RunOfthree
        Else
            LastCombinationChecked = RunOfMore
        End If
        'Total up the score for the cards
        For Counter = LBound(Cards) To UBound(Cards)
            CheckCardCombination = CheckCardCombination + ConvertCardtoScore(Cards(Counter))
        Next
        Exit Function
        'This should never happen...
        MsgBox CurrentCard.GetCardName & " - " & ConvertValueToName(DetermineNextCard(CurrentCard))
    End If
    
End Function

Private Function DetermineNextCard(ObjCard As clsCard) As CardValues
'This function is used to find the next card expected. It takes a card object
'and returns a value between 1 - 13 (Fourteen is a joker and is never expected)
    If ObjCard.CardValue < King Then
        DetermineNextCard = ObjCard.CardValue + 1
    ElseIf ObjCard.CardValue = King Then
        DetermineNextCard = Ace
    End If
End Function

Private Function intDetermineNextCard(ObjCard As Integer) As CardValues
'Very similiar to the DetermineNextCard function except for the fact that
'it only accepts an integer. Due to complications involving Aces and jokers
'this is a separate procedure from the one above
        If ObjCard < King Then
            intDetermineNextCard = ObjCard + 1
        ElseIf ObjCard = King Then
            intDetermineNextCard = Ace
        End If
End Function

Public Property Get LastCheckCombinationError() As String
    LastCheckCombinationError = LastCombinationError
End Property

Public Property Get LastCheckCombinationType() As PlayTypes
    LastCheckCombinationType = LastCombinationChecked
End Property


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                     -+-+-+ Mapped Functions +-+-+-
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'This function is just mapping the one supplied by the internal deck
Public Function GetCardFromPosition(sID As Integer) As clsCard
    Set GetCardFromPosition = InternalDeck.GetCardFromPosition(sID)
End Function

'This function is just mapping the one provided by the deck
Public Function GetDeck() As clsCard()
    GetDeck = InternalDeck.GetDeck
End Function

'This function is just a mapping of the one provided by the deck
Public Function GetNoRemainingCards() As Integer
    GetNoRemainingCards = InternalDeck.GetRemainingCardNo
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                     -+-+-+ Event Responders +-+-+-
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Private Sub InternalDeck_OutOfCards()
    MsgBox "There are no more cards in the Deck! The Game is Over!"
    WinnerAbortGame = True
    RaiseEvent GameOver
End Sub

'I suppose this is the crux of the whole game. This is the subroutine that causes
'either a human turn or a computer turn to commence. Surprisingly straightforward
Public Sub DetermineNextPlayer()
    If WinnerAbortGame = True Then
        'Stop playing
        RaiseEvent GameOver
    Else
        If mPlayersTurn = 3 Then
            'It is now the humans turn to play
            mPlayersTurn = 0
            RaiseEvent HumanTurn
        Else
            'Otherwise let one of our insanely stupid bots have a turn
            'Well, they were insanely stupid... until they started beating me...
            mPlayersTurn = mPlayersTurn + 1
            Players(mPlayersTurn).Play
        End If
    End If
End Sub


Public Sub BeginGame()
    DetermineNextPlayer
End Sub


Public Sub DealFromDeck(toPlayer As clsPlayer)
    Dim TempCard As clsCard
    
    Set TempCard = InternalDeck.DealNextCard
    If WinnerAbortGame = True Then Exit Sub
    Set TempCard.CardOwner = toPlayer
    toPlayer.AddCard TempCard
End Sub


Public Function DeterminePlayerScore(Player As clsPlayer) As Integer
    DeterminePlayerScore = DeterminePlayerScoreOnTable(Player) - DeterminePlayerScoreInHand(Player)
End Function

Public Function DeterminePlayerScoreInHand(Player As clsPlayer) As Integer
    Dim Counter As Integer
    Dim Score As Integer
    Dim Cards() As clsCard
    
    Cards = Player.GetHand
     
    For Counter = LBound(Cards) To UBound(Cards)
        Score = Score + ConvertCardtoScore(Cards(Counter))
    Next
    
    DeterminePlayerScoreInHand = Score
End Function

Public Function DeterminePlayerScoreOnTable(Player As clsPlayer) As Integer
Dim Counter As Integer
Dim Innercounter As Integer
Dim Cards() As clsCard
Dim Score As Integer
'Dim Combos() As clsCardCombo
On Error Resume Next

For Counter = LBound(CardsPlayed) To UBound(CardsPlayed)
    Cards = CardsPlayed(Counter).GetCombo
    For Innercounter = LBound(Cards) To UBound(Cards)
        If Cards(Innercounter).CardOwner.PlayerID = Player.PlayerID Then
            Score = Score + ConvertCardtoScore(Cards(Innercounter))
        End If
    Next
Next
    DeterminePlayerScoreOnTable = Score

End Function

Public Sub PlayerWin(Player As clsPlayer)
    MsgBox Player.PlayerName & " has ended the game"
    WinnerAbortGame = True
End Sub
