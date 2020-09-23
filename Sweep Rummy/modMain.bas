Attribute VB_Name = "modMain"

Public Enum CardValues
    Ace = 1
    Two = 2
    Three = 3
    Four = 4
    five = 5
    Six = 6
    Seven = 7
    Eight = 8
    Nine = 9
    Ten = 10
    Jack = 11
    Queen = 12
    King = 13
    joker = 14
End Enum

Public Enum CardSuits
    clubs = 1
    diamonds = 2
    Hearts = 3
    Spades = 4
End Enum

Public Enum PlayTypes
    Trip = 1
    Quad = 2
    RunOfthree = 3
    RunOfMore = 4
    Error = 0
End Enum

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum AILevels
    UtterlyUselessMindlessDrone = 0     'I once actually lost to this guy, took me seventy two rounds...
    SemiConsciousFumblingIdiot = 1
    MildlyIntelligentFrothingMoron = 2
    NotMensaMaterial = 3
    
    TelepathicOmipotentBeing = 5
    'All seeing and all cheating entity (NOT IMPLEMENTED, Kinda spoiled the game somewhat
End Enum

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const ONTOP = 0         ' for setting z-order
Public Const ONBOTTOM = 1
Public Const conFaces = 0
Public Const conBacks = 1
Public Const conInvert = 2
Public Const conCrossHatch = 53   'This is the design for the discard pile marker
Public Const conPlaid = 54
Public Const conWeave = 55
Public Const conRobot = 56
Public Const conRoses = 57
Public Const conIvyBlack = 58
Public Const conIvyBlue = 59
Public Const conFishCyan = 60
Public Const conFishBlue = 61
Public Const conShell = 62
Public Const conCastle = 63
Public Const conBeach = 64
Public Const conCardHand = 65
Public Const conUnused = 66
Public Const conX = 67            'big red X
Public Const conO = 68            'big green O
Public Const NOTDRAGGING = -1
Public Const DRAGGINGSINGLE = 0
Public Const DRAGGINGMULTIPLE = 1



'Public Declare Function cdtInit Lib "Cards32.Dll" (dX As Long, dY As Long) As Long
'Public Declare Function cdtDrawExt Lib "Cards32.Dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal conCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
'Public Declare Function cdtDraw Lib "Cards32.Dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal iCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
'Public Declare Function cdtAnimate Lib "Cards32.Dll" (ByVal hDC As Long, ByVal iCardBack As Long, ByVal X As Long, ByVal Y As Long, ByVal iState As Long) As Long
'Public Declare Function cdtTerm Lib "Cards32.Dll" () As Long

Public Declare Function cdtInit Lib "Cards.Dll" (dX As Long, dY As Long) As Long
Public Declare Function cdtDrawExt Lib "Cards.Dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal conCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
Public Declare Function cdtDraw Lib "Cards.Dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal iCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
Public Declare Function cdtAnimate Lib "Cards.Dll" (ByVal hDC As Long, ByVal iCardBack As Long, ByVal X As Long, ByVal Y As Long, ByVal iState As Long) As Long
Public Declare Function cdtTerm Lib "Cards.Dll" () As Long
'If you have trouble with the above declarations, the switch the the CARDS32.DLL Version above


Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long


Public Function ConvertValueToName(sValue As Integer) As String
    ConvertValueToName = Choose(sValue, "Ace", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Jack", "Queen", "King", "Joker")
End Function

Public Function ConvertSuitToName(sValue As Integer) As String
    If sValue = 0 Then Exit Function 'This is for jokers
    ConvertSuitToName = Choose(sValue, "Clubs", "Diamonds", "Hearts", "Spades")
End Function

Public Function ConvertCardToName(ObjCard As clsCard)
    If ObjCard.CardValue = joker Then
        ConvertCardToName = "Joker"
    Else
        ConvertCardToName = ConvertValueToName(ObjCard.CardValue) & " of " & ConvertSuitToName(ObjCard.Suit)
    End If
End Function

Public Function ConvertCardtoScore(ObjCard As clsCard) As Integer
On Error Resume Next
    ConvertCardtoScore = Choose(ObjCard.CardValue, 15, 2, 3, 4, 5, 6, 7, 8, 9, 10, 10, 10, 10, 50)
End Function

Public Function ConvertPlayTypeToName(sType As PlayTypes) As String
    ConvertPlayTypeToName = Choose(sType, "Trip", "Closed Trip", "Suit Run", "Suit Run", "Something not quite rigt")
End Function
