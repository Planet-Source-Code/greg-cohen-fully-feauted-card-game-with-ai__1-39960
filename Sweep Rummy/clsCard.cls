VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'These are the two critical properties of the class
Public CardValue As CardValues
Public Suit As CardSuits
'Stores the position of the card in the deck
Private m_AbsolutePosition As Integer   'Position in deck
Public CardDrawValue As Integer         'Number between 0 and 51 reprisenting the card to be drawn
Public CardOwner As clsPlayer
'Public JokerValue As CardValues
'Public JokerSuit As CardSuits
'Public JokerValuesCertain As Boolean

Public Property Get AbsolutePosition() As Integer
    AbsolutePosition = m_AbsolutePosition
End Property

Public Property Let AbsolutePosition(ByVal vNewValue As Integer)
    m_AbsolutePosition = vNewValue
End Property

Public Function GetCardName()
    GetCardName = ConvertCardToName(Me)
End Function


Public Sub DrawCard(hDC As Long, Optional BackgroundColour As ColorConstants = vbGreen)
    Dim X As Long, Y As Long
    'Initilize the cards dll
    cdtInit X, Y
    'Draw the Card
    If CardValue = Joker Then
        Call cdtDraw(hDC, 0, 0, conRobot, conBacks, 0)
    Else
        Call cdtDraw(hDC, 0, 0, CardDrawValue, 0, BackgroundColour)
    End If
End Sub

