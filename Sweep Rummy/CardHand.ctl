VERSION 5.00
Begin VB.UserControl CardHand 
   BackColor       =   &H0000C000&
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
   ScaleHeight     =   2010
   ScaleWidth      =   6240
   Begin VB.HScrollBar CardScroll 
      Enabled         =   0   'False
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   1710
      Width           =   6135
   End
   Begin VB.PictureBox picCard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   0
      Left            =   60
      ScaleHeight     =   1455
      ScaleWidth      =   1065
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "CardHand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ReferencedPlayer As clsPlayer
Attribute ReferencedPlayer.VB_VarHelpID = -1
Private InitilizedDone As Boolean
Private Cards() As clsCard
Private Selected() As Boolean
Private StartCard As Integer
Private MappedID(15) As Integer
Private SelectedCard() As clsCard

Const CardNo = 15

Public Event CardClicked(Card As clsCard, Selected As Boolean)
Public Event CardCleared()

Public Sub SetPlayer(cPlayer As clsPlayer)
    Set ReferencedPlayer = cPlayer
    InitilizedDone = True
End Sub



Private Sub CardScroll_Change()
    StartCard = CardScroll.Value
    
    DrawCards
End Sub

Private Sub picCard_Click(Index As Integer)
    'MsgBox ConvertCardToName(GetCardFromPosition(CInt(picCard(Index).Tag)))
    
    '210 and 60
    If Selected(MappedID(Index)) = True Then
        Selected(MappedID(Index)) = False
        picCard(Index).Top = 210
        RaiseEvent CardClicked(Cards(MappedID(Index)), False)
    Else
        picCard(Index).Top = 60
        Selected(MappedID(Index)) = True
        RaiseEvent CardClicked(Cards(MappedID(Index)), True)
    End If
    'Raise the event
    
End Sub

Private Sub ReferencedPlayer_HandChange()
'    Dim Counter As Integer
'    Dim A() As clsCard
'    Dim Temp As String
'    A = ReferencedPlayer.GetHand
'    For Counter = LBound(A) To UBound(A)
'        Temp = Temp & Chr(13) & Chr(10) & ConvertCardToName(A(Counter))
'    Next
'
'    MsgBox Temp
    Cards = ReferencedPlayer.GetHand
    ReDim Selected(UBound(Cards))
    ReDim SelectedCard(0)
    RaiseEvent CardCleared
    InitCards
    DrawCards
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
    On Error Resume Next
    Dim Counter As Integer
    Dim BreakoutCounter As Integer
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
        If Selected(Counter) = True Then
            picCard(BreakoutCounter).Top = 60
        End If
        'Increase the counter
        BreakoutCounter = BreakoutCounter + 1
        'MsgBox ConvertCardToName(Cards(Counter))
    Next
End Sub

Private Sub UserControl_Resize()
    CardScroll.Width = UserControl.Width - CardScroll.Left
    UserControl.Height = 2010
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

Public Function GetSelected() As clsCard()
Dim Counter As Integer
Dim Temp() As clsCard
Dim Pos As Integer
Dim First As Boolean
Dim Changed As Boolean
ReDim Temp(0)
First = True
Changed = False
For Counter = LBound(Cards) To UBound(Cards)
    If Selected(Counter) = True Then
        If First = True Then
            Set Temp(0) = Cards(Counter)
            Changed = True
            First = False
        Else
            Pos = UBound(Temp) + 1
            ReDim Preserve Temp(Pos)
            Set Temp(Pos) = Cards(Counter)
            Changed = True
        End If
    End If
Next

If Changed = True Then GetSelected = Temp
End Function

Public Function SelectedCount() As Integer
    Dim Counter As Integer
    ReDim SelectedCard(0)
    If InitilizedDone = False Then Exit Function
    
    
    For Counter = LBound(Cards) To UBound(Cards)
        If Selected(Counter) = True Then
            If SelectedCard(0) Is Nothing Then
                Set SelectedCard(0) = Cards(Counter)
            Else
                ReDim Preserve SelectedCard(UBound(SelectedCard) + 1)
                Set SelectedCard(UBound(SelectedCard)) = Cards(Counter)
            End If
        End If
    Next
    
    If SelectedCard(0) Is Nothing Then
        SelectedCount = 0
    Else
        SelectedCount = UBound(SelectedCard)
    End If
End Function

Public Function GetSelectedCard(Index As Integer) As clsCard
On Error GoTo EndFunction
    If Index > UBound(SelectedCard) Then Exit Function
    
    Set GetSelectedCard = SelectedCard(Index)
EndFunction:
End Function

