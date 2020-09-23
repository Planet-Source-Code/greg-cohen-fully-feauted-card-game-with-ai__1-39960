VERSION 5.00
Begin VB.UserControl CardCombo 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2145
   ScaleHeight     =   1500
   ScaleWidth      =   2145
   Begin VB.PictureBox picCard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   0
      Left            =   30
      ScaleHeight     =   1455
      ScaleWidth      =   1065
      TabIndex        =   0
      Top             =   30
      Width           =   1065
   End
End
Attribute VB_Name = "CardCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ReferencedCardCombo As clsCardCombo
Attribute ReferencedCardCombo.VB_VarHelpID = -1

Public Event CardClicked(CardPos As Integer)
Public Event DoubleClick()
Public Event MouseMove(CardPos As Integer)
Public Event HideMe(sTExt As String)

Public Sub SetRefence(mCombo As clsCardCombo)
    Dim Cards() As clsCard
    Set ReferencedCardCombo = mCombo
    Cards = ReferencedCardCombo.GetCombo
    DrawCards Cards
End Sub

Public Function GetReference() As clsCardCombo
    Set GetReference = ReferencedCardCombo
End Function


Public Sub DrawCards(mCards As Variant)
    Dim Counter As Integer
    Dim NewWidth As Integer
    Dim Temp As Integer
    Dim Other As Integer
    'Increment left by 300
    NewWidth = picCard(0).Left + picCard(0).Width + ((UBound(mCards) + 1) * 300)
    'Resize the window
    UserControl.Width = NewWidth
    
    Temp = (UBound(mCards) + 1) - picCard.Count
    
    If Temp > 0 Then
        For Counter = 1 To Temp
            Other = picCard.Count
            Load picCard(Other)
            picCard(Other).Left = picCard(0).Left + (300 * Other)
            picCard(Other).Visible = True
            picCard(Other).ZOrder ONTOP
            
        Next
    End If
    
    For Counter = LBound(mCards) To UBound(mCards)
        mCards(Counter).DrawCard picCard(Counter).hDC
        picCard(Counter).Refresh
        picCard(Counter).Tag = mCards(Counter).AbsolutePosition
    Next
    
    'MsgBox Temp
End Sub

Private Sub picCard_Click(Index As Integer)
    'RaiseEvent CardClicked(picCard(Index).Tag)
End Sub

Private Sub picCard_DblClick(Index As Integer)
    'RaiseEvent DoubleClick
End Sub

Private Sub picCard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        RaiseEvent DoubleClick
    Else
        RaiseEvent CardClicked(picCard(Index).Tag)
    End If
End Sub

Private Sub picCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(picCard(Index).Tag)
End Sub

Private Sub ReferencedCardCombo_CardsChanged(Cards() As clsCard)
    DrawCards Cards
End Sub

Private Sub ReferencedCardCombo_HideMe()
    'UserControl.Width = 1
    'UserControl.Height = 1
    Dim Temp As String
    Dim Cards() As clsCard
    Cards = ReferencedCardCombo.GetCombo
    Temp = "Quad " & ConvertValueToName(Cards(LBound(Cards)).CardValue)
    RaiseEvent HideMe(Temp)
End Sub

