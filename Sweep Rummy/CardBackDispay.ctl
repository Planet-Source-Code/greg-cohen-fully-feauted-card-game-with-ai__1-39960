VERSION 5.00
Begin VB.UserControl CardBackDispay 
   BackColor       =   &H0000C000&
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   ScaleHeight     =   1830
   ScaleWidth      =   6015
   Begin VB.PictureBox picCard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1755
      Index           =   0
      Left            =   30
      ScaleHeight     =   1755
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "CardBackDispay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_Alignment = 1
Const m_def_Back = conPlaid
Const m_WidthDiff = 300
'Property Variables:
Dim m_Alignment As Long
Dim m_Back As Long
Public WithEvents ReferencedPlayer As clsPlayer
Attribute ReferencedPlayer.VB_VarHelpID = -1
Public Event MouseMove()
Public Event AIPlayCardCombo(Cards() As clsCard)

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get Alignment() As Long
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Long)
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function DrawCards(NumberOfCards As Integer) As Variant
    Dim Counter As Long
    Dim CurrentLeft As Long
    Dim CurrentHeight As Long
    
    For Counter = 0 To picCard.Count - 1
        picCard(Counter).Visible = False
        picCard(Counter).ZOrder 0
    Next
    
    If NumberOfCards > 15 Then NumberOfCards = 15
    CurrentLeft = 5
    CurrentHeight = 5
    For Counter = 1 To NumberOfCards
        picCard(Counter - 1).Visible = True
         If m_Alignment = 1 Then
            'Width
            picCard(Counter - 1).Left = CurrentLeft
            CurrentLeft = CurrentLeft + 300
        Else
            'Height
            picCard(Counter - 1).Top = CurrentHeight
            CurrentHeight = CurrentHeight + 300
        End If
        DrawCardBack picCard(Counter - 1)
    Next
    
    
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Back() As Long
    Back = m_Back
End Property

Public Property Let Back(ByVal New_Back As Long)
    m_Back = New_Back
    PropertyChanged "Back"
End Property

Private Sub picCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove
End Sub

Private Sub ReferencedPlayer_AIPlayCardCombo(Cards() As clsCard)
    RaiseEvent AIPlayCardCombo(Cards)
End Sub

Private Sub ReferencedPlayer_HandChange()
    DrawCards ReferencedPlayer.GetNumberOfCards
End Sub

Private Sub UserControl_Initialize()
    Dim Counter As Integer
    picCard(0).BackColor = UserControl.BackColor
    For Counter = 1 To 20
        Load picCard(Counter)
    Next
    
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Alignment = m_def_Alignment
    m_Back = m_def_Back
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    m_Back = PropBag.ReadProperty("Back", m_def_Back)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
    Call PropBag.WriteProperty("Back", m_Back, m_def_Back)
End Sub


Private Sub DrawCardBack(sTarget As PictureBox)
    Dim X As Long, Y As Long
    cdtInit X, Y
    Call cdtDraw(sTarget.hDC, 0, 0, m_Back, conBacks, 0)    'Draws card backs
    
    sTarget.Refresh
End Sub

Public Sub SetPlayer(cPlayer As clsPlayer)
    Set ReferencedPlayer = cPlayer
    
    DrawCards cPlayer.GetNumberOfCards
End Sub

