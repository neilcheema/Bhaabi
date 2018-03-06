''///////////////////////////////////////////////////////////
''// The clsCardDeck Class
''///////////////////////////////////////////////////////////
'Option Explicit On 
'Option Strict On

'Imports System
'Imports System.Drawing
Public Class clsCardDeck
#Region "Variables and Structures"

    'Other Private Variables
    Private PlayerCount As Integer 'Holds the number of players in the game
    'Cards Variables
    Private m_playerID As Integer = 1
    Private m_cardID As Integer = 0
    Private cards As clsHand 'The cards currently in the deck.    
    Private rand As Random 'Used to generate random numbers.
    
#End Region

    Public ReadOnly Property Count() As Integer
        Get
            Return Me.cards.Count
        End Get
    End Property
    Public Sub New()
        Me.cards = New clsHand         'Create a full deck with no Jokers.        
        For Each suit As Suit In [Enum].GetValues(GetType(Suit))
            If suit <> suit.None Then
                For Each face As Face In [Enum].GetValues(GetType(Face))
                    If face <> face.None Then
                        Me.cards.Add(New Card(face, suit))
                    End If
                Next face
            End If
        Next suit
        Me.rand = New Random
    End Sub

    Public Sub Shuffle()
        Dim shuffledCards As New clsHand
        While Me.cards.Count > 0
            Dim index As Integer = Me.rand.Next(0, Me.cards.Count)             'Add a card at random to the new deck from the existing deck.           
            shuffledCards.Add(Me.cards(index))             'Remove the chosen card from the existing deck.            
            Me.cards.RemoveAt(index)
        End While
        Me.cards = shuffledCards
    End Sub
    Public Sub Cut(ByVal upperCount As Integer)        'Move the first upperCount cards to the bottom.        
        For i As Integer = 1 To upperCount
            Me.cards.Add(Me.cards(0))
            Me.cards.RemoveAt(0)
        Next i
    End Sub     'Returns the next card without removing it from the deck.    
    Public Function PeekNextCard() As Card
        Return Me.cards(0)
    End Function     'Returns the next card and removes it from the deck.    
    Public Function GetNextCard(ByVal flipped As Boolean, ByVal index As Integer) As Card
        Dim nextCard As Card = Me.cards(0)
        Me.cards.RemoveAt(0)
        nextCard.Flipped = flipped
        nextCard.Index = index
        Return nextCard
    End Function     'Returns the specified number of hands, each with the specified number of cards.    
    Public Function Deal(ByVal cardCount As Integer, ByVal flipped As Boolean) As clsHand
        Dim hand As clsHand
        hand = New clsHand
        For i As Integer = 1 To cardCount Step 1
            hand.Insert(hand.Count, Me.GetNextCard(flipped, i))
        Next i
        Return hand
    End Function
    Public Sub Remove(ByVal card As Card)
        'this sub collects cards in from the designated hand
        Me.Remove(card)
    End Sub
End Class
''    Public Sub setUp(ByVal nPlayers As Integer)
'        'if we are allready set we dont need to start again
'        If isSet = True Then
'            Exit Sub
'        End If

'        Dim i As Integer 'counter
'        Dim t As Integer 'counter
'        Dim CardCount As Integer 'Another counter to determine what card number we are up to

'        PlayerCount = nPlayers

'        CardCount = 0

'        'these loops fill our array with all the cards
'        For t = 0 To 3
'            For i = 0 To 12
'                Select Case t
'                    Case 0
'                        deck(CardCount) = "h." & i + 1
'                    Case 1
'                        deck(CardCount) = "d." & i + 1
'                    Case 2
'                        deck(CardCount) = "s." & i + 1
'                    Case 3
'                        deck(CardCount) = "c." & i + 1
'                End Select
'                CardCount = CardCount + 1
'            Next
'        Next

'        'Now we must set up the drawing section
'        'cdtInit(0, 0)
'        Application.DoEvents()

'        isSet = True ' set the variable to true so we can tell that the class has been initialised!
'    End Sub

'    Public Sub Shuffle(ByVal nTimes As Integer)
'        'this sub shuffles the deck and moves the values to activeDeck

'        Dim x As Integer
'        Dim swap As String
'        Dim r As Random = New Random

'        For p As Integer = 0 To nTimes
'            For i As Integer = 0 To deck.GetUpperBound(0)
'                x = r.Next(0, i)

'                swap = deck(x)
'                deck(x) = deck(i)
'                deck(i) = swap
'            Next i
'        Next p

'        For i As Integer = 0 To 51
'            m_cardDeck(i).index = i
'            m_cardDeck(i).CardType = deck(i)
'            m_cardDeck(i).Location = "."
'            m_cardDeck(i).PlayerID = "."
'        Next

'    End Sub

'    Public Sub collect(ByVal nPlayer As Integer)
'        'this sub collects cards in from the designated player id and sets location values to .
'        For i As Integer = 0 To 51
'            If m_cardDeck(i).playerID = CType(nPlayer, String) Then
'                m_cardDeck(i).location = "."
'                m_cardDeck(i).playerID = "."
'            End If
'        Next
'    End Sub

'    Public Function Deal(ByVal nPlayer As Integer) As String
'        'this sub deals a card to a nominated player and stores the player id in the active deck array
'        Dim i As Integer = currentIndex

'        If m_cardDeck(i).location = "." Then
'            If currentIndex >= 51 Then
'                currentIndex = 0
'            Else
'                currentIndex = currentIndex + 1
'            End If
'            If m_playerID <> nPlayer Then
'                m_cardID = 1
'            Else
'                m_cardID = m_cardID + 1
'            End If
'            m_cardDeck(i).location = "+"
'            m_cardDeck(i).playerID = CType(nPlayer, String)
'            m_cardDeck(i).cardID = CType(m_cardID, String)
'            Return m_cardDeck(i).cardType
'        End If
'    End Function

'    Public Function returnCards(ByVal nplayer As Integer) As String
'        Dim temp As String

'        For i As Integer = 0 To 51
'            If m_cardDeck(i).playerID = CType(nplayer, String) Then
'                temp = temp & "," & m_cardDeck(i).cardType
'            End If
'        Next

'        Return temp
'    End Function

'    'Public Sub drawBack(ByVal style As cardBack, ByVal g As Graphics, ByVal x As Integer, ByVal y As Integer)
'    '    Dim hdc As IntPtr = g.GetHdc()
'    '    Try


'    '        cdtDraw(hdc, x, y, backStyle, 1, 0)
'    '    Finally
'    '        g.ReleaseHdc(hdc)
'    '    End Try
'    'End Sub
'    'Public Sub DrawCard(ByVal nCard As String, ByVal g As Graphics, ByVal face As orientation, ByVal x As Integer, ByVal y As Integer)
'    '    Dim index As Integer
'    '    For i As Integer = 0 To 51
'    '        If m_cardDeck(i).cardType = CType(nCard, String) Then
'    '            index = i
'    '        End If
'    '    Next
'    '    m_cardDeck(index).Display(nCard, g, face, x, y)

'    'End Sub
'    Public Sub Close()
'        'cdtTerm()
'    End Sub


'End Class
