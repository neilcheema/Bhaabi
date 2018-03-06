Public Module common
#Region "Variables and Structures"
    Public player1Clicking As Boolean = True
    Public Enum Face
        None 'Used for Jokers    
        Ace
        Two
        Three
        Four
        Five
        Six
        Seven
        Eight
        Nine
        Ten
        Jack
        Queen
        King
    End Enum
    Public Enum Suit
        None 'Used for Jokers    
        Hearts
        Clubs
        Diamonds
        Spades
    End Enum
    Public Enum Orientation
        Flipped
        NotFlipped
        NoCard
    End Enum
   
    Public pileID As Integer
    Public Structure Card
        Private m_Face As Face
        Private m_Suit As Suit
        Private m_Flipped As Boolean
        Private m_Index As Integer
        Public ReadOnly Property Face() As Face
            Get
                Return Me.m_Face
            End Get
        End Property
        Public ReadOnly Property Suit() As Suit
            Get
                Return Me.m_Suit
            End Get
        End Property
        Public ReadOnly Property IsJoker() As Boolean
            Get
                Return Me.m_Face = Face.None AndAlso Me.m_Suit = Suit.None
            End Get
        End Property
        Public Property Flipped() As Boolean
            Get
                Flipped = m_Flipped
            End Get
            Set(ByVal Value As Boolean)
                m_Flipped = Value
            End Set
        End Property
        Public Property Index() As Integer
            Get
                Index = m_Index
            End Get
            Set(ByVal Value As Integer)
                m_Index = Value
            End Set
        End Property
        Public Sub New(ByVal face As Face, ByVal suit As Suit)
            If face = face.None Xor suit = suit.None Then
                Throw New ApplicationException("If either face or suit is None then the other must also be None.")
            Else
                Me.m_Face = face
                Me.m_Suit = suit
                Me.m_Index = -1
            End If
        End Sub
    End Structure

#End Region

#Region "Common Functions"
    Public Function ReturnOrientationText(ByVal Orient As Orientation) As String
        Select Case Orient
            Case Orientation.Flipped
                Return "FLIPPED"
            Case Orientation.NotFlipped
                Return "NOT FLIPPED"
            Case Orientation.NoCard
                Return "NO CARD"
        End Select
    End Function
    Public Function ConvertSuit(ByVal suit As Suit) As String
        Select Case suit
            Case suit.None
                Return ("o")
            Case suit.Clubs
                Return ("c")
            Case suit.Hearts
                Return ("h")
            Case suit.Spades
                Return ("s")
            Case suit.Diamonds
                Return ("d")
        End Select
    End Function
    Public Function convertFace(ByVal cardType As String) As Integer
        'If cardType.Empty Then Exit Function
        Dim faceValue As Integer
        cardType = cardType.TrimEnd.TrimStart
        cardType = cardType.Substring(1, cardType.Length - 1)
        faceValue = CType(cardType, Integer)
        Return faceValue
    End Function
    Public Function returnSuit(ByVal cardType As String) As Suit
        'If cardType.Empty Then Exit Function
        Dim faceValue As Integer
        cardType = cardType.TrimEnd.TrimStart
        cardType = cardType.Substring(0, 1)
        faceValue = CType(cardType, Integer)
        Select Case faceValue
            Case "c"
                Return Suit.Clubs
            Case "d"
                Return Suit.Diamonds
            Case "s"
                Return Suit.Spades
            Case "h"
                Return Suit.Hearts
        End Select
        Return faceValue
    End Function
    Public Function returnFace(ByVal cardType As String) As Face
        'If cardType.Empty Then Exit Function
        Dim faceValue As Integer
        cardType = cardType.TrimEnd.TrimStart
        cardType = cardType.Substring(1, cardType.Length - 1)
        faceValue = CType(cardType, Integer)
        Select Case faceValue
            Case 0
                Return Face.None
            Case 1
                Return Face.Ace
            Case 2
                Return Face.Two
            Case 3
                Return Face.Three
            Case 4
                Return Face.Four
            Case 5
                Return Face.Five
            Case 6
                Return Face.Six
            Case 7
                Return Face.Seven
            Case 8
                Return Face.Eight
            Case 9
                Return Face.Nine
            Case 10
                Return Face.Ten
            Case 11
                Return Face.Jack
            Case 12
                Return Face.Queen
            Case 13
                Return Face.King
        End Select
    End Function
    Public Function ConvertFace(ByVal Face As Face) As String
        Select Case Face
            Case Face.None
                Return 0
            Case Face.Ace
                Return 1
            Case Face.Two
                Return 2
            Case Face.Three
                Return 3
            Case Face.Four
                Return 4
            Case Face.Five
                Return 5
            Case Face.Six
                Return 6
            Case Face.Seven
                Return 7
            Case Face.Eight
                Return 8
            Case Face.Nine
                Return 9
            Case Face.Ten
                Return 10
            Case Face.Jack
                Return 11
            Case Face.Queen
                Return 12
            Case Face.King
                Return 13
        End Select
    End Function
    Public Function ConvertScore(ByVal face As Face, ByVal suit As Suit) As Integer
        Select Case suit
            Case suit.Spades
                Select Case face
                    Case face.None
                        Return 0
                    Case face.Ace
                        Return 1
                    Case face.Two
                        Return 2
                    Case face.Three
                        Return 3
                    Case face.Four
                        Return 4
                    Case face.Five
                        Return 5
                    Case face.Six
                        Return 6
                    Case face.Seven
                        Return 7
                    Case face.Eight
                        Return 8
                    Case face.Nine
                        Return 9
                    Case face.Ten
                        Return 10
                    Case face.Jack
                        Return 11
                    Case face.Queen
                        Return 12
                    Case face.King
                        Return 13
                End Select
            Case suit.Clubs
                Select Case face
                    Case face.None
                        Return 0
                    Case face.Ace
                        Return 1
                    Case Else
                        Return 0
                End Select
            Case suit.Diamonds
                Select Case face
                    Case face.None
                        Return 0
                    Case face.Ace
                        Return 1
                    Case face.Ten
                        Return 6
                    Case Else
                        Return 0
                End Select
            Case suit.Hearts
                Select Case face
                    Case face.None
                        Return 0
                    Case face.Ace
                        Return 1
                    Case Else
                        Return 0
                End Select
        End Select
    End Function
    Public Function ConvertPileValue(ByVal value As String) As Integer
        Select Case value
            Case "Nine"
                Return 9
            Case "Ten"
                Return 10
            Case "Eleven"
                Return 11
            Case "Twelve"
                Return 12
            Case "Thirteen"
                Return 13
        End Select

    End Function



    'Public Function becomingPile(ByVal cardPoint As Point, ByVal paRect() As playAreaRectangles) As Boolean


    '        For i As Integer = 0 To paRect.Length - 1

    '            'If ((cardPoint.X > paRect(i).paRectPoint.X) AndAlso _
    '            '            (cardPoint.X < (paRect(i).paRectPoint.X + paRect(i).paRectW)) AndAlso _
    '            '            (cardPoint.Y > paRect(i).paRectPoint.Y) AndAlso _
    '            '            ((cardPoint.Y < (paRect(i).paRectPoint.Y + paRect(i).paRectH) > 380))) Then Return True

    '            If (cardPoint.X > paRect(i).paRectPoint.X) AndAlso _
    '            (cardPoint.Y > paRect(i).paRectPoint.Y) Then Return True
    '        Next
    '        Return False
    '    End Function
#End Region ' Common Functions

End Module
