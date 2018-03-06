Public Class clsHand
    Inherits CollectionBase
    
    Default Public Property Item(ByVal index As Integer) As Card
        Get
            Return CType(Me.InnerList(index), Card)
        End Get
        Set(ByVal Value As Card)
            Me.InnerList(index) = Value
        End Set
    End Property
    Public Function Add(ByVal card As Card) As Integer
        Return Me.InnerList.Add(card)
    End Function
    Public Sub Insert(ByVal index As Integer, ByVal card As Card)
        Me.InnerList.Insert(index, card)
    End Sub
    Public Sub Remove(ByVal card As Card)
        Me.InnerList.Remove(card)
    End Sub

    Public Sub New()
        MyBase.New()
    End Sub

End Class
