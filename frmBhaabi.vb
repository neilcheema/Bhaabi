Public Class frmBhaabi
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.


    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents lblPlayer2Score As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblPlayer1Score As System.Windows.Forms.Label
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents playerCard As System.Windows.Forms.PictureBox
    Friend WithEvents middleCard As System.Windows.Forms.PictureBox
    Friend WithEvents computerCard As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.lblPlayer2Score = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.lblPlayer1Score = New System.Windows.Forms.Label
        Me.btnStart = New System.Windows.Forms.Button
        Me.playerCard = New System.Windows.Forms.PictureBox
        Me.middleCard = New System.Windows.Forms.PictureBox
        Me.computerCard = New System.Windows.Forms.PictureBox
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.lblPlayer2Score)
        Me.GroupBox5.Location = New System.Drawing.Point(536, 168)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(152, 48)
        Me.GroupBox5.TabIndex = 19
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Computer Cards"
        '
        'lblPlayer2Score
        '
        Me.lblPlayer2Score.Location = New System.Drawing.Point(64, 16)
        Me.lblPlayer2Score.Name = "lblPlayer2Score"
        Me.lblPlayer2Score.Size = New System.Drawing.Size(80, 24)
        Me.lblPlayer2Score.TabIndex = 8
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lblPlayer1Score)
        Me.GroupBox3.Location = New System.Drawing.Point(16, 168)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(152, 48)
        Me.GroupBox3.TabIndex = 92
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Player Cards"
        '
        'lblPlayer1Score
        '
        Me.lblPlayer1Score.Location = New System.Drawing.Point(64, 16)
        Me.lblPlayer1Score.Name = "lblPlayer1Score"
        Me.lblPlayer1Score.Size = New System.Drawing.Size(80, 24)
        Me.lblPlayer1Score.TabIndex = 7
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(296, 176)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.TabIndex = 93
        Me.btnStart.Text = "Play Again"
        '
        'playerCard
        '
        Me.playerCard.Location = New System.Drawing.Point(48, 32)
        Me.playerCard.Name = "playerCard"
        Me.playerCard.Size = New System.Drawing.Size(70, 100)
        Me.playerCard.TabIndex = 94
        Me.playerCard.TabStop = False
        '
        'middleCard
        '
        Me.middleCard.Location = New System.Drawing.Point(304, 24)
        Me.middleCard.Name = "middleCard"
        Me.middleCard.Size = New System.Drawing.Size(70, 100)
        Me.middleCard.TabIndex = 95
        Me.middleCard.TabStop = False
        '
        'computerCard
        '
        Me.computerCard.Location = New System.Drawing.Point(576, 24)
        Me.computerCard.Name = "computerCard"
        Me.computerCard.Size = New System.Drawing.Size(70, 100)
        Me.computerCard.TabIndex = 96
        Me.computerCard.TabStop = False
        '
        'frmMain
        '
        Me.AllowDrop = True
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(692, 266)
        Me.Controls.Add(Me.computerCard)
        Me.Controls.Add(Me.middleCard)
        Me.Controls.Add(Me.playerCard)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox5)
        Me.Name = "frmBhaabi"
        Me.Text = "Card Collector"
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Variables and Structures"

    Private player1Score, player2Score As Integer
    Public myDeck As clsCardDeck
    Public myHands As clsHand
    Public compHands As clsHand
    Public middleHands As clsHand




    Private Structure ControlPositionType
        Dim Left As Single
        Dim Top As Single
        Dim Width As Single
        Dim Height As Single
        Dim FontSize As Single
    End Structure

    Private m_ControlPositions() As ControlPositionType
    Private m_FormWid As Single
    Private m_FormHgt As Single






#End Region       ' Variables and Structures

#Region "Event Handlers"



    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        reset()
        btnStart.Enabled = False
    End Sub
    Private Sub playerCard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles playerCard.Click


        If myHands.Count = 0 AndAlso compHands.Count <> 0 Then
            Dim strMessage As String = "You won, Computer has: " + CType(compHands.Count, String) + " Cards and cards in the middle: " + CType(middleHands.Count, String)
            MessageBox.Show(strMessage, "Two Player Bhaabi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            playerCard.Image = Nothing
            btnStart.Enabled = True
            Exit Sub
        ElseIf myHands.Count = 0 AndAlso compHands.Count = 0 Then
            Dim strMessage As String = "Draw Game and cards in the middle: " + CType(middleHands.Count, String)
            MessageBox.Show(strMessage, "Two Player Bhaabi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            playerCard.Image = Nothing
            computerCard.Image = Nothing
            btnStart.Enabled = True
            Exit Sub
        ElseIf compHands.Count = 0 Then
            Dim strMessage As String = "You lost please press Start Again Button"
            MessageBox.Show(strMessage, "Two Player Bhaabi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            btnStart.Enabled = True
            Exit Sub
        End If



        If (FlipPlayerCard(0)) Then
            player1Clicking = Not (player1Clicking)
            computerPlays()
        End If

        lblPlayer1Score.Text = myHands.Count
        lblPlayer2Score.Text = compHands.Count
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SaveSizes()
        reset()
    End Sub
    Private Sub frmMain_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        ResizeControls()
    End Sub




#End Region     ' Event Handlers

#Region "Initial Two Player Bhaabi Functions"

    ' Save the form's and controls' dimensions.
    Private Sub SaveSizes()
        Dim i As Integer
        Dim ctl As Control

        ' Save the controls' positions and sizes.
        ReDim m_ControlPositions(Controls.Count)
        i = 1
        For Each ctl In Controls
            With m_ControlPositions(i)
                .Left = ctl.Left
                .Top = ctl.Top
                .Width = ctl.Width
                .Height = ctl.Height
                On Error Resume Next
                .FontSize = ctl.Font.Size
                On Error GoTo 0
            End With
            i = i + 1
        Next ctl

        ' Save the form's size.
        m_FormWid = Me.Width
        m_FormHgt = Me.Height
    End Sub
    ' Arrange the controls for the new size.
    Private Sub ResizeControls()
        Dim i As Integer
        Dim ctl As Control
        Dim x_scale As Single
        Dim y_scale As Single

        ' Don't bother if we are minimized.
        If WindowState = FormWindowState.Maximized Then Exit Sub

        ' Get the form's current scale factors.
        x_scale = Me.Width / m_FormWid
        y_scale = Me.Height / m_FormHgt

        ' Position the controls.
        i = 1
        For Each ctl In Controls
            With m_ControlPositions(i)

                ctl.Left = x_scale * .Left
                ctl.Top = y_scale * .Top
                ctl.Width = x_scale * .Width
                If Not (TypeOf ctl Is ComboBox) Then
                    ' Cannot change height of ComboBoxes.
                    ctl.Height = y_scale * .Height
                End If
                On Error Resume Next
                'ctl.Font.Size = y_scale * .FontSize
                On Error GoTo 0
            End With
            i = i + 1
        Next ctl
    End Sub



    Private Sub reset()
        player1Score = 0
        player2Score = 0
        myHands = Nothing
        compHands = Nothing
        middleHands = Nothing
        Me.lblPlayer1Score.Text = player1Score
        Me.lblPlayer2Score.Text = player1Score
        btnStart.Enabled = False
        middleCard.Image = Nothing
        FlipFirstHand()
    End Sub
    Private Sub FlipFirstHand()
        Dim rm As New Resources.ResourceManager("twoPlayerBhaabi.images", System.Reflection.Assembly.GetExecutingAssembly)

        Dim cardType As String

        myDeck = New clsCardDeck
        myDeck.Shuffle()

        myHands = myDeck.Deal(26, False)
        compHands = myDeck.Deal(26, False)
        Dim tempImage As Image = rm.GetObject("b1fv.gif")

        playerCard.Image = tempImage
        computerCard.Image = tempImage

    End Sub

#End Region     ' Intial Two Player Bhaabi Functions




#Region "General Functions"

    Private Function FlipPlayerCard(ByVal index As Integer) As Boolean
        Dim result As Boolean
        Dim rm As New Resources.ResourceManager("twoPlayerBhaabi.images", System.Reflection.Assembly.GetExecutingAssembly)
        Dim cardType As String = ConvertSuit(myHands.Item(index).Suit) + ConvertFace(myHands.Item(index).Face)
        Dim cardFace As String = ConvertFace(myHands.Item(index).Face)
        Dim fileName As String = cardType + ".gif"
        Dim tempImage As Image = rm.GetObject(fileName)
        Dim middleCardFace As String = middleCard.Text
        If middleCardFace <> "" Then middleCardFace = middleCardFace.Substring(1, middleCardFace.Length - 1)

        If middleHands Is Nothing Then middleHands = New clsHand
        Dim newCard As Card
        newCard = New Card
        newCard = myHands.Item(index)
        newCard.Flipped = True
        myHands.Remove(myHands.Item(index))
        middleHands.Insert(middleHands.Count, newCard)

        If middleCardFace = cardFace Then
            middleCard.Image = tempImage
            middleCard.Text = cardType.ToUpper
            Me.Refresh()
            MyBase.Refresh()
            Threading.Thread.Sleep(1000)
            AddMiddleCardsToComputer()
            middleCard.Image = Nothing
            middleCard.Text = ""
            result = True
            Me.Refresh()
            MyBase.Refresh()
        Else
            middleCard.Image = tempImage
            middleCard.Text = cardType.ToUpper
            result = True
            Me.Refresh()
            MyBase.Refresh()
            Threading.Thread.Sleep(500)
        End If




        Return result
    End Function
    Private Sub AddMiddleCardsToComputer()
        Dim length As Integer
        length = middleHands.Count

        For i As Integer = length - 1 To 0 Step -1
            Dim newCard As Card
            newCard = New Card
            newCard = middleHands.Item(i)
            newCard.Flipped = False
            middleHands.Remove(middleHands.Item(i))
            compHands.Insert(compHands.Count, newCard)
        Next i
    End Sub

    Private Sub computerPlays()


        If (FlipComputerCard(0)) Then
            player1Clicking = Not (player1Clicking)
        End If

        If compHands.Count = 0 AndAlso myHands.Count <> 0 Then
            Dim strMessage As String = "Computer won, You still have: " + CType(myHands.Count, String) + " Cards and cards in the middle: " + CType(middleHands.Count, String)
            MessageBox.Show(strMessage, "Two Player Bhaabi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            computerCard.Image = Nothing
            btnStart.Enabled = True
            Exit Sub
        ElseIf compHands.Count = 0 AndAlso myHands.Count = 0 Then
            Dim strMessage As String = "Draw Game and cards in the middle: " + CType(middleHands.Count, String)
            MessageBox.Show(strMessage, "Two Player Bhaabi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            computerCard.Image = Nothing
            playerCard.Image = Nothing
            btnStart.Enabled = True
            Exit Sub
        End If

    End Sub
    Private Function FlipComputerCard(ByVal index As Integer) As Boolean
        Dim result As Boolean
        Dim rm As New Resources.ResourceManager("twoPlayerBhaabi.images", System.Reflection.Assembly.GetExecutingAssembly)
        Dim cardType As String = ConvertSuit(compHands.Item(index).Suit) + ConvertFace(compHands.Item(index).Face)
        Dim cardFace As String = ConvertFace(compHands.Item(index).Face)
        Dim fileName As String = cardType + ".gif"
        Dim tempImage As Image = rm.GetObject(fileName)
        Dim middleCardFace As String = middleCard.Text
        If middleCardFace <> "" Then middleCardFace = middleCardFace.Substring(1, middleCardFace.Length - 1)


        Dim newCard As Card
        newCard = New Card
        newCard = compHands.Item(index)
        newCard.Flipped = True
        compHands.Remove(compHands.Item(index))
        middleHands.Insert(middleHands.Count, newCard)


        If middleCardFace = cardFace Then
            middleCard.Image = tempImage
            middleCard.Text = cardType.ToUpper
            Me.Refresh()
            MyBase.Refresh()
            Threading.Thread.Sleep(1000)
            AddMiddleCardsToPlayer()
            middleCard.Image = Nothing
            middleCard.Text = ""
            result = True
            Me.Refresh()
            MyBase.Refresh()
        Else
            middleCard.Image = tempImage
            middleCard.Text = cardType.ToUpper
            result = True
            Me.Refresh()
            MyBase.Refresh()
            Threading.Thread.Sleep(500)
        End If


        Return result
    End Function
    Private Sub AddMiddleCardsToPlayer()
        Dim length As Integer
        length = middleHands.Count

        For i As Integer = length - 1 To 0 Step -1
            Dim newCard As Card
            newCard = New Card
            newCard = middleHands.Item(i)
            newCard.Flipped = False
            middleHands.Remove(middleHands.Item(i))
            myHands.Insert(myHands.Count, newCard)
        Next i
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region     ' General Functions





End Class