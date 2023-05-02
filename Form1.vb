Public Class Form1

    Dim rnd = New Random()
    Dim codeWord As String
    Dim lettersDecoded As New List(Of Boolean)
    Dim lives As Integer
    Dim wordList As New List(Of String)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        endGame()
        radColors.Checked = False
        radShapes.Checked = False
        radFruits.Checked = False
        radSpace.Checked = False
        radAll.Checked = True
    End Sub

    Private Sub newGame()
        getWordList()
        codeWord = wordList(rnd.Next(0, wordList.Count))
        lettersDecoded = New List(Of Boolean)
        lives = 8
        lblLivesNum.Text = lives
        For i As Integer = 1 To Len(codeWord)
            lettersDecoded.Add(False)
        Next
        printCodeWord()
        btnA.Enabled = True
        btnB.Enabled = True
        btnC.Enabled = True
        btnD.Enabled = True
        btnE.Enabled = True
        btnF.Enabled = True
        btnG.Enabled = True
        btnH.Enabled = True
        btnI.Enabled = True
        btnJ.Enabled = True
        btnK.Enabled = True
        btnL.Enabled = True
        btnM.Enabled = True
        btnN.Enabled = True
        btnO.Enabled = True
        btnP.Enabled = True
        btnQ.Enabled = True
        btnR.Enabled = True
        btnS.Enabled = True
        btnT.Enabled = True
        btnU.Enabled = True
        btnV.Enabled = True
        btnW.Enabled = True
        btnX.Enabled = True
        btnY.Enabled = True
        btnZ.Enabled = True
    End Sub

    Private Sub getWordList()
        If radColors.Checked Then
            wordList = New List(Of String)({"black", "blue", "brown", "cyan", "gold", "green", "indigo", "magenta", "maroon", "orange", "peach", "pink", "purple", "red", "scarlet", "silver", "tan", "teal", "turquoise", "violet", "white", "yellow"})
        ElseIf radShapes.Checked Then
            wordList = New List(Of String)({"circle", "cone", "cube", "cylinder", "decagon", "ellipse", "heart", "heptagon", "hexagon", "kite", "nonagon", "octagon", "oval", "pentagon", "polygon", "pyramid", "rectangle", "rhombus", "sphere", "trapezoid", "triangle"})
        ElseIf radFruits.Checked Then
            wordList = New List(Of String)({"apple", "apricot", "avocado", "banana", "blackberry", "blueberry", "cantaloupe", "cherry", "coconut", "cranberry", "date", "grape", "grapefruit", "guava", "honeydew", "kiwi", "lemon", "lime", "mango", "melon", "orange", "papaya", "peach", "pear", "pineapple", "plum", "prune", "raisin", "raspberry", "strawberry", "tangerine", "watermelon"})
        ElseIf radSpace.Checked Then
            wordList = New List(Of String)({"asteroid", "astronaut", "astronomy", "comet", "crater", "earth", "eclipse", "galaxy", "gamma", "gravity", "helium", "hydrogen", "jupiter", "lightyear", "lunar", "mars", "mercury", "meteor", "meteorite", "meteoroid", "moon", "nebula", "neptune", "nova", "orbit", "planet", "pluto", "probe", "quasar", "radiation", "rocket", "satellite", "saturn", "solar", "space", "star", "sun", "supernova", "telescope", "universe", "uranus", "vacuum", "venus", "wormhole"})
        ElseIf radAll.Checked Then
            wordList = New List(Of String)({"black", "blue", "brown", "cyan", "gold", "green", "indigo", "magenta", "maroon", "orange", "peach", "pink", "purple", "red", "scarlet", "silver", "tan", "teal", "turquoise", "violet", "white", "yellow", "circle", "cone", "cube", "cylinder", "decagon", "ellipse", "heart", "heptagon", "hexagon", "kite", "nonagon", "octagon", "oval", "pentagon", "polygon", "pyramid", "rectangle", "rhombus", "sphere", "trapezoid", "triangle", "apple", "apricot", "avocado", "banana", "blackberry", "blueberry", "cantaloupe", "cherry", "coconut", "cranberry", "date", "grape", "grapefruit", "guava", "honeydew", "kiwi", "lemon", "lime", "mango", "melon", "orange", "papaya", "peach", "pear", "pineapple", "plum", "prune", "raisin", "raspberry", "strawberry", "tangerine", "watermelon", "asteroid", "astronaut", "astronomy", "comet", "crater", "earth", "eclipse", "galaxy", "gamma", "gravity", "helium", "hydrogen", "jupiter", "lightyear", "lunar", "mars", "mercury", "meteor", "meteorite", "meteoroid", "moon", "nebula", "neptune", "nova", "orbit", "planet", "pluto", "probe", "quasar", "radiation", "rocket", "satellite", "saturn", "solar", "space", "star", "sun", "supernova", "telescope", "universe", "uranus", "vacuum", "venus", "wormhole"})
        End If
    End Sub

    Private Sub printCodeWord()
        Dim currentWord As String = ""
        For i As Integer = 1 To Len(codeWord)
            If (lettersDecoded(i - 1)) Then
                currentWord = currentWord + codeWord(i - 1) + " "
            Else
                currentWord = currentWord + "_ "
            End If
        Next
        lblWord.Text = currentWord
    End Sub

    Private Sub updateGame()
        lblLivesNum.Text = lives

        If (lives < 1) Then
            Dim msg As String
            msg = "YOU LOSE - Try Again?" + vbNewLine + vbNewLine
            msg = msg + "The code word was: [" + codeWord + "]"
            MsgBox(msg)
            endGame()
        End If

        If (lettersDecoded.Contains(False) = False) Then
            Dim msg As String
            msg = "YOU WIN - Congratulations!" + vbNewLine + vbNewLine
            msg = msg + "The code word was: [" + codeWord + "]"
            MsgBox(msg)
            endGame()
        End If
    End Sub

    Private Sub endGame()
        btnA.Enabled = False
        btnB.Enabled = False
        btnC.Enabled = False
        btnD.Enabled = False
        btnE.Enabled = False
        btnF.Enabled = False
        btnG.Enabled = False
        btnH.Enabled = False
        btnI.Enabled = False
        btnJ.Enabled = False
        btnK.Enabled = False
        btnL.Enabled = False
        btnM.Enabled = False
        btnN.Enabled = False
        btnO.Enabled = False
        btnP.Enabled = False
        btnQ.Enabled = False
        btnR.Enabled = False
        btnS.Enabled = False
        btnT.Enabled = False
        btnU.Enabled = False
        btnV.Enabled = False
        btnW.Enabled = False
        btnX.Enabled = False
        btnY.Enabled = False
        btnZ.Enabled = False
        lblLivesNum.Text = ""
        lblWord.Text = "New Game?"
    End Sub

    Private Sub btnNewGame_Click(sender As Object, e As EventArgs) Handles btnNewGame.Click
        newGame()
    End Sub

    Private Sub btnA_Click(sender As Object, e As EventArgs) Handles btnA.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "a") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnA.Enabled = False
    End Sub

    Private Sub btnB_Click(sender As Object, e As EventArgs) Handles btnB.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "b") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnB.Enabled = False
    End Sub

    Private Sub btnC_Click(sender As Object, e As EventArgs) Handles btnC.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "c") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnC.Enabled = False
    End Sub

    Private Sub btnD_Click(sender As Object, e As EventArgs) Handles btnD.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "d") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnD.Enabled = False
    End Sub

    Private Sub btnE_Click(sender As Object, e As EventArgs) Handles btnE.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "e") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnE.Enabled = False
    End Sub

    Private Sub btnF_Click(sender As Object, e As EventArgs) Handles btnF.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "f") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnF.Enabled = False
    End Sub

    Private Sub btnG_Click(sender As Object, e As EventArgs) Handles btnG.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "g") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnG.Enabled = False
    End Sub

    Private Sub btnH_Click(sender As Object, e As EventArgs) Handles btnH.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "h") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnH.Enabled = False
    End Sub

    Private Sub btnI_Click(sender As Object, e As EventArgs) Handles btnI.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "i") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnI.Enabled = False
    End Sub

    Private Sub btnJ_Click(sender As Object, e As EventArgs) Handles btnJ.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "j") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnJ.Enabled = False
    End Sub

    Private Sub btnK_Click(sender As Object, e As EventArgs) Handles btnK.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "k") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnK.Enabled = False
    End Sub

    Private Sub btnL_Click(sender As Object, e As EventArgs) Handles btnL.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "l") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnL.Enabled = False
    End Sub

    Private Sub btnM_Click(sender As Object, e As EventArgs) Handles btnM.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "m") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnM.Enabled = False
    End Sub

    Private Sub btnN_Click(sender As Object, e As EventArgs) Handles btnN.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "n") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnN.Enabled = False
    End Sub

    Private Sub btnO_Click(sender As Object, e As EventArgs) Handles btnO.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "o") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnO.Enabled = False
    End Sub

    Private Sub btnP_Click(sender As Object, e As EventArgs) Handles btnP.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "p") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnP.Enabled = False
    End Sub

    Private Sub btnQ_Click(sender As Object, e As EventArgs) Handles btnQ.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "q") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnQ.Enabled = False
    End Sub

    Private Sub btnR_Click(sender As Object, e As EventArgs) Handles btnR.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "r") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnR.Enabled = False
    End Sub

    Private Sub btnS_Click(sender As Object, e As EventArgs) Handles btnS.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "s") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnS.Enabled = False
    End Sub

    Private Sub btnT_Click(sender As Object, e As EventArgs) Handles btnT.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "t") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnT.Enabled = False
    End Sub

    Private Sub btnU_Click(sender As Object, e As EventArgs) Handles btnU.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "u") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnU.Enabled = False
    End Sub

    Private Sub btnV_Click(sender As Object, e As EventArgs) Handles btnV.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "v") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnV.Enabled = False
    End Sub

    Private Sub btnW_Click(sender As Object, e As EventArgs) Handles btnW.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "w") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnW.Enabled = False
    End Sub

    Private Sub btnX_Click(sender As Object, e As EventArgs) Handles btnX.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "x") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnX.Enabled = False
    End Sub

    Private Sub btnY_Click(sender As Object, e As EventArgs) Handles btnY.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "y") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnY.Enabled = False
    End Sub

    Private Sub btnZ_Click(sender As Object, e As EventArgs) Handles btnZ.Click
        Dim letterUsed As Boolean = False
        For i As Integer = 1 To Len(codeWord)
            If (codeWord(i - 1) = "z") Then
                lettersDecoded(i - 1) = True
                letterUsed = True
            End If
        Next
        If letterUsed = False Then
            lives = lives - 1
        End If
        printCodeWord()
        updateGame()
        btnZ.Enabled = False
    End Sub

    Private Sub btnHowToPlay_Click(sender As Object, e As EventArgs) Handles btnHowToPlay.Click
        Dim msg As String
        msg = "HOW TO PLAY:" + vbNewLine + vbNewLine
        msg = msg + "Guess letters to fill in the secret code word, but you only have 8 lives." + vbNewLine + vbNewLine
        msg = msg + "If you guess a letter that is in the code word, all instances of that letter will be revealed in the code word. If you guess a letter that is not in the code word, you lose one of your lives." + vbNewLine + vbNewLine
        msg = msg + "If you lose all your lives before all letters in the code word are revealed, you lose the game! If you reveal all the letters in the code word without losing all your lives, you win the game!" + vbNewLine + vbNewLine
        msg = msg + "You can use the radio buttons to limit the code words to a specific word category (or from all categories). This will not affect an in-progress game, but will take effect when you start a new game." + vbNewLine + vbNewLine
        MsgBox(msg)
    End Sub
End Class
