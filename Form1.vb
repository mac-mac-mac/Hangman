Public Class Form1

    Dim rnd = New Random()
    Dim codeWord As String
    Dim lettersDecoded As New List(Of Boolean)
    Dim lives As Integer
    Dim wordList As New List(Of String)({"apple", "banana", "carrots"})

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        endGame()
    End Sub

    Private Sub newGame()
        codeWord = wordList(rnd.Next(0, wordList.Count))
        lettersDecoded = New List(Of Boolean)
        lives = 5
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
End Class
