' Name:         Letter Guess Project
' Purpose:      Guess a random letter.
' Programmer:   <celeste delgado> on <4/12/23>

Option Explicit On
Option Strict On
Option Infer Off

Public Class frmMain
    ' Class-level variable.
    Private strRandLetter As String
    Private Sub btnNewGame_Click(sender As Object, e As EventArgs) Handles btnNewGame.Click
        ' Select a random letter and prepare for a new game.

        Const strALPHABET As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim randGem As New Random
        Dim intRandNum As Integer
        ' Generate a random number and use it to select a letter.
        intRandNum = randGem.Next(0, 26)
        strRandLetter = strALPHABET(intRandNum)
        ' Clear lblGuesses, enable btnCheck, send focus to txtGuess.
        lblGuesses.Text = String.Empty
        btnCheck.Enabled = True
        txtGuess.Focus()
    End Sub

    Private Sub btnCheck_Click(sender As Object, e As EventArgs) Handles btnCheck.Click
        ' Determine whether the user guessed the random letter.

        Dim strGuess As String

        strGuess = txtGuess.Text.Trim.ToUpper
        ' Display guess in lblGuesses.
        lblGuesses.Text = lblGuesses.Text & " " & strGuess
        If strGuess = strRandLetter Then
            MessageBox.Show("You guessed the correct letter: " & strGuess,
                            "Guess a Letter", MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            btnCheck.Enabled = False
        Else
            MessageBox.Show("Guess again!", "Guess a Letter",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        txtGuess.Text = String.Empty
        txtGuess.Focus()
    End Sub

    Private Sub txtGuess_Enter(sender As Object, e As EventArgs) Handles txtGuess.Enter
        txtGuess.SelectAll()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
