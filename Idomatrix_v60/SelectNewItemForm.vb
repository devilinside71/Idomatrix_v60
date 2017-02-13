Public Class SelectNewItemForm
    Private ItemChoice As Integer

    Public ReadOnly Property ItemChoiced As Integer
        Get
            Return ItemChoice
        End Get
    End Property

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ItemChoice = 1
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ItemChoice = 2
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ItemChoice = 3
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ItemChoice = 4
        Me.Close()
    End Sub
End Class