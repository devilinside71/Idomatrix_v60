Imports System.Diagnostics
Imports System.Windows.Forms

Public Class IGMPane
    Private Suspended As Boolean
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Suspended = True
        DateTimePicker1.Value = DateTime.Today.Date
        DateTimePicker2.Value = DateTime.Today.Date
        Suspended = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Suspended = True
        DateTimePicker1.Value = DateTime.Today.AddDays(1).Date
        DateTimePicker2.Value = DateTime.Today.AddDays(1).Date
        Suspended = False
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Suspended = True
        DateTimePicker1.Value = Today.AddDays((Today.DayOfWeek - DayOfWeek.Monday) * -1).Date
        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(4).Date
        Suspended = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call RefreshData()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MessageBox.Show("RIPORT")
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If Not Suspended Then
            If DateTimePicker2.Value < DateTimePicker1.Value Then
                DateTimePicker1.Value = DateTimePicker2.Value
            End If
            Call RefreshData()
        End If
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        If Not Suspended Then
            If DateTimePicker2.Value < DateTimePicker1.Value Then
                DateTimePicker1.Value = DateTimePicker2.Value
            End If
            Call RefreshData()
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

    End Sub
    Private Sub IGMPane_Load(sender As Object, e As EventArgs) Handles Me.Load
        Debug.Print("Pane betöltése")
    End Sub
    Public Sub RefreshData()

    End Sub
End Class
