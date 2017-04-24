Imports System.Diagnostics

Public Class RecursiveForm
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.ComboBox7.SelectedIndex = 3
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Debug.Print(DateValue(DateTimePicker1.Value.ToString) + vbCrLf + TimeValue(DateTimePicker3.Value.ToString))
        Call WriteAppt(ComboBox1.Text, TextBox1.Text, DateValue(DateTimePicker1.Value.ToString) + TimeValue(DateTimePicker3.Value.ToString), DateValue(DateTimePicker2.Value.ToString) + TimeValue(DateTimePicker3.Value.ToString), NumericUpDown1.Value)


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.ComboBox1.SelectedIndex = 0 Then
            Me.ComboBox1.BackColor = Drawing.Color.CornflowerBlue
            Me.TextBox1.BackColor = Drawing.Color.CornflowerBlue
        End If
        If Me.ComboBox1.SelectedIndex = 1 Then
            Me.ComboBox1.BackColor = Drawing.Color.LightGreen
            Me.TextBox1.BackColor = Drawing.Color.LightGreen
        End If
        If Me.ComboBox1.SelectedIndex = 2 Then
            Me.ComboBox1.BackColor = Drawing.Color.Yellow
            Me.TextBox1.BackColor = Drawing.Color.Yellow
        End If
        If Me.ComboBox1.SelectedIndex = 3 Then
            Me.ComboBox1.BackColor = Drawing.Color.Salmon
            Me.TextBox1.BackColor = Drawing.Color.Salmon
        End If
    End Sub

    Private Sub WriteAppt(catStr As String, subject As String, startDate As Date, endDate As Date, duration As Integer)
        Dim myItem As Object
        'Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient
        Dim cCount As Integer = 0

        Do While startDate <= endDate
            Try
                myItem = Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
                'myItem.Subject = subject + " (r)"
                myItem.Subject = subject
                'myItem.Location = "Conference Room B"
                'myItem.Start = #9/24/2015 1:30:00 PM#
                myItem.Start = startDate
                myItem.Duration = duration
                'myItem.End = DateValue(startDate) + TimeValue(duration)
                'myItem.ReminderSet = False
                'myRequiredAttendee = myItem.Recipients.Add("Nate Sun")
                'myOptionalAttendee = myItem.Recipients.Add("Kevin Kennedy")
                'myResourceAttendee = myItem.Recipients.Add("Conference Room B")
                myItem.Categories = catStr
                'myItem.Display(True)
                myItem.Save()
                startDate = startDate.AddDays(1)
                cCount = cCount + 1

            Catch ex As Exception
                Debug.WriteLine("WriteAppt hiba")

            End Try
        Loop
        Windows.Forms.MessageBox.Show(cCount.ToString + " találkozó bevitele kész!", subject)
    End Sub


    Private Sub WriteTask(catStr As String, subject As String, startDate As Date, endDate As Date, duration As Integer)
        Dim myItem As Object
        Dim cCount As Integer = 0

        Do While startDate <= endDate
            Try
                myItem = Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olTaskItem)
                'myItem.Subject = subject + " (r)"
                myItem.Subject = subject
                myItem.StartDate = startDate
                myItem.DueDate = startDate
                myItem.ActualWork = duration
                myItem.ReminderSet = False
                myItem.Categories = catStr
                'myItem.Display(True)
                myItem.Save()
                startDate = startDate.AddDays(1)
                cCount = cCount + 1
            Catch ex As Exception
                Debug.WriteLine("WriteTask hiba")
            End Try


        Loop
        Windows.Forms.MessageBox.Show(cCount.ToString + " feladat bevitele kész!", subject)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If Me.ComboBox2.SelectedIndex = 0 Then
            Me.ComboBox2.BackColor = Drawing.Color.CornflowerBlue
            Me.TextBox2.BackColor = Drawing.Color.CornflowerBlue
        End If
        If Me.ComboBox2.SelectedIndex = 1 Then
            Me.ComboBox2.BackColor = Drawing.Color.LightGreen
            Me.TextBox2.BackColor = Drawing.Color.LightGreen
        End If
        If Me.ComboBox2.SelectedIndex = 2 Then
            Me.ComboBox2.BackColor = Drawing.Color.Yellow
            Me.TextBox2.BackColor = Drawing.Color.Yellow
        End If
        If Me.ComboBox2.SelectedIndex = 3 Then
            Me.ComboBox2.BackColor = Drawing.Color.Salmon
            Me.TextBox2.BackColor = Drawing.Color.Salmon
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call WriteTask(ComboBox2.Text, TextBox2.Text, DateValue(DateTimePicker4.Value.ToString), DateValue(DateTimePicker5.Value.ToString), NumericUpDown2.Value)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call WriteAppt(ComboBox3.Text, TextBox3.Text, DateValue(DateTimePicker6.Value.ToString) + TimeValue(DateTimePicker8.Value.ToString), DateValue(DateTimePicker7.Value.ToString) + TimeValue(DateTimePicker8.Value.ToString), NumericUpDown3.Value)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Me.ComboBox3.SelectedIndex = 0 Then
            Me.ComboBox3.BackColor = Drawing.Color.CornflowerBlue
            Me.TextBox3.BackColor = Drawing.Color.CornflowerBlue
        End If
        If Me.ComboBox3.SelectedIndex = 1 Then
            Me.ComboBox3.BackColor = Drawing.Color.LightGreen
            Me.TextBox3.BackColor = Drawing.Color.LightGreen
        End If
        If Me.ComboBox3.SelectedIndex = 2 Then
            Me.ComboBox3.BackColor = Drawing.Color.Yellow
            Me.TextBox3.BackColor = Drawing.Color.Yellow
        End If
        If Me.ComboBox3.SelectedIndex = 3 Then
            Me.ComboBox3.BackColor = Drawing.Color.Salmon
            Me.TextBox3.BackColor = Drawing.Color.Salmon
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If Me.ComboBox4.SelectedIndex = 0 Then
            Me.ComboBox4.BackColor = Drawing.Color.CornflowerBlue
            Me.TextBox4.BackColor = Drawing.Color.CornflowerBlue
        End If
        If Me.ComboBox4.SelectedIndex = 1 Then
            Me.ComboBox4.BackColor = Drawing.Color.LightGreen
            Me.TextBox4.BackColor = Drawing.Color.LightGreen
        End If
        If Me.ComboBox4.SelectedIndex = 2 Then
            Me.ComboBox4.BackColor = Drawing.Color.Yellow
            Me.TextBox4.BackColor = Drawing.Color.Yellow
        End If
        If Me.ComboBox4.SelectedIndex = 3 Then
            Me.ComboBox4.BackColor = Drawing.Color.Salmon
            Me.TextBox4.BackColor = Drawing.Color.Salmon
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call WriteTask(ComboBox4.Text, TextBox4.Text, DateValue(DateTimePicker9.Value.ToString), DateValue(DateTimePicker10.Value.ToString), NumericUpDown4.Value)
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If Me.ComboBox5.SelectedIndex = 0 Then
            Me.ComboBox5.BackColor = Drawing.Color.CornflowerBlue
            Me.TextBox5.BackColor = Drawing.Color.CornflowerBlue
        End If
        If Me.ComboBox5.SelectedIndex = 1 Then
            Me.ComboBox5.BackColor = Drawing.Color.LightGreen
            Me.TextBox5.BackColor = Drawing.Color.LightGreen
        End If
        If Me.ComboBox5.SelectedIndex = 2 Then
            Me.ComboBox5.BackColor = Drawing.Color.Yellow
            Me.TextBox5.BackColor = Drawing.Color.Yellow
        End If
        If Me.ComboBox5.SelectedIndex = 3 Then
            Me.ComboBox5.BackColor = Drawing.Color.Salmon
            Me.TextBox5.BackColor = Drawing.Color.Salmon
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call WriteTask(ComboBox5.Text, TextBox5.Text, DateValue(DateTimePicker11.Value.ToString), DateValue(DateTimePicker12.Value.ToString), NumericUpDown5.Value)
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        If Me.ComboBox6.SelectedIndex = 0 Then
            Me.ComboBox6.BackColor = Drawing.Color.CornflowerBlue
            Me.TextBox6.BackColor = Drawing.Color.CornflowerBlue
        End If
        If Me.ComboBox6.SelectedIndex = 1 Then
            Me.ComboBox6.BackColor = Drawing.Color.LightGreen
            Me.TextBox6.BackColor = Drawing.Color.LightGreen
        End If
        If Me.ComboBox6.SelectedIndex = 2 Then
            Me.ComboBox6.BackColor = Drawing.Color.Yellow
            Me.TextBox6.BackColor = Drawing.Color.Yellow
        End If
        If Me.ComboBox6.SelectedIndex = 3 Then
            Me.ComboBox6.BackColor = Drawing.Color.Salmon
            Me.TextBox6.BackColor = Drawing.Color.Salmon
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call WriteTask(ComboBox6.Text, TextBox6.Text, DateValue(DateTimePicker13.Value.ToString), DateValue(DateTimePicker14.Value.ToString), NumericUpDown6.Value)
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        If Me.ComboBox7.SelectedIndex = 0 Then
            Me.ComboBox7.BackColor = Drawing.Color.CornflowerBlue
            Me.TextBox7.BackColor = Drawing.Color.CornflowerBlue
        End If
        If Me.ComboBox7.SelectedIndex = 1 Then
            Me.ComboBox7.BackColor = Drawing.Color.LightGreen
            Me.TextBox7.BackColor = Drawing.Color.LightGreen
        End If
        If Me.ComboBox7.SelectedIndex = 2 Then
            Me.ComboBox7.BackColor = Drawing.Color.Yellow
            Me.TextBox7.BackColor = Drawing.Color.Yellow
        End If
        If Me.ComboBox7.SelectedIndex = 3 Then
            Me.ComboBox7.BackColor = Drawing.Color.Salmon
            Me.TextBox7.BackColor = Drawing.Color.Salmon
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call WriteTask(ComboBox7.Text, TextBox7.Text, DateValue(DateTimePicker16.Value.ToString), DateValue(DateTimePicker15.Value.ToString), NumericUpDown7.Value)
    End Sub
End Class