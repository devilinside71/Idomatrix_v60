Imports System.Diagnostics

Public Class RecursiveForm
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Debug.Print(DateValue(DateTimePicker1.Value.ToString) + vbCrLf + TimeValue(DateTimePicker3.Value.ToString))
        Call WriteAppt(ComboBox1.Text, TextBox1.Text, DateValue(DateTimePicker1.Value.ToString) + TimeValue(DateTimePicker3.Value.ToString), DateValue(DateTimePicker2.Value.ToString) + TimeValue(DateTimePicker3.Value.ToString), NumericUpDown1.Value)


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.ComboBox1.SelectedIndex = 0 Then
            Me.ComboBox1.BackColor = Drawing.Color.CornflowerBlue
        End If
        If Me.ComboBox1.SelectedIndex = 1 Then
            Me.ComboBox1.BackColor = Drawing.Color.LightGreen
        End If
        If Me.ComboBox1.SelectedIndex = 2 Then
            Me.ComboBox1.BackColor = Drawing.Color.Yellow
        End If
        If Me.ComboBox1.SelectedIndex = 3 Then
            Me.ComboBox1.BackColor = Drawing.Color.Salmon
        End If
    End Sub

    Private Sub WriteAppt(catStr As String, subject As String, startDate As Date, endDate As Date, duration As Integer)
        Dim myItem As Object
        'Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient

        Do While startDate <= endDate
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
        Loop
        Windows.Forms.MessageBox.Show("Bevitel kész!", subject)
    End Sub


    Private Sub WriteTask(catStr As String, subject As String, startDate As Date, endDate As Date, duration As Integer)
        Dim myItem As Object

        Do While startDate <= endDate
            myItem = Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olTaskItem)
            'myItem.Subject = subject + " (r)"
            myItem.Subject = subject
            myItem.StartDate = startDate

            myItem.ActualWork = duration
            myItem.ReminderSet = False
            myItem.Categories = catStr
            myItem.Display(True)
            'myItem.Save()
            startDate = startDate.AddDays(1)

        Loop
        Windows.Forms.MessageBox.Show("Bevitel kész!", subject)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If Me.ComboBox2.SelectedIndex = 0 Then
            Me.ComboBox2.BackColor = Drawing.Color.CornflowerBlue
        End If
        If Me.ComboBox2.SelectedIndex = 1 Then
            Me.ComboBox2.BackColor = Drawing.Color.LightGreen
        End If
        If Me.ComboBox2.SelectedIndex = 2 Then
            Me.ComboBox2.BackColor = Drawing.Color.Yellow
        End If
        If Me.ComboBox2.SelectedIndex = 3 Then
            Me.ComboBox2.BackColor = Drawing.Color.Salmon
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call WriteTask(ComboBox2.Text, TextBox2.Text, DateValue(DateTimePicker4.Value.ToString), DateValue(DateTimePicker5.Value.ToString), NumericUpDown2.Value)
    End Sub
End Class