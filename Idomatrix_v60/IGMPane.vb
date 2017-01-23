Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class IGMPane
    Private Suspended As Boolean
    Private resultSF As Double = 0
    Private resultSNF As Double = 0
    Private resultNSF As Double = 0
    Private resultNSNF As Double = 0
    Private resultSum As Double = 0
    Private resultEvalSF As Integer = 0
    Private resultEvalSNF As Integer = 0
    Private resultEvalNSF As Integer = 0
    Private resultEvalNSNF As Integer = 0

    Private resultT_SF As Double = 0
    Private resultT_SNF As Double = 0
    Private resultT_NSF As Double = 0
    Private resultT_NSNF As Double = 0
    Private resultT_sum As Double = 0
    Private resultT_EvalSF As Double = 0
    Private resultT_EvalNSF As Double = 0
    Private resultT_EvalSNF As Double = 0
    Private resultT_EvalNSNF As Double = 0
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
        Call RefreshData()
    End Sub
    Public Sub RefreshData()
        Call ClearLists()
        Call SetEmailTasksInRange()
    End Sub
    Private Sub ClearLists()
        Me.ListView1.Items.Clear()
        Me.ListView2.Items.Clear()
        Me.ListView3.Items.Clear()
        Me.ListView4.Items.Clear()
        Me.ListView5.Items.Clear()
        Me.ListView6.Items.Clear()
    End Sub
    Private Sub SetEmailTasksInRange()
        Dim resultMin As String
        Dim resultMin2 As String
        Dim calFolder As Outlook.Folder = TryCast(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox), Outlook.Folder)
        Dim startTime As DateTime = DateTimePicker1.Value.Date
        'A nullaóra miatt 1 nap eltérés kell
        Dim endTime As DateTime = DateTimePicker2.Value.AddDays(1).Date
        Dim rangeAppts As Outlook.Items = GetEmailTasksInRange(calFolder, startTime, endTime)
        If rangeAppts IsNot Nothing Then
            For Each appt As Outlook.MailItem In rangeAppts
                'Debug.WriteLine("EmailTask Subject: " + appt.Subject)

                If (String.IsNullOrEmpty(appt.Categories) = False) Then
                    If appt.Categories.Contains("@Sürgős - Fontos") Then
                        With Me.ListView1.Items.Add("E")
                            .SubItems.Add(appt.Subject)
                            .SubItems.Add(Format(appt.CreationTime, "yyyy/MM/dd"))
                            .SubItems.Add(appt.EntryID)
                            'resultMin = Me.NumericUpDown2.Value.ToString
                            'resultMin2 = Me.NumericUpDown2.Value.ToString
                            Try
                                Dim tervTeny As String() = Split(appt.Companies, "@")
                                If tervTeny(0) <> "" Then
                                    resultMin = tervTeny(0)
                                Else
                                    resultMin = "0"
                                End If
                                Try
                                    If tervTeny(1) <> "" Then
                                        resultMin2 = tervTeny(1)
                                    Else
                                        resultMin2 = "0"
                                    End If
                                Catch ex As Exception
                                    resultMin2 = "0"
                                End Try
                            Catch ex As Exception
                                resultMin = "0"
                            End Try
                            .SubItems.Add(resultMin)
                            resultT_SF = resultT_SF + CDbl(resultMin2)
                            .SubItems.Add(resultMin2)
                            resultSF = resultSF + CDbl(resultMin)
                            .SubItems.Add("")
                        End With
                    End If
                    If appt.Categories.Contains("@Sürgős - Nem fontos") Then
                        With Me.ListView3.Items.Add("E")
                            .SubItems.Add(appt.Subject)
                            .SubItems.Add(Format(appt.CreationTime, "yyyy/MM/dd"))
                            .SubItems.Add(appt.EntryID)
                            'resultMin = Me.NumericUpDown3.Value.ToString
                            'resultMin2 = Me.NumericUpDown2.Value.ToString
                            Try
                                Dim tervTeny As String() = Split(appt.Companies, "@")
                                If tervTeny(0) <> "" Then
                                    resultMin = tervTeny(0)
                                Else
                                    resultMin = "0"
                                End If
                                Try
                                    If tervTeny(1) <> "" Then
                                        resultMin2 = tervTeny(1)
                                    Else
                                        resultMin2 = "0"
                                    End If
                                Catch ex As Exception
                                    resultMin2 = "0"
                                End Try
                            Catch ex As Exception
                                resultMin = "0"
                            End Try
                            .SubItems.Add(resultMin)
                            resultT_SNF = resultT_SNF + CDbl(resultMin2)
                            .SubItems.Add(resultMin2)
                            resultSNF = resultSNF + CDbl(resultMin)
                            .SubItems.Add("")
                        End With
                    End If
                    If appt.Categories.Contains("@Nem sürgős - Fontos") Then
                        With Me.ListView2.Items.Add("E")
                            .SubItems.Add(appt.Subject)
                            .SubItems.Add(Format(appt.CreationTime, "yyyy/MM/dd"))
                            .SubItems.Add(appt.EntryID)
                            'resultMin = Me.NumericUpDown1.Value.ToString
                            'resultMin2 = Me.NumericUpDown2.Value.ToString
                            Try
                                Dim tervTeny As String() = Split(appt.Companies, "@")
                                If tervTeny(0) <> "" Then
                                    resultMin = tervTeny(0)
                                Else
                                    resultMin = "0"
                                End If
                                Try
                                    If tervTeny(1) <> "" Then
                                        resultMin2 = tervTeny(1)
                                    Else
                                        resultMin2 = "0"
                                    End If
                                Catch ex As Exception
                                    resultMin2 = "0"
                                End Try
                            Catch ex As Exception
                                resultMin = "0"
                            End Try
                            .SubItems.Add(resultMin)
                            resultT_NSF = resultT_NSF + CDbl(resultMin2)
                            .SubItems.Add(resultMin2)
                            resultNSF = resultNSF + CDbl(resultMin)
                            .SubItems.Add("")
                        End With
                    End If
                    If appt.Categories.Contains("@Nem sürgős - Nem fontos") Then
                        With Me.ListView4.Items.Add("E")
                            .SubItems.Add(appt.Subject)
                            .SubItems.Add(Format(appt.CreationTime, "yyyy/MM/dd"))
                            .SubItems.Add(appt.EntryID)
                            'resultMin = Me.NumericUpDown4.Value.ToString
                            'resultMin2 = Me.NumericUpDown2.Value.ToString
                            Try
                                Dim tervTeny As String() = Split(appt.Companies, "@")
                                If tervTeny(0) <> "" Then
                                    resultMin = tervTeny(0)
                                Else
                                    resultMin = "0"
                                End If
                                Try
                                    If tervTeny(1) <> "" Then
                                        resultMin2 = tervTeny(1)
                                    Else
                                        resultMin2 = "0"
                                    End If
                                Catch ex As Exception
                                    resultMin2 = "0"
                                End Try
                            Catch ex As Exception
                                resultMin = "0"
                            End Try
                            .SubItems.Add(resultMin)
                            resultT_NSNF = resultT_NSNF + CDbl(resultMin2)
                            .SubItems.Add(resultMin2)
                            resultNSNF = resultNSNF + CDbl(resultMin)
                            .SubItems.Add("")
                        End With
                    End If
                End If

            Next
            If rangeAppts IsNot Nothing Then Marshal.ReleaseComObject(rangeAppts)

        End If
    End Sub
    Private Function GetEmailTasksInRange(folder As Outlook.Folder, startTime As DateTime, endTime As DateTime) As Outlook.Items
        'Dim filter As String = "[DueDate] >= '" + Format(startTime, "yyyy/MM/dd") + "' AND [DueDate] <= '" + Format(endTime, "yyyy/MM/dd") + "'"
        'Dim filter As String = "[DueDate] >= '" + Format(startTime, "yyyy/MM/dd") + "'"
        'Dim filter As String = "[CreationTime] >= '" + Format(startTime, "yyyy/MM/dd") + "' AND [CreationTime] <= '" + Format(endTime, "yyyy/MM/dd") + "'"
        'Dim filter As String = "[CreationTime] >= '" + Format(startTime, "yyyy/MM/dd") + "' AND [CreationTime] <= '" + Format(endTime, "yyyy/MM/dd") + "' OR [IsMarkedAsTask] = True"
        'Dim filter As String = "[CreationTime] >= '" + Format(startTime, "yyyy/MM/dd")
        'Dim filter As String = "[CreationTime] >= '" + startTime.ToString("g") + "' AND [CreationTime] <= '" + endTime.ToString("g") + "'"
        'Dim filter As String = "[IsMarkedAsTask] = True"
        Dim filter As String = "[TaskDueDate] >= '" + Format(startTime.AddDays(-1), "yyyy/MM/dd") + "'"
        'Debug.WriteLine(filter)
        Try
            Dim calItems As Outlook.Items = folder.Items
            calItems.IncludeRecurrences = True
            calItems.Sort("[CreationTime]", Type.Missing)

            Dim restrictItems As Outlook.Items = calItems.Restrict(filter)
            If restrictItems.Count > 0 Then
                Return restrictItems
            Else
                Return Nothing
            End If
        Catch
            Return Nothing
        End Try
    End Function
End Class
