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
            Debug.Print(DateTimePicker1.Value)
            Debug.Print(DateTimePicker2.Value)
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
        Suspended = True
        DateTimePicker1.Value = DateTime.Today.Date
        DateTimePicker2.Value = DateTime.Today.Date
        Suspended = False
        Call RefreshData()
    End Sub
    Public Sub RefreshData()
        Call ClearLists()
        Call SetEmailTasksInRange()
        Call SetTasksInRange()
        Call SetAppointmentsInRange()
    End Sub
    Private Sub ClearLists()
        Me.ListView1.Items.Clear()
        Me.ListView2.Items.Clear()
        Me.ListView3.Items.Clear()
        Me.ListView4.Items.Clear()
        Me.ListView5.Items.Clear()
        Me.ListView6.Items.Clear()
    End Sub
    Private Function GetMinutes(startTime As DateTime, endTime As DateTime)
        Dim elapsedTime As TimeSpan = endTime.Subtract(startTime)
        Dim elapsedMinutesText As String = elapsedTime.TotalMinutes.ToString()
        Return elapsedMinutesText
    End Function
#Region "Read Outlook items into Listviews"
    Private Sub SetAppointmentsInRange()
        Dim resultMin As String
        Dim calFolder As Outlook.Folder = TryCast(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar), Outlook.Folder)
        Dim startTime As DateTime = DateTimePicker1.Value
        'A nullaóra miatt 1 nap eltérés kell
        Dim endTime As DateTime = DateTimePicker2.Value.AddDays(1)
        Dim rangeAppts As Outlook.Items = GetAppointmentsInRange(calFolder, startTime, endTime)
        If rangeAppts IsNot Nothing Then
            For Each appt As Outlook.AppointmentItem In rangeAppts
                'Debug.WriteLine("Subject: " + appt.Subject + " Start: " + appt.Start.ToString("g"))

                If (String.IsNullOrEmpty(appt.Categories) = False) Then
                    If appt.Categories.Contains("@Sürgős - Fontos") Then
                        With Me.ListView1.Items.Add("T")
                            .SubItems.Add(appt.Subject)
                            .SubItems.Add(Format(appt.Start, "yyyy/MM/dd"))
                            .SubItems.Add(appt.EntryID)
                            resultMin = GetMinutes(appt.Start, appt.End)
                            .SubItems.Add(resultMin)
                            resultSF = resultSF + CDbl(resultMin)
                            Dim tervTeny As String() = Split(appt.Companies, "@")
                            .SubItems.Add(tervTeny(1))
                            Try
                                resultT_SF = resultT_SF + CDbl(tervTeny(1))
                            Catch ex As Exception
                            End Try
                            .SubItems.Add("")
                        End With
                    End If
                    If appt.Categories.Contains("@Sürgős - Nem fontos") Then
                        With Me.ListView3.Items.Add("T")
                            .SubItems.Add(appt.Subject)
                            .SubItems.Add(Format(appt.Start, "yyyy/MM/dd"))
                            .SubItems.Add(appt.EntryID)
                            resultMin = GetMinutes(appt.Start, appt.End)
                            .SubItems.Add(resultMin)
                            resultSNF = resultSNF + CDbl(resultMin)
                            Dim tervTeny As String() = Split(appt.Companies, "@")
                            .SubItems.Add(tervTeny(1))
                            Try
                                resultT_SNF = resultT_SNF + CDbl(tervTeny(1))
                            Catch ex As Exception
                            End Try
                            .SubItems.Add("")
                        End With
                    End If
                    If appt.Categories.Contains("@Nem sürgős - Fontos") Then
                        With Me.ListView2.Items.Add("T")
                            .SubItems.Add(appt.Subject)
                            .SubItems.Add(Format(appt.Start, "yyyy/MM/dd"))
                            .SubItems.Add(appt.EntryID)
                            resultMin = GetMinutes(appt.Start, appt.End)
                            .SubItems.Add(resultMin)
                            resultNSF = resultNSF + CDbl(resultMin)
                            Dim tervTeny As String() = Split(appt.Companies, "@")
                            .SubItems.Add(tervTeny(1))
                            Try
                                resultT_NSF = resultT_NSF + CDbl(tervTeny(1))
                            Catch ex As Exception
                            End Try
                            .SubItems.Add("")
                        End With
                    End If
                    If appt.Categories.Contains("@Nem sürgős - Nem fontos") Then
                        With Me.ListView4.Items.Add("T")
                            .SubItems.Add(appt.Subject)
                            .SubItems.Add(Format(appt.Start, "yyyy/MM/dd"))
                            .SubItems.Add(appt.EntryID)
                            resultMin = GetMinutes(appt.Start, appt.End)
                            .SubItems.Add(resultMin)
                            resultNSNF = resultNSNF + CDbl(resultMin)
                            Dim tervTeny As String() = Split(appt.Companies, "@")
                            .SubItems.Add(tervTeny(1))
                            Try
                                resultT_NSNF = resultT_NSNF + CDbl(tervTeny(1))
                            Catch ex As Exception
                            End Try
                            .SubItems.Add("")
                        End With
                    End If
                End If
            Next
            If rangeAppts IsNot Nothing Then Marshal.ReleaseComObject(rangeAppts)
        End If
    End Sub
    ''' <summary>
    ''' Get appointments in date range.
    ''' </summary>
    ''' <param name="folder"></param>
    ''' <param name="startTime"></param>
    ''' <param name="endTime"></param>
    ''' <returns>Outlook.Items</returns>
    Private Function GetAppointmentsInRange(folder As Outlook.Folder, startTime As DateTime, endTime As DateTime) As Outlook.Items
        Dim filter As String = "[Start] >= '" + startTime.ToString("g") + "' AND [End] <= '" + endTime.ToString("g") + "'"
        'Debug.WriteLine(filter)
        Try
            Dim calItems As Outlook.Items = folder.Items
            calItems.IncludeRecurrences = True
            calItems.Sort("[Start]", Type.Missing)
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
    Private Sub SetTasksInRange()
        Dim resultMin As String
        Dim resultMinInt As Integer
        Dim maxYear As Long = 2040
        Dim calFolder As Outlook.Folder = TryCast(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks), Outlook.Folder)
        Dim startTime As DateTime = DateTimePicker1.Value.Date
        'A nullaóra miatt 1 nap eltérés kell
        Dim endTime As DateTime = DateTimePicker2.Value.AddDays(1).Date
        Dim rangeAppts As Outlook.Items = GetTasksInRange(calFolder, startTime, endTime)
        If rangeAppts IsNot Nothing Then
            For Each appt As Outlook.TaskItem In rangeAppts
                'Debug.WriteLine("Task Subject: " + appt.Subject + " Start: " + appt.StartDate.ToString("g"))

                If (String.IsNullOrEmpty(appt.Categories) = False) Then
                    If appt.Categories.Contains("@Sürgős - Fontos") Then
                        With Me.ListView1.Items.Add("F")
                            .SubItems.Add(appt.Subject)
                            Dim evStr = Format(appt.DueDate, "yyyy/MM/dd")
                            If CLng(Strings.Left(evStr, 4)) > maxYear Then
                                evStr = "Nincs"
                            End If
                            .SubItems.Add(evStr)
                            .SubItems.Add(appt.EntryID)
                            'resultMin = Me.NumericUpDown2.Value.ToString
                            resultMinInt = appt.ActualWork
                            If resultMinInt < 1 Then
                                resultMinInt = 20
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)
                            resultSF = resultSF + resultMinInt
                            resultMinInt = appt.TotalWork
                            If resultMinInt < 1 Then
                                resultMinInt = 0
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)
                            resultT_SF = resultT_SF + resultMinInt

                            .SubItems.Add(appt.PercentComplete.ToString + "%")
                        End With
                    End If
                    If appt.Categories.Contains("@Sürgős - Nem fontos") Then
                        With Me.ListView3.Items.Add("F")
                            .SubItems.Add(appt.Subject)
                            Dim evStr = Format(appt.DueDate, "yyyy/MM/dd")
                            If CLng(Strings.Left(evStr, 4)) > maxYear Then
                                evStr = "Nincs"
                            End If
                            .SubItems.Add(evStr)
                            .SubItems.Add(appt.EntryID)
                            'resultMin = Me.NumericUpDown2.Value.ToString
                            resultMinInt = appt.ActualWork
                            If resultMinInt < 1 Then
                                resultMinInt = 15
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)
                            resultSNF = resultSNF + resultMinInt
                            resultMinInt = appt.TotalWork
                            If resultMinInt < 1 Then
                                resultMinInt = 0
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)
                            resultT_SNF = resultT_SNF + resultMinInt
                            .SubItems.Add(appt.PercentComplete.ToString + "%")
                        End With
                    End If
                    If appt.Categories.Contains("@Nem sürgős - Fontos") Then
                        With Me.ListView2.Items.Add("F")
                            .SubItems.Add(appt.Subject)
                            Dim evStr = Format(appt.DueDate, "yyyy/MM/dd")
                            If CLng(Strings.Left(evStr, 4)) > maxYear Then
                                evStr = "Nincs"
                            End If
                            .SubItems.Add(evStr)
                            .SubItems.Add(appt.EntryID)
                            'resultMin = Me.NumericUpDown2.Value.ToString
                            resultMinInt = appt.ActualWork
                            If resultMinInt < 1 Then
                                resultMinInt = 25
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)
                            resultNSF = resultNSF + resultMinInt
                            resultMinInt = appt.TotalWork
                            If resultMinInt < 1 Then
                                resultMinInt = 0
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)
                            resultT_NSF = resultT_NSF + resultMinInt
                            .SubItems.Add(appt.PercentComplete.ToString + "%")
                        End With
                    End If
                    If appt.Categories.Contains("@Nem sürgős - Nem fontos") Then
                        With Me.ListView4.Items.Add("F")
                            .SubItems.Add(appt.Subject)
                            Dim evStr = Format(appt.DueDate, "yyyy/MM/dd")
                            If CLng(Strings.Left(evStr, 4)) > maxYear Then
                                evStr = "Nincs"
                            End If
                            .SubItems.Add(evStr)
                            .SubItems.Add(appt.EntryID)
                            'resultMin = Me.NumericUpDown2.Value.ToString
                            resultMinInt = appt.ActualWork
                            If resultMinInt < 1 Then
                                resultMinInt = 10
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)
                            resultNSNF = resultNSNF + resultMinInt
                            resultMinInt = appt.TotalWork
                            If resultMinInt < 1 Then
                                resultMinInt = 0
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)
                            resultT_NSNF = resultT_NSNF + resultMinInt
                            .SubItems.Add(appt.PercentComplete.ToString + "%")
                        End With
                    End If
                    If appt.Categories.Contains("@Havi feladat") Then
                        With Me.ListView6.Items.Add("F")
                            .SubItems.Add(appt.Subject)
                            Dim evStr = Format(appt.DueDate, "yyyy/MM/dd")
                            If CLng(Strings.Left(evStr, 4)) > maxYear Then
                                evStr = "Nincs"
                            End If
                            .SubItems.Add(evStr)
                            .SubItems.Add(appt.EntryID)
                            '.SubItems.Add("")
                            resultMinInt = appt.ActualWork
                            If resultMinInt < 1 Then
                                resultMinInt = 20
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)
                            resultMinInt = appt.TotalWork
                            If resultMinInt < 1 Then
                                resultMinInt = 0
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)

                            .SubItems.Add(appt.PercentComplete.ToString + "%")
                        End With
                    End If
                    If appt.Categories.Contains("@Havi cél") Then
                        With Me.ListView5.Items.Add("F")
                            .SubItems.Add(appt.Subject)
                            Dim evStr = Format(appt.DueDate, "yyyy/MM/dd")
                            If CLng(Strings.Left(evStr, 4)) > maxYear Then
                                evStr = "Nincs"
                            End If
                            .SubItems.Add(evStr)
                            .SubItems.Add(appt.EntryID)
                            '.SubItems.Add("")
                            resultMinInt = appt.ActualWork
                            If resultMinInt < 1 Then
                                resultMinInt = 20
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)
                            resultMinInt = appt.TotalWork
                            If resultMinInt < 1 Then
                                resultMinInt = 0
                            End If
                            resultMin = resultMinInt.ToString
                            .SubItems.Add(resultMin)

                            .SubItems.Add(appt.PercentComplete.ToString + "%")
                        End With
                    End If
                End If

            Next
            If rangeAppts IsNot Nothing Then Marshal.ReleaseComObject(rangeAppts)
        End If
    End Sub
    ''' <summary>
    ''' Get Tasks in range
    ''' </summary>
    ''' <param name="folder"></param>
    ''' <param name="startTime">Start date</param>
    ''' <param name="endTime">Due date</param>
    ''' <returns></returns>
    Private Function GetTasksInRange(folder As Outlook.Folder, startTime As DateTime, endTime As DateTime) As Outlook.Items
        'Dim filter As String = "[DueDate] >= '" + Format(startTime, "yyyy/MM/dd") + "' AND [DueDate] <= '" + Format(endTime, "yyyy/MM/dd") + "'"
        'Dim filter As String = "[DueDate] >= '" + Format(startTime, "yyyy/MM/dd") + "'"
        Dim filter As String = "[DueDate] >= '" + Format(startTime, "yyyy/MM/dd") + "' OR [Complete] <> True"

        'Debug.WriteLine(filter)
        Try
            Dim calItems As Outlook.Items = folder.Items
            calItems.IncludeRecurrences = True
            calItems.Sort("[DueDate]", Type.Missing)
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
#End Region
End Class
