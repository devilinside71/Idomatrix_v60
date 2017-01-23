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
    Private Sub OpenItem(sender As Object)
        Dim myItem As Object

        myItem = Globals.ThisAddIn.Application.Session.GetItemFromID(sender.SelectedItems.Item(0).SubItems.Item(3).Text)
        myItem.Display(True)
        Call RefreshData()

    End Sub

    Public Sub RefreshData()
        Dim sumdata As Double
        Dim resSF As Integer
        Dim resNSF As Integer
        Dim resSNF As Integer
        Dim resNSNF As Integer
        Dim resTSF As Integer
        Dim resTNSF As Integer
        Dim resTSNF As Integer
        Dim resTNSNF As Integer

        resultNSF = 0
        resultNSNF = 0
        resultSNF = 0
        resultSF = 0
        resultT_NSF = 0
        resultT_SF = 0
        resultT_SNF = 0
        resultT_NSNF = 0
        resTNSF = 0
        resTNSNF = 0
        resTSF = 0
        resTSNF = 0

        Call ClearLists()
        Call SetEmailTasksInRange()
        Call SetTasksInRange()
        Call SetAppointmentsInRange()

        Label11.Text = resultSF.ToString
        Label10.Text = resultSNF.ToString
        Label12.Text = resultNSF.ToString
        Label13.Text = resultNSNF.ToString
        Label16.Text = Math.Round(resultSF / 60, 2)
        Label15.Text = Math.Round(resultSNF / 60, 2)
        Label17.Text = Math.Round(resultNSF / 60, 2)
        Label18.Text = Math.Round(resultNSNF / 60, 2)
        Label20.Text = resultT_NSF.ToString
        Label21.Text = resultT_SF.ToString
        Label22.Text = resultT_SNF.ToString
        Label23.Text = resultT_NSNF.ToString
        Label25.Text = Math.Round(resultT_NSF / 60, 2)
        Label26.Text = Math.Round(resultT_SF / 60, 2)
        Label27.Text = Math.Round(resultT_SNF / 60, 2)
        Label28.Text = Math.Round(resultT_NSNF / 60, 2)

        resultSum = resultSF + resultNSF + resultNSNF + resultSNF
        Label14.Text = resultSum.ToString
        sumdata = Math.Round(resultSum / 60, 2)
        Label19.Text = sumdata.ToString

        resultT_sum = resultT_SF + resultT_NSF + resultT_SNF + resultT_NSNF
        Label24.Text = resultT_sum.ToString
        sumdata = Math.Round(resultT_sum / 60, 2)
        Label29.Text = sumdata.ToString

        resNSF = CInt(resultNSF / resultSum * 100)
        resSF = CInt(resultSF / resultSum * 100)
        resSNF = CInt(resultSNF / resultSum * 100)
        resNSNF = 100 - resSF - resNSF - resSNF
        If resNSNF < 0 Then
            resNSNF = 0
        End If
        Label37.Text = resNSF.ToString + "%"
        Label38.Text = resSF.ToString + "%"
        Label39.Text = resSNF.ToString + "%"
        Label40.Text = resNSNF.ToString + "%"

        resultEvalSF = resSF
        resultEvalSNF = resSNF
        resultEvalNSF = resNSF
        resultEvalNSNF = resNSNF

        resTNSF = CInt(resultT_NSF / resultT_sum * 100)
        resTSF = CInt(resultT_SF / resultT_sum * 100)
        resTSNF = CInt(resultT_SNF / resultT_sum * 100)
        resTNSNF = CInt(resultT_NSNF / resultT_sum * 100)
        If resTNSNF < 0 Then
            resTNSNF = 0
        End If
        Label41.Text = resTNSF.ToString + "%"
        Label42.Text = resTSF.ToString + "%"
        Label43.Text = resTSNF.ToString + "%"
        Label44.Text = resTNSNF.ToString + "%"
        resultT_EvalNSF = resTNSF
        resultT_EvalSF = resTSF
        resultT_EvalSNF = resTSNF
        resultT_EvalNSNF = resTNSNF

        Dim styles As TableLayoutColumnStyleCollection =
Me.TableLayoutPanel14.ColumnStyles

        Dim styles2 As TableLayoutColumnStyleCollection =
    Me.TableLayoutPanel15.ColumnStyles

        styles(0).SizeType = SizeType.Percent
        styles(0).Width = resNSF
        styles(1).SizeType = SizeType.Percent
        styles(1).Width = resSF
        styles(2).SizeType = SizeType.Percent
        styles(2).Width = resSNF
        styles(3).SizeType = SizeType.Percent
        styles(3).Width = resNSNF

        styles2(0).SizeType = SizeType.Percent
        styles2(0).Width = resTNSF
        styles2(1).SizeType = SizeType.Percent
        styles2(1).Width = resTSF
        styles2(2).SizeType = SizeType.Percent
        styles2(2).Width = resTSNF
        styles2(3).SizeType = SizeType.Percent
        styles2(3).Width = resTNSNF

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

    Private Sub ListView5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView5.SelectedIndexChanged

    End Sub

    Private Sub ListView5_DoubleClick(sender As Object, e As EventArgs) Handles ListView5.DoubleClick
        Call OpenItem(sender)
    End Sub

    Private Sub ListView6_DoubleClick(sender As Object, e As EventArgs) Handles ListView6.DoubleClick
        Call OpenItem(sender)
    End Sub

    Private Sub ListView4_DoubleClick(sender As Object, e As EventArgs) Handles ListView4.DoubleClick
        Call OpenItem(sender)
    End Sub

    Private Sub ListView3_DoubleClick(sender As Object, e As EventArgs) Handles ListView3.DoubleClick
        Call OpenItem(sender)
    End Sub

    Private Sub ListView2_DoubleClick(sender As Object, e As EventArgs) Handles ListView2.DoubleClick
        Call OpenItem(sender)
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        Call OpenItem(sender)
    End Sub

    Private Sub ListView1_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView1.MouseUp
        Select Case e.Button
            Case MouseButtons.Right
                Call AddNewItem("@Sürgős - Fontos")
            Case MouseButtons.Middle
        End Select
    End Sub

#Region "Add new items"

    Private Sub AddNewItem(catStr As String)
        'Dim res As MsgBoxResult = MessageBox.Show("Új találkozó hozzáadása." + vbCrLf + catStr, "Új elem", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        'Select Case res
        '    Case vbYes
        '        Call CreateAppt(catStr)
        '    Case vbNo
        '        Dim res2 As MsgBoxResult = MessageBox.Show("Új feladat hozzáadása" + vbCrLf + catStr, "Új elem", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        '        Select Case res2
        '            Case vbYes
        '                Call CreateTask(catStr)
        '        End Select
        'End Select
        Dim newItemForm As SelectNewItemForm
        newItemForm = New SelectNewItemForm
        newItemForm.Text = "Új elem"
        newItemForm.Label1.Text = catStr
        'newItemForm.Top = (My.Computer.Screen.WorkingArea.Height) / 2 - (newItemForm.Height \ 2)
        'newItemForm.Left = (My.Computer.Screen.WorkingArea.Width \ 2) - (newItemForm.Width \ 2)
        If catStr = "@Sürgős - Fontos" Then
            newItemForm.Button1.BackColor = Drawing.Color.CornflowerBlue
            newItemForm.Button2.BackColor = Drawing.Color.CornflowerBlue
            newItemForm.Button3.BackColor = Drawing.Color.CornflowerBlue
            newItemForm.Button3.Visible = False
        End If
        If catStr = "@Sürgős - Nem fontos" Then
            newItemForm.Button1.BackColor = Drawing.Color.Yellow
            newItemForm.Button2.BackColor = Drawing.Color.Yellow
            newItemForm.Button3.BackColor = Drawing.Color.Yellow
            newItemForm.Button3.Visible = True
        End If
        If catStr = "@Nem sürgős - Fontos" Then
            newItemForm.Button1.BackColor = Drawing.Color.LimeGreen
            newItemForm.Button2.BackColor = Drawing.Color.LimeGreen
            newItemForm.Button3.BackColor = Drawing.Color.LimeGreen
            newItemForm.Button3.Visible = False
        End If
        If catStr = "@Nem sürgős - Nem fontos" Then
            newItemForm.Button1.BackColor = Drawing.Color.Salmon
            newItemForm.Button2.BackColor = Drawing.Color.Salmon
            newItemForm.Button3.BackColor = Drawing.Color.Salmon
            newItemForm.Button3.Visible = False
        End If



        newItemForm.ShowDialog()

        If newItemForm.ItemChoiced = 1 Then
            Call CreateAppt(catStr)
        End If
        If newItemForm.ItemChoiced = 2 Then
            Call CreateTask(catStr)
        End If
        If newItemForm.ItemChoiced = 3 Then
            'Spec levelezés
            'Call CreateSpecAppt(catStr, "Levelezés átnézése", 90)
            Call CreateSpecTask(catStr, "Levelezés átnézése", 90)
            'MessageBox.Show("Ide a spec jön")
        End If
    End Sub
    Private Sub AddNewItemMonthly(catStr As String)
        Dim res2 As MsgBoxResult = MessageBox.Show("Új havi feladat hozzáadása", "Új elem", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        Select Case res2
            Case vbYes
                Call CreateTaskMonthly(catStr)
        End Select
    End Sub
    Private Sub AddNewItemMonthlyGoal(catStr As String)
        Dim res2 As MsgBoxResult = MessageBox.Show("Új havi cél hozzáadása", "Új elem", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        Select Case res2
            Case vbYes
                Call CreateTaskMonthly(catStr)
        End Select
    End Sub
    Private Sub CreateSpecAppt(catStr As String, subject As String, duration As Integer)
        Dim myItem As Object
        'Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient
        myItem = Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
        myItem.Subject = subject + " (r)"
        'myItem.Location = "Conference Room B"
        'myItem.Start = #9/24/2015 1:30:00 PM#
        myItem.Start = DateTime.Today
        myItem.ActualWork = duration
        myItem.ReminderSet = False
        'myRequiredAttendee = myItem.Recipients.Add("Nate Sun")
        'myOptionalAttendee = myItem.Recipients.Add("Kevin Kennedy")
        'myResourceAttendee = myItem.Recipients.Add("Conference Room B")
        myItem.Categories = catStr
        myItem.Display(True)

        Call RefreshData()


    End Sub
    Private Sub CreateSpecTask(catStr As String, subject As String, duration As Integer)
        Dim myItem As Object

        myItem = Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olTaskItem)
        myItem.Subject = subject + " (r)"
        myItem.StartDate = DateTime.Today

        myItem.ActualWork = duration
        myItem.ReminderSet = False
        myItem.Categories = catStr
        myItem.Display(True)

        Call RefreshData()


    End Sub
    Friend Sub CreateAppointment(title As String)
        'Dim apptItem As Outlook.AppointmentItem = Nothing
        'apptItem =
        '        OutlookApp.Session.Application.CreateItem(
        '        Outlook.OlItemType.olAppointmentItem)

        'With apptItem
        '    .Subject = title
        '    .Start = DateTime.Now
        '    .End = Date.Now.AddHours(1)
        '    .Save()
        '    .ReminderSet = False
        'End With

        ''Release COM Objects
        'If apptItem IsNot Nothing Then Marshal.ReleaseComObject(apptItem)
    End Sub
    Private Sub CreateAppt(catStr As String)
        Dim myItem As Object
        'Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient
        myItem = Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
        'myItem.Subject = "Strategy Meeting"
        'myItem.Location = "Conference Room B"
        'myItem.Start = #9/24/2015 1:30:00 PM#
        'myItem.Duration = 90
        'myRequiredAttendee = myItem.Recipients.Add("Nate Sun")
        'myOptionalAttendee = myItem.Recipients.Add("Kevin Kennedy")
        'myResourceAttendee = myItem.Recipients.Add("Conference Room B")
        myItem.Categories = catStr
        myItem.Display(True)

        Call RefreshData()

    End Sub
    Private Sub CreateTask(catStr As String)
        Dim myItem As Object
        myItem = Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olTaskItem)
        myItem.Categories = catStr
        myItem.ActualWork = 30
        myItem.Display(True)
        'MessageBox.Show("Feladat")

        Call RefreshData()
    End Sub
    Private Sub CreateTaskMonthly(catStr As String)
        Dim myItem As Object
        myItem = Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olTaskItem)
        myItem.Categories = catStr
        Dim FirstDay As Date
        'This year, this month, first day
        FirstDay = DateSerial(Today.Year, Today.Month, 1)
        Dim LastDay As Date
        'This year, next month, 0th day is this month's last day
        LastDay = DateSerial(Today.Year, Today.Month + 1, 0)
        myItem.StartDate = FirstDay
        myItem.DueDate = LastDay
        myItem.ActualWork = 20
        myItem.Display(True)
        'MessageBox.Show("Feladat")

        Call RefreshData()

    End Sub

#End Region

    Private Sub ListView2_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView2.MouseUp
        Select Case e.Button
            Case MouseButtons.Right
                Call AddNewItem("@Nem sürgős - Fontos")
            Case MouseButtons.Middle
        End Select
    End Sub

    Private Sub ListView3_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView3.MouseUp
        Select Case e.Button
            Case MouseButtons.Right
                Call AddNewItem("@Sürgős - Nem fontos")
            Case MouseButtons.Middle
        End Select
    End Sub

    Private Sub ListView4_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView4.MouseUp
        Select Case e.Button
            Case MouseButtons.Right
                Call AddNewItem("@Nem sürgős - Nem fontos")
            Case MouseButtons.Middle
        End Select
    End Sub

    Private Sub ListView5_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView5.MouseUp
        Select Case e.Button
            Case MouseButtons.Right
                Call AddNewItemMonthlyGoal("@Havi cél")
            Case MouseButtons.Middle
        End Select
    End Sub

    Private Sub ListView6_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView6.MouseUp
        Select Case e.Button
            Case MouseButtons.Right
                Call AddNewItemMonthly("@Havi feladat")
            Case MouseButtons.Middle
        End Select
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Call RefreshData()
    End Sub

    Private Sub SendReport(Optional Test As Boolean = False)

    End Sub
End Class
