Imports System.Diagnostics
Imports System.Windows.Forms

Public Class FormRegion1

#Region "Form Region Factory"

    <Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Appointment)>
    <Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)>
    <Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Task)>
    <Microsoft.Office.Tools.Outlook.FormRegionName("Idomatrix_v60.FormRegion1")>
    Partial Public Class FormRegion1Factory

        ' Occurs before the form region is initialized.
        ' To prevent the form region from appearing, set e.Cancel to true.
        ' Use e.OutlookItem to get a reference to the current Outlook item.
        Private Sub FormRegion1Factory_FormRegionInitializing(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs) Handles Me.FormRegionInitializing

        End Sub

    End Class

#End Region

    'Occurs before the form region is displayed. 
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub FormRegion1_FormRegionShowing(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionShowing
        If (TypeOf Me.OutlookItem Is Outlook.TaskItem) Then
            Me.Button6.Visible = True
            Me.Button7.Visible = True
        Else
            Me.Button6.Visible = False
            Me.Button7.Visible = False
        End If
#Region "Setup Terv-Tény"
        If (TypeOf Me.OutlookItem Is Outlook.MailItem) Then
            Try
                Dim tervTeny As String() = Split(Me.OutlookItem.Companies, "@")
                'Debug.WriteLine(Me.OutlookItem.Companies)
                'Debug.WriteLine(tervTeny(0))
                'Debug.WriteLine(tervTeny(1))
                Me.NumericUpDown1.Value = CInt(tervTeny(0))
                Me.NumericUpDown2.Value = CInt(tervTeny(1))
            Catch ex2 As Exception
                'Me.NumericUpDown1.Value = 0
                Debug.WriteLine("FR:Mail:NoCompanies:" + Me.OutlookItem.Subject)
            End Try
            Me.NumericUpDown1.Enabled = True
        ElseIf (TypeOf Me.OutlookItem Is Outlook.AppointmentItem) Then
            Try
                Dim tervTeny As String() = Split(Me.OutlookItem.Companies, "@")
                'Debug.WriteLine(Me.OutlookItem.Companies)
                'Debug.WriteLine(tervTeny(0))
                'Debug.WriteLine(tervTeny(1))
                'Me.NumericUpDown1.Value = CInt(tervTeny(0))
                Me.NumericUpDown2.Value = CInt(tervTeny(1))
            Catch ex2 As Exception
                Debug.WriteLine("FR:Appt:NoCompanies:" + Me.OutlookItem.Subject)
                Me.NumericUpDown2.Value = 0
            End Try
            Try
                Dim elapsedTime As TimeSpan = Me.OutlookItem.End.Subtract(Me.OutlookItem.Start)
                Me.NumericUpDown1.Value = elapsedTime.TotalMinutes
            Catch ex As Exception
                Debug.WriteLine("FR:Mail:Subtract:" + Me.OutlookItem.Subject)
            End Try
            Me.NumericUpDown1.Enabled = False
        ElseIf (TypeOf Me.OutlookItem Is Outlook.MeetingItem) Then
            Try
                Dim tervTeny As String() = Split(Me.OutlookItem.Companies, "@")
                'Debug.WriteLine(Me.OutlookItem.Companies)
                'Debug.WriteLine(tervTeny(0))
                'Debug.WriteLine(tervTeny(1))
                'Me.NumericUpDown1.Value = CInt(tervTeny(0))
                Me.NumericUpDown2.Value = CInt(tervTeny(1))
            Catch ex2 As Exception
                Debug.WriteLine("FR:Meeting:NoCompanies:" + Me.OutlookItem.Subject)
                Me.NumericUpDown2.Value = 0
            End Try
            Try
                Dim elapsedTime As TimeSpan = Me.OutlookItem.End.Subtract(Me.OutlookItem.Start)
                Me.NumericUpDown1.Value = elapsedTime.TotalMinutes
            Catch ex As Exception
                Debug.WriteLine("FR:Meeting:Subtract:" + Me.OutlookItem.Subject)
            End Try
            Me.NumericUpDown1.Enabled = False
        ElseIf (TypeOf Me.OutlookItem Is Outlook.TaskItem) Then

            Try
                If Me.OutlookItem.ActualWork = 0 Then
                    Me.OutlookItem.ActualWork = 30
                End If
                Me.NumericUpDown1.Value = Me.OutlookItem.ActualWork
                'Me.NumericUpDown2.Value = Me.OutlookItem.TotalWork
            Catch ex As Exception
                Debug.WriteLine("FR:Task:Actual:" + Me.OutlookItem.Subject)
            End Try
            Try
                Me.NumericUpDown2.Value = Me.OutlookItem.TotalWork
            Catch ex As Exception
                Debug.WriteLine("FR:Task:Total:" + Me.OutlookItem.Subject)
            End Try
            'ElseIf (TypeOf Me.OutlookItem Is Outlook.AppointmentItem) Or (TypeOf Me.OutlookItem Is Outlook.MeetingItem) Then
            Me.NumericUpDown1.Enabled = True
        End If

#End Region
    End Sub

    'Occurs when the form region is closed.   
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub FormRegion1_FormRegionClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionClosed
        Call Globals.ThisAddIn.RefreshTaskpane()
    End Sub


    Private Sub SetItem(catStr As String)

        If (TypeOf Me.OutlookItem Is Outlook.MailItem) Then
            Dim mailItem As Outlook.MailItem = TryCast(Me.OutlookItem, Outlook.MailItem)
            Dim existingCategories = mailItem.Categories
            If (String.IsNullOrEmpty(existingCategories)) Then
                mailItem.Categories = catStr
            Else
                existingCategories = StripEisenhowerCats(existingCategories)
                If (mailItem.Categories.Contains(catStr) = False) Then
                    mailItem.Categories = existingCategories + ", " + catStr
                End If
            End If
            mailItem.FlagStatus = Outlook.OlFlagStatus.olFlagMarked
            mailItem.FlagIcon = Outlook.OlFlagIcon.olRedFlagIcon
            mailItem.MarkAsTask(Microsoft.Office.Interop.Outlook.OlMarkInterval.olMarkToday)
            mailItem.Save()
            Call Globals.ThisAddIn.RefreshTaskpane()
        ElseIf (TypeOf Me.OutlookItem Is Outlook.AppointmentItem) Then
            Dim apptItem As Outlook.AppointmentItem = TryCast(Me.OutlookItem, Outlook.AppointmentItem)
            Dim existingCategories = apptItem.Categories
            If (String.IsNullOrEmpty(existingCategories)) Then
                apptItem.Categories = catStr
            Else
                existingCategories = StripEisenhowerCats(existingCategories)
                If (apptItem.Categories.Contains(catStr) = False) Then
                    apptItem.Categories = existingCategories + ", " + catStr
                End If
            End If
            apptItem.Save()
            Call Globals.ThisAddIn.RefreshTaskpane()
        ElseIf (TypeOf Me.OutlookItem Is Outlook.MeetingItem) Then
            Dim meetingItem As Outlook.MeetingItem = TryCast(Me.OutlookItem, Outlook.MeetingItem)
            Dim existingCategories = meetingItem.Categories
            If (String.IsNullOrEmpty(existingCategories)) Then
                meetingItem.Categories = catStr
            Else
                existingCategories = StripEisenhowerCats(existingCategories)
                If (meetingItem.Categories.Contains(catStr) = False) Then
                    meetingItem.Categories = existingCategories + ", " + catStr
                End If
            End If
            meetingItem.Save()
            Call Globals.ThisAddIn.RefreshTaskpane()
        ElseIf (TypeOf Me.OutlookItem Is Outlook.TaskItem) Then
            Dim taskItem As Outlook.TaskItem = TryCast(Me.OutlookItem, Outlook.TaskItem)
            Dim existingCategories = taskItem.Categories
            If (String.IsNullOrEmpty(existingCategories)) Then
                taskItem.Categories = catStr
            Else
                existingCategories = StripEisenhowerCats(existingCategories)
                If (taskItem.Categories.Contains(catStr) = False) Then
                    taskItem.Categories = existingCategories + ", " + catStr
                End If
            End If
            taskItem.Save()
            Call Globals.ThisAddIn.RefreshTaskpane()
        End If
    End Sub
    Private Sub SetMonthlyItem(catStr As String)
        If (TypeOf Me.OutlookItem Is Outlook.TaskItem) Then
            Dim res As MsgBoxResult = MessageBox.Show("Ebben az esetben nem csak a kategória, hanem az IDŐ (kezdés, befejezés) is változik!" + vbCrLf + "Folytatod?", "FIGYELEM!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If res = MsgBoxResult.Yes Then
                Dim taskItem As Outlook.TaskItem = TryCast(Me.OutlookItem, Outlook.TaskItem)
                Dim existingCategories = taskItem.Categories
                If (String.IsNullOrEmpty(existingCategories)) Then
                    taskItem.Categories = catStr
                Else
                    existingCategories = StripEisenhowerCats(existingCategories)
                    If (taskItem.Categories.Contains(catStr) = False) Then
                        taskItem.Categories = existingCategories + ", " + catStr
                    End If
                End If
                Dim FirstDay As Date
                'This year, this month, first day
                FirstDay = DateSerial(Today.Year, Today.Month, 1)
                Dim LastDay As Date
                'This year, next month, 0th day is this month's last day
                LastDay = DateSerial(Today.Year, Today.Month + 1, 0)
                taskItem.StartDate = FirstDay
                taskItem.DueDate = LastDay
                taskItem.Save()
                Call Globals.ThisAddIn.RefreshTaskpane()
            End If
        End If
    End Sub
    Private Function StripEisenhowerCats(catStr As String) As String
        catStr = catStr.Replace("@Nem sürgős - Fontos", vbNullString)
        catStr = catStr.Replace("@Sürgős - Fontos", vbNullString)
        catStr = catStr.Replace("@Sürgős - Nem fontos", vbNullString)
        catStr = catStr.Replace("@Nem sürgős - Nem fontos", vbNullString)
        catStr = catStr.Replace("@Havi feladat", vbNullString)
        catStr = catStr.Replace("@Havi cél", vbNullString)

        Return catStr
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call SetItem("@Sürgős - Fontos")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call SetItem("@Nem sürgős - Fontos")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call SetItem("@Sürgős - Nem fontos")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call SetItem("@Nem sürgős - Nem fontos")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If (TypeOf Me.OutlookItem Is Outlook.MailItem) Then
            Dim mailItem As Outlook.MailItem = TryCast(Me.OutlookItem, Outlook.MailItem)
            Dim existingCategories = mailItem.Categories
            If (String.IsNullOrEmpty(existingCategories)) Then
            Else
                existingCategories = StripEisenhowerCats(existingCategories)
                mailItem.Categories = existingCategories
                mailItem.Save()
            End If
            Call Globals.ThisAddIn.RefreshTaskpane()
        ElseIf (TypeOf Me.OutlookItem Is Outlook.AppointmentItem) Then
            Dim apptItem As Outlook.AppointmentItem = TryCast(Me.OutlookItem, Outlook.AppointmentItem)
            Dim existingCategories = apptItem.Categories
            If (String.IsNullOrEmpty(existingCategories)) Then
            Else
                existingCategories = StripEisenhowerCats(existingCategories)
                apptItem.Categories = existingCategories
                apptItem.Save()
            End If
            Call Globals.ThisAddIn.RefreshTaskpane()
        ElseIf (TypeOf Me.OutlookItem Is Outlook.MeetingItem) Then
            Dim meetingItem As Outlook.MeetingItem = TryCast(Me.OutlookItem, Outlook.MeetingItem)
            Dim existingCategories = meetingItem.Categories
            If (String.IsNullOrEmpty(existingCategories)) Then
            Else
                existingCategories = StripEisenhowerCats(existingCategories)
                meetingItem.Categories = existingCategories
                meetingItem.Save()
            End If
            Call Globals.ThisAddIn.RefreshTaskpane()
        ElseIf (TypeOf Me.OutlookItem Is Outlook.TaskItem) Then
            Dim taskItem As Outlook.TaskItem = TryCast(Me.OutlookItem, Outlook.TaskItem)
            Dim existingCategories = taskItem.Categories
            If (String.IsNullOrEmpty(existingCategories)) Then
            Else
                existingCategories = StripEisenhowerCats(existingCategories)
                taskItem.Categories = existingCategories
                taskItem.Save()
            End If
            Call Globals.ThisAddIn.RefreshTaskpane()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call SetMonthlyItem("@Havi cél")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call SetMonthlyItem("@Havi feladat")
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        Try
            Me.OutlookItem.Companies = CStr(NumericUpDown1.Value) + "@" + CStr(NumericUpDown2.Value)
        Catch ex2 As Exception
        End Try
        If (TypeOf Me.OutlookItem Is Outlook.TaskItem) Then
            Try
                Me.OutlookItem.ActualWork = Me.NumericUpDown1.Value
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        Try
            Me.OutlookItem.Companies = CStr(NumericUpDown1.Value) + "@" + CStr(NumericUpDown2.Value)
        Catch ex2 As Exception
        End Try

        If (TypeOf Me.OutlookItem Is Outlook.TaskItem) Then
            Try
                Me.OutlookItem.TotalWork = Me.NumericUpDown2.Value
            Catch ex As Exception
            End Try
        End If
    End Sub
End Class
