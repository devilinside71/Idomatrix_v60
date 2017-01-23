'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook

<Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("Idomatrix_v60.Ribbon.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub
    Public Function imageRibbon_GetImage_DelMatrixCat(ByVal control As Office.IRibbonControl) As Drawing.Bitmap
        Return My.Resources.Idomatrix_Delete_32
    End Function
    Public Function imageRibbon_GetImage_MatrixCat(ByVal control As Office.IRibbonControl) As Drawing.Bitmap
        Return My.Resources.Idomatrix_32
    End Function


    ''' <summary>
    ''' Context menu and ribbon category button action
    ''' </summary>
    ''' <param name="control">Ribbon control item</param>
    Public Sub OnMyButtonClick(ByVal control As Office.IRibbonControl)
        Dim catStr As String = vbNullString
        Dim explorer As Outlook.Explorer
        explorer = Globals.ThisAddIn.Application.ActiveExplorer


        If explorer IsNot Nothing AndAlso explorer.Selection IsNot Nothing AndAlso explorer.Selection.Count > 0 Then
            'Setup category text
            If control.Id = "ButtonCatSF" Or control.Id = "MyContextMenuMailItemSF" Or control.Id = "MyContextMenuCalendarItemSF" Or control.Id = "MyContextMenuTaskItemSF" Or control.Id = "MyContextMenuFlaggedMailItemSF" Then
                catStr = "@Sürgős - Fontos"
            ElseIf control.Id = "ButtonCatSNF" Or control.Id = "MyContextMenuMailItemSNF" Or control.Id = "MyContextMenuCalendarItemSNF" Or control.Id = "MyContextMenuTaskItemSNF" Or control.Id = "MyContextMenuFlaggedMailItemSNF" Then
                catStr = "@Sürgős - Nem fontos"
            ElseIf control.Id = "ButtonCatNSF" Or control.Id = "MyContextMenuMailItemNSF" Or control.Id = "MyContextMenuCalendarItemNSF" Or control.Id = "MyContextMenuTaskItemNSF" Or control.Id = "MyContextMenuFlaggedMailItemNSF" Then
                catStr = "@Nem sürgős - Fontos"
            ElseIf control.Id = "ButtonCatNSNF" Or control.Id = "MyContextMenuMailItemNSNF" Or control.Id = "MyContextMenuCalendarItemNSNF" Or control.Id = "MyContextMenuTaskItemNSNF" Or control.Id = "MyContextMenuFlaggedMailItemNSNF" Then
                catStr = "@Nem sürgős - Nem fontos"
            ElseIf control.Id = "ButtonCatMonthly" Or control.Id = "MyContextMenuTaskItemMonthly" Then
                catStr = "@Havi feladat"
            ElseIf control.Id = "ButtonCatMonthlyGoal" Or control.Id = "MyContextMenuTaskItemMonthlyGoal" Then
                catStr = "@Havi cél"
            End If

            Dim item As Object = explorer.Selection(1)
            If TypeOf item Is MailItem Then
                If control.Id <> "ButtonCatMonthly" And control.Id <> "ButtonCatMonthlyGoal" Then
                    Dim mailItem As MailItem = TryCast(item, MailItem)
                    Dim existingCategories = mailItem.Categories
                    If (String.IsNullOrEmpty(existingCategories)) Then
                        mailItem.Categories = catStr
                    Else
                        existingCategories = StripEisenhowerCats(existingCategories)
                        If (mailItem.Categories.Contains(catStr) = False) Then
                            mailItem.Categories = existingCategories + ", " + catStr
                        End If
                    End If
                    mailItem.FlagStatus = OlFlagStatus.olFlagMarked
                    mailItem.FlagIcon = OlFlagIcon.olRedFlagIcon
                    mailItem.MarkAsTask(OlMarkInterval.olMarkToday)
                    mailItem.Save()
                    Call Globals.ThisAddIn.RefreshTaskpane()
                    If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem)
                End If
            ElseIf (TypeOf item Is AppointmentItem) Then
                If control.Id <> "ButtonCatMonthly" And control.Id <> "ButtonCatMonthlyGoal" Then
                    Dim apptItem As AppointmentItem = TryCast(item, AppointmentItem)
                    Dim existingCategories = apptItem.Categories
                    'If apptItem.IsRecurring Then
                    '    Dim res = MessageBox.Show("Ez egy ismétlődő találkozó, csak megnyitás után kategorizálható." + vbCrLf + "Megnyitod?", "FIGYELEM!", MessageBoxButtons.YesNo)
                    '    If res = vbYes Then

                    '    End If
                    'Else
                    If (String.IsNullOrEmpty(existingCategories)) Then
                        apptItem.Categories = catStr
                    Else
                        existingCategories = StripEisenhowerCats(existingCategories)
                        If (apptItem.Categories.Contains(catStr) = False) Then
                            apptItem.Categories = existingCategories + ", " + catStr
                        End If
                    End If
                    apptItem.Save()
                    'End If
                    Call Globals.ThisAddIn.RefreshTaskpane()
                    If apptItem IsNot Nothing Then Marshal.ReleaseComObject(apptItem)
                End If
            ElseIf (TypeOf item Is MeetingItem) Then
                If control.Id <> "ButtonCatMonthly" And control.Id <> "ButtonCatMonthlyGoal" Then
                    Dim meetingItem As MeetingItem = TryCast(item, MeetingItem)
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
                    If meetingItem IsNot Nothing Then Marshal.ReleaseComObject(meetingItem)
                End If
            ElseIf (TypeOf item Is TaskItem) Then
                If catStr = "@Havi feladat" Or catStr = "@Havi cél" Then
                    Dim res As MsgBoxResult = MessageBox.Show("Ebben az esetben nem csak a kategória, hanem a KEZDÉS DÁTUMA és a HATÁRIDŐ is változik!" + vbCrLf + "Folytatod?", "FIGYELEM!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                    If res = MsgBoxResult.Yes Then
                        Dim taskItem As TaskItem = TryCast(item, TaskItem)
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
                        If taskItem IsNot Nothing Then Marshal.ReleaseComObject(taskItem)
                    End If
                Else
                    Dim taskItem As TaskItem = TryCast(item, TaskItem)
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
                    If taskItem IsNot Nothing Then Marshal.ReleaseComObject(taskItem)
                End If
            End If
            'If (control.Id = "MyContextMenuMailItemSF") Or (control.Id = "ButtonCatSF") Then
            '    MessageBox.Show("You clicked " + control.Id)
            'End If
            'If My.Settings.AutoRefresh Then
            '    Dim sss As IdomatrixCustomPane
            '    sss = New IdomatrixCustomPane
            '    Call sss.RefreshData()
            'End If
        End If
    End Sub
    ''' <summary>
    ''' Delete Matrix categories
    ''' </summary>
    ''' <param name="control">Ribbon or Contect menu item</param>
    Public Sub OnMyButtonClickDel(ByVal control As Office.IRibbonControl)
        Dim catStr As String = vbNullString
        Dim explorer As Outlook.Explorer
        explorer = Globals.ThisAddIn.Application.ActiveExplorer

        Dim item As Object = explorer.Selection(1)
        If explorer IsNot Nothing AndAlso explorer.Selection IsNot Nothing AndAlso explorer.Selection.Count > 0 Then
            If TypeOf item Is MailItem Then
                Dim mailItem As MailItem = TryCast(item, MailItem)
                Dim existingCategories = mailItem.Categories
                If (String.IsNullOrEmpty(existingCategories)) Then
                Else
                    mailItem.Categories = StripEisenhowerCats(existingCategories)
                    mailItem.Save()
                End If

                Call Globals.ThisAddIn.RefreshTaskpane()

                If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem)
            ElseIf (TypeOf item Is AppointmentItem) Then
                Dim apptItem As AppointmentItem = TryCast(item, AppointmentItem)
                Dim existingCategories = apptItem.Categories
                If (String.IsNullOrEmpty(existingCategories)) Then
                Else
                    apptItem.Categories = StripEisenhowerCats(existingCategories)
                    apptItem.Save()
                End If
                Call Globals.ThisAddIn.RefreshTaskpane()
                If apptItem IsNot Nothing Then Marshal.ReleaseComObject(apptItem)
            ElseIf (TypeOf item Is MeetingItem) Then
                Dim meetingItem As MeetingItem = TryCast(item, MeetingItem)
                Dim existingCategories = meetingItem.Categories
                If (String.IsNullOrEmpty(existingCategories)) Then
                Else
                    meetingItem.Categories = StripEisenhowerCats(existingCategories)
                    meetingItem.Save()
                End If
                Call Globals.ThisAddIn.RefreshTaskpane()
                If meetingItem IsNot Nothing Then Marshal.ReleaseComObject(meetingItem)
            ElseIf (TypeOf item Is TaskItem) Then
                Dim taskItem As TaskItem = TryCast(item, TaskItem)
                Dim existingCategories = taskItem.Categories
                If (String.IsNullOrEmpty(existingCategories)) Then
                Else
                    taskItem.Categories = StripEisenhowerCats(existingCategories)
                    taskItem.Save()
                End If
                Call Globals.ThisAddIn.RefreshTaskpane()
                If taskItem IsNot Nothing Then Marshal.ReleaseComObject(taskItem)
            End If
        End If
    End Sub
    Public Sub ButtonMOnOff_Click(ByVal control As Office.IRibbonControl)
        Globals.ThisAddIn.TaskPane.Visible = True
    End Sub


#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function
    Private Function StripEisenhowerCats(catStr As String) As String
        catStr = catStr.Replace("@Nem sürgős - Fontos", vbNullString)
        catStr = catStr.Replace("@Sürgős - Fontos", vbNullString)
        catStr = catStr.Replace("@Sürgős - Nem fontos", vbNullString)
        catStr = catStr.Replace("@Nem sürgős - Nem fontos", vbNullString)
        catStr = catStr.Replace("@Havi feladat", vbNullString)
        catStr = catStr.Replace("@Havi cél", vbNullString)

        Return catStr
    End Function
#End Region

End Class
