Imports Microsoft.Office.Interop.Outlook

Public Class ThisAddIn
    Private taskPaneControl1 As IGMPane
    Public Shared WithEvents taskPaneValue As Microsoft.Office.Tools.CustomTaskPane

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Call AddCategories()
        'TaskPane
        taskPaneControl1 = New IGMPane()
        taskPaneValue = Me.CustomTaskPanes.Add(taskPaneControl1, "Időmátrix")
        taskPaneValue.Visible = True
        taskPaneValue.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon()
    End Function
    Public ReadOnly Property TaskPane() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return taskPaneValue
        End Get
    End Property
    Public Sub RefreshTaskpane()
        Call taskPaneControl1.RefreshData()
    End Sub
    ''' <summary>
    ''' Kategóriák hozzáadása, ha még nem léteznek
    ''' </summary>
    Private Sub AddCategories()
        Dim categories As Categories = Me.Application.Session.Categories
        Try
            categories.Add("@Nem sürgős - Fontos", Outlook.OlCategoryColor.olCategoryColorGreen)
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        End Try
        Try
            categories.Add("@Sürgős - Fontos", Outlook.OlCategoryColor.olCategoryColorBlue)
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        End Try
        Try
            categories.Add("@Sürgős - Nem fontos", Outlook.OlCategoryColor.olCategoryColorYellow)
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        End Try
        Try
            categories.Add("@Nem sürgős - Nem fontos", Outlook.OlCategoryColor.olCategoryColorRed)
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        End Try
        Try
            categories.Add("@Havi feladat", Outlook.OlCategoryColor.olCategoryColorSteel)
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        End Try
        Try
            categories.Add("@Havi cél", Outlook.OlCategoryColor.olCategoryColorGray)
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

End Class
