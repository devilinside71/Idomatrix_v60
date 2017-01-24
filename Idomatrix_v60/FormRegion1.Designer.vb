<System.ComponentModel.ToolboxItemAttribute(False)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormRegion1
    Inherits Microsoft.Office.Tools.Outlook.FormRegionBase

    Public Sub New(ByVal formRegion As Microsoft.Office.Interop.Outlook.FormRegion)
        MyBase.New(Globals.Factory, formRegion)
        Me.InitializeComponent()
    End Sub

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.NumericUpDown2 = New System.Windows.Forms.NumericUpDown()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel3.SuspendLayout()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 5
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 48.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Button1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Button2, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Button3, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Button4, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Button5, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Button6, 4, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Button7, 4, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 3, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel3, 3, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(465, 71)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.CornflowerBlue
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button1.Location = New System.Drawing.Point(3, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(54, 29)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "SF"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.LightGreen
        Me.Button2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button2.Location = New System.Drawing.Point(63, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(54, 29)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "NSF"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Yellow
        Me.Button3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button3.Location = New System.Drawing.Point(3, 38)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(54, 30)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "SNF"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.Salmon
        Me.Button4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button4.Location = New System.Drawing.Point(63, 38)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(54, 30)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "NSNF"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button5.Location = New System.Drawing.Point(123, 3)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(54, 29)
        Me.Button5.TabIndex = 4
        Me.Button5.Text = "Törlés"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Gray
        Me.Button6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button6.Location = New System.Drawing.Point(406, 3)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(56, 29)
        Me.Button6.TabIndex = 5
        Me.Button6.Text = "Havi cél"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.Silver
        Me.Button7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button7.Location = New System.Drawing.Point(406, 38)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(56, 30)
        Me.Button7.TabIndex = 6
        Me.Button7.Text = "Havi feladat"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 84.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.NumericUpDown1, 1, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(183, 3)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(217, 29)
        Me.TableLayoutPanel2.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(28, 29)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Terv"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(37, 3)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(120, 18)
        Me.NumericUpDown1.TabIndex = 1
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 2
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 84.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.NumericUpDown2, 1, 0)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(183, 38)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(217, 30)
        Me.TableLayoutPanel3.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label2.Location = New System.Drawing.Point(3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(28, 30)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Tény"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'NumericUpDown2
        '
        Me.NumericUpDown2.Location = New System.Drawing.Point(37, 3)
        Me.NumericUpDown2.Name = "NumericUpDown2"
        Me.NumericUpDown2.Size = New System.Drawing.Size(120, 18)
        Me.NumericUpDown2.TabIndex = 1
        '
        'FormRegion1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Name = "FormRegion1"
        Me.Size = New System.Drawing.Size(465, 71)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel3.PerformLayout()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    'NOTE: The following procedure is required by the Form Regions Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Shared Sub InitializeManifest(ByVal manifest As Microsoft.Office.Tools.Outlook.FormRegionManifest, ByVal factory As Microsoft.Office.Tools.Outlook.Factory)
        manifest.FormRegionName = "Idomatrix_v60"
        manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Adjoining

    End Sub

    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel

    Friend WithEvents Button1 As Windows.Forms.Button

    Friend WithEvents Button2 As Windows.Forms.Button

    Friend WithEvents Button3 As Windows.Forms.Button

    Friend WithEvents Button4 As Windows.Forms.Button

    Friend WithEvents Button5 As Windows.Forms.Button

    Friend WithEvents Button6 As Windows.Forms.Button

    Friend WithEvents Button7 As Windows.Forms.Button
    Friend WithEvents TableLayoutPanel2 As Windows.Forms.TableLayoutPanel

    Friend WithEvents Label1 As Windows.Forms.Label

    Friend WithEvents NumericUpDown1 As Windows.Forms.NumericUpDown

    Friend WithEvents TableLayoutPanel3 As Windows.Forms.TableLayoutPanel

    Friend WithEvents Label2 As Windows.Forms.Label

    Friend WithEvents NumericUpDown2 As Windows.Forms.NumericUpDown

    Partial Public Class FormRegion1Factory
        Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory

        Public Event FormRegionInitializing As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler

        Private _Manifest As Microsoft.Office.Tools.Outlook.FormRegionManifest


        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Public Sub New()
            Me._Manifest = Globals.Factory.CreateFormRegionManifest()
            FormRegion1.InitializeManifest(Me._Manifest, Globals.Factory)
        End Sub

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        ReadOnly Property Manifest() As Microsoft.Office.Tools.Outlook.FormRegionManifest Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.Manifest
            Get
                Return Me._Manifest
            End Get
        End Property

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Function CreateFormRegion(ByVal formRegion As Microsoft.Office.Interop.Outlook.FormRegion) As Microsoft.Office.Tools.Outlook.IFormRegion Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion
            Dim form As FormRegion1 = New FormRegion1(formRegion)
            form.Factory = Me
            Return form
        End Function

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Function GetFormRegionStorage(ByVal outlookItem As Object, ByVal formRegionMode As Microsoft.Office.Interop.Outlook.OlFormRegionMode, ByVal formRegionSize As Microsoft.Office.Interop.Outlook.OlFormRegionSize) As Byte() Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage
            Throw New System.NotSupportedException()
        End Function

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Function IsDisplayedForItem(ByVal outlookItem As Object, ByVal formRegionMode As Microsoft.Office.Interop.Outlook.OlFormRegionMode, ByVal formRegionSize As Microsoft.Office.Interop.Outlook.OlFormRegionSize) As Boolean Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem
            Dim cancelArgs As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, False)
            cancelArgs.Cancel = False
            RaiseEvent FormRegionInitializing(Me, cancelArgs)
            Return Not cancelArgs.Cancel
        End Function

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        ReadOnly Property Kind() As Microsoft.Office.Tools.Outlook.FormRegionKindConstants Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            Get
                Return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms
            End Get
        End Property
    End Class
End Class

Partial Class WindowFormRegionCollection

    Friend ReadOnly Property FormRegion1() As FormRegion1
        Get
            For Each Item As Object In Me
                If (TypeOf (Item) Is FormRegion1) Then
                    Return Item
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class