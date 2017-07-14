<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.grbSearch = New System.Windows.Forms.GroupBox
        Me.txtSearch = New System.Windows.Forms.TextBox
        Me.cmdSearch = New System.Windows.Forms.Button
        Me.optAll = New System.Windows.Forms.RadioButton
        Me.optYear = New System.Windows.Forms.RadioButton
        Me.optCompany = New System.Windows.Forms.RadioButton
        Me.optRefNo = New System.Windows.Forms.RadioButton
        Me.grbBrower = New System.Windows.Forms.GroupBox
        Me.cmdBrowser = New System.Windows.Forms.Button
        Me.cmdFinder = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboDLforYA = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFilePath = New System.Windows.Forms.TextBox
        Me.cmdDLFull = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.grbCommand = New System.Windows.Forms.GroupBox
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grbSearch.SuspendLayout()
        Me.grbBrower.SuspendLayout()
        Me.grbCommand.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(3, 54)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(474, 236)
        Me.DataGridView1.TabIndex = 0
        '
        'grbSearch
        '
        Me.grbSearch.Controls.Add(Me.txtSearch)
        Me.grbSearch.Controls.Add(Me.cmdSearch)
        Me.grbSearch.Controls.Add(Me.optAll)
        Me.grbSearch.Controls.Add(Me.optYear)
        Me.grbSearch.Controls.Add(Me.optCompany)
        Me.grbSearch.Controls.Add(Me.optRefNo)
        Me.grbSearch.Location = New System.Drawing.Point(5, 3)
        Me.grbSearch.Name = "grbSearch"
        Me.grbSearch.Size = New System.Drawing.Size(611, 45)
        Me.grbSearch.TabIndex = 2
        Me.grbSearch.TabStop = False
        Me.grbSearch.Text = "Search"
        '
        'txtSearch
        '
        Me.txtSearch.Location = New System.Drawing.Point(397, 18)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(169, 20)
        Me.txtSearch.TabIndex = 5
        '
        'cmdSearch
        '
        Me.cmdSearch.Location = New System.Drawing.Point(572, 15)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(33, 23)
        Me.cmdSearch.TabIndex = 4
        Me.cmdSearch.Text = "..."
        Me.cmdSearch.UseVisualStyleBackColor = True
        '
        'optAll
        '
        Me.optAll.AutoSize = True
        Me.optAll.Location = New System.Drawing.Point(291, 19)
        Me.optAll.Name = "optAll"
        Me.optAll.Size = New System.Drawing.Size(74, 17)
        Me.optAll.TabIndex = 3
        Me.optAll.TabStop = True
        Me.optAll.Text = "All Record"
        Me.optAll.UseVisualStyleBackColor = True
        '
        'optYear
        '
        Me.optYear.AutoSize = True
        Me.optYear.Location = New System.Drawing.Point(224, 19)
        Me.optYear.Name = "optYear"
        Me.optYear.Size = New System.Drawing.Size(47, 17)
        Me.optYear.TabIndex = 2
        Me.optYear.TabStop = True
        Me.optYear.Text = "Year"
        Me.optYear.UseVisualStyleBackColor = True
        '
        'optCompany
        '
        Me.optCompany.AutoSize = True
        Me.optCompany.Location = New System.Drawing.Point(107, 19)
        Me.optCompany.Name = "optCompany"
        Me.optCompany.Size = New System.Drawing.Size(100, 17)
        Me.optCompany.TabIndex = 1
        Me.optCompany.TabStop = True
        Me.optCompany.Text = "Company Name"
        Me.optCompany.UseVisualStyleBackColor = True
        '
        'optRefNo
        '
        Me.optRefNo.AutoSize = True
        Me.optRefNo.Location = New System.Drawing.Point(9, 19)
        Me.optRefNo.Name = "optRefNo"
        Me.optRefNo.Size = New System.Drawing.Size(92, 17)
        Me.optRefNo.TabIndex = 0
        Me.optRefNo.TabStop = True
        Me.optRefNo.Text = "Reference No"
        Me.optRefNo.UseVisualStyleBackColor = True
        '
        'grbBrower
        '
        Me.grbBrower.Controls.Add(Me.cmdBrowser)
        Me.grbBrower.Controls.Add(Me.cmdFinder)
        Me.grbBrower.Controls.Add(Me.Label3)
        Me.grbBrower.Controls.Add(Me.cboDLforYA)
        Me.grbBrower.Controls.Add(Me.Label2)
        Me.grbBrower.Controls.Add(Me.txtFilePath)
        Me.grbBrower.Location = New System.Drawing.Point(5, 288)
        Me.grbBrower.Name = "grbBrower"
        Me.grbBrower.Size = New System.Drawing.Size(611, 67)
        Me.grbBrower.TabIndex = 3
        Me.grbBrower.TabStop = False
        '
        'cmdBrowser
        '
        Me.cmdBrowser.Location = New System.Drawing.Point(478, 19)
        Me.cmdBrowser.Name = "cmdBrowser"
        Me.cmdBrowser.Size = New System.Drawing.Size(127, 44)
        Me.cmdBrowser.TabIndex = 28
        Me.cmdBrowser.Text = "Browser"
        Me.cmdBrowser.UseVisualStyleBackColor = True
        '
        'cmdFinder
        '
        Me.cmdFinder.BackColor = System.Drawing.Color.Transparent
        Me.cmdFinder.FlatAppearance.BorderSize = 0
        Me.cmdFinder.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdFinder.Image = Global.Integration.My.Resources.Resources.btnSearch_Image
        Me.cmdFinder.Location = New System.Drawing.Point(450, 40)
        Me.cmdFinder.Name = "cmdFinder"
        Me.cmdFinder.Size = New System.Drawing.Size(23, 21)
        Me.cmdFinder.TabIndex = 27
        Me.cmdFinder.Text = "cmdFinder"
        Me.cmdFinder.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(256, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(87, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Download for YA"
        '
        'cboDLforYA
        '
        Me.cboDLforYA.FormattingEnabled = True
        Me.cboDLforYA.Location = New System.Drawing.Point(347, 40)
        Me.cboDLforYA.Name = "cboDLforYA"
        Me.cboDLforYA.Size = New System.Drawing.Size(107, 21)
        Me.cboDLforYA.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "File Path"
        '
        'txtFilePath
        '
        Me.txtFilePath.Location = New System.Drawing.Point(63, 17)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(409, 20)
        Me.txtFilePath.TabIndex = 0
        '
        'cmdDLFull
        '
        Me.cmdDLFull.Location = New System.Drawing.Point(3, 65)
        Me.cmdDLFull.Name = "cmdDLFull"
        Me.cmdDLFull.Size = New System.Drawing.Size(126, 56)
        Me.cmdDLFull.TabIndex = 2
        Me.cmdDLFull.Text = "Download Full Form C"
        Me.cmdDLFull.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(2, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(131, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "PDF BORANG C 2008"
        '
        'grbCommand
        '
        Me.grbCommand.Controls.Add(Me.Label1)
        Me.grbCommand.Controls.Add(Me.cmdDLFull)
        Me.grbCommand.Location = New System.Drawing.Point(483, 48)
        Me.grbCommand.Name = "grbCommand"
        Me.grbCommand.Size = New System.Drawing.Size(133, 242)
        Me.grbCommand.TabIndex = 1
        Me.grbCommand.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(623, 358)
        Me.Controls.Add(Me.grbSearch)
        Me.Controls.Add(Me.grbCommand)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.grbBrower)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grbSearch.ResumeLayout(False)
        Me.grbSearch.PerformLayout()
        Me.grbBrower.ResumeLayout(False)
        Me.grbBrower.PerformLayout()
        Me.grbCommand.ResumeLayout(False)
        Me.grbCommand.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents grbSearch As System.Windows.Forms.GroupBox
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents optAll As System.Windows.Forms.RadioButton
    Friend WithEvents optYear As System.Windows.Forms.RadioButton
    Friend WithEvents optCompany As System.Windows.Forms.RadioButton
    Friend WithEvents optRefNo As System.Windows.Forms.RadioButton
    Friend WithEvents grbBrower As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboDLforYA As System.Windows.Forms.ComboBox
    Friend WithEvents cmdFinder As System.Windows.Forms.Button
    Friend WithEvents cmdBrowser As System.Windows.Forms.Button
    Friend WithEvents cmdDLFull As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grbCommand As System.Windows.Forms.GroupBox

End Class
