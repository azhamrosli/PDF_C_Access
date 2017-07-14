<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDownloadDetails
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
        Me.grbAuditor = New System.Windows.Forms.GroupBox
        Me.lstViewAuditor = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.grbDirector = New System.Windows.Forms.GroupBox
        Me.lstViewDirector = New System.Windows.Forms.ListView
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.grbAkuan = New System.Windows.Forms.GroupBox
        Me.dtpPrintDateA = New System.Windows.Forms.DateTimePicker
        Me.lblTaxAgent = New System.Windows.Forms.Label
        Me.dtpPrintDate = New System.Windows.Forms.DateTimePicker
        Me.grbRKST = New System.Windows.Forms.GroupBox
        Me.optMenyerah = New System.Windows.Forms.RadioButton
        Me.optMenuntut = New System.Windows.Forms.RadioButton
        Me.chkRKST = New System.Windows.Forms.CheckBox
        Me.chkKeepRecord = New System.Windows.Forms.CheckBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtPosition = New System.Windows.Forms.TextBox
        Me.txtIC = New System.Windows.Forms.TextBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.tsbtnPost = New System.Windows.Forms.ToolStripButton
        Me.grbTaxAgent = New System.Windows.Forms.GroupBox
        Me.lstViewTaxAgent = New System.Windows.Forms.ListView
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.grbAuditor.SuspendLayout()
        Me.grbDirector.SuspendLayout()
        Me.grbAkuan.SuspendLayout()
        Me.grbRKST.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.grbTaxAgent.SuspendLayout()
        Me.SuspendLayout()
        '
        'grbAuditor
        '
        Me.grbAuditor.Controls.Add(Me.lstViewAuditor)
        Me.grbAuditor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grbAuditor.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.grbAuditor.Location = New System.Drawing.Point(10, 45)
        Me.grbAuditor.Name = "grbAuditor"
        Me.grbAuditor.Size = New System.Drawing.Size(546, 118)
        Me.grbAuditor.TabIndex = 2
        Me.grbAuditor.TabStop = False
        Me.grbAuditor.Text = "Select Auditor Profile to export into Borang C :"
        '
        'lstViewAuditor
        '
        Me.lstViewAuditor.AllowColumnReorder = True
        Me.lstViewAuditor.AutoArrange = False
        Me.lstViewAuditor.CheckBoxes = True
        Me.lstViewAuditor.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.lstViewAuditor.FullRowSelect = True
        Me.lstViewAuditor.HideSelection = False
        Me.lstViewAuditor.Location = New System.Drawing.Point(6, 19)
        Me.lstViewAuditor.MultiSelect = False
        Me.lstViewAuditor.Name = "lstViewAuditor"
        Me.lstViewAuditor.Size = New System.Drawing.Size(528, 90)
        Me.lstViewAuditor.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lstViewAuditor.TabIndex = 0
        Me.lstViewAuditor.TabStop = False
        Me.lstViewAuditor.UseCompatibleStateImageBehavior = False
        Me.lstViewAuditor.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "No"
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Company Name"
        Me.ColumnHeader2.Width = 230
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Address"
        Me.ColumnHeader3.Width = 228
        '
        'grbDirector
        '
        Me.grbDirector.Controls.Add(Me.lstViewDirector)
        Me.grbDirector.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grbDirector.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.grbDirector.Location = New System.Drawing.Point(9, 300)
        Me.grbDirector.Name = "grbDirector"
        Me.grbDirector.Size = New System.Drawing.Size(547, 113)
        Me.grbDirector.TabIndex = 3
        Me.grbDirector.TabStop = False
        Me.grbDirector.Text = "Select Director Profile to export into the Borang C Akuan:"
        '
        'lstViewDirector
        '
        Me.lstViewDirector.CheckBoxes = True
        Me.lstViewDirector.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.lstViewDirector.FullRowSelect = True
        Me.lstViewDirector.Location = New System.Drawing.Point(7, 19)
        Me.lstViewDirector.MultiSelect = False
        Me.lstViewDirector.Name = "lstViewDirector"
        Me.lstViewDirector.Size = New System.Drawing.Size(528, 90)
        Me.lstViewDirector.TabIndex = 0
        Me.lstViewDirector.TabStop = False
        Me.lstViewDirector.UseCompatibleStateImageBehavior = False
        Me.lstViewDirector.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "No"
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Director Name"
        Me.ColumnHeader5.Width = 230
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "IC No"
        Me.ColumnHeader6.Width = 229
        '
        'grbAkuan
        '
        Me.grbAkuan.Controls.Add(Me.dtpPrintDateA)
        Me.grbAkuan.Controls.Add(Me.lblTaxAgent)
        Me.grbAkuan.Controls.Add(Me.dtpPrintDate)
        Me.grbAkuan.Controls.Add(Me.grbRKST)
        Me.grbAkuan.Controls.Add(Me.chkKeepRecord)
        Me.grbAkuan.Controls.Add(Me.Label5)
        Me.grbAkuan.Controls.Add(Me.Label4)
        Me.grbAkuan.Controls.Add(Me.txtPosition)
        Me.grbAkuan.Controls.Add(Me.txtIC)
        Me.grbAkuan.Controls.Add(Me.txtName)
        Me.grbAkuan.Controls.Add(Me.Label3)
        Me.grbAkuan.Controls.Add(Me.Label2)
        Me.grbAkuan.Controls.Add(Me.Label1)
        Me.grbAkuan.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grbAkuan.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.grbAkuan.Location = New System.Drawing.Point(10, 419)
        Me.grbAkuan.Name = "grbAkuan"
        Me.grbAkuan.Size = New System.Drawing.Size(546, 191)
        Me.grbAkuan.TabIndex = 4
        Me.grbAkuan.TabStop = False
        Me.grbAkuan.Text = "Borang C Akuan"
        '
        'dtpPrintDateA
        '
        Me.dtpPrintDateA.Checked = False
        Me.dtpPrintDateA.CustomFormat = "dd MMM yyyy"
        Me.dtpPrintDateA.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPrintDateA.Location = New System.Drawing.Point(101, 134)
        Me.dtpPrintDateA.Name = "dtpPrintDateA"
        Me.dtpPrintDateA.ShowCheckBox = True
        Me.dtpPrintDateA.Size = New System.Drawing.Size(117, 20)
        Me.dtpPrintDateA.TabIndex = 7
        Me.dtpPrintDateA.Value = New Date(2008, 8, 7, 0, 0, 0, 0)
        '
        'lblTaxAgent
        '
        Me.lblTaxAgent.AutoSize = True
        Me.lblTaxAgent.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTaxAgent.Location = New System.Drawing.Point(10, 138)
        Me.lblTaxAgent.Name = "lblTaxAgent"
        Me.lblTaxAgent.Size = New System.Drawing.Size(56, 13)
        Me.lblTaxAgent.TabIndex = 12
        Me.lblTaxAgent.Text = "Tax Agent"
        '
        'dtpPrintDate
        '
        Me.dtpPrintDate.CustomFormat = "dd MMM yyyy"
        Me.dtpPrintDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPrintDate.Location = New System.Drawing.Point(101, 108)
        Me.dtpPrintDate.Name = "dtpPrintDate"
        Me.dtpPrintDate.ShowCheckBox = True
        Me.dtpPrintDate.Size = New System.Drawing.Size(117, 20)
        Me.dtpPrintDate.TabIndex = 4
        '
        'grbRKST
        '
        Me.grbRKST.Controls.Add(Me.optMenyerah)
        Me.grbRKST.Controls.Add(Me.optMenuntut)
        Me.grbRKST.Controls.Add(Me.chkRKST)
        Me.grbRKST.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grbRKST.Location = New System.Drawing.Point(248, 102)
        Me.grbRKST.Name = "grbRKST"
        Me.grbRKST.Size = New System.Drawing.Size(225, 49)
        Me.grbRKST.TabIndex = 10
        Me.grbRKST.TabStop = False
        '
        'optMenyerah
        '
        Me.optMenyerah.AutoSize = True
        Me.optMenyerah.Location = New System.Drawing.Point(123, 23)
        Me.optMenyerah.Name = "optMenyerah"
        Me.optMenyerah.Size = New System.Drawing.Size(72, 17)
        Me.optMenyerah.TabIndex = 2
        Me.optMenyerah.Text = "Menyerah"
        Me.optMenyerah.UseVisualStyleBackColor = True
        '
        'optMenuntut
        '
        Me.optMenuntut.AutoSize = True
        Me.optMenuntut.Checked = True
        Me.optMenuntut.Location = New System.Drawing.Point(36, 23)
        Me.optMenuntut.Name = "optMenuntut"
        Me.optMenuntut.Size = New System.Drawing.Size(70, 17)
        Me.optMenuntut.TabIndex = 1
        Me.optMenuntut.TabStop = True
        Me.optMenuntut.Text = "Menuntut"
        Me.optMenuntut.UseVisualStyleBackColor = True
        '
        'chkRKST
        '
        Me.chkRKST.AutoSize = True
        Me.chkRKST.Location = New System.Drawing.Point(6, 9)
        Me.chkRKST.Name = "chkRKST"
        Me.chkRKST.Size = New System.Drawing.Size(91, 17)
        Me.chkRKST.TabIndex = 6
        Me.chkRKST.Text = "RK-S or RK-T"
        Me.chkRKST.UseVisualStyleBackColor = True
        '
        'chkKeepRecord
        '
        Me.chkKeepRecord.AutoSize = True
        Me.chkKeepRecord.Checked = True
        Me.chkKeepRecord.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkKeepRecord.Location = New System.Drawing.Point(101, 166)
        Me.chkKeepRecord.Name = "chkKeepRecord"
        Me.chkKeepRecord.Size = New System.Drawing.Size(15, 14)
        Me.chkKeepRecord.TabIndex = 5
        Me.chkKeepRecord.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(10, 166)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(70, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Keep Record"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(10, 112)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(33, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Date:"
        '
        'txtPosition
        '
        Me.txtPosition.Location = New System.Drawing.Point(101, 66)
        Me.txtPosition.Name = "txtPosition"
        Me.txtPosition.Size = New System.Drawing.Size(212, 20)
        Me.txtPosition.TabIndex = 3
        '
        'txtIC
        '
        Me.txtIC.Location = New System.Drawing.Point(101, 44)
        Me.txtIC.Name = "txtIC"
        Me.txtIC.Size = New System.Drawing.Size(212, 20)
        Me.txtIC.TabIndex = 2
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(101, 22)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(212, 20)
        Me.txtName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(10, 69)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(47, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Position:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(10, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "IC No:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(10, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Name:"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbtnPost})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(570, 25)
        Me.ToolStrip1.TabIndex = 6
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'tsbtnPost
        '
        Me.tsbtnPost.Image = Global.Integration.My.Resources.Resources.export
        Me.tsbtnPost.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtnPost.Name = "tsbtnPost"
        Me.tsbtnPost.Size = New System.Drawing.Size(110, 22)
        Me.tsbtnPost.Text = "  Post Full Form C"
        '
        'grbTaxAgent
        '
        Me.grbTaxAgent.Controls.Add(Me.lstViewTaxAgent)
        Me.grbTaxAgent.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grbTaxAgent.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.grbTaxAgent.Location = New System.Drawing.Point(10, 169)
        Me.grbTaxAgent.Name = "grbTaxAgent"
        Me.grbTaxAgent.Size = New System.Drawing.Size(546, 118)
        Me.grbTaxAgent.TabIndex = 3
        Me.grbTaxAgent.TabStop = False
        Me.grbTaxAgent.Text = "Select Tax Agent Profile to export into Borang C :"
        '
        'lstViewTaxAgent
        '
        Me.lstViewTaxAgent.AllowColumnReorder = True
        Me.lstViewTaxAgent.AutoArrange = False
        Me.lstViewTaxAgent.CheckBoxes = True
        Me.lstViewTaxAgent.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9})
        Me.lstViewTaxAgent.FullRowSelect = True
        Me.lstViewTaxAgent.HideSelection = False
        Me.lstViewTaxAgent.Location = New System.Drawing.Point(6, 19)
        Me.lstViewTaxAgent.MultiSelect = False
        Me.lstViewTaxAgent.Name = "lstViewTaxAgent"
        Me.lstViewTaxAgent.Size = New System.Drawing.Size(528, 90)
        Me.lstViewTaxAgent.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lstViewTaxAgent.TabIndex = 0
        Me.lstViewTaxAgent.TabStop = False
        Me.lstViewTaxAgent.UseCompatibleStateImageBehavior = False
        Me.lstViewTaxAgent.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "No"
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Company Name"
        Me.ColumnHeader8.Width = 230
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Address"
        Me.ColumnHeader9.Width = 228
        '
        'frmDownloadDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(570, 620)
        Me.Controls.Add(Me.grbTaxAgent)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.grbAkuan)
        Me.Controls.Add(Me.grbDirector)
        Me.Controls.Add(Me.grbAuditor)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmDownloadDetails"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Download - Details"
        Me.grbAuditor.ResumeLayout(False)
        Me.grbDirector.ResumeLayout(False)
        Me.grbAkuan.ResumeLayout(False)
        Me.grbAkuan.PerformLayout()
        Me.grbRKST.ResumeLayout(False)
        Me.grbRKST.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.grbTaxAgent.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grbAuditor As System.Windows.Forms.GroupBox
    Friend WithEvents grbDirector As System.Windows.Forms.GroupBox
    Friend WithEvents grbAkuan As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtPosition As System.Windows.Forms.TextBox
    Friend WithEvents txtIC As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents grbRKST As System.Windows.Forms.GroupBox
    Friend WithEvents optMenuntut As System.Windows.Forms.RadioButton
    Friend WithEvents chkRKST As System.Windows.Forms.CheckBox
    Friend WithEvents chkKeepRecord As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents optMenyerah As System.Windows.Forms.RadioButton
    Friend WithEvents lstViewAuditor As System.Windows.Forms.ListView
    Friend WithEvents lstViewDirector As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents dtpPrintDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbtnPost As System.Windows.Forms.ToolStripButton
    Friend WithEvents dtpPrintDateA As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTaxAgent As System.Windows.Forms.Label
    Friend WithEvents grbTaxAgent As System.Windows.Forms.GroupBox
    Friend WithEvents lstViewTaxAgent As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
End Class
