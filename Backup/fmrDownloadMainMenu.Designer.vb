<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDownloadMainMenu
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
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.tsbtnPost = New System.Windows.Forms.ToolStripButton
        Me.tsbtnSearch = New System.Windows.Forms.ToolStripSplitButton
        Me.tsddlManual = New System.Windows.Forms.ToolStripMenuItem
        Me.tsddlAuto = New System.Windows.Forms.ToolStripMenuItem
        Me.cboSearchCriteria = New System.Windows.Forms.ToolStripComboBox
        Me.ToolStripLabel2 = New System.Windows.Forms.ToolStripLabel
        Me.txtSearchContent = New System.Windows.Forms.ToolStripTextBox
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel
        Me.dgdDownload = New System.Windows.Forms.DataGridView
        Me.No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CReferenceNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.YA = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CompanyName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ToolStrip1.SuspendLayout()
        CType(Me.dgdDownload, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.FileName = "OpenFileDialog"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbtnPost, Me.tsbtnSearch, Me.cboSearchCriteria, Me.ToolStripLabel2, Me.txtSearchContent, Me.ToolStripLabel1})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(712, 25)
        Me.ToolStrip1.TabIndex = 9
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'tsbtnPost
        '
        Me.tsbtnPost.Image = Global.Integration.My.Resources.Resources.export
        Me.tsbtnPost.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtnPost.Name = "tsbtnPost"
        Me.tsbtnPost.Size = New System.Drawing.Size(48, 22)
        Me.tsbtnPost.Text = "Post"
        '
        'tsbtnSearch
        '
        Me.tsbtnSearch.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbtnSearch.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbtnSearch.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsddlManual, Me.tsddlAuto})
        Me.tsbtnSearch.Image = Global.Integration.My.Resources.Resources.pngSearch
        Me.tsbtnSearch.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtnSearch.Name = "tsbtnSearch"
        Me.tsbtnSearch.Size = New System.Drawing.Size(32, 22)
        Me.tsbtnSearch.Text = "ToolStripSplitButton1"
        '
        'tsddlManual
        '
        Me.tsddlManual.Name = "tsddlManual"
        Me.tsddlManual.Size = New System.Drawing.Size(144, 22)
        Me.tsddlManual.Text = "Manual Search"
        '
        'tsddlAuto
        '
        Me.tsddlAuto.Name = "tsddlAuto"
        Me.tsddlAuto.Size = New System.Drawing.Size(144, 22)
        Me.tsddlAuto.Text = "Auto Search"
        '
        'cboSearchCriteria
        '
        Me.cboSearchCriteria.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.cboSearchCriteria.Items.AddRange(New Object() {"-All Record-", "C Reference No.", "Year of Assessment"})
        Me.cboSearchCriteria.Name = "cboSearchCriteria"
        Me.cboSearchCriteria.Size = New System.Drawing.Size(121, 25)
        '
        'ToolStripLabel2
        '
        Me.ToolStripLabel2.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripLabel2.Name = "ToolStripLabel2"
        Me.ToolStripLabel2.Size = New System.Drawing.Size(59, 22)
        Me.ToolStripLabel2.Text = "Search For"
        '
        'txtSearchContent
        '
        Me.txtSearchContent.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.txtSearchContent.Name = "txtSearchContent"
        Me.txtSearchContent.Size = New System.Drawing.Size(100, 25)
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(43, 22)
        Me.ToolStripLabel1.Text = "Search "
        '
        'dgdDownload
        '
        Me.dgdDownload.AllowUserToAddRows = False
        Me.dgdDownload.AllowUserToDeleteRows = False
        Me.dgdDownload.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgdDownload.BackgroundColor = System.Drawing.Color.White
        Me.dgdDownload.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgdDownload.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.No, Me.CReferenceNo, Me.YA, Me.CompanyName})
        Me.dgdDownload.Location = New System.Drawing.Point(6, 33)
        Me.dgdDownload.MultiSelect = False
        Me.dgdDownload.Name = "dgdDownload"
        Me.dgdDownload.ReadOnly = True
        Me.dgdDownload.RowHeadersVisible = False
        Me.dgdDownload.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgdDownload.Size = New System.Drawing.Size(698, 360)
        Me.dgdDownload.TabIndex = 10
        Me.dgdDownload.TabStop = False
        '
        'No
        '
        Me.No.DataPropertyName = "RECORDCOUNT"
        Me.No.FillWeight = 56.97945!
        Me.No.HeaderText = "No."
        Me.No.Name = "No"
        Me.No.ReadOnly = True
        '
        'CReferenceNo
        '
        Me.CReferenceNo.DataPropertyName = "RefNO"
        Me.CReferenceNo.FillWeight = 92.1349!
        Me.CReferenceNo.HeaderText = "C Reference No"
        Me.CReferenceNo.Name = "CReferenceNo"
        Me.CReferenceNo.ReadOnly = True
        '
        'YA
        '
        Me.YA.DataPropertyName = "YA"
        Me.YA.FillWeight = 81.21828!
        Me.YA.HeaderText = "Year of Assessment"
        Me.YA.Name = "YA"
        Me.YA.ReadOnly = True
        '
        'CompanyName
        '
        Me.CompanyName.DataPropertyName = "comName"
        Me.CompanyName.FillWeight = 169.6674!
        Me.CompanyName.HeaderText = "Company Name"
        Me.CompanyName.Name = "CompanyName"
        Me.CompanyName.ReadOnly = True
        '
        'frmDownloadMainMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(712, 400)
        Me.Controls.Add(Me.dgdDownload)
        Me.Controls.Add(Me.ToolStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmDownloadMainMenu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Download"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        CType(Me.dgdDownload, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbtnPost As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbtnSearch As System.Windows.Forms.ToolStripSplitButton
    Friend WithEvents tsddlManual As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsddlAuto As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cboSearchCriteria As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents ToolStripLabel2 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents txtSearchContent As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents dgdDownload As System.Windows.Forms.DataGridView
    Friend WithEvents No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CReferenceNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents YA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CompanyName As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
