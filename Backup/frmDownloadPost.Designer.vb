<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDownloadPost
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
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtSaveFile = New System.Windows.Forms.TextBox
        Me.btnSaveFile = New System.Windows.Forms.Button
        Me.btnSearchFile = New System.Windows.Forms.Button
        Me.txtOpenFile = New System.Windows.Forms.TextBox
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.tsbtnPost = New System.Windows.Forms.ToolStripButton
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog = New System.Windows.Forms.SaveFileDialog
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblProgress = New System.Windows.Forms.Label
        Me.lblRemark = New System.Windows.Forms.Label
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 90)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Save File :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 65)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Open File :"
        '
        'txtSaveFile
        '
        Me.txtSaveFile.BackColor = System.Drawing.SystemColors.Control
        Me.txtSaveFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtSaveFile.Location = New System.Drawing.Point(82, 90)
        Me.txtSaveFile.Name = "txtSaveFile"
        Me.txtSaveFile.ReadOnly = True
        Me.txtSaveFile.Size = New System.Drawing.Size(285, 20)
        Me.txtSaveFile.TabIndex = 3
        Me.txtSaveFile.TabStop = False
        '
        'btnSaveFile
        '
        Me.btnSaveFile.Location = New System.Drawing.Point(377, 88)
        Me.btnSaveFile.Name = "btnSaveFile"
        Me.btnSaveFile.Size = New System.Drawing.Size(25, 25)
        Me.btnSaveFile.TabIndex = 2
        Me.btnSaveFile.Text = "..."
        Me.btnSaveFile.UseVisualStyleBackColor = True
        '
        'btnSearchFile
        '
        Me.btnSearchFile.Location = New System.Drawing.Point(377, 61)
        Me.btnSearchFile.Name = "btnSearchFile"
        Me.btnSearchFile.Size = New System.Drawing.Size(25, 25)
        Me.btnSearchFile.TabIndex = 1
        Me.btnSearchFile.Text = "..."
        Me.btnSearchFile.UseVisualStyleBackColor = True
        '
        'txtOpenFile
        '
        Me.txtOpenFile.BackColor = System.Drawing.SystemColors.Control
        Me.txtOpenFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtOpenFile.Location = New System.Drawing.Point(81, 63)
        Me.txtOpenFile.Name = "txtOpenFile"
        Me.txtOpenFile.ReadOnly = True
        Me.txtOpenFile.Size = New System.Drawing.Size(286, 20)
        Me.txtOpenFile.TabIndex = 2
        Me.txtOpenFile.TabStop = False
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbtnPost})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(435, 25)
        Me.ToolStrip1.TabIndex = 3
        Me.ToolStrip1.TabStop = True
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
        'OpenFileDialog
        '
        Me.OpenFileDialog.FileName = "OpenFileDialog"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label3.Location = New System.Drawing.Point(12, 37)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Export to PDF"
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(12, 130)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(0, 13)
        Me.lblProgress.TabIndex = 7
        '
        'lblRemark
        '
        Me.lblRemark.BackColor = System.Drawing.Color.Transparent
        Me.lblRemark.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRemark.ForeColor = System.Drawing.Color.Red
        Me.lblRemark.Location = New System.Drawing.Point(17, 122)
        Me.lblRemark.Name = "lblRemark"
        Me.lblRemark.Size = New System.Drawing.Size(405, 39)
        Me.lblRemark.TabIndex = 14
        Me.lblRemark.Text = "Remark: YA 2009 and onwards, Form C PDF format is not for submission to LHDNM by " & _
            "Tax Agent"
        '
        'frmDownloadPost
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(435, 163)
        Me.Controls.Add(Me.lblRemark)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnSaveFile)
        Me.Controls.Add(Me.btnSearchFile)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.txtSaveFile)
        Me.Controls.Add(Me.txtOpenFile)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmDownloadPost"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Download - Post"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSaveFile As System.Windows.Forms.TextBox
    Friend WithEvents btnSaveFile As System.Windows.Forms.Button
    Friend WithEvents btnSearchFile As System.Windows.Forms.Button
    Friend WithEvents txtOpenFile As System.Windows.Forms.TextBox
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbtnPost As System.Windows.Forms.ToolStripButton
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents lblRemark As System.Windows.Forms.Label
End Class
