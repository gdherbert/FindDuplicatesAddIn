<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFindDups
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnGetFields = New System.Windows.Forms.Button
        Me.btnFindDups = New System.Windows.Forms.Button
        Me.cbVisExtent = New System.Windows.Forms.CheckBox
        Me.cbCaseSensitive = New System.Windows.Forms.CheckBox
        Me.cbSelectAll = New System.Windows.Forms.CheckBox
        Me.lbFields = New System.Windows.Forms.ListBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnHelp = New System.Windows.Forms.Button
        Me.btnRestart = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnGetFields
        '
        Me.btnGetFields.Location = New System.Drawing.Point(12, 289)
        Me.btnGetFields.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnGetFields.Name = "btnGetFields"
        Me.btnGetFields.Size = New System.Drawing.Size(109, 28)
        Me.btnGetFields.TabIndex = 0
        Me.btnGetFields.Text = "Get Field List"
        Me.btnGetFields.UseVisualStyleBackColor = True
        '
        'btnFindDups
        '
        Me.btnFindDups.Location = New System.Drawing.Point(135, 289)
        Me.btnFindDups.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFindDups.Name = "btnFindDups"
        Me.btnFindDups.Size = New System.Drawing.Size(124, 28)
        Me.btnFindDups.TabIndex = 1
        Me.btnFindDups.Text = "Find Duplicates"
        Me.btnFindDups.UseVisualStyleBackColor = True
        '
        'cbVisExtent
        '
        Me.cbVisExtent.AutoSize = True
        Me.cbVisExtent.Location = New System.Drawing.Point(17, 261)
        Me.cbVisExtent.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbVisExtent.Name = "cbVisExtent"
        Me.cbVisExtent.Size = New System.Drawing.Size(182, 21)
        Me.cbVisExtent.TabIndex = 2
        Me.cbVisExtent.Text = "Restrict to Visible Extent"
        Me.cbVisExtent.UseVisualStyleBackColor = True
        '
        'cbCaseSensitive
        '
        Me.cbCaseSensitive.AutoSize = True
        Me.cbCaseSensitive.Checked = True
        Me.cbCaseSensitive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbCaseSensitive.Location = New System.Drawing.Point(17, 233)
        Me.cbCaseSensitive.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbCaseSensitive.Name = "cbCaseSensitive"
        Me.cbCaseSensitive.Size = New System.Drawing.Size(123, 21)
        Me.cbCaseSensitive.TabIndex = 3
        Me.cbCaseSensitive.Text = "Case Sensitive"
        Me.cbCaseSensitive.UseVisualStyleBackColor = True
        '
        'cbSelectAll
        '
        Me.cbSelectAll.AutoSize = True
        Me.cbSelectAll.Location = New System.Drawing.Point(153, 233)
        Me.cbSelectAll.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbSelectAll.Name = "cbSelectAll"
        Me.cbSelectAll.Size = New System.Drawing.Size(88, 21)
        Me.cbSelectAll.TabIndex = 4
        Me.cbSelectAll.Text = "Select All"
        Me.cbSelectAll.UseVisualStyleBackColor = True
        '
        'lbFields
        '
        Me.lbFields.FormattingEnabled = True
        Me.lbFields.ItemHeight = 16
        Me.lbFields.Location = New System.Drawing.Point(15, 2)
        Me.lbFields.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.lbFields.Name = "lbFields"
        Me.lbFields.Size = New System.Drawing.Size(340, 212)
        Me.lbFields.TabIndex = 5
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 332)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(372, 27)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 6
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(50, 22)
        Me.ToolStripStatusLabel1.Text = "Ready"
        '
        'btnHelp
        '
        Me.btnHelp.Location = New System.Drawing.Point(272, 225)
        Me.btnHelp.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.Size = New System.Drawing.Size(84, 28)
        Me.btnHelp.TabIndex = 7
        Me.btnHelp.Text = "Help"
        Me.btnHelp.UseVisualStyleBackColor = True
        '
        'btnRestart
        '
        Me.btnRestart.Location = New System.Drawing.Point(272, 289)
        Me.btnRestart.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnRestart.Name = "btnRestart"
        Me.btnRestart.Size = New System.Drawing.Size(84, 28)
        Me.btnRestart.TabIndex = 8
        Me.btnRestart.Text = "Reset"
        Me.btnRestart.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(268, 256)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 36)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "(Reset before running again)"
        '
        'frmFindDups
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(372, 359)
        Me.Controls.Add(Me.btnRestart)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnHelp)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.lbFields)
        Me.Controls.Add(Me.cbSelectAll)
        Me.Controls.Add(Me.cbCaseSensitive)
        Me.Controls.Add(Me.cbVisExtent)
        Me.Controls.Add(Me.btnFindDups)
        Me.Controls.Add(Me.btnGetFields)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFindDups"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select Duplicate Records in Table"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnGetFields As System.Windows.Forms.Button
    Friend WithEvents btnFindDups As System.Windows.Forms.Button
    Friend WithEvents cbVisExtent As System.Windows.Forms.CheckBox
    Friend WithEvents cbCaseSensitive As System.Windows.Forms.CheckBox
    Friend WithEvents cbSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents lbFields As System.Windows.Forms.ListBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnHelp As System.Windows.Forms.Button
    Friend WithEvents btnRestart As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
