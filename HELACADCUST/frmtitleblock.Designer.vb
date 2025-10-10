<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTitleblock
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
        Me.Cbosize = New System.Windows.Forms.ComboBox()
        Me.cmdGo = New System.Windows.Forms.Button()
        Me.cmdCAN = New System.Windows.Forms.Button()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Cbosize
        '
        Me.Cbosize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cbosize.FormattingEnabled = True
        Me.Cbosize.Location = New System.Drawing.Point(12, 12)
        Me.Cbosize.Name = "Cbosize"
        Me.Cbosize.Size = New System.Drawing.Size(277, 21)
        Me.Cbosize.TabIndex = 0
        '
        'cmdGo
        '
        Me.cmdGo.Location = New System.Drawing.Point(184, 78)
        Me.cmdGo.Name = "cmdGo"
        Me.cmdGo.Size = New System.Drawing.Size(75, 23)
        Me.cmdGo.TabIndex = 19
        Me.cmdGo.Text = "Go!"
        Me.cmdGo.UseVisualStyleBackColor = True
        '
        'cmdCAN
        '
        Me.cmdCAN.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCAN.Location = New System.Drawing.Point(51, 78)
        Me.cmdCAN.Name = "cmdCAN"
        Me.cmdCAN.Size = New System.Drawing.Size(75, 23)
        Me.cmdCAN.TabIndex = 20
        Me.cmdCAN.Text = "Cancel"
        Me.cmdCAN.UseVisualStyleBackColor = True
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(12, 52)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(187, 13)
        Me.Label19.TabIndex = 39
        Me.Label19.Text = "Note:  These are only HEL Titleblocks"
        '
        'FrmTitleblock
        '
        Me.AcceptButton = Me.cmdGo
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCAN
        Me.ClientSize = New System.Drawing.Size(311, 127)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.cmdCAN)
        Me.Controls.Add(Me.cmdGo)
        Me.Controls.Add(Me.Cbosize)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTitleblock"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Titleblock"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Cbosize As System.Windows.Forms.ComboBox
    Friend WithEvents cmdGo As System.Windows.Forms.Button
    Friend WithEvents cmdCAN As System.Windows.Forms.Button
    Friend WithEvents Label19 As System.Windows.Forms.Label
End Class
