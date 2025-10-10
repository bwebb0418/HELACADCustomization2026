<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmdynamicblock
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TXTline1 = New System.Windows.Forms.TextBox()
        Me.TXTLine2 = New System.Windows.Forms.TextBox()
        Me.Cbovisopt = New System.Windows.Forms.ComboBox()
        Me.Cbolookopt = New System.Windows.Forms.ComboBox()
        Me.cmdfinish = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(26, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Line 1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(26, 84)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Line 2"
        '
        'TXTline1
        '
        Me.TXTline1.Location = New System.Drawing.Point(68, 41)
        Me.TXTline1.Name = "TXTline1"
        Me.TXTline1.Size = New System.Drawing.Size(100, 20)
        Me.TXTline1.TabIndex = 2
        '
        'TXTLine2
        '
        Me.TXTLine2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TXTLine2.Location = New System.Drawing.Point(68, 81)
        Me.TXTLine2.Name = "TXTLine2"
        Me.TXTLine2.Size = New System.Drawing.Size(100, 20)
        Me.TXTLine2.TabIndex = 3
        '
        'Cbovisopt
        '
        Me.Cbovisopt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cbovisopt.FormattingEnabled = True
        Me.Cbovisopt.Location = New System.Drawing.Point(29, 145)
        Me.Cbovisopt.Name = "Cbovisopt"
        Me.Cbovisopt.Size = New System.Drawing.Size(232, 21)
        Me.Cbovisopt.TabIndex = 4
        '
        'Cbolookopt
        '
        Me.Cbolookopt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cbolookopt.FormattingEnabled = True
        Me.Cbolookopt.Location = New System.Drawing.Point(29, 187)
        Me.Cbolookopt.Name = "Cbolookopt"
        Me.Cbolookopt.Size = New System.Drawing.Size(232, 21)
        Me.Cbolookopt.TabIndex = 5
        '
        'cmdfinish
        '
        Me.cmdfinish.Location = New System.Drawing.Point(197, 229)
        Me.cmdfinish.Name = "cmdfinish"
        Me.cmdfinish.Size = New System.Drawing.Size(75, 23)
        Me.cmdfinish.TabIndex = 6
        Me.cmdfinish.Text = "Finish"
        Me.cmdfinish.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(29, 228)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'frmdynamicblock
        '
        Me.AcceptButton = Me.cmdfinish
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(284, 264)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cmdfinish)
        Me.Controls.Add(Me.Cbolookopt)
        Me.Controls.Add(Me.Cbovisopt)
        Me.Controls.Add(Me.TXTLine2)
        Me.Controls.Add(Me.TXTline1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmdynamicblock"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TXTline1 As System.Windows.Forms.TextBox
    Friend WithEvents TXTLine2 As System.Windows.Forms.TextBox
    Friend WithEvents Cbovisopt As System.Windows.Forms.ComboBox
    Friend WithEvents Cbolookopt As System.Windows.Forms.ComboBox
    Friend WithEvents cmdfinish As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
