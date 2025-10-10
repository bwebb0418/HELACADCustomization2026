<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class uscMasonry
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.Cbounit = New System.Windows.Forms.ComboBox()
        Me.cmdinsert = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Cbounit
        '
        Me.Cbounit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cbounit.FormattingEnabled = True
        Me.Cbounit.Location = New System.Drawing.Point(3, 18)
        Me.Cbounit.Name = "Cbounit"
        Me.Cbounit.Size = New System.Drawing.Size(200, 21)
        Me.Cbounit.TabIndex = 0
        '
        'cmdinsert
        '
        Me.cmdinsert.Location = New System.Drawing.Point(3, 56)
        Me.cmdinsert.Name = "cmdinsert"
        Me.cmdinsert.Size = New System.Drawing.Size(200, 23)
        Me.cmdinsert.TabIndex = 1
        Me.cmdinsert.Text = "Insert"
        Me.cmdinsert.UseVisualStyleBackColor = True
        '
        'uscMasonry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.cmdinsert)
        Me.Controls.Add(Me.Cbounit)
        Me.Name = "uscMasonry"
        Me.Size = New System.Drawing.Size(207, 150)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Cbounit As System.Windows.Forms.ComboBox
    Friend WithEvents cmdinsert As System.Windows.Forms.Button

End Class
