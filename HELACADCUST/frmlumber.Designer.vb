<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmlumber
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rdoDLEDGE = New System.Windows.Forms.RadioButton()
        Me.rdoDLFLAT = New System.Windows.Forms.RadioButton()
        Me.CboInsert = New System.Windows.Forms.ComboBox()
        Me.CboDLSIZE = New System.Windows.Forms.ComboBox()
        Me.CboRSSIZE = New System.Windows.Forms.ComboBox()
        Me.CboRSINSERT = New System.Windows.Forms.ComboBox()
        Me.rdoRSFLAT = New System.Windows.Forms.RadioButton()
        Me.rdoRSEDGE = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CboDLSIZE)
        Me.GroupBox1.Controls.Add(Me.CboInsert)
        Me.GroupBox1.Controls.Add(Me.rdoDLFLAT)
        Me.GroupBox1.Controls.Add(Me.rdoDLEDGE)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 11)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(270, 109)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Dimensional Lumber"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.CboRSSIZE)
        Me.GroupBox2.Controls.Add(Me.CboRSINSERT)
        Me.GroupBox2.Controls.Add(Me.rdoRSFLAT)
        Me.GroupBox2.Controls.Add(Me.rdoRSEDGE)
        Me.GroupBox2.Location = New System.Drawing.Point(7, 120)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(270, 109)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Rough Sawn Lumber"
        '
        'rdoDLEDGE
        '
        Me.rdoDLEDGE.AutoSize = True
        Me.rdoDLEDGE.Location = New System.Drawing.Point(6, 86)
        Me.rdoDLEDGE.Name = "rdoDLEDGE"
        Me.rdoDLEDGE.Size = New System.Drawing.Size(67, 17)
        Me.rdoDLEDGE.TabIndex = 0
        Me.rdoDLEDGE.TabStop = True
        Me.rdoDLEDGE.Text = "On Edge"
        Me.rdoDLEDGE.UseVisualStyleBackColor = True
        '
        'rdoDLFLAT
        '
        Me.rdoDLFLAT.AutoSize = True
        Me.rdoDLFLAT.Location = New System.Drawing.Point(79, 86)
        Me.rdoDLFLAT.Name = "rdoDLFLAT"
        Me.rdoDLFLAT.Size = New System.Drawing.Size(59, 17)
        Me.rdoDLFLAT.TabIndex = 1
        Me.rdoDLFLAT.TabStop = True
        Me.rdoDLFLAT.Text = "On Flat"
        Me.rdoDLFLAT.UseVisualStyleBackColor = True
        '
        'CboInsert
        '
        Me.CboInsert.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboInsert.FormattingEnabled = True
        Me.CboInsert.Location = New System.Drawing.Point(6, 59)
        Me.CboInsert.Name = "CboInsert"
        Me.CboInsert.Size = New System.Drawing.Size(239, 21)
        Me.CboInsert.TabIndex = 2
        '
        'CboDLSIZE
        '
        Me.CboDLSIZE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboDLSIZE.FormattingEnabled = True
        Me.CboDLSIZE.Location = New System.Drawing.Point(6, 19)
        Me.CboDLSIZE.Name = "CboDLSIZE"
        Me.CboDLSIZE.Size = New System.Drawing.Size(239, 21)
        Me.CboDLSIZE.TabIndex = 3
        '
        'CboRSSIZE
        '
        Me.CboRSSIZE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboRSSIZE.FormattingEnabled = True
        Me.CboRSSIZE.Location = New System.Drawing.Point(6, 19)
        Me.CboRSSIZE.Name = "CboRSSIZE"
        Me.CboRSSIZE.Size = New System.Drawing.Size(239, 21)
        Me.CboRSSIZE.TabIndex = 7
        '
        'CboRSINSERT
        '
        Me.CboRSINSERT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboRSINSERT.FormattingEnabled = True
        Me.CboRSINSERT.Location = New System.Drawing.Point(6, 59)
        Me.CboRSINSERT.Name = "CboRSINSERT"
        Me.CboRSINSERT.Size = New System.Drawing.Size(239, 21)
        Me.CboRSINSERT.TabIndex = 6
        '
        'rdoRSFLAT
        '
        Me.rdoRSFLAT.AutoSize = True
        Me.rdoRSFLAT.Location = New System.Drawing.Point(79, 86)
        Me.rdoRSFLAT.Name = "rdoRSFLAT"
        Me.rdoRSFLAT.Size = New System.Drawing.Size(59, 17)
        Me.rdoRSFLAT.TabIndex = 5
        Me.rdoRSFLAT.TabStop = True
        Me.rdoRSFLAT.Text = "On Flat"
        Me.rdoRSFLAT.UseVisualStyleBackColor = True
        '
        'rdoRSEDGE
        '
        Me.rdoRSEDGE.AutoSize = True
        Me.rdoRSEDGE.Location = New System.Drawing.Point(6, 86)
        Me.rdoRSEDGE.Name = "rdoRSEDGE"
        Me.rdoRSEDGE.Size = New System.Drawing.Size(67, 17)
        Me.rdoRSEDGE.TabIndex = 4
        Me.rdoRSEDGE.TabStop = True
        Me.rdoRSEDGE.Text = "On Edge"
        Me.rdoRSEDGE.UseVisualStyleBackColor = True
        '
        'frmlumber
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 237)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmlumber"
        Me.Text = "Form1"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoDLEDGE As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents CboInsert As System.Windows.Forms.ComboBox
    Friend WithEvents rdoDLFLAT As System.Windows.Forms.RadioButton
    Friend WithEvents CboDLSIZE As System.Windows.Forms.ComboBox
    Friend WithEvents CboRSSIZE As System.Windows.Forms.ComboBox
    Friend WithEvents CboRSINSERT As System.Windows.Forms.ComboBox
    Friend WithEvents rdoRSFLAT As System.Windows.Forms.RadioButton
    Friend WithEvents rdoRSEDGE As System.Windows.Forms.RadioButton
End Class
