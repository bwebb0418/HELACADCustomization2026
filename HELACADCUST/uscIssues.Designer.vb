<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class uscIssues
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
        Me.cmdIFRS = New System.Windows.Forms.Button()
        Me.cmdift = New System.Windows.Forms.Button()
        Me.cmdIFC = New System.Windows.Forms.Button()
        Me.cmdIFI = New System.Windows.Forms.Button()
        Me.cmdifr = New System.Windows.Forms.Button()
        Me.rdo50 = New System.Windows.Forms.RadioButton()
        Me.rdo100 = New System.Windows.Forms.RadioButton()
        Me.rdo75 = New System.Windows.Forms.RadioButton()
        Me.rdofinal = New System.Windows.Forms.RadioButton()
        Me.rdoother = New System.Windows.Forms.RadioButton()
        Me.cmdifa = New System.Windows.Forms.Button()
        Me.dtpIssueDate = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdIFRS
        '
        Me.cmdIFRS.Location = New System.Drawing.Point(3, 118)
        Me.cmdIFRS.Name = "cmdIFRS"
        Me.cmdIFRS.Size = New System.Drawing.Size(147, 23)
        Me.cmdIFRS.TabIndex = 0
        Me.cmdIFRS.Text = "Review w/ Stamp"
        Me.cmdIFRS.UseVisualStyleBackColor = True
        '
        'cmdift
        '
        Me.cmdift.Location = New System.Drawing.Point(3, 225)
        Me.cmdift.Name = "cmdift"
        Me.cmdift.Size = New System.Drawing.Size(147, 23)
        Me.cmdift.TabIndex = 1
        Me.cmdift.Text = "Tender"
        Me.cmdift.UseVisualStyleBackColor = True
        '
        'cmdIFC
        '
        Me.cmdIFC.Location = New System.Drawing.Point(3, 267)
        Me.cmdIFC.Name = "cmdIFC"
        Me.cmdIFC.Size = New System.Drawing.Size(147, 23)
        Me.cmdIFC.TabIndex = 2
        Me.cmdIFC.Text = "Construction"
        Me.cmdIFC.UseVisualStyleBackColor = True
        '
        'cmdIFI
        '
        Me.cmdIFI.Location = New System.Drawing.Point(3, 307)
        Me.cmdIFI.Name = "cmdIFI"
        Me.cmdIFI.Size = New System.Drawing.Size(147, 23)
        Me.cmdIFI.TabIndex = 3
        Me.cmdIFI.Text = "Information"
        Me.cmdIFI.UseVisualStyleBackColor = True
        '
        'cmdifr
        '
        Me.cmdifr.Location = New System.Drawing.Point(3, 147)
        Me.cmdifr.Name = "cmdifr"
        Me.cmdifr.Size = New System.Drawing.Size(147, 23)
        Me.cmdifr.TabIndex = 4
        Me.cmdifr.Text = "Review"
        Me.cmdifr.UseVisualStyleBackColor = True
        '
        'rdo50
        '
        Me.rdo50.AutoSize = True
        Me.rdo50.Location = New System.Drawing.Point(3, 26)
        Me.rdo50.Name = "rdo50"
        Me.rdo50.Size = New System.Drawing.Size(113, 17)
        Me.rdo50.TabIndex = 6
        Me.rdo50.TabStop = True
        Me.rdo50.Text = "50% Client Review"
        Me.rdo50.UseVisualStyleBackColor = True
        '
        'rdo100
        '
        Me.rdo100.AutoSize = True
        Me.rdo100.Location = New System.Drawing.Point(3, 72)
        Me.rdo100.Name = "rdo100"
        Me.rdo100.Size = New System.Drawing.Size(119, 17)
        Me.rdo100.TabIndex = 7
        Me.rdo100.TabStop = True
        Me.rdo100.Text = "100% Client Review"
        Me.rdo100.UseVisualStyleBackColor = True
        '
        'rdo75
        '
        Me.rdo75.AutoSize = True
        Me.rdo75.Location = New System.Drawing.Point(3, 49)
        Me.rdo75.Name = "rdo75"
        Me.rdo75.Size = New System.Drawing.Size(113, 17)
        Me.rdo75.TabIndex = 8
        Me.rdo75.TabStop = True
        Me.rdo75.Text = "75% Client Review"
        Me.rdo75.UseVisualStyleBackColor = True
        '
        'rdofinal
        '
        Me.rdofinal.AutoSize = True
        Me.rdofinal.Location = New System.Drawing.Point(3, 95)
        Me.rdofinal.Name = "rdofinal"
        Me.rdofinal.Size = New System.Drawing.Size(115, 17)
        Me.rdofinal.TabIndex = 9
        Me.rdofinal.TabStop = True
        Me.rdofinal.Text = "Final Client Review"
        Me.rdofinal.UseVisualStyleBackColor = True
        '
        'rdoother
        '
        Me.rdoother.AutoSize = True
        Me.rdoother.Location = New System.Drawing.Point(3, 3)
        Me.rdoother.Name = "rdoother"
        Me.rdoother.Size = New System.Drawing.Size(51, 17)
        Me.rdoother.TabIndex = 10
        Me.rdoother.TabStop = True
        Me.rdoother.Text = "Other"
        Me.rdoother.UseVisualStyleBackColor = True
        '
        'cmdifa
        '
        Me.cmdifa.Location = New System.Drawing.Point(3, 186)
        Me.cmdifa.Name = "cmdifa"
        Me.cmdifa.Size = New System.Drawing.Size(147, 23)
        Me.cmdifa.TabIndex = 11
        Me.cmdifa.Text = "Approval"
        Me.cmdifa.UseVisualStyleBackColor = True
        '
        'dtpIssueDate
        '
        Me.dtpIssueDate.CustomFormat = "yyyy.MM.dd"
        Me.dtpIssueDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpIssueDate.Location = New System.Drawing.Point(3, 364)
        Me.dtpIssueDate.Name = "dtpIssueDate"
        Me.dtpIssueDate.Size = New System.Drawing.Size(147, 20)
        Me.dtpIssueDate.TabIndex = 12
        Me.dtpIssueDate.Value = New Date(2017, 1, 17, 9, 16, 4, 0)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 345)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Issue Date"
        '
        'uscIssues
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpIssueDate)
        Me.Controls.Add(Me.cmdifa)
        Me.Controls.Add(Me.rdoother)
        Me.Controls.Add(Me.rdofinal)
        Me.Controls.Add(Me.rdo75)
        Me.Controls.Add(Me.rdo100)
        Me.Controls.Add(Me.rdo50)
        Me.Controls.Add(Me.cmdifr)
        Me.Controls.Add(Me.cmdIFI)
        Me.Controls.Add(Me.cmdIFC)
        Me.Controls.Add(Me.cmdift)
        Me.Controls.Add(Me.cmdIFRS)
        Me.Name = "uscIssues"
        Me.Size = New System.Drawing.Size(184, 493)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdIFRS As System.Windows.Forms.Button
    Friend WithEvents cmdift As System.Windows.Forms.Button
    Friend WithEvents cmdIFC As System.Windows.Forms.Button
    Friend WithEvents cmdIFI As System.Windows.Forms.Button
    Friend WithEvents cmdifr As System.Windows.Forms.Button
    Friend WithEvents rdo50 As System.Windows.Forms.RadioButton
    Friend WithEvents rdo100 As System.Windows.Forms.RadioButton
    Friend WithEvents rdo75 As System.Windows.Forms.RadioButton
    Friend WithEvents rdofinal As System.Windows.Forms.RadioButton
    Friend WithEvents rdoother As System.Windows.Forms.RadioButton
    Friend WithEvents cmdifa As System.Windows.Forms.Button
    Friend WithEvents dtpIssueDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
