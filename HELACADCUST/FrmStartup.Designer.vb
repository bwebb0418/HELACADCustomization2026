<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmStartup
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tboCSC = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CboSCL = New System.Windows.Forms.ComboBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.rdo20 = New System.Windows.Forms.RadioButton()
        Me.rdo25 = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.rdoARW = New System.Windows.Forms.RadioButton()
        Me.rdoTCK = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rdoMMS = New System.Windows.Forms.RadioButton()
        Me.rdoFTS = New System.Windows.Forms.RadioButton()
        Me.rdoMTS = New System.Windows.Forms.RadioButton()
        Me.rdoINS = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rdoARC = New System.Windows.Forms.RadioButton()
        Me.rdoIND = New System.Windows.Forms.RadioButton()
        Me.rdoBDE = New System.Windows.Forms.RadioButton()
        Me.rdoBDG = New System.Windows.Forms.RadioButton()
        Me.butSTR = New System.Windows.Forms.Button()
        Me.butEND = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnClient = New System.Windows.Forms.Button()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Label2)
        Me.GroupBox5.Controls.Add(Me.tboCSC)
        Me.GroupBox5.Controls.Add(Me.Label1)
        Me.GroupBox5.Controls.Add(Me.CboSCL)
        Me.GroupBox5.Font = New System.Drawing.Font("Arial", 7.8!)
        Me.GroupBox5.Location = New System.Drawing.Point(21, 300)
        Me.GroupBox5.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Padding = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox5.Size = New System.Drawing.Size(189, 133)
        Me.GroupBox5.TabIndex = 35
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Select Scale"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(27, 65)
        Me.Label2.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(89, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Custom Scale"
        '
        'tboCSC
        '
        Me.tboCSC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tboCSC.Location = New System.Drawing.Point(57, 90)
        Me.tboCSC.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.tboCSC.Name = "tboCSC"
        Me.tboCSC.Size = New System.Drawing.Size(111, 22)
        Me.tboCSC.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 94)
        Me.Label1.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(18, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "1:"
        '
        'CboSCL
        '
        Me.CboSCL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboSCL.FormattingEnabled = True
        Me.CboSCL.Location = New System.Drawing.Point(29, 28)
        Me.CboSCL.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.CboSCL.Name = "CboSCL"
        Me.CboSCL.Size = New System.Drawing.Size(140, 24)
        Me.CboSCL.TabIndex = 0
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.rdo20)
        Me.GroupBox4.Controls.Add(Me.rdo25)
        Me.GroupBox4.Font = New System.Drawing.Font("Arial", 7.8!)
        Me.GroupBox4.Location = New System.Drawing.Point(219, 195)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox4.Size = New System.Drawing.Size(189, 100)
        Me.GroupBox4.TabIndex = 34
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Text Size"
        '
        'rdo20
        '
        Me.rdo20.AutoSize = True
        Me.rdo20.FlatAppearance.CheckedBackColor = System.Drawing.Color.Transparent
        Me.rdo20.Location = New System.Drawing.Point(32, 36)
        Me.rdo20.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdo20.Name = "rdo20"
        Me.rdo20.Size = New System.Drawing.Size(113, 20)
        Me.rdo20.TabIndex = 19
        Me.rdo20.Text = "2.0mm Plotted"
        Me.rdo20.UseVisualStyleBackColor = True
        '
        'rdo25
        '
        Me.rdo25.AutoSize = True
        Me.rdo25.Location = New System.Drawing.Point(32, 66)
        Me.rdo25.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdo25.Name = "rdo25"
        Me.rdo25.Size = New System.Drawing.Size(113, 20)
        Me.rdo25.TabIndex = 21
        Me.rdo25.Text = "2.5mm Plotted"
        Me.rdo25.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.rdoARW)
        Me.GroupBox3.Controls.Add(Me.rdoTCK)
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 7.8!)
        Me.GroupBox3.Location = New System.Drawing.Point(21, 194)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox3.Size = New System.Drawing.Size(189, 98)
        Me.GroupBox3.TabIndex = 33
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Dimension Style"
        '
        'rdoARW
        '
        Me.rdoARW.AutoSize = True
        Me.rdoARW.Location = New System.Drawing.Point(32, 36)
        Me.rdoARW.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdoARW.Name = "rdoARW"
        Me.rdoARW.Size = New System.Drawing.Size(68, 20)
        Me.rdoARW.TabIndex = 19
        Me.rdoARW.Text = "Arrows"
        Me.rdoARW.UseVisualStyleBackColor = True
        '
        'rdoTCK
        '
        Me.rdoTCK.AutoSize = True
        Me.rdoTCK.Location = New System.Drawing.Point(32, 66)
        Me.rdoTCK.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdoTCK.Name = "rdoTCK"
        Me.rdoTCK.Size = New System.Drawing.Size(59, 20)
        Me.rdoTCK.TabIndex = 21
        Me.rdoTCK.Text = "Ticks"
        Me.rdoTCK.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox2.Controls.Add(Me.rdoMMS)
        Me.GroupBox2.Controls.Add(Me.rdoFTS)
        Me.GroupBox2.Controls.Add(Me.rdoMTS)
        Me.GroupBox2.Controls.Add(Me.rdoINS)
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 7.8!)
        Me.GroupBox2.Location = New System.Drawing.Point(219, 17)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(189, 170)
        Me.GroupBox2.TabIndex = 32
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Units"
        '
        'rdoMMS
        '
        Me.rdoMMS.AutoSize = True
        Me.rdoMMS.Location = New System.Drawing.Point(32, 36)
        Me.rdoMMS.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdoMMS.Name = "rdoMMS"
        Me.rdoMMS.Size = New System.Drawing.Size(139, 20)
        Me.rdoMMS.TabIndex = 19
        Me.rdoMMS.TabStop = True
        Me.rdoMMS.Text = "Metric - Millimeters"
        Me.rdoMMS.UseVisualStyleBackColor = True
        '
        'rdoFTS
        '
        Me.rdoFTS.AutoSize = True
        Me.rdoFTS.Location = New System.Drawing.Point(32, 126)
        Me.rdoFTS.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdoFTS.Name = "rdoFTS"
        Me.rdoFTS.Size = New System.Drawing.Size(111, 20)
        Me.rdoFTS.TabIndex = 22
        Me.rdoFTS.TabStop = True
        Me.rdoFTS.Text = "Imperial - Feet"
        Me.rdoFTS.UseVisualStyleBackColor = True
        '
        'rdoMTS
        '
        Me.rdoMTS.AutoSize = True
        Me.rdoMTS.Location = New System.Drawing.Point(32, 95)
        Me.rdoMTS.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdoMTS.Name = "rdoMTS"
        Me.rdoMTS.Size = New System.Drawing.Size(116, 20)
        Me.rdoMTS.TabIndex = 20
        Me.rdoMTS.TabStop = True
        Me.rdoMTS.Text = "Metric - Meters"
        Me.rdoMTS.UseVisualStyleBackColor = True
        '
        'rdoINS
        '
        Me.rdoINS.AutoSize = True
        Me.rdoINS.Location = New System.Drawing.Point(32, 66)
        Me.rdoINS.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdoINS.Name = "rdoINS"
        Me.rdoINS.Size = New System.Drawing.Size(123, 20)
        Me.rdoINS.TabIndex = 21
        Me.rdoINS.TabStop = True
        Me.rdoINS.Text = "Imperial - Inches"
        Me.rdoINS.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdoARC)
        Me.GroupBox1.Controls.Add(Me.rdoIND)
        Me.GroupBox1.Controls.Add(Me.rdoBDE)
        Me.GroupBox1.Controls.Add(Me.rdoBDG)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 7.8!)
        Me.GroupBox1.Location = New System.Drawing.Point(21, 16)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(189, 170)
        Me.GroupBox1.TabIndex = 31
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Discipline"
        '
        'rdoARC
        '
        Me.rdoARC.AutoSize = True
        Me.rdoARC.Location = New System.Drawing.Point(32, 36)
        Me.rdoARC.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdoARC.Name = "rdoARC"
        Me.rdoARC.Size = New System.Drawing.Size(101, 20)
        Me.rdoARC.TabIndex = 19
        Me.rdoARC.TabStop = True
        Me.rdoARC.Text = "Architectural"
        Me.rdoARC.UseVisualStyleBackColor = True
        '
        'rdoIND
        '
        Me.rdoIND.AutoSize = True
        Me.rdoIND.Location = New System.Drawing.Point(32, 126)
        Me.rdoIND.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdoIND.Name = "rdoIND"
        Me.rdoIND.Size = New System.Drawing.Size(80, 20)
        Me.rdoIND.TabIndex = 22
        Me.rdoIND.TabStop = True
        Me.rdoIND.Text = "Industrial"
        Me.rdoIND.UseVisualStyleBackColor = True
        '
        'rdoBDE
        '
        Me.rdoBDE.AutoSize = True
        Me.rdoBDE.Location = New System.Drawing.Point(32, 66)
        Me.rdoBDE.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdoBDE.Name = "rdoBDE"
        Me.rdoBDE.Size = New System.Drawing.Size(130, 20)
        Me.rdoBDE.TabIndex = 20
        Me.rdoBDE.TabStop = True
        Me.rdoBDE.Text = "Building Envelope"
        Me.rdoBDE.UseVisualStyleBackColor = True
        '
        'rdoBDG
        '
        Me.rdoBDG.AutoSize = True
        Me.rdoBDG.Location = New System.Drawing.Point(32, 97)
        Me.rdoBDG.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.rdoBDG.Name = "rdoBDG"
        Me.rdoBDG.Size = New System.Drawing.Size(81, 20)
        Me.rdoBDG.TabIndex = 21
        Me.rdoBDG.TabStop = True
        Me.rdoBDG.Text = "Buildings"
        Me.rdoBDG.UseVisualStyleBackColor = True
        '
        'butSTR
        '
        Me.butSTR.Font = New System.Drawing.Font("Arial", 7.8!)
        Me.butSTR.Location = New System.Drawing.Point(219, 358)
        Me.butSTR.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.butSTR.Name = "butSTR"
        Me.butSTR.Size = New System.Drawing.Size(189, 28)
        Me.butSTR.TabIndex = 30
        Me.butSTR.Text = "Start"
        Me.butSTR.UseVisualStyleBackColor = True
        '
        'butEND
        '
        Me.butEND.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.butEND.Font = New System.Drawing.Font("Arial", 7.8!)
        Me.butEND.Location = New System.Drawing.Point(219, 394)
        Me.butEND.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.butEND.Name = "butEND"
        Me.butEND.Size = New System.Drawing.Size(191, 28)
        Me.butEND.TabIndex = 28
        Me.butEND.Text = "Cancel"
        Me.butEND.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial Narrow", 7.8!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(216, 449)
        Me.Label3.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 16)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "Label3"
        '
        'btnClient
        '
        Me.btnClient.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClient.Font = New System.Drawing.Font("Arial", 7.8!)
        Me.btnClient.Location = New System.Drawing.Point(219, 322)
        Me.btnClient.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.btnClient.Name = "btnClient"
        Me.btnClient.Size = New System.Drawing.Size(189, 28)
        Me.btnClient.TabIndex = 38
        Me.btnClient.Text = "Client Specific Setup"
        Me.btnClient.UseVisualStyleBackColor = True
        '
        'FrmStartup
        '
        Me.AcceptButton = Me.butSTR
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.butEND
        Me.ClientSize = New System.Drawing.Size(427, 474)
        Me.Controls.Add(Me.btnClient)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.butSTR)
        Me.Controls.Add(Me.butEND)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmStartup"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Herold Engineering Drawing Startup"
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tboCSC As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CboSCL As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents rdo20 As System.Windows.Forms.RadioButton
    Friend WithEvents rdo25 As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoARW As System.Windows.Forms.RadioButton
    Friend WithEvents rdoTCK As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoMMS As System.Windows.Forms.RadioButton
    Friend WithEvents rdoFTS As System.Windows.Forms.RadioButton
    Friend WithEvents rdoMTS As System.Windows.Forms.RadioButton
    Friend WithEvents rdoINS As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoARC As System.Windows.Forms.RadioButton
    Friend WithEvents rdoIND As System.Windows.Forms.RadioButton
    Friend WithEvents rdoBDE As System.Windows.Forms.RadioButton
    Friend WithEvents rdoBDG As System.Windows.Forms.RadioButton
    Friend WithEvents butSTR As System.Windows.Forms.Button
    Friend WithEvents butEND As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnClient As System.Windows.Forms.Button
End Class
