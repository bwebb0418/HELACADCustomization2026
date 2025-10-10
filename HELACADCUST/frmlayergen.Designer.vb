<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frmlayergen
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
        Me.rdoAll = New System.Windows.Forms.RadioButton()
        Me.CboBDR = New System.Windows.Forms.CheckBox()
        Me.Cbogen = New System.Windows.Forms.CheckBox()
        Me.Cboarc = New System.Windows.Forms.CheckBox()
        Me.Cbobev = New System.Windows.Forms.CheckBox()
        Me.Cbogrd = New System.Windows.Forms.CheckBox()
        Me.Cbotxt = New System.Windows.Forms.CheckBox()
        Me.Cbofdn = New System.Windows.Forms.CheckBox()
        Me.Cbocon = New System.Windows.Forms.CheckBox()
        Me.Cborei = New System.Windows.Forms.CheckBox()
        Me.Cbostl = New System.Windows.Forms.CheckBox()
        Me.Cbomas = New System.Windows.Forms.CheckBox()
        Me.Cbowod = New System.Windows.Forms.CheckBox()
        Me.Cboind = New System.Windows.Forms.CheckBox()
        Me.Cbobri = New System.Windows.Forms.CheckBox()
        Me.Cbodes = New System.Windows.Forms.CheckBox()
        Me.butCAN = New System.Windows.Forms.Button()
        Me.butCLR = New System.Windows.Forms.Button()
        Me.butGEN = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'rdoAll
        '
        Me.rdoAll.AutoSize = True
        Me.rdoAll.Location = New System.Drawing.Point(20, 10)
        Me.rdoAll.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rdoAll.Name = "rdoAll"
        Me.rdoAll.Size = New System.Drawing.Size(86, 20)
        Me.rdoAll.TabIndex = 0
        Me.rdoAll.Text = "All Layers"
        Me.rdoAll.UseVisualStyleBackColor = True
        '
        'CboBDR
        '
        Me.CboBDR.AutoSize = True
        Me.CboBDR.Location = New System.Drawing.Point(30, 66)
        Me.CboBDR.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CboBDR.Name = "CboBDR"
        Me.CboBDR.Size = New System.Drawing.Size(110, 20)
        Me.CboBDR.TabIndex = 1
        Me.CboBDR.Tag = "BRDR"
        Me.CboBDR.Text = "Border Layers"
        Me.CboBDR.UseVisualStyleBackColor = True
        '
        'Cbogen
        '
        Me.Cbogen.AutoSize = True
        Me.Cbogen.FlatAppearance.CheckedBackColor = System.Drawing.Color.Transparent
        Me.Cbogen.Location = New System.Drawing.Point(30, 38)
        Me.Cbogen.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbogen.Name = "Cbogen"
        Me.Cbogen.Size = New System.Drawing.Size(117, 20)
        Me.Cbogen.TabIndex = 2
        Me.Cbogen.Tag = "GNRL"
        Me.Cbogen.Text = "General Layers"
        Me.Cbogen.UseVisualStyleBackColor = True
        '
        'Cboarc
        '
        Me.Cboarc.AutoSize = True
        Me.Cboarc.Location = New System.Drawing.Point(30, 150)
        Me.Cboarc.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cboarc.Name = "Cboarc"
        Me.Cboarc.Size = New System.Drawing.Size(145, 20)
        Me.Cboarc.TabIndex = 3
        Me.Cboarc.Tag = "ARCH"
        Me.Cboarc.Text = "Architectural Layers"
        Me.Cboarc.UseVisualStyleBackColor = True
        '
        'Cbobev
        '
        Me.Cbobev.AutoSize = True
        Me.Cbobev.Location = New System.Drawing.Point(30, 178)
        Me.Cbobev.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbobev.Name = "Cbobev"
        Me.Cbobev.Size = New System.Drawing.Size(174, 20)
        Me.Cbobev.TabIndex = 4
        Me.Cbobev.Tag = "BDEN"
        Me.Cbobev.Text = "Building Envelope Layers"
        Me.Cbobev.UseVisualStyleBackColor = True
        '
        'Cbogrd
        '
        Me.Cbogrd.AutoSize = True
        Me.Cbogrd.Location = New System.Drawing.Point(30, 122)
        Me.Cbogrd.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbogrd.Name = "Cbogrd"
        Me.Cbogrd.Size = New System.Drawing.Size(96, 20)
        Me.Cbogrd.TabIndex = 5
        Me.Cbogrd.Tag = "GRID"
        Me.Cbogrd.Text = "Grid Layers"
        Me.Cbogrd.UseVisualStyleBackColor = True
        '
        'Cbotxt
        '
        Me.Cbotxt.AutoSize = True
        Me.Cbotxt.Location = New System.Drawing.Point(30, 94)
        Me.Cbotxt.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbotxt.Name = "Cbotxt"
        Me.Cbotxt.Size = New System.Drawing.Size(96, 20)
        Me.Cbotxt.TabIndex = 6
        Me.Cbotxt.Tag = "TEXT"
        Me.Cbotxt.Text = "Text Layers"
        Me.Cbotxt.UseVisualStyleBackColor = True
        '
        'Cbofdn
        '
        Me.Cbofdn.AutoSize = True
        Me.Cbofdn.Location = New System.Drawing.Point(30, 234)
        Me.Cbofdn.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbofdn.Name = "Cbofdn"
        Me.Cbofdn.Size = New System.Drawing.Size(136, 20)
        Me.Cbofdn.TabIndex = 7
        Me.Cbofdn.Tag = "FNDN"
        Me.Cbofdn.Text = "Foundation Layers"
        Me.Cbofdn.UseVisualStyleBackColor = True
        '
        'Cbocon
        '
        Me.Cbocon.AutoSize = True
        Me.Cbocon.Location = New System.Drawing.Point(30, 206)
        Me.Cbocon.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbocon.Name = "Cbocon"
        Me.Cbocon.Size = New System.Drawing.Size(124, 20)
        Me.Cbocon.TabIndex = 8
        Me.Cbocon.Tag = "CONC"
        Me.Cbocon.Text = "Concrete Layers"
        Me.Cbocon.UseVisualStyleBackColor = True
        '
        'Cborei
        '
        Me.Cborei.AutoSize = True
        Me.Cborei.Location = New System.Drawing.Point(30, 262)
        Me.Cborei.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cborei.Name = "Cborei"
        Me.Cborei.Size = New System.Drawing.Size(136, 20)
        Me.Cborei.TabIndex = 9
        Me.Cborei.Tag = "REBR"
        Me.Cborei.Text = "Reinforcing Layers"
        Me.Cborei.UseVisualStyleBackColor = True
        '
        'Cbostl
        '
        Me.Cbostl.AutoSize = True
        Me.Cbostl.Location = New System.Drawing.Point(30, 318)
        Me.Cbostl.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbostl.Name = "Cbostl"
        Me.Cbostl.Size = New System.Drawing.Size(102, 20)
        Me.Cbostl.TabIndex = 10
        Me.Cbostl.Tag = "STEL"
        Me.Cbostl.Text = "Steel Layers"
        Me.Cbostl.UseVisualStyleBackColor = True
        '
        'Cbomas
        '
        Me.Cbomas.AutoSize = True
        Me.Cbomas.Location = New System.Drawing.Point(30, 290)
        Me.Cbomas.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbomas.Name = "Cbomas"
        Me.Cbomas.Size = New System.Drawing.Size(122, 20)
        Me.Cbomas.TabIndex = 11
        Me.Cbomas.Tag = "MSNR"
        Me.Cbomas.Text = "Masonry Layers"
        Me.Cbomas.UseVisualStyleBackColor = True
        '
        'Cbowod
        '
        Me.Cbowod.AutoSize = True
        Me.Cbowod.Location = New System.Drawing.Point(30, 374)
        Me.Cbowod.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbowod.Name = "Cbowod"
        Me.Cbowod.Size = New System.Drawing.Size(106, 20)
        Me.Cbowod.TabIndex = 12
        Me.Cbowod.Tag = "WOOD"
        Me.Cbowod.Text = "Wood Layers"
        Me.Cbowod.UseVisualStyleBackColor = True
        '
        'Cboind
        '
        Me.Cboind.AutoSize = True
        Me.Cboind.Location = New System.Drawing.Point(30, 346)
        Me.Cboind.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cboind.Name = "Cboind"
        Me.Cboind.Size = New System.Drawing.Size(156, 20)
        Me.Cboind.TabIndex = 13
        Me.Cboind.Tag = "INDS"
        Me.Cboind.Text = "Misc Industrial Layers"
        Me.Cboind.UseVisualStyleBackColor = True
        '
        'Cbobri
        '
        Me.Cbobri.AutoSize = True
        Me.Cbobri.Location = New System.Drawing.Point(30, 402)
        Me.Cbobri.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbobri.Name = "Cbobri"
        Me.Cbobri.Size = New System.Drawing.Size(109, 20)
        Me.Cbobri.TabIndex = 14
        Me.Cbobri.Tag = "BRDG"
        Me.Cbobri.Text = "Bridge Layers"
        Me.Cbobri.UseVisualStyleBackColor = True
        '
        'Cbodes
        '
        Me.Cbodes.AutoSize = True
        Me.Cbodes.Location = New System.Drawing.Point(30, 430)
        Me.Cbodes.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cbodes.Name = "Cbodes"
        Me.Cbodes.Size = New System.Drawing.Size(112, 20)
        Me.Cbodes.TabIndex = 15
        Me.Cbodes.Tag = "DSGN"
        Me.Cbodes.Text = "Design Layers"
        Me.Cbodes.UseVisualStyleBackColor = True
        '
        'butCAN
        '
        Me.butCAN.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.butCAN.Location = New System.Drawing.Point(20, 543)
        Me.butCAN.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.butCAN.Name = "butCAN"
        Me.butCAN.Size = New System.Drawing.Size(203, 28)
        Me.butCAN.TabIndex = 16
        Me.butCAN.Text = "Cancel"
        Me.butCAN.UseVisualStyleBackColor = True
        '
        'butCLR
        '
        Me.butCLR.Location = New System.Drawing.Point(20, 471)
        Me.butCLR.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.butCLR.Name = "butCLR"
        Me.butCLR.Size = New System.Drawing.Size(203, 28)
        Me.butCLR.TabIndex = 17
        Me.butCLR.Text = "Clear Selection"
        Me.butCLR.UseVisualStyleBackColor = True
        '
        'butGEN
        '
        Me.butGEN.Location = New System.Drawing.Point(20, 507)
        Me.butGEN.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.butGEN.Name = "butGEN"
        Me.butGEN.Size = New System.Drawing.Size(203, 28)
        Me.butGEN.TabIndex = 18
        Me.butGEN.Text = "Generate Layers"
        Me.butGEN.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(20, 579)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ProgressBar1.MarqueeAnimationSpeed = 20
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(204, 28)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar1.TabIndex = 19
        Me.ProgressBar1.Visible = False
        '
        'Frmlayergen
        '
        Me.AcceptButton = Me.butGEN
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.butCAN
        Me.ClientSize = New System.Drawing.Size(242, 620)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.butGEN)
        Me.Controls.Add(Me.butCLR)
        Me.Controls.Add(Me.butCAN)
        Me.Controls.Add(Me.Cbodes)
        Me.Controls.Add(Me.Cbobri)
        Me.Controls.Add(Me.Cboind)
        Me.Controls.Add(Me.Cbowod)
        Me.Controls.Add(Me.Cbomas)
        Me.Controls.Add(Me.Cbostl)
        Me.Controls.Add(Me.Cborei)
        Me.Controls.Add(Me.Cbocon)
        Me.Controls.Add(Me.Cbofdn)
        Me.Controls.Add(Me.Cbotxt)
        Me.Controls.Add(Me.Cbogrd)
        Me.Controls.Add(Me.Cbobev)
        Me.Controls.Add(Me.Cboarc)
        Me.Controls.Add(Me.Cbogen)
        Me.Controls.Add(Me.CboBDR)
        Me.Controls.Add(Me.rdoAll)
        Me.Font = New System.Drawing.Font("Arial", 7.8!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frmlayergen"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Layer Generator"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rdoAll As System.Windows.Forms.RadioButton
    Friend WithEvents CboBDR As System.Windows.Forms.CheckBox
    Friend WithEvents Cbogen As System.Windows.Forms.CheckBox
    Friend WithEvents Cboarc As System.Windows.Forms.CheckBox
    Friend WithEvents Cbobev As System.Windows.Forms.CheckBox
    Friend WithEvents Cbogrd As System.Windows.Forms.CheckBox
    Friend WithEvents Cbotxt As System.Windows.Forms.CheckBox
    Friend WithEvents Cbofdn As System.Windows.Forms.CheckBox
    Friend WithEvents Cbocon As System.Windows.Forms.CheckBox
    Friend WithEvents Cborei As System.Windows.Forms.CheckBox
    Friend WithEvents Cbostl As System.Windows.Forms.CheckBox
    Friend WithEvents Cbomas As System.Windows.Forms.CheckBox
    Friend WithEvents Cbowod As System.Windows.Forms.CheckBox
    Friend WithEvents Cboind As System.Windows.Forms.CheckBox
    Friend WithEvents Cbobri As System.Windows.Forms.CheckBox
    Friend WithEvents Cbodes As System.Windows.Forms.CheckBox
    Friend WithEvents butCAN As System.Windows.Forms.Button
    Friend WithEvents butCLR As System.Windows.Forms.Button
    Friend WithEvents butGEN As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
End Class
