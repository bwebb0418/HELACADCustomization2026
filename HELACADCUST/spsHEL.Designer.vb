<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class spsHEL
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
    Friend WithEvents ApplicationTitle As System.Windows.Forms.Label
    Friend WithEvents AutoCloselbl As System.Windows.Forms.Label
    Friend WithEvents Copyright As System.Windows.Forms.Label
    Friend WithEvents MainLayoutPanel As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents DetailsLayoutPanel As System.Windows.Forms.TableLayoutPanel

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(spsHEL))
        Me.MainLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
        Me.ApplicationTitle = New System.Windows.Forms.Label()
        Me.DetailsLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
        Me.Versiontext = New System.Windows.Forms.Label()
        Me.Copyright = New System.Windows.Forms.Label()
        Me.AutoCloselbl = New System.Windows.Forms.Label()
        Me.MainLayoutPanel.SuspendLayout()
        Me.DetailsLayoutPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainLayoutPanel
        '
        Me.MainLayoutPanel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.MainLayoutPanel.BackgroundImage = CType(resources.GetObject("MainLayoutPanel.BackgroundImage"), System.Drawing.Image)
        Me.MainLayoutPanel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.MainLayoutPanel.ColumnCount = 2
        Me.MainLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 435.0!))
        Me.MainLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 23.0!))
        Me.MainLayoutPanel.Controls.Add(Me.ApplicationTitle, 0, 1)
        Me.MainLayoutPanel.Controls.Add(Me.DetailsLayoutPanel, 1, 1)
        Me.MainLayoutPanel.Font = New System.Drawing.Font("Arial", 7.8!)
        Me.MainLayoutPanel.Location = New System.Drawing.Point(0, 0)
        Me.MainLayoutPanel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MainLayoutPanel.Name = "MainLayoutPanel"
        Me.MainLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 276.0!))
        Me.MainLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 39.0!))
        Me.MainLayoutPanel.Size = New System.Drawing.Size(661, 373)
        Me.MainLayoutPanel.TabIndex = 0
        '
        'ApplicationTitle
        '
        Me.ApplicationTitle.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.ApplicationTitle.BackColor = System.Drawing.Color.Transparent
        Me.ApplicationTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ApplicationTitle.Location = New System.Drawing.Point(4, 276)
        Me.ApplicationTitle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.ApplicationTitle.Name = "ApplicationTitle"
        Me.ApplicationTitle.Size = New System.Drawing.Size(427, 97)
        Me.ApplicationTitle.TabIndex = 0
        Me.ApplicationTitle.Text = "AutoCAD 2026 Customization"
        Me.ApplicationTitle.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'DetailsLayoutPanel
        '
        Me.DetailsLayoutPanel.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.DetailsLayoutPanel.BackColor = System.Drawing.Color.Transparent
        Me.DetailsLayoutPanel.ColumnCount = 1
        Me.DetailsLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 219.0!))
        Me.DetailsLayoutPanel.Controls.Add(Me.Versiontext, 0, 1)
        Me.DetailsLayoutPanel.Controls.Add(Me.Copyright, 0, 1)
        Me.DetailsLayoutPanel.Controls.Add(Me.AutoCloselbl, 0, 0)
        Me.DetailsLayoutPanel.Location = New System.Drawing.Point(439, 280)
        Me.DetailsLayoutPanel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DetailsLayoutPanel.Name = "DetailsLayoutPanel"
        Me.DetailsLayoutPanel.RowCount = 2
        Me.DetailsLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.11037!))
        Me.DetailsLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.44482!))
        Me.DetailsLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.44482!))
        Me.DetailsLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25.0!))
        Me.DetailsLayoutPanel.Size = New System.Drawing.Size(218, 89)
        Me.DetailsLayoutPanel.TabIndex = 1
        '
        'Versiontext
        '
        Me.Versiontext.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Versiontext.BackColor = System.Drawing.Color.Transparent
        Me.Versiontext.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Versiontext.Location = New System.Drawing.Point(4, 31)
        Me.Versiontext.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Versiontext.Name = "Versiontext"
        Me.Versiontext.Size = New System.Drawing.Size(211, 25)
        Me.Versiontext.TabIndex = 3
        Me.Versiontext.Text = "Version {0}.{1}.{2}"
        '
        'Copyright
        '
        Me.Copyright.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Copyright.BackColor = System.Drawing.Color.Transparent
        Me.Copyright.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Copyright.Location = New System.Drawing.Point(4, 65)
        Me.Copyright.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Copyright.Name = "Copyright"
        Me.Copyright.Size = New System.Drawing.Size(211, 17)
        Me.Copyright.TabIndex = 2
        Me.Copyright.Text = "Copyright"
        '
        'AutoCloselbl
        '
        Me.AutoCloselbl.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.AutoCloselbl.BackColor = System.Drawing.Color.Transparent
        Me.AutoCloselbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AutoCloselbl.Location = New System.Drawing.Point(4, 2)
        Me.AutoCloselbl.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.AutoCloselbl.Name = "AutoCloselbl"
        Me.AutoCloselbl.Size = New System.Drawing.Size(211, 25)
        Me.AutoCloselbl.TabIndex = 1
        Me.AutoCloselbl.Text = "AutoClose in 5..."
        Me.AutoCloselbl.Visible = False
        '
        'spsHEL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(661, 373)
        Me.ControlBox = False
        Me.Controls.Add(Me.MainLayoutPanel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "spsHEL"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.MainLayoutPanel.ResumeLayout(False)
        Me.DetailsLayoutPanel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Versiontext As System.Windows.Forms.Label

End Class
