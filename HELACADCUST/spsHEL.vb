Imports System.Timers
Public NotInheritable Class spsHEL
    Private Shared WithEvents mytimer As New System.Windows.Forms.Timer
    Private Shared alarmcounter As Integer = 1
    Private Shared exitflag As Boolean = False

    Private Sub spsHEL_Leave(sender As Object, e As System.EventArgs) Handles Me.Leave

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        '        Me.Hide()
        Me.Dispose()
    End Sub

    'TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
    '  of the Project Designer ("Properties" under the "Project" menu).


    Private Sub SplashScreen2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set up the dialog text at runtime according to the application's assembly information.  

        'TODO: Customize the application's assembly information in the "Application" pane of the project 
        '  properties dialog (under the "Project" menu).

        'Application title
        If My.Application.Info.Title <> "" Then
            ApplicationTitle.Text = My.Application.Info.Title
        Else
            'If the application title is missing, use the application name, without the extension
            ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If

        'Format the version information using the text set into the Version control at design time as the
        '  formatting string.  This allows for effective localization if desired.
        '  Build and revision information could be included by using the following code and changing the 
        '  Version control's designtime text to "Version {0}.{1:00}.{2}.{3}" or something similar.  See
        '  String.Format() in Help for more information.
        '
        Versiontext.Text = System.String.Format(Versiontext.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build) ', My.Application.Info.Version.Revision)

        'Versiontext.Text = System.String.Format(Versiontext.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Revision)

        'Copyright info
        Copyright.Text = My.Application.Info.Copyright
        mytimer.Interval = 1000
        mytimer.Enabled = True
        mytimer.Start()
        AddHandler mytimer.Tick, AddressOf ontimedevent
        ' While exitflag = False
        ' System.Windows.Forms.Application.DoEvents()
        '
        '        End While

        'Me.Hide()
        AutoCloselbl.Text = "AutoClose in 5..."

    End Sub



    Private Sub MainLayoutPanel_Click(sender As Object, e As System.EventArgs) Handles MainLayoutPanel.Click

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Hide()
    End Sub

    Private Sub MainLayoutPanel_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles MainLayoutPanel.Paint

    End Sub

    Private Sub ontimedevent(ByVal sender As Object, ByVal e As EventArgs)
        'If AutoCloselbl.Text = "AutoClose in 0..." Then AutoCloselbl.Text = "AutoClose in 5..."
        Dim numstring As String = Strings.Right(AutoCloselbl.Text, 4)
        numstring = Strings.Left(numstring, 1)
        Dim num As Integer = Convert.ToInt32(numstring)
        num = num - 1
        AutoCloselbl.Text = "AutoClose in " & num & "..."
        If num <= 0 Then
            mytimer.Stop()
            mytimer.Enabled = False
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            'Me.Hide()
            Me.Dispose()
            Exit Sub
        End If

        Me.Refresh()
    End Sub

    Private Sub DetailsLayoutPanel_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles DetailsLayoutPanel.Paint

    End Sub


    Private Sub AutoCloselbl_Click(sender As Object, e As EventArgs) Handles AutoCloselbl.Click

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
