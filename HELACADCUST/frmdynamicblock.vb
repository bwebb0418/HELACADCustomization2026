Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Windows


Public Class Frmdynamicblock
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property
    Public check_routines As New clsCheckroutines
    Public utilities As New Clsutilities
    Dim attribs, dynprop

    Dim blk As AcadBlockReference
    Dim units As Integer

    <CommandMethod("Stamp")> Sub Stamp()
        Dim inspoint
        Dim rotation
        Dim blkname
        Dim blkscl
        Dim prop As AcadDynamicBlockReferenceProperty

        TXTline1.Text = ""
        TXTLine2.Text = ""
        Cbolookopt.Items.Clear()
        Cbovisopt.Items.Clear()

        inspoint = utilities.Getpoint("Select Stamp Insertion Point: ")
        rotation = 0

        'blkscl = 1

        If check_routines.blkchk("hel_stamps-2026") = True Then
            blkname = "hel_stamps-2026"
        Else
            blkname = "hel_stamps-2026.dwg"
        End If

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            blkscl = utilities.getscale
            If blkscl = 0 Then blkscl = 1
            blk = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)
        Else
            blkscl = Thisdrawing.GetVariable("Userr3")
            If blkscl = 0 Then blkscl = 1
            blk = Thisdrawing.PaperSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)
        End If


        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Visibility" Then
                prop = dynprop(i)
                For j = 0 To UBound(prop.AllowedValues)
                    Me.Cbovisopt.Items.Add(prop.AllowedValues(j))
                Next j
            End If
        Next i

        Me.ShowDialog()

        blk.Layer = "0"
        blk.Update()


    End Sub

    Private Sub Frmdynamicblock_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If Me.Text <> "Form1" Then Exit Sub
        Cbovisopt.Visible = True
        Cbolookopt.Visible = True
        TXTline1.Visible = True
        TXTLine2.Visible = True
        Label1.Visible = True
        Label2.Visible = True
        If blk.Name = "hel_stamps-2026" Then
            Cbolookopt.Visible = False
            Me.Width = 250
            Me.Height = 300
            Me.Text = "Stamp Block Options"
            Cbovisopt.Left = 25
            Cbovisopt.Top = 25
            Cbovisopt.MaxDropDownItems = 12
            cmdfinish.Left = 100
            cmdfinish.Top = 100
            Label1.Visible = False
            TXTline1.Visible = False
            Label2.Visible = False
            TXTLine2.Visible = False
        ElseIf blk.Name = "hel_elev_tag-2026" Then
            Dim feet
            Dim neg As Boolean
            Dim inches
            Dim fractions, denom
            Dim total
            Cbovisopt.Visible = False
            Me.Text = "Elevation Tag Options"
            Me.Width = 250
            Me.Height = 300
            Label1.Top = 16
            Label1.Text = "Elevation"
            TXTline1.Top = 10
            units = Thisdrawing.GetVariable("Userr3")
            total = blk.InsertionPoint(1)
            If units <> 1 Then
                If total < 0 Then
                    total *= -1
                    neg = True
                End If
                feet = Fix(total / 12)
                inches = total - (feet * 12)
                fractions = inches - Fix(inches)
                fractions = Math.Round(fractions * 16)
                denom = 16
                If fractions / 2 = Fix(fractions / 2) Then
                    denom = 8
                End If
                If fractions / 4 = Fix(fractions / 4) Then
                    denom = 4
                End If
                If fractions / 8 = Fix(fractions / 8) Then
                    denom = 2
                End If
                If fractions / 16 = Fix(fractions / 16) Then
                    If fractions / 16 > 0.5 Then inches += +1
                    fractions = 0
                End If
                inches = Fix(inches)
                If denom <> 16 Then fractions /= (16 / denom)
                If fractions <> 0 Then
                    fractions = " " & fractions & "/" & denom
                ElseIf fractions = 0 Then
                    fractions = ""
                End If
                If neg = True Then feet *= -1
                TXTline1.Text = "EL. " & feet & "'-" & inches & fractions & Chr(34)
            Else
                If total > 1000 Or total < -1000 Then
                    total = Math.Round((total / 1000), 3) & "m"
                Else
                    total = Math.Round(total) & "mm"
                End If
                TXTline1.Text = "EL. " & total
            End If
            Label2.Top = 44
            Label2.Text = "Ref."
            TXTLine2.Top = 38
            TXTLine2.Text = "T.O. STEEL"
            cmdfinish.Left = TXTline1.Left
            Cbolookopt.Left = 10
            Cbolookopt.Top = 80
            cmdfinish.Left = 49
            cmdfinish.Top = 106


        ElseIf blk.Name = "hel_clmark-2026" Then
            Cbovisopt.Visible = False
            Cbolookopt.Visible = False
            Me.Width = 170
            Me.Height = 110
            Me.Text = "Centerline Mark Block Options"
            'Cbolookopt.Left = 10
            'Cbolookopt.top = 10
            Label1.Top = 16
            TXTline1.Top = 10
            Label2.Top = 44
            TXTLine2.Top = 38
            cmdfinish.Left = TXTline1.Left
            cmdfinish.Top = 60
            'Cbovisopt.Value = ""
        ElseIf blk.Name = "hel_decking-2026" Or blk.Name = "hel_decktag-2026" Then
            Cbolookopt.Visible = False
            Me.Width = 170
            Me.Height = 100
            Me.Text = "Decking Block Options"
            Cbovisopt.Visible = True
            Cbovisopt.Left = 10
            Cbovisopt.Top = 10
            cmdfinish.Left = 49
            cmdfinish.Top = 36
            Label1.Visible = False
            TXTline1.Visible = False
            Label2.Visible = False
            TXTLine2.Visible = False
        End If

        '    Cbolookopt.SelectedValue = Cbolookopt.Items(0)
        '    Cbovisopt.SelectedValue = Cbovisopt.Items(0)
    End Sub

    Private Sub Frmdynamicblock_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Me.Text = "Form1"
    End Sub

    Private Sub Frmdynamicblock_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
    End Sub



    Private Sub Frmdynamicblock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If blk.Name = "hel_stamps-2026" Then
        '    Cbolookopt.Visible = False
        '    Me.Width = 250
        '    Me.Height = 300
        '    Me.Text = "Stamp Block Options"
        '    Cbovisopt.Left = 25
        '    Cbovisopt.Top = 25
        '    Cbovisopt.MaxDropDownItems = 12
        '    cmdfinish.Left = 100
        '    cmdfinish.Top = 100
        '    Label1.Visible = False
        '    TXTline1.Visible = False
        '    Label2.Visible = False
        '    TXTLine2.Visible = False
        'ElseIf blk.Name = "hel_elev_tag-2026" Then
        '    Dim feet, neg
        '    Dim inches
        '    Dim fractions, denom
        '    Dim total
        '    Cbovisopt.Visible = False
        '    Me.Text = "Elevation Tag Options"
        '    Me.Width = 250
        '    Me.Height = 300
        '    Label1.Top = 16
        '    Label1.Text = "Elevation"
        '    TXTline1.Top = 10
        '    units = Thisdrawing.GetVariable("Userr3")
        '    total = blk.InsertionPoint(1)
        '    If units <> 1 Then
        '        If total < 0 Then
        '            total = total * -1
        '            neg = True
        '        End If
        '        feet = Fix(total / 12)
        '        inches = total - (feet * 12)
        '        fractions = inches - Fix(inches)
        '        fractions = Math.Round(fractions * 16)
        '        denom = 16
        '        If fractions / 2 = Fix(fractions / 2) Then
        '            denom = 8
        '        End If
        '        If fractions / 4 = Fix(fractions / 4) Then
        '            denom = 4
        '        End If
        '        If fractions / 8 = Fix(fractions / 8) Then
        '            denom = 2
        '        End If
        '        If fractions / 16 = Fix(fractions / 16) Then
        '            If fractions / 16 > 0.5 Then inches = inches + 1
        '            fractions = 0
        '        End If
        '        inches = Fix(inches)
        '        If denom <> 16 Then fractions = fractions / (16 / denom)
        '        If fractions <> 0 Then
        '            fractions = " " & fractions & "/" & denom
        '        ElseIf fractions = 0 Then
        '            fractions = ""
        '        End If
        '        If neg = True Then feet = feet * -1
        '        TXTline1.Text = "EL. " & feet & "'-" & inches & fractions & Chr(34)
        '    Else
        '        If total > 1000 Or total < -1000 Then
        '            total = Math.Round((total / 1000), 3) & "m"
        '        Else
        '            total = Math.Round(total) & "mm"
        '        End If
        '        TXTline1.Text = "EL. " & total
        '    End If
        '    Label2.Top = 44
        '    Label2.Text = "Ref."
        '    TXTLine2.Top = 38
        '    TXTLine2.Text = "T.O. STEEL"
        '    cmdfinish.Left = TXTline1.Left
        '    Cbolookopt.Left = 10
        '    Cbolookopt.Top = 80
        '    cmdfinish.Left = 49
        '    cmdfinish.Top = 106
        '    Cbovisopt.SelectedValue = ""
        '    Cbolookopt.SelectedValue = Cbolookopt.Items(0)
        'ElseIf blk.Name = "hel_clmark-2026" Then
        '    Cbovisopt.Visible = False
        '    Cbolookopt.Visible = False
        '    Me.Width = 170
        '    Me.Height = 110
        '    Me.Text = "Centerline Mark Block Options"
        '    'Cbolookopt.Left = 10
        '    'Cbolookopt.top = 10
        '    Label1.Top = 16
        '    TXTline1.Top = 10
        '    Label2.Top = 44
        '    TXTLine2.Top = 38
        '    cmdfinish.Left = TXTline1.Left
        '    cmdfinish.Top = 60
        '    'Cbovisopt.Value = ""
        'ElseIf blk.Name = "hel_decking-2026" Or blk.Name = "hel_decktag-2026" Then
        '    MsgBox("hH")
        '    Cbolookopt.Visible = False
        '    Me.Width = 170
        '    Me.Height = 100
        '    Me.Text = "Decking Block Options"
        '    Cbovisopt.Left = 10
        '    Cbovisopt.Top = 10
        '    cmdfinish.Left = 49
        '    cmdfinish.Top = 36
        '    Label1.Visible = False
        '    TXTline1.Visible = False
        '    Label2.Visible = False
        '    TXTLine2.Visible = False
        'End If
    End Sub

    Private Sub Cmdfinish_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdfinish.Click
        Dim minpoint1 As Object = Nothing
        Dim maxpoint1 As Object = Nothing
        Dim minpoint2 As Object = Nothing
        Dim maxpoint2 As Object = Nothing
        Dim i As Integer
        Dim j As Integer
        Dim K As Integer
        For i = 0 To UBound(dynprop)
            If blk.Name = "hel_elev_tag-2026" Then
                For j = 0 To UBound(attribs)
                    If attribs(j).TextString = "EL." Then
                        If TXTline1.Text = "" Then TXTline1 = attribs(j).TextString
                        attribs(j).TextString = TXTline1.Text
                        attribs(j).GetBoundingBox(minpoint1, maxpoint1)
                    ElseIf attribs(j).TextString = "T.O." Then
                        If TXTLine2.Text = "" Then TXTLine2 = attribs(j).TextString
                        attribs(j).TextString = TXTLine2.Text
                        attribs(j).GetBoundingBox(minpoint2, maxpoint2)
                    End If
                Next j
                For K = 0 To UBound(dynprop)
                    If dynprop(K).PropertyName = "Lookup" Then
                        If Cbolookopt.SelectedIndex = -1 Then
                            dynprop(K).Value = Cbolookopt.Items(1).ToString
                        Else
                            dynprop(K).Value = Cbolookopt.SelectedItem.ToString
                        End If
                    ElseIf dynprop(K).PropertyName = "Left Distance" Then
                        Dim units, scl
                        attribs(0).GetBoundingBox(minpoint1, maxpoint1)
                        attribs(1).GetBoundingBox(minpoint2, maxpoint2)
                        units = Thisdrawing.GetVariable("userr3")
                        scl = Thisdrawing.GetVariable("userr2")
                        If maxpoint1(0) > maxpoint2(0) Then
                            dynprop(K).Value = (maxpoint1(0) - minpoint1(0)) + (10 * scl * units)
                        Else
                            dynprop(K).Value = (maxpoint2(0) - minpoint1(0)) + (10 * scl * units)
                        End If
                    End If
                Next K
            ElseIf blk.Name = "hel_stamps-2026" Or blk.Name = "hel_decking-2026" Or blk.Name = "hel_decktag-2026" Then
                If dynprop(i).PropertyName = "Visibility" Then dynprop(i).Value = Me.Cbovisopt.SelectedItem.ToString
            ElseIf blk.Name = "hel_clmark-2026" Then
                If dynprop(i).PropertyName = "Visibility" Then
                    If TXTline1.Text = "" Or TXTLine2.Text = "" Then
                        dynprop(i).Value = "1-LINE"
                        For j = 0 To UBound(attribs)
                            If attribs(j).TagString = "COL." Then
                                attribs(j).TextString = TXTline1.Text & TXTLine2.Text
                            End If
                        Next j
                    Else
                        dynprop(i).Value = "2-LINE"
                        For j = 0 To UBound(attribs)
                            If attribs(j).TagString = "CL2-1" Then
                                attribs(j).TextString = TXTline1.Text
                            ElseIf attribs(j).TagString = "CL2-2" Then
                                attribs(j).TextString = TXTLine2.Text
                            End If
                        Next j
                    End If
                End If
            End If
        Next i
        Me.Hide()

        TXTline1.Text = ""
        TXTLine2.Text = ""
        Cbolookopt.Items.Clear()
        Cbovisopt.Items.Clear()
        Me.Text = "Form1"

    End Sub

    <CommandMethod("ElevTag")> Sub Elev_tag()
        Dim inspoint
        Dim blkname
        Dim blkscl
        Dim prop As AcadDynamicBlockReferenceProperty
        Dim blockinsert As New ClsBlock_insert
        TXTline1.Text = ""
        TXTLine2.Text = ""
        Cbolookopt.Items.Clear()
        Cbovisopt.Items.Clear()

        inspoint = utilities.Getpoint("Select Elevation Tag Insertion Point: ")



        blkname = "hel_elev_tag-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.Loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If

        blkscl = utilities.Getscale
        If blkscl = 0 Then blkscl = 1

        blk = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        attribs = blk.GetAttributes
        For i = LBound(attribs) To UBound(attribs)
            On Error Resume Next
            attribs(i).StyleName = utilities.Gettxtstyle
            attribs(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Strings.Right(utilities.Gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Strings.Right(utilities.Gettxtstyle, 2))
            End If
            attribs(i).Layer = "~-TXT" & Strings.Right(utilities.Gettxtstyle, 2)
            On Error GoTo 0
        Next i

        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Lookup" Then
                prop = dynprop(i)
                For j = 0 To UBound(prop.AllowedValues)
                    Me.Cbolookopt.Items.Add(prop.AllowedValues(j))
                Next j
            End If
        Next i

        Me.ShowDialog()

        If check_routines.Laychk("~-TSYM3") = False Then Thisdrawing.Layers.Add("~-TSYM3")
        blk.Layer = "~-TSYM3"
        blk.Update()



    End Sub

    <CommandMethod("decking")> Sub Decking()
        Dim space
        Dim inspoint
        Dim rotation
        Dim blkname
        Dim blkscl
        Dim prop As AcadDynamicBlockReferenceProperty
        Dim tagtype As String = Nothing
        Dim ATTRIBS
        Dim blockinsert As New ClsBlock_insert

        TXTline1.Text = ""
        TXTLine2.Text = ""
        Cbolookopt.Items.Clear()
        Cbovisopt.Items.Clear()

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If

        inspoint = utilities.Getpoint("Select Decking Symbol Insertion Point: ")
        rotation = 0
        blkscl = Thisdrawing.GetVariable("Userr3")
        If blkscl = 0 Then blkscl = 1

tagpick:


        Try
            tagtype = Thisdrawing.Utility.GetString(False, "Input deck number: ")
        Catch

        End Try

        If tagtype = Nothing Then
            blkname = "hel_decking-2026"
        Else
            blkname = "hel_decktag-2026"
        End If

        If check_routines.Blkchk(blkname) = True Then

        Else
            If blockinsert.Loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If


        blk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Visibility" Then
                prop = dynprop(i)
                For j = 0 To UBound(prop.AllowedValues)
                    Me.Cbovisopt.Items.Add(prop.AllowedValues(j))
                Next j
            End If
        Next i

        Me.ShowDialog()


        If tagtype IsNot Nothing Then
            ATTRIBS = blk.GetAttributes
            For i = LBound(ATTRIBS) To UBound(ATTRIBS)
                Try
                    ATTRIBS(i).StyleName = utilities.Gettxtstyle
                    Try
                        ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                        If check_routines.Laychk("~-TXT" & Strings.Right(utilities.Gettxtstyle, 2)) = False Then
                            Thisdrawing.Layers.Add("~-TXT" & Strings.Right(utilities.Gettxtstyle, 2))
                        End If
                        ATTRIBS(i).Layer = "~-TXT" & Strings.Right(utilities.Gettxtstyle, 2)
                    Catch
                    End Try
                    If ATTRIBS(i).Tagstring = "D1" Then
                            ATTRIBS(i).TextString = tagtype
                        End If
                    Catch

                    End Try
            Next i
        End If

        blk.Layer = "0"
        blk.Update()

    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        Me.Text = "Form1"
    End Sub


End Class