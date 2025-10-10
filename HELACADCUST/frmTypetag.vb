Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows

Public Class FrmTypetag
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    <CommandMethod("typetag2")> Public Sub Runtypeform()
        Me.ShowDialog()
    End Sub

    Private Sub FrmTypetag_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Cmbobox As System.Windows.Forms.ComboBox
        Dim contrl As System.Windows.Forms.Control


        'contrls = frmtypeform.Controls


        For Each contrl In Me.Controls
            If TypeOf contrl Is System.Windows.Forms.ComboBox Then
                Cmbobox = contrl
                Cmbobox.Items.Add("BP")
                Cmbobox.Items.Add("BR")
                Cmbobox.Items.Add("CB")
                Cmbobox.Items.Add("CC")
                Cmbobox.Items.Add("CP")
                Cmbobox.Items.Add("CW")
                Cmbobox.Items.Add("GB")
                Cmbobox.Items.Add("MF")
                Cmbobox.Items.Add("ML")
                Cmbobox.Items.Add("MW")
                Cmbobox.Items.Add("P")
                Cmbobox.Items.Add("PC")
                Cmbobox.Items.Add("PF")
                Cmbobox.Items.Add("SC")
                Cmbobox.Items.Add("SF")
                Cmbobox.Items.Add("SP")
                Cmbobox.Items.Add("ST")
                Cmbobox.Items.Add("WP")
                Cmbobox.Items.Add("Z")
                Cmbobox.Items.Add("Other")
            End If
        Next contrl

    End Sub

    Private Sub CmdINSERT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdINSERT.Click
        Dim Cmbobox As System.Windows.Forms.ComboBox
        Dim contrl As System.Windows.Forms.Control
        'Dim contrls As Controls
        Dim count As Integer
        Dim inspoint
        Dim blkscl
        Dim blkname
        Dim dynprop
        Dim prop As AcadDynamicBlockReferenceProperty
        Dim blkref As AcadBlockReference
        Dim attrib As AcadAttributeReference
        Dim ATTRIBS


        Me.Hide()

        count = 0

        If txttop.Text = "" Then
            Me.Show()
            MsgBox("Please enter a value in the 'Top' box", vbOKOnly, "Nothing to do!")
        End If

        'contrls = frmtypeform.Controls
        For Each contrl In Me.Controls
            If TypeOf contrl Is System.Windows.Forms.ComboBox Then
                Cmbobox = contrl
                If Cmbobox.SelectedItem <> Nothing Then
                    count += 1
                End If

            End If
        Next contrl
        Dim utilities As New Clsutilities
        inspoint = utilities.Getpoint("Select Insertion Point for Type Tag:")
        blkscl = utilities.Getscale * Thisdrawing.GetVariable("userr1")
        If blkscl = 0 Then blkscl = 1
        Dim check_routines As New ClsCheckroutines
        If check_routines.Blkchk("hel_type_tag-2026") = True Then
            blkname = "hel_type_tag-2026"
        Else
            blkname = "hel_type_tag-2026.dwg"
        End If

        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        If count > 1 Then
            dynprop = blkref.GetDynamicBlockProperties
            For i = LBound(dynprop) To UBound(dynprop)
                If dynprop(i).PropertyName = "Visibility" Then
                    prop = dynprop(i)
                    prop.Value = count & " Tags"
                End If
            Next i
        End If

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next

            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height

            If check_routines.Laychk("~-TXT" & Strings.Right(utilities.Gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Strings.Right(utilities.Gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Strings.Right(utilities.Gettxtstyle, 2)
            If ATTRIBS(i).TagString = "BOTTAG" Then
                If Cbobot.SelectedItem <> "Other" Then
                    ATTRIBS(i).TextString = Cbobot.SelectedItem.ToString & txtbot.Text
                Else
                    ATTRIBS(i).TextString = txtbot.Text
                End If
            ElseIf ATTRIBS(i).TagString = "1UPTAG" Then
                If Cbo1up.SelectedItem <> "Other" Then
                    ATTRIBS(i).TextString = Cbo1up.SelectedItem.ToString & txt1up.Text
                Else
                    ATTRIBS(i).TextString = txt1up.Text
                End If
            ElseIf ATTRIBS(i).TagString = "2UPTAG" Then
                If Cbo2up.SelectedItem <> "Other" Then
                    ATTRIBS(i).TextString = Cbo2up.SelectedItem.ToString & txt2up.Text
                Else
                    ATTRIBS(i).TextString = txt2up.Text
                End If
            ElseIf ATTRIBS(i).TagString = "3UPTAG" Then
                If Cbo3up.SelectedItem <> "Other" And Cbo3up.SelectedItem <> Nothing Then
                    ATTRIBS(i).TextString = Cbo3up.SelectedItem.ToString & txt3up.Text
                Else
                    ATTRIBS(i).TextString = txt3up.Text
                End If
            ElseIf ATTRIBS(i).TagString = "4UPTAG" Then
                If Cbo4up.SelectedItem <> "Other" Then
                    ATTRIBS(i).TextString = Cbo4up.SelectedItem.ToString & txt4up.Text
                Else
                    ATTRIBS(i).TextString = txt4up.Text
                End If
            ElseIf ATTRIBS(i).TagString = "TOPTAG" Then
                If Cbotop.SelectedItem <> "Other" Then
                    ATTRIBS(i).TextString = Cbotop.SelectedItem.ToString & txttop.Text
                Else
                    ATTRIBS(i).TextString = txttop.Text
                End If
            End If
            On Error GoTo 0
        Next i

        If count = 0 Then
            Dim minpt = Nothing
            Dim maxpt = Nothing
            ATTRIBS(0).GetBoundingBox(minpt, maxpt)
            maxpt(1) = minpt(1)
            dynprop = blkref.GetDynamicBlockProperties
            For i = LBound(dynprop) To UBound(dynprop)
                If dynprop(i).PropertyName = "Distance" Then
                    dynprop(i).Value = utilities.Getlength(minpt, maxpt) / 2
                ElseIf dynprop(i).PropertyName = "Distance1" Then
                    dynprop(i).Value = utilities.Getlength(minpt, maxpt) / 2
                End If
            Next i
        End If
        Me.Dispose()
    End Sub

    Private Sub Txttop_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txttop.TextChanged
        If txttop.Text <> "" Then
            Cbo4up.Enabled = True
            txt4up.Enabled = True
        Else
            Cbo4up.Enabled = False
            txt4up.Enabled = False
        End If
    End Sub


    Private Sub Txt4up_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt4up.TextChanged
        If txt4up.Text <> "" Then
            Cbo3up.Enabled = True
            txt3up.Enabled = True
        Else
            Cbo3up.Enabled = False
            txt3up.Enabled = False
        End If
    End Sub

    Private Sub Txt3up_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt3up.TextChanged
        If txt3up.Text <> "" Then
            Cbo2up.Enabled = True
            txt2up.Enabled = True
        Else
            Cbo2up.Enabled = False
            txt2up.Enabled = False
        End If
    End Sub

    Private Sub Txt2up_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt2up.TextChanged
        If txt2up.Text <> "" Then
            Cbo1up.Enabled = True
            txt1up.Enabled = True
        Else
            Cbo1up.Enabled = False
            txt1up.Enabled = False
        End If
    End Sub

    Private Sub Txt1up_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt1up.TextChanged
        If txt1up.Text <> "" Then
            Cbobot.Enabled = True
            txtbot.Enabled = True
        Else
            Cbobot.Enabled = False
            txtbot.Enabled = False
        End If
    End Sub

    Private Sub CmdCAN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCAN.Click
        Me.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Cmbobox As System.Windows.Forms.ComboBox
        Dim contrl As System.Windows.Forms.Control
        Dim textbox As System.Windows.Forms.TextBox

        'contrls = frmtypeform.Controls


        For Each contrl In Me.Controls
            If TypeOf contrl Is System.Windows.Forms.ComboBox Then
                Cmbobox = contrl
                Cmbobox.Items.Clear()
                Cmbobox.Items.Add("BP")
                Cmbobox.Items.Add("BR")
                Cmbobox.Items.Add("CB")
                Cmbobox.Items.Add("CC")
                Cmbobox.Items.Add("CP")
                Cmbobox.Items.Add("CW")
                Cmbobox.Items.Add("GB")
                Cmbobox.Items.Add("MF")
                Cmbobox.Items.Add("ML")
                Cmbobox.Items.Add("MW")
                Cmbobox.Items.Add("P")
                Cmbobox.Items.Add("PC")
                Cmbobox.Items.Add("PF")
                Cmbobox.Items.Add("SC")
                Cmbobox.Items.Add("SF")
                Cmbobox.Items.Add("SP")
                Cmbobox.Items.Add("ST")
                Cmbobox.Items.Add("WP")
                Cmbobox.Items.Add("Z")
                Cmbobox.Items.Add("Other")
            ElseIf TypeOf contrl Is System.Windows.Forms.TextBox Then
                textbox = contrl
                textbox.Text = ""
            End If

        Next contrl
    End Sub
End Class