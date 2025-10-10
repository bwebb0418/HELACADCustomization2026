Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Windows


Public Class Frmlayergen


    Dim layerfrm As frmlayergen
    Dim checker As New ClsCheckroutines

    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Private Sub Layergen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Client As New ClsXdata
        If Client.Check_client = True Then
            Dim writetocmd As New Clsutilities
            writetocmd.WritetoCMD("Client specific setup.  Exiting")
            Me.Dispose()
        End If

        ProgressBar1.Visible = False
        rdoAll.Checked = False
        Cbotxt.Checked = True
        Cbogen.Checked = True
        ' ''If startup.rdoIND.Checked = True Then
        ' ''    Cboind.Checked = True
        ' ''ElseIf startup.rdoBDE.Checked = True Then
        ' ''    Cbobev.Checked = True
        ' ''ElseIf startup.rdoARC.Checked = True Then
        ' ''    Cboarc.Checked = True
        ' ''End If
    End Sub

    Private Sub RdoAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoAll.CheckedChanged
        If rdoAll.Checked = True Then
            layerfrm = Me
            For Each c As System.Windows.Forms.Control In Me.Controls
                Dim t As System.Windows.Forms.CheckBox = TryCast(c, System.Windows.Forms.CheckBox)
                If t IsNot Nothing Then
                    t.Checked = True
                End If
            Next
        ElseIf rdoAll.Checked = False Then
            layerfrm = Me
            For Each c As System.Windows.Forms.Control In Me.Controls
                Dim t As System.Windows.Forms.CheckBox = TryCast(c, System.Windows.Forms.CheckBox)
                If t IsNot Nothing Then
                    t.Checked = False
                End If
            Next
        End If
    End Sub

    Private Sub ButCAN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butCAN.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Abort
        layerfrm = Me
        layerfrm.Hide()

    End Sub

    Private Sub ButGEN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butGEN.Click
        Try
            Dim NUMCHEK As Integer
            Dim NUMDONE As Integer

            NUMCHEK = 0
            NUMDONE = 0
            ProgressBar1.Visible = True

            layerfrm = Me
            For Each c As System.Windows.Forms.Control In Me.Controls
                Dim t As System.Windows.Forms.CheckBox = TryCast(c, System.Windows.Forms.CheckBox)
                If t IsNot Nothing Then
                    If t.Checked = True Then
                        NUMCHEK += 1
                    End If
                End If
            Next

            For Each c As System.Windows.Forms.Control In Me.Controls
                Dim t As System.Windows.Forms.CheckBox = TryCast(c, System.Windows.Forms.CheckBox)
                If t IsNot Nothing Then
                    If t.Checked = True Then
                        Generator(t.Tag)
                        NUMDONE += 1
                        ProgressBar1.Value = NUMDONE / NUMCHEK * 100
                    End If
                End If
            Next
            layerfrm.Hide()
            If Me.DialogResult <> System.Windows.Forms.DialogResult.Ignore Then
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Else
                Me.DialogResult = System.Windows.Forms.DialogResult.Abort
            End If
        Catch
            layerfrm.Hide()
            Me.DialogResult = System.Windows.Forms.DialogResult.Abort
        End Try

    End Sub


    Private Sub ButCLR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butCLR.Click
        rdoAll.Checked = False
    End Sub
    Private Sub Generator(ByVal catagory As String)
        Dim fh As Integer, tmp
        Dim layname, laycol, layline, layplot
        Dim objlay As AcadLayer
        Dim loadlt As New Clsutilities


        Try
            Dim Fpath As New ClsOtherRoutines
            Dim FileFullPath As String = Fpath.get_path("Layers.csv")
            Dim writetocmd As New Clsutilities
            If FileFullPath = Nothing Then

                writetocmd.WritetoCMD("File not found, check support path")
                Exit Sub
            Else
                'writetocmd.WritetoCMD("File found, " & FileFullPath)
            End If
            fh = Microsoft.VisualBasic.FreeFile
            Microsoft.VisualBasic.FileOpen(fh, FileFullPath.ToString, Microsoft.VisualBasic.OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

            Do While Not EOF(fh)
                tmp = Microsoft.VisualBasic.FileSystem.LineInput(fh)
                Dim layprops As String() = tmp.Split(New Char() {","c})
                If layprops(0) = catagory Then
                    layname = Trim(layprops(1))
                    laycol = Trim(layprops(2))
                    layline = Trim(layprops(3))
                    layplot = Trim(layprops(4))

                    objlay = Thisdrawing.Layers.Add(layname)
                    objlay.color = laycol
                    loadlt.Loadlinetype(UCase(layline))
                    'Thisdrawing.Linetypes.Load(UCase(layline), "acad.lin")
                    objlay.Linetype = layline
                    objlay.Plottable = layplot
                End If

            Loop
            FileClose(fh)
            fh = Nothing
            If Thisdrawing.GetVariable("UserS3") <> "AR_" Then
                For Each Item In Thisdrawing.Layers
                    If Strings.Left(Item.Name, 2) = "A-" Then Item.Delete()
                Next Item
            Else
                For Each Item In Thisdrawing.Layers
                    If Strings.Left(Item.Name, 6) = "S-ARCH" Then Item.Delete()
                Next Item
            End If

            If checker.Laychk("~-VPORT") = True Then
                Thisdrawing.Layers.Item("~-VPORT").Plottable = False
            End If
            Fpath = Nothing
            FileFullPath = Nothing
        Catch ex As Exception
            Me.DialogResult = System.Windows.Forms.DialogResult.Ignore

        Finally


        End Try

    End Sub

    Public Sub Newlayerconvert()
        On Error Resume Next
        Dim obj As AcadObject
        Dim lyer As AcadLayer
        For Each Item In Thisdrawing.Layers
            If Strings.Left(Item.Name, 3) = "S-L" Or Strings.Left(Item.Name, 3) = "S-T" Then
                lyer = Item
                lyer.Name = "~-" & Strings.Right(Item.Name, (Len(Item.Name)) - 2)
            ElseIf Strings.Left(Item.Name, 7) = "S-SHADE" Then
                lyer = Item
                lyer.Name = "~-" & Strings.Right(Item.Name, (Len(Item.Name)) - 2)
            ElseIf Strings.Left(Item.Name, 6) = "S-GRID" Then
                lyer = Item
                lyer.Name = "~-" & Strings.Right(Item.Name, (Len(Item.Name)) - 2)
            ElseIf Strings.Left(Item.Name, 6) = "S-XREF" Then
                lyer = Item
                lyer.Name = "~-" & Strings.Right(Item.Name, (Len(Item.Name)) - 2)
            ElseIf Strings.Left(Item.Name, 6) = "S-VPORT" Then
                lyer = Item
                lyer.Name = "~-" & Strings.Right(Item.Name, (Len(Item.Name)) - 2)
            End If
        Next Item
        On Error GoTo 0
        For Each obj In Thisdrawing.ModelSpace
            If Strings.Left(obj.Layer, 3) = "S-L" Or Strings.Left(obj.Layer, 3) = "S-T" Then
                obj.Layer = "~-" & Strings.Right(obj.Layer, (Len(obj.Layer)) - 2)
            ElseIf Strings.Left(obj.Layer, 7) = "S-SHADE" Then
                obj.Layer = "~-" & Strings.Right(obj.Layer, (Len(obj.Layer)) - 2)
            ElseIf Strings.Left(obj.Layer, 6) = "S-GRID" Then
                obj.Layer = "~-" & Strings.Right(obj.Layer, (Len(obj.Layer)) - 2)
            ElseIf Strings.Left(obj.Layer, 6) = "S-XREF" Then
                obj.Layer = "~-" & Strings.Right(obj.Layer, (Len(obj.Layer)) - 2)
            ElseIf Strings.Left(obj.Layer, 6) = "S-VPORT" Then
                obj.Layer = "~-" & Strings.Right(obj.Layer, (Len(obj.Layer)) - 2)
            End If
        Next obj

    End Sub

End Class