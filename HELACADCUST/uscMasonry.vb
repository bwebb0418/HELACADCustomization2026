Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows


Public Class uscMasonry


    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        Dim inspoint(0 To 2) As Double
        Dim rotation
        Dim blkname
        Dim blkscl = 0
        Dim prop As AcadDynamicBlockReferenceProperty
        Dim check_routines As New ClsCheckroutines
        Dim utilities As New Clsutilities
        Dim blk As AcadBlockReference
        Dim dynprop

        inspoint(0) = 0
        inspoint(1) = 0
        inspoint(2) = 0
        rotation = 0

        If blkscl = 0 Then blkscl = 1

        blkname = "hel_masonry-2023"
        If check_routines.blkchk(blkname) = True Then

        Else
            blkname &= ".dwg"
        End If

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then

            'MsgBox("Test")
            blk = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)
        Else
            blkscl = Thisdrawing.GetVariable("Userr3")
            If blkscl = 0 Then blkscl = 1
            blk = Thisdrawing.PaperSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)
        End If

        Me.Cbounit.Items.Add("Metric Units")

        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Visibility" Then
                prop = dynprop(i)
                For j = 0 To UBound(prop.AllowedValues)
                    Me.Cbounit.Items.Add(prop.AllowedValues(j))
                Next j
            End If
        Next i

        blk.Delete()

        blkname = "hel_masonry_i-2023"
        If check_routines.blkchk(blkname) = True Then

        Else
            blkname &= ".dwg"
        End If

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then


            blk = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)
        Else
            blkscl = Thisdrawing.GetVariable("Userr3")
            If blkscl = 0 Then blkscl = 1
            blk = Thisdrawing.PaperSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)
        End If

        Me.Cbounit.Items.Add("Imperial Units")

        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Visibility" Then
                prop = dynprop(i)
                For j = 0 To UBound(prop.AllowedValues)
                    Me.Cbounit.Items.Add(prop.AllowedValues(j))
                Next j
            End If
        Next i

        blk.Delete()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Function Rollup()
        'Dim ps As customization = New customization
        If HELACADCustomization.customization.m_ps.Dock = DockSides.None Then
            'yep so check for Snappable and turn off

            If HELACADCustomization.customization.m_ps.Style.Equals(32) Then
                HELACADCustomization.customization.m_ps.Style = 0
            End If
            'roll it up and toggle visibility so palette resets
            With HELACADCustomization.customization.m_ps
                .AutoRollUp = True
                .Visible = False
                .Visible = True
            End With
        Else

        End If
        Rollup = Nothing
    End Function



    Private Sub Cmdinsert_Click(sender As System.Object, e As System.EventArgs) Handles cmdinsert.Click
        Dim Blockinsert As New ClsBlock_insert

        Dim blkid As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim blkname As String = ""
        Dim check_routines As New ClsCheckroutines
        Dim Blkscl As New Clsutilities

        Rollup()

        If Cbounit.SelectedItem = Nothing Then Exit Sub

        If Strings.Left(Cbounit.SelectedItem.ToString, 3) = "8"" " Or Strings.Left(Cbounit.SelectedItem.ToString, 3) = "10""" Then
            blkname = "hel_masonry_i-2023"
        ElseIf Strings.Left(Cbounit.SelectedItem.ToString, 3) = "200" Or Strings.Left(Cbounit.SelectedItem.ToString, 3) = "250" Then
            blkname = "hel_masonry-2023"
        ElseIf LCase(Cbounit.SelectedItem.ToString) = "imperial units" Or LCase(Cbounit.SelectedItem.ToString) = "metric untis" Then
            Exit Sub
        ElseIf Cbounit.SelectedItem.ToString.Contains("Brick") Then
            Dim idex As Integer = Cbounit.SelectedIndex
            If Cbounit.Items(idex - 1) = "Metric Units" Then
                blkname = "hel_masonry-2023"
            ElseIf Cbounit.Items(idex - 2) = "Metric Units" Then
                blkname = "hel_masonry-2023"
            ElseIf Cbounit.Items(idex - 3) = "Metric Units" Then
                blkname = "hel_masonry-2023"
            ElseIf Cbounit.Items(idex - 1) = "Imperial Units" Or Cbounit.Items(idex - 2) = "Imperial Units" Or Cbounit.Items(idex - 3) = "Imperial Units" Then
                blkname = "hel_masonry_i-2023"
            End If

        End If



        If check_routines.blkchk(blkname) = True Then

        Else
            If Blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If

        blkid = Blockinsert.InsertBlockWithJig(blkname, False, "Masonry Unit Insertion Point: ", False, Blkscl.getobjectscale)
        Dim tr As Transaction = doc.Database.TransactionManager.StartTransaction
        Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
            Dim blkref As BlockReference


            blkref = tr.GetObject(blkid, OpenMode.ForWrite)

            Dim defcol As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
            For Each dbpid As DynamicBlockReferenceProperty In defcol

                If dbpid.PropertyName = "Visibility" Then
                    Try
                        dbpid.Value = Cbounit.SelectedItem.ToString
                    Catch ex As Exception

                    End Try
                End If
            Next
        End Using
        tr.Commit()
    End Sub
End Class
