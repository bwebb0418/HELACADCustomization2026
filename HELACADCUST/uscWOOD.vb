Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows


Public Class uscWOOD


    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        With Me.CboDLSIZE.Items
            .Add("38x38 / 2x2")
            .Add("38x64 / 2x3")
            .Add("38x89 / 2x4")
            .Add("38x140 / 2x6")
            .Add("38x184 / 2x8")
            .Add("38x235 / 2x10")
            .Add("38x286 / 2x12")
            .Add("64x64 / 3x3")
            .Add("64x89 / 3x4")
            .Add("64x140 / 3x6")
            .Add("64x184 / 3x8")
            .Add("64x235 / 3x10")
            .Add("64x286 / 3x12")
            .Add("89x89 / 4x4")
            .Add("89x140 / 4x6")
            .Add("89x184 / 4x8")
            .Add("89x235 / 4x10")
            .Add("89x286 / 4x12")
            .Add("140x140 / 6x6")
            .Add("140x184 / 6x8")
            .Add("140x235 / 6x10")
            .Add("140x286 / 6x12")
            .Add("140x336 / 6x14")
            .Add("140x387 / 6x16")
            .Add("184x184 / 8x8")
            .Add("184x235 / 8x10")
            .Add("184x286 / 8x12")
            .Add("184x336 / 8x14")
            .Add("184x387 / 8x16")
            .Add("235x235 / 10x10")
            .Add("235x286 / 10x12")
            .Add("235x336 / 10x14")
            .Add("235x387 / 10x16")
            .Add("286x286 / 12x12")
            .Add("286x336 / 12x14")
            .Add("286x387 / 12x16")
        End With

        With Me.CboRSSIZE.Items
            .Add("51x51 / 2x2")
            .Add("51x76 / 2x3")
            .Add("51x102 / 2x4")
            .Add("51x152 / 2x6")
            .Add("51x203 / 2x8")
            .Add("51x254 / 2x10")
            .Add("51x305 / 2x12")
            .Add("76x76 / 3x3")
            .Add("76x102 / 3x4")
            .Add("76x152 / 3x6")
            .Add("76x203 / 3x8")
            .Add("76x254 / 3x10")
            .Add("76x305 / 3x12")
            .Add("102x102 / 4x4")
            .Add("102x152 / 4x6")
            .Add("102x203 / 4x8")
            .Add("102x254 / 4x10")
            .Add("102x305 / 4x12")
            .Add("152x152 / 6x6")
            .Add("152x203 / 6x8")
            .Add("152x254 / 6x10")
            .Add("152x305 / 6x12")
            .Add("152x356 / 6x14")
            .Add("152x406 / 6x16")
            .Add("203x203 / 8x8")
            .Add("203x254 / 8x10")
            .Add("203x305 / 8x12")
            .Add("203x356 / 8x14")
            .Add("203x406 / 8x16")
            .Add("254x254 / 10x10")
            .Add("254x305 / 10x12")
            .Add("254x356 / 10x14")
            .Add("254x406 / 10x16")
            .Add("305x305 / 12x12")
            .Add("305x356 / 12x14")
            .Add("305x406 / 12x16")
        End With

        With Me.Cboglwidth.Items
            .Add("80 / 3""")
            .Add("130 / 5""")
            .Add("175 / 6 7/8""")
            .Add("215 / 8 1/2""")
            .Add("265 / 10 1/4""")
            .Add("315 / 12 1/4""")
            .Add("365 / 14 1/4""")
        End With

        With Me.Cbomlwidth.Items
            .Add("1 3/4""")
        End With

        With Me.Cboplwidth.Items
            .Add("1 3/4""")
            .Add("3 1/2""")
            .Add("5 1/4""")
            .Add("7""")
        End With

        With Me.Cbotjiseries.Items
            .Add("TJI 230")
            .Add("TJI 300C")
            .Add("TJI 400C")
            .Add("TJI 560")
            .Add("TJI L65")
        End With

        With Me.Cbotswidth.Items
            .Add("1 1/4""")
            .Add("1 1/2""")
            .Add("1 3/4""")
            .Add("2""")
            .Add("3 1/2""")
        End With
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Function Mettoimp(ByVal size)
        Mettoimp = ""
        Select Case size
            Case 38
                Mettoimp = 1.5
            Case 51
                Mettoimp = 2
            Case 64
                Mettoimp = 2.5
            Case 76
                Mettoimp = 3
            Case 89
                Mettoimp = 3.5
            Case 102
                Mettoimp = 4
            Case 140
                Mettoimp = 5.5
            Case 152
                Mettoimp = 6
            Case 184
                Mettoimp = 7.25
            Case 203
                Mettoimp = 8
            Case 235
                Mettoimp = 9.25
            Case 254
                Mettoimp = 10
            Case 286
                Mettoimp = 11.25
            Case 305
                Mettoimp = 12
            Case 337
                Mettoimp = 13.25
            Case 356
                Mettoimp = 14
            Case 387
                Mettoimp = 15.25
            Case 406
                Mettoimp = 16
        End Select
        Return Mettoimp
    End Function
    Function GetJ(ByVal i As Integer)
        Dim j As Double
        j = (i / 38) * 1.5
        If Strings.Right(j, 2) = ".5" Then
            GetJ = Fix(j) & " 1/2"
        Else
            GetJ = j
        End If
    End Function

    Private Sub Cboglwidth_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cboglwidth.DropDownClosed
        Dim i As Integer


        Me.Cbogldepth.Items.Clear()
        Me.Cbogldepth.SelectedValue = ""
        With Cbogldepth
            If Cboglwidth.Items.IndexOf(0) = True Then
                i = 114
                Do Until i > 570
                    Cbogldepth.Items.Add(i & " / " & getJ(i) & """")
                    i += 38
                Loop

            ElseIf Cboglwidth.Items.IndexOf(1) = True Then
                i = 152
                Do Until i > 950
                    Cbogldepth.Items.Add(i & " / " & getJ(i) & """")
                    i += 38
                Loop
            ElseIf Cboglwidth.Items.IndexOf(2) = True Then
                i = 190
                Do Until i > 1254
                    Cbogldepth.Items.Add(i & " / " & getJ(i) & """")
                    i += 38
                Loop
            ElseIf Cboglwidth.Items.IndexOf(3) = True Then
                i = 266
                Do Until i > 1596
                    Cbogldepth.Items.Add(i & " / " & getJ(i) & """")
                    i += 38
                Loop
            ElseIf Cboglwidth.Items.IndexOf(4) = True Then
                i = 342
                Do Until i > 1976
                    Cbogldepth.Items.Add(i & " / " & getJ(i) & """")
                    i += 38
                Loop
            ElseIf Cboglwidth.Items.IndexOf(5) = True Or Cboglwidth.Items.IndexOf(6) = True Then
                i = 380
                Do Until i > 2128
                    Cbogldepth.Items.Add(i & " / " & getJ(i) & """")
                    i += 38
                Loop
            End If
        End With
    End Sub

    Private Sub Cboplwidth_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cboplwidth.DropDownClosed
        Cbopldepth.Items.Clear()
        Cbopldepth.SelectedValue = ""
        With Cbopldepth
            .Items.Add("9 1/4""")
            .Items.Add("9 1/2""")
            .Items.Add("11 7/8""")
            .Items.Add("14""")
            .Items.Add("16""")
            .Items.Add("19""")
        End With
    End Sub

    Private Sub Cbomlwidth_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cbomlwidth.DropDownClosed
        Cbomldepth.Items.Clear()
        Cbomldepth.SelectedValue = ""
        With Cbomldepth
            .Items.Add("9 1/2""")
            .Items.Add("11 7/8""")
            .Items.Add("14""")
            .Items.Add("16""")
            .Items.Add("18 3/4""")
        End With
    End Sub

    Private Sub Cbotjiseries_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cbotjiseries.DropDownClosed
        Cbotjidepth.Items.Clear()
        Cbotjidepth.SelectedValue = ""
        If Cbotjiseries.SelectedItem.ToString = "TJI 230" Then
            With Cbotjidepth
                .Items.Add("9 1/2""")
                .Items.Add("11 7/8""")
                .Items.Add("14""")
                .Items.Add("16""")
            End With
        ElseIf Cbotjiseries.SelectedItem.ToString = "TJI 300C" Or Cbotjiseries.SelectedItem.ToString = "TJI 400C" Then
            With Cbotjidepth
                .Items.Add("9 1/2""")
                .Items.Add("11 7/8""")
                .Items.Add("14""")
                .Items.Add("16""")
                .Items.Add("18""")
                .Items.Add("20""")
            End With
        ElseIf Cbotjiseries.SelectedItem.ToString = "TJI 560" Then
            With Cbotjidepth
                .Items.Add("11 7/8""")
                .Items.Add("14""")
                .Items.Add("16""")
                .Items.Add("18""")
                .Items.Add("20""")
            End With
        ElseIf Cbotjiseries.SelectedItem.ToString = "TJI L65" Then
            With Cbotjidepth
                .Items.Add("22""")
                .Items.Add("24""")
                .Items.Add("26""")
                .Items.Add("28""")
            End With
        End If
    End Sub

    Private Sub Cbotswidth_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cbotswidth.DropDownClosed
        Cbotsdepth.Items.Clear()
        Cbotsdepth.SelectedValue = ""
        With Cbotsdepth
            If Cbotswidth.SelectedItem.ToString = "1 1/4""" Then
                .Items.Add("9 1/2""")
                .Items.Add("11 7/8""")
                .Items.Add("14""")
                .Items.Add("16""")
                .Items.Add("18""")
                .Items.Add("20""")
            ElseIf Cbotswidth.SelectedItem.ToString = "1 1/2""" Then
                .Items.Add("22""")
                .Items.Add("24""")
            ElseIf Cbotswidth.SelectedItem.ToString = "1 3/4""" Or Cbotswidth.SelectedItem.ToString = "3 1/2""" Then
                .Items.Add("9 1/2""")
                .Items.Add("11 7/8""")
                .Items.Add("14""")
                .Items.Add("16""")
            ElseIf Cbotswidth.SelectedItem.ToString = "2""" Then
                .Items.Add("4""")
                .Items.Add("6""")
                .Items.Add("8""")
            End If
        End With
    End Sub

    Private Sub CmdTJ_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdTJ.Click

        Dim blockinsert As New ClsBlock_insert
        Dim width As Double
        Dim Depth As Double
        Dim blkid As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim blkname As String
        Dim check_routines As New ClsCheckroutines
        Dim blkscl As New Clsutilities
        Dim tr As Transaction = doc.Database.TransactionManager.StartTransaction

        Using tr
            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                Try
                    rollup()

                    blkname = "hel_TJI-2023"
                    If check_routines.Blkchk(blkname) = True Then

                    Else
                        If blockinsert.Loadblock(blkname & ".dwg") = False Then
                            Exit Sub
                        End If
                    End If

                    If Cbotjiseries.SelectedItem = Nothing Or Cbotjidepth.SelectedItem = Nothing Then Exit Sub

                    If Cbotjiseries.SelectedItem.ToString = "TJI 230" Then
                        width = 2.3125 * 25.4
                    ElseIf Cbotjiseries.SelectedItem.ToString = "TJI 300C" Then
                        width = 2.5 * 25.4
                    ElseIf Cbotjiseries.SelectedItem.ToString = "TJI 400C" Then
                        width = 3.5 * 25.4
                    ElseIf Cbotjiseries.SelectedItem.ToString = "TJI 560" Then
                        width = (3.5 * 25.4)
                    ElseIf Cbotjiseries.SelectedItem.ToString = "TJI L65" Then
                        width = 2.5 * 25.4
                    End If

                    If Cbotjidepth.SelectedItem.ToString = "9 1/2""" Then
                        Depth = (9.5 * 25.4)
                    ElseIf Cbotjidepth.SelectedItem.ToString = "11 7/8""" Then
                        Depth = (11.875 * 25.4)
                    ElseIf Cbotjidepth.SelectedItem.ToString = "14""" Then
                        Depth = (14 * 25.4)
                    ElseIf Cbotjidepth.SelectedItem.ToString = "16""" Then
                        Depth = (16 * 25.4)
                    ElseIf Cbotjidepth.SelectedItem.ToString = "18""" Then
                        Depth = (18 * 25.4)
                    ElseIf Cbotjidepth.SelectedItem.ToString = "20""" Then
                        Depth = (20 * 25.4)
                    ElseIf Cbotjidepth.SelectedItem.ToString = "22""" Then
                        Depth = (22 * 25.4)
                    ElseIf Cbotjidepth.SelectedItem.ToString = "24""" Then
                        Depth = (24 * 25.4)
                    ElseIf Cbotjidepth.SelectedItem.ToString = "26""" Then
                        Depth = (26 * 25.4)
                    ElseIf Cbotjidepth.SelectedItem.ToString = "28""" Then
                        Depth = (28 * 25.4)
                    End If

                    blkid = blockinsert.InsertBlockWithJig(blkname, False, "Select Trus Joist Insertion Point: ", False, blkscl.getobjectscale)


                    Dim blkref As BlockReference



                    blkref = tr.GetObject(blkid, OpenMode.ForWrite)
                    Dim defcol As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
                    For Each dbpid As DynamicBlockReferenceProperty In defcol

                        If dbpid.PropertyName = "Halfwidth" Then
                            Try
                                dbpid.Value = width / 2 * blkscl.getobjectscale
                            Catch ex As Exception

                            End Try
                        ElseIf dbpid.PropertyName = "Depth" Then
                            Try
                                dbpid.Value = Depth * blkscl.getobjectscale
                            Catch ex As Exception

                            End Try
                        End If

                    Next

                Catch ex As Exception
                    tr.Abort()
                End Try
                tr.Commit()
            End Using
        End Using
    End Sub


    Function Rollup()
        'Dim ps As customization = New customization
        If HELACADCustomization.Customization.m_ps.Dock = DockSides.None Then
            'yep so check for Snappable and turn off

            If HELACADCustomization.Customization.m_ps.Style.Equals(32) Then
                HELACADCustomization.Customization.m_ps.Style = 0
            End If
            'roll it up and toggle visibility so palette resets
            With HELACADCustomization.Customization.m_ps
                .AutoRollUp = True
                .Visible = False
                .Visible = True
            End With
        Else

        End If
        Rollup = Nothing
    End Function
    Private Shared Clock As System.Windows.Forms.Timer
    Friend Shared Sub CreateTimer()
        Clock = New System.Windows.Forms.Timer With {
        .Interval = 500
        }
        Clock.Start()

        AddHandler Clock.Tick, AddressOf Timer_Tick
    End Sub

    Friend Shared Sub Timer_Tick(ByVal sender As Object, ByVal eArgs As EventArgs)

        If sender Is Clock Then
            Try
                With HELACADCustomization.customization.m_ps
                    '.AutoRollUp = True
                    .RolledUp = True
                    '.Visible = False
                    '.Visible = True
                End With
                'stop the clock and destroy it
                Clock.Stop()
                Clock.Dispose()
            Catch ex As Exception
                If HELACADCustomization.customization.m_ps.AutoRollUp.Equals(True) Then
                    'stop the clock and destroy it
                    Clock.Stop()
                    Clock.Dispose()
                End If
            End Try
        End If
    End Sub

    Private Sub CmdGL_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdGL.Click
        Dim blockinsert As New ClsBlock_insert
        Dim width As Double
        Dim Depth As Double
        Dim blkid As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim blkname As String
        Dim check_routines As New ClsCheckroutines
        Dim blkscl As New Clsutilities
        Dim tr As Transaction = doc.Database.TransactionManager.StartTransaction
        Using tr
            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                Try

                    Rollup()


                    blkname = "hel_glulam-2023"
                    If check_routines.Blkchk(blkname) = True Then

                    Else
                        If blockinsert.Loadblock(blkname & ".dwg") = False Then
                            Exit Sub
                        End If
                    End If
                    If Cboglwidth.SelectedItem = Nothing Or Cbogldepth.SelectedItem = Nothing Then Exit Sub
                    blkid = blockinsert.InsertBlockWithJig(blkname, False, "Select Glulam Insertion Point: ", False, blkscl.Getobjectscale)


                    Dim blkref As BlockReference


                    width = CDbl(Trim(Strings.Left(Cboglwidth.SelectedItem.ToString, 3)))
                    Depth = CDbl(Trim(Strings.Left(Cbogldepth.SelectedItem.ToString, 4)))

                    blkref = tr.GetObject(blkid, OpenMode.ForWrite)
                    Dim defcol As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
                    For Each dbpid As DynamicBlockReferenceProperty In defcol

                        If dbpid.PropertyName = "Halfwidth" Then
                            Try
                                dbpid.Value = width / 2 * blkscl.Getobjectscale
                            Catch ex As Exception

                            End Try
                        ElseIf dbpid.PropertyName = "Depth" Then
                            Try
                                dbpid.Value = (Depth - (38 * 2)) * blkscl.Getobjectscale
                            Catch ex As Exception

                            End Try
                        End If

                    Next



                Catch ex As Exception

                    tr.Abort()
                End Try
                tr.Commit()

            End Using
        End Using
    End Sub

    Private Sub CmdPL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPL.Click
        Dim blockinsert As New ClsBlock_insert
        Dim width As Double
        Dim Depth As Double
        Dim blkid As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim blkname As String
        Dim check_routines As New ClsCheckroutines
        Dim blkscl As New Clsutilities
        Dim tr As Transaction = doc.Database.TransactionManager.StartTransaction
        Using tr
            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                Try
                    Rollup()


                    blkname = "hel_parabeam-2023"
                    If check_routines.Blkchk(blkname) = True Then

                    Else
                        If blockinsert.Loadblock(blkname & ".dwg") = False Then
                            Exit Sub
                        End If
                    End If

                    If Cbopldepth.SelectedItem = Nothing Or Cboplwidth.SelectedItem = Nothing Then Exit Sub

                    If Cboplwidth.SelectedItem.ToString = "1 3/4""" Then
                        width = 1.75 * 25.4
                    ElseIf Cboplwidth.SelectedItem.ToString = "3 1/2""" Then
                        width = 3.5 * 25.4
                    ElseIf Cboplwidth.SelectedItem.ToString = "5 1/4""" Then
                        width = 5.25 * 25.4
                    ElseIf Cboplwidth.SelectedItem.ToString = "7""" Then
                        width = (7 * 25.4)
                    End If

                    If Cbopldepth.SelectedItem.ToString = "9 1/4""" Then
                        Depth = (9.25 * 25.4)
                    ElseIf Cbopldepth.SelectedItem.ToString = "9 1/2""" Then
                        Depth = (9.5 * 25.4)
                    ElseIf Cbopldepth.SelectedItem.ToString = "11 7/8""" Then
                        Depth = (11.875 * 25.4)
                    ElseIf Cbopldepth.SelectedItem.ToString = "14""" Then
                        Depth = (14 * 25.4)
                    ElseIf Cbopldepth.SelectedItem.ToString = "16""" Then
                        Depth = (16 * 25.4)
                    ElseIf Cbopldepth.SelectedItem.ToString = "19""" Then
                        Depth = (19 * 25.4)
                    End If

                    blkid = blockinsert.InsertBlockWithJig(blkname, False, "Select Paralam Insertion Point: ", False, blkscl.Getobjectscale)


                    Dim blkref As BlockReference


                    blkref = tr.GetObject(blkid, OpenMode.ForWrite)
                    Dim defcol As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
                    For Each dbpid As DynamicBlockReferenceProperty In defcol

                        If dbpid.PropertyName = "Halfwidth" Then
                            Try
                                dbpid.Value = width / 2 * blkscl.Getobjectscale
                            Catch ex As Exception

                            End Try
                        ElseIf dbpid.PropertyName = "Depth" Then
                            Try
                                dbpid.Value = Depth * blkscl.Getobjectscale
                            Catch ex As Exception

                            End Try
                        End If

                    Next

                Catch ex As Exception
                    tr.Abort()
                End Try
                tr.Commit()
            End Using
        End Using
    End Sub

    Private Sub CmdML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdML.Click
        Dim blockinsert As New ClsBlock_insert
        Dim width As Double
        Dim Depth As Double
        Dim blkid As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim blkname As String
        Dim check_routines As New ClsCheckroutines
        Dim blkscl As New Clsutilities
        Dim tr As Transaction = doc.Database.TransactionManager.StartTransaction
        Using tr
            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                Try
                    Rollup()

                    If Cbomldepth.SelectedItem = Nothing Or Cbomlwidth.SelectedItem = Nothing Then Exit Sub

                    If Cbomlwidth.SelectedItem.ToString = "1 3/4""" Then
                        width = 1.75 * 25.4
                    End If

                    If Cbomldepth.SelectedItem.ToString = "9 1/2""" Then
                        Depth = (9.5 * 25.4)
                    ElseIf Cbomldepth.SelectedItem.ToString = "11 7/8""" Then
                        Depth = (11.875 * 25.4)
                    ElseIf Cbomldepth.SelectedItem.ToString = "14""" Then
                        Depth = (14 * 25.4)
                    ElseIf Cbomldepth.SelectedItem.ToString = "16""" Then
                        Depth = (16 * 25.4)
                    ElseIf Cbomldepth.SelectedItem.ToString = "18 3/4""" Then
                        Depth = (18.75 * 25.4)
                    End If

                    blkname = "hel_microbeam-2023"
                    If check_routines.Blkchk(blkname) = True Then

                    Else
                        If blockinsert.Loadblock(blkname & ".dwg") = False Then
                            Exit Sub
                        End If
                    End If

                    blkid = blockinsert.InsertBlockWithJig(blkname, False, "Select Microlam Insertion Point: ", False, blkscl.Getobjectscale)


                    Dim blkref As BlockReference

                    blkref = tr.GetObject(blkid, OpenMode.ForWrite)
                    Dim defcol As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
                    For Each dbpid As DynamicBlockReferenceProperty In defcol

                        If dbpid.PropertyName = "Halfwidth" Then
                            Try
                                dbpid.Value = width / 2 * blkscl.Getobjectscale
                            Catch ex As Exception

                            End Try
                        ElseIf dbpid.PropertyName = "Depth" Then
                            Try
                                dbpid.Value = Depth * blkscl.Getobjectscale
                            Catch ex As Exception

                            End Try
                        End If

                    Next

                Catch ex As Exception
                    tr.Abort()
                End Try
                tr.Commit()
            End Using
        End Using
    End Sub

    Private Sub CmdTS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTS.Click
        Dim blockinsert As New ClsBlock_insert
        Dim width As Double
        Dim Depth As Double
        Dim blkid As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim blkname As String
        Dim check_routines As New ClsCheckroutines
        Dim blkscl As New Clsutilities
        Dim tr As Transaction = doc.Database.TransactionManager.StartTransaction
        Using tr
            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                Try
                    Rollup()

                    If Cbotsdepth.SelectedItem = Nothing Or Cbotswidth.SelectedItem = Nothing Then Exit Sub

                    If Cbotswidth.SelectedItem.ToString = "1 1/4""" Then
                        width = 1.25 * 25.4
                    ElseIf Cbotswidth.SelectedItem.ToString = "1 1/2""" Then
                        width = 1.5 * 25.4
                    ElseIf Cbotswidth.SelectedItem.ToString = "1 3/4""" Then
                        width = 1.75 * 25.4
                    ElseIf Cbotswidth.SelectedItem.ToString = "2""" Then
                        width = 2 * 25.4
                    ElseIf Cbotswidth.SelectedItem.ToString = "3 1/2""" Then
                        width = 3.5 * 25.4
                    End If

                    If Cbotsdepth.SelectedItem.ToString = "4""" Then
                        Depth = (4 * 25.4)
                    ElseIf Cbotsdepth.SelectedItem.ToString = "6""" Then
                        Depth = (6 * 25.4)
                    ElseIf Cbotsdepth.SelectedItem.ToString = "8""" Then
                        Depth = (8 * 25.4)
                    ElseIf Cbotsdepth.SelectedItem.ToString = "9 1/2""" Then
                        Depth = (9.5 * 25.4)
                    ElseIf Cbotsdepth.SelectedItem.ToString = "11 7/8""" Then
                        Depth = (11.875 * 25.4)
                    ElseIf Cbotsdepth.SelectedItem.ToString = "14""" Then
                        Depth = (14 * 25.4)
                    ElseIf Cbotsdepth.SelectedItem.ToString = "16""" Then
                        Depth = (16 * 25.4)
                    ElseIf Cbotsdepth.SelectedItem.ToString = "18""" Then
                        Depth = (18 * 25.4)
                    ElseIf Cbotsdepth.SelectedItem.ToString = "20""" Then
                        Depth = (20 * 25.4)
                    ElseIf Cbotsdepth.SelectedItem.ToString = "22""" Then
                        Depth = (22 * 25.4)
                    ElseIf Cbotsdepth.SelectedItem.ToString = "24""" Then
                        Depth = (24 * 25.4)

                    End If

                    blkname = "hel_tmbrstrnd-2023"
                    If check_routines.Blkchk(blkname) = True Then

                    Else
                        If blockinsert.Loadblock(blkname & ".dwg") = False Then
                            Exit Sub
                        End If
                    End If

                    blkid = blockinsert.InsertBlockWithJig(blkname, False, "Select Timberstrand Insertion Point: ", False, blkscl.Getobjectscale)


                    Dim blkref As BlockReference

                    blkref = tr.GetObject(blkid, OpenMode.ForWrite)
                    Dim defcol As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
                    For Each dbpid As DynamicBlockReferenceProperty In defcol

                        If dbpid.PropertyName = "Halfwidth" Then
                            Try
                                dbpid.Value = width / 2 * blkscl.Getobjectscale
                            Catch ex As Exception

                            End Try
                        ElseIf dbpid.PropertyName = "Depth" Then
                            Try
                                dbpid.Value = Depth * blkscl.Getobjectscale
                            Catch ex As Exception

                            End Try
                        End If

                    Next

                Catch ex As Exception
                    tr.Abort()
                End Try
                tr.Commit()
            End Using
        End Using
    End Sub

    Private Sub CmdDL_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDLFLAT.Click
        Insertlumber(sender.name)
    End Sub

    Function Insertlumber(ByVal sender As String)
        Dim blockinsert As New ClsBlock_insert
        Dim width As Double
        Dim Depth As Double
        Dim blkid As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim blkname As String
        Dim check_routines As New ClsCheckroutines
        Dim blkscl As New Clsutilities
        Dim size
        Dim tr As Transaction = doc.Database.TransactionManager.StartTransaction
        Insertlumber = Nothing
        Rollup()
        Using tr
            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                Try
                    If Strings.Left(sender, 5) = "cmdDL" Then
                        If CboDLSIZE.SelectedItem = Nothing Then Exit Function
                        size = Trim(Strings.Left(CboDLSIZE.SelectedItem.ToString, 7))
                    Else
                        If CboRSSIZE.SelectedItem = Nothing Then Exit Function
                        size = Trim(Strings.Left(CboRSSIZE.SelectedItem.ToString, 7))
                    End If

                    If Strings.Right(size, 1) = "/" Then size = Trim(Strings.Left(size, 6))

                    If Len(size) = 5 Then
                        width = CDbl(Strings.Left(size, 2))
                        Depth = CDbl(Strings.Right(size, 2))
                    ElseIf Len(size) = 6 Then
                        width = CDbl(Strings.Left(size, 2))
                        Depth = CDbl(Strings.Right(size, 3))
                    ElseIf Len(size) = 7 Then
                        width = CDbl(Strings.Left(size, 3))
                        Depth = CDbl(Strings.Right(size, 3))
                    End If

                    blkname = "hel_lumber-2023"
                    If check_routines.Blkchk(blkname) = True Then

                    Else
                        If blockinsert.Loadblock(blkname & ".dwg") = False Then
                            Exit Function
                        End If
                    End If

                    blkid = blockinsert.InsertBlockWithJig(blkname, False, "Select Lumber Insertion Point: ", False, blkscl.Getobjectscale)
                    Dim blkref As BlockReference

                    blkref = tr.GetObject(blkid, OpenMode.ForWrite)
                    Dim defcol As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
                    For Each dbpid As DynamicBlockReferenceProperty In defcol

                        If dbpid.PropertyName = "Width" Then
                            Try
                                If Strings.Right(sender, 4) = "EDGE" Then
                                    dbpid.Value = width * blkscl.Getobjectscale
                                Else
                                    dbpid.Value = Depth * blkscl.Getobjectscale
                                End If
                            Catch ex As Exception

                            End Try
                        ElseIf dbpid.PropertyName = "Depth" Then
                            Try
                                If Strings.Right(sender, 4) = "EDGE" Then
                                    dbpid.Value = Depth * blkscl.Getobjectscale
                                Else
                                    dbpid.Value = width * blkscl.Getobjectscale
                                End If
                            Catch ex As Exception

                            End Try
                        End If

                    Next
                    tr.Commit()
                    Insertlumber = Nothing
                Catch ex As Exception
                    tr.Abort()
                    Insertlumber = Nothing
                End Try

            End Using
        End Using

    End Function

    Private Sub CmdDLEDGE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDLEDGE.Click

        Insertlumber(sender.name)
    End Sub

    Private Sub CmdRSFLAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRSFLAT.Click

        Insertlumber(sender.name)
    End Sub

    Private Sub CmdRSEDGE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRSEDGE.Click

        Insertlumber(sender.name)
    End Sub

    Private Sub UscWOOD_Disposed(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.Windows.PalettePersistEventArgs) Handles Me.Disposed
        e.ConfigurationSection.WriteProperty("WoodRoutines", 20)
    End Sub

    'Private Sub uscWOOD_Load(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.Windows.PalettePersistEventArgs) Handles Me.Load
    '    Dim a As Double = CType(e.ConfigurationSection.ReadProperty("WoodRoutines", 20), Double)
    'End Sub
End Class
