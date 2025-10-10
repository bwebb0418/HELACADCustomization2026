Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
''' <summary>
''' Class for inserting various structural blocks and symbols in AutoCAD.
''' Includes commands for tags, rebar, arrows, and other engineering elements.
''' </summary>

Public Class ClsBlocks1
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    ''' <summary>
    ''' Inserts a shearwall holddown block at user-selected point.
    ''' </summary>
    End Property

    Public check_routines As New clsCheckroutines
    Public utilities As New clsutilities

    <CommandMethod("SWHD")> Public Sub ShearwallHD()
        Dim blk As AcadBlockReference
        Dim blkname As String
        Dim inspoint
        Dim blkscl
        Dim blockinsert As New clsBlock_insert
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        blkname = "hel_holddown-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If

        blkscl = Thisdrawing.GetVariable("Userr3")
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.getpoint("Select Insertion Point of Shearwall Holddown: ")
        On Error GoTo nd
        blk = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl,
            Thisdrawing.Utility.GetAngle(inspoint, "Select Holddown Rotation"))
        On Error Resume Next
        blk.Layer = "S-WHD"
nd:
    End Sub

    <CommandMethod("WATER")> Public Sub Water()
        Dim blkref As BlockReference
        Dim blkname As String
        'Dim inspoint
        Dim blkid As ObjectId
        Dim blockinsert As New clsBlock_insert
        Dim blkscl As New clsutilities
        blkname = "hel_water-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        blkid = blockinsert.InsertBlockWithJig(blkname, False, "Select Water Symbol Insertion Point: ", False, blkscl.getscale)
        Dim tr As Transaction = doc.TransactionManager.StartTransaction
        Try
            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument

                blkref = tr.GetObject(blkid, OpenMode.ForWrite)

                blkref.LayerId = check_routines.layercheck("~-L18")
            End Using
            tr.Commit()
        Catch ex As Exception
            tr.Abort()
        End Try
    End Sub

    <CommandMethod("revtri")> Public Sub Revtri()
        Dim blkref As BlockReference

        Dim blkname
        Dim blkscl As New clsutilities
        Dim blockinsert As New clsBlock_insert
        Dim blkid As ObjectId


        blkname = "hel_revtri-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim tr As Transaction = doc.TransactionManager.StartTransaction
        Try
            blkid = blockinsert.InsertBlockWithJig(blkname, False, "Select Revision Triangle Insertion Point: ", False, blkscl.getscale * Thisdrawing.GetVariable("Userr1"))

            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument

                blkref = tr.GetObject(blkid, OpenMode.ForWrite)

                blkref.LayerId = check_routines.layercheck("~-LREV")
            End Using
            tr.Commit()
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            tr = doc.TransactionManager.StartTransaction
            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                Dim attcol As AttributeCollection = blkref.AttributeCollection
                Dim tst As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
                'Dim tstyle As TextStyleTableRecord
                For Each attid As ObjectId In attcol
                    Dim attref As AttributeReference = tr.GetObject(attid, OpenMode.ForWrite)
                    attref.TextString = (ed.GetString("Drawing Revision? ")).StringResult
                    attref.TextStyleId = check_routines.textcheck(utilities.gettxtstyle)
                    attref.LayerId = check_routines.layercheck("~-TXT" & Right(utilities.gettxtstyle, 2))
                Next

            End Using
            tr.Commit()
        Catch ex As Exception
            tr.Abort()
        End Try

    End Sub

    <CommandMethod("rebar")> Public Sub Rebar()
        Dim blk As AcadBlockReference
        Dim inspoint
        Dim barsize As String
        barsize = Thisdrawing.Utility.GetString(False)

        If check_routines.blkchk(barsize) = False Then
            Build_Rebar(barsize)
        End If
        Try
            inspoint = utilities.getpoint("Select Insertion Point of " & barsize & " Bar")
            'On Error GoTo nd

            blk = Thisdrawing.ModelSpace.InsertBlock(inspoint, barsize, 1, 1, 1, 0)
        Catch
            Exit Sub
        End Try

        Try
            blk.Layer = "S-RSEC"
        Catch
            Thisdrawing.Layers.Add("S-RSEC")
            blk.Layer = "S-RSEC"
        End Try
    End Sub

    <CommandMethod("dblarrow")> Public Sub Dbl_arrow()
        Dim space
        Dim inspoint, inspoint2
        Dim blk As AcadBlockReference
        Dim ATTRIBS
        Dim rotation As Double
        Dim blkname
        Dim blkscl
        Dim blockinsert As New clsBlock_insert

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If

        inspoint = utilities.getpoint("Select Double Arrow Insertion Point: ")

        blkscl = utilities.getscale
        If blkscl = 0 Then blkscl = 1

        blkname = "hel_dbl_arrow-2026"

        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If




        On Error GoTo nd


        blk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        On Error GoTo Err
        rotation = Thisdrawing.Utility.GetAngle(inspoint, "Select Double Arrow Rotation")
        If rotation = Nothing Then rotation = utilities.degtorad(0)
        If utilities.radtodeg(rotation) > 105 And utilities.radtodeg(rotation) < 285 Then
            rotation += rotation + utilities.degtorad(180)
        End If
        On Error GoTo 0
        blk.Rotate(inspoint, rotation)

        ATTRIBS = blk.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.gettxtstyle, 2)
            If ATTRIBS(i).Tagstring = "VALUE" Then
                ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(True, "Input Attribute Value: ")
                If ATTRIBS(i).TextString = "*" Then ATTRIBS(i).TextString = ""
            End If
            On Error GoTo 0
        Next i

        blk.Layer = "0"
        blk.Update()
        Exit Sub
Err:
        blk.Delete()
nd:

    End Sub
    <CommandMethod("dia_tag")> Public Sub Dia_tag()
        Dim inspoint
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim blkscl
        Dim ATTRIBS
        Dim blockinsert As New clsBlock_insert
        blkscl = utilities.getscale
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.getpoint("Select Tag Insertion Point: ")
        blkname = "hel_dia_tag-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If
        On Error GoTo nd
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.gettxtstyle, 2)
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Tag Value: ")
            On Error GoTo 0
        Next i
nd:
    End Sub
    <CommandMethod("hex_tag")> Public Sub Hex_tag()
        Dim inspoint
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim blkscl
        Dim ATTRIBS
        Dim blockinsert As New clsBlock_insert
        blkscl = utilities.getscale
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.getpoint("Select Tag Insertion Point: ")
        blkname = "hel_hex_tag-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If

        On Error GoTo nd
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.gettxtstyle, 2)
            If ATTRIBS(i).TagString = "T1" Then
                ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Inside Value: ")
            ElseIf ATTRIBS(i).TagString = "T2" Then
                ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Outside Value: ")
            End If
            On Error GoTo 0
        Next i
nd:
    End Sub
    <CommandMethod("cir_tag")> Public Sub Cir_tag()
        Dim inspoint
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim blkscl
        Dim ATTRIBS
        Dim blockinsert As New clsBlock_insert
        blkscl = utilities.getscale
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.getpoint("Select Tag Insertion Point: ")
        blkname = "hel_cir_tag-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If
        On Error GoTo nd
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.gettxtstyle, 2)
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Tag Value: ")
            On Error GoTo 0
        Next i
nd:
    End Sub

    <CommandMethod("squ_tag")> Public Sub Squ_tag()
        Dim inspoint
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim blkscl
        Dim ATTRIBS
        Dim blockinsert As New clsBlock_insert

        blkscl = utilities.getscale
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.getpoint("Select Tag Insertion Point: ")


        blkname = "hel_squ_tag-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If
        On Error GoTo nd
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.gettxtstyle, 2)
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Tag Value: ")
            On Error GoTo 0
        Next i
nd:
    End Sub
    <CommandMethod("tri_tag")> Public Sub Tri_tag()
        Dim inspoint
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim blkscl
        Dim ATTRIBS
        Dim blockinsert As New clsBlock_insert

        blkscl = utilities.getscale
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.getpoint("Select Tag Insertion Point: ")


        blkname = "hel_tri_tag-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If
        On Error GoTo nd
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.gettxtstyle, 2)
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Tag Value: ")
            On Error GoTo 0
        Next i
nd:
    End Sub

    <CommandMethod("elip_tag")> Public Sub Ellipse_tag()
        Dim inspoint
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim blkscl
        Dim ATTRIBS
        Dim blockinsert As New clsBlock_insert

        blkscl = utilities.getscale
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.getpoint("Select Tag Insertion Point: ")


        blkname = "hel_elip_tag-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If
        On Error GoTo nd
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.gettxtstyle, 2)
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Tag Value: ")
            On Error GoTo 0
        Next i
nd:
    End Sub

    <CommandMethod("gravel")> Public Sub Gravel()
        Dim inspoint
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim blkscl
        Dim blockinsert As New clsBlock_insert

        blkscl = Thisdrawing.GetVariable("Userr3")
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.getpoint("Select Gravel Insertion point: ")


        blkname = "hel_gravel-2026"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If
        On Error Resume Next
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

    End Sub


    Public Function Build_Rebar(ByVal barsize As String)

        Dim uni As Double
        Dim circ(0) As AcadEntity
        Dim point(0 To 2) As Double
        Dim circhatch As AcadHatch
        Dim cirhat(0) As AcadEntity
        Dim lyr As String
        Dim blk As AcadBlock
        Dim blks As AcadBlocks

        uni = Thisdrawing.GetVariable("UserR3")
        If uni = 0.0# Then uni = 1
        lyr = Thisdrawing.ActiveLayer.Name
        Thisdrawing.ActiveLayer = Thisdrawing.Layers(CType("0", Index))
        point(0) = 0 : point(1) = 0 : point(2) = 0

        If UCase(barsize) = "10M" Then
            circ(0) = Thisdrawing.ModelSpace.AddCircle(point, (11.3 / 2) * uni)
        ElseIf UCase(barsize) = "15M" Then
            circ(0) = Thisdrawing.ModelSpace.AddCircle(point, (16 / 2) * uni)
        ElseIf UCase(barsize) = "20M" Then
            circ(0) = Thisdrawing.ModelSpace.AddCircle(point, (19.5 / 2) * uni)
        ElseIf UCase(barsize) = "25M" Then
            circ(0) = Thisdrawing.ModelSpace.AddCircle(point, (25.2 / 2) * uni)
        ElseIf UCase(barsize) = "30M" Then
            circ(0) = Thisdrawing.ModelSpace.AddCircle(point, (29.9 / 2) * uni)
        ElseIf UCase(barsize) = "35M" Then
            circ(0) = Thisdrawing.ModelSpace.AddCircle(point, (35.7 / 2) * uni)
        ElseIf UCase(barsize) = "45M" Then
            circ(0) = Thisdrawing.ModelSpace.AddCircle(point, (43.7 / 2) * uni)
        ElseIf UCase(barsize) = "55M" Then
            circ(0) = Thisdrawing.ModelSpace.AddCircle(point, (56.4 / 2) * uni)
        End If

        circhatch = Thisdrawing.ModelSpace.AddHatch(AcPatternType.acHatchPatternTypePreDefined, "solid", True)
        circhatch.AppendOuterLoop(circ)
        circhatch.Evaluate()

        blks = Thisdrawing.Blocks
        For Each blk In blks
            If blk.Name = barsize Then
                Dim ENTITY As AcadEntity
                For Each ENTITY In blk
                    ENTITY.Delete()
                Next ENTITY
                GoTo Rebuildblk
            End If
        Next blk
Rebuildblk:
        blk = Thisdrawing.Blocks.Add(point, barsize)
        Thisdrawing.CopyObjects(circ, blk)
        cirhat(0) = circhatch
        Thisdrawing.CopyObjects(cirhat, blk)
        circ(0).Delete()
        circhatch.Delete()

        Thisdrawing.ActiveLayer = Thisdrawing.Layers.Item(lyr)
        Thisdrawing.Regen(AcRegenType.acAllViewports)
        Build_Rebar = Nothing
    End Function



End Class
