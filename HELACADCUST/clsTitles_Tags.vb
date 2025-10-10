Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry

Public Class ClsTitles_Tags

    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public utilities As New clsutilities
    Public check_routines As New clsCheckroutines
    Public text_dimensions As New ClsXdata



    Public Shared Function GETSCL() As String
        Dim SCL As Double
        Dim xdata As New ClsXdata

        GETSCL = Nothing

        If xdata.Check_xdata = False Then


            If Application.GetSystemVariable("measurement") = 1 Then
                GETSCL = "1:" & Application.GetSystemVariable("Dimscale")
            Else
                If Application.GetSystemVariable("Dimscale") > 8 Then
                    GETSCL = "1/" & SCL / 12 & """ = 1'-0"""
                ElseIf Application.GetSystemVariable("Dimscale") = 8 Then
                    GETSCL = "1 1/2"" = 1'-0"""
                ElseIf Application.GetSystemVariable("Dimscale") = 12 Then
                    GETSCL = "1"" = 1'-0"""
                ElseIf Application.GetSystemVariable("Dimscale") = 4 Then
                    GETSCL = "3"" = 1'-0"""
                End If
            End If

        Else
            SCL = Application.GetSystemVariable("UserR2")
            Dim xd2
            xd2 = xdata.Readxdata("Units")
            If Strings.Right(xd2, Strings.Len(xd2) - 6) = "fts" Then
                GETSCL = "1"" = " & SCL & "'"
            ElseIf Strings.Right(xd2, Strings.Len(xd2) - 6) = "ins" Then
                If SCL = 64 Or SCL = 32 Or SCL = 16 Then
                    GETSCL = "3/" & SCL / 4 & """ = 1'-0"""
                ElseIf SCL = 192 Or SCL = 96 Or SCL = 48 Or SCL = 24 Then
                    GETSCL = "1/" & SCL / 12 & """ = 1'-0"""
                ElseIf SCL = 8 Then
                    GETSCL = "1 1/2"" = 1'-0"""
                ElseIf SCL = 12 Then
                    GETSCL = "1"" = 1'-0"""
                ElseIf SCL = 4 Then
                    GETSCL = "3"" = 1'-0"""
                Else
                    GETSCL = ""
                End If
            ElseIf Strings.Right(xd2, Strings.Len(xd2) - 6) = "mms" Then
                GETSCL = "1:" & SCL
            ElseIf Strings.Right(xd2, Strings.Len(xd2) - 6) = "mts" Then
                GETSCL = "1:" & SCL * 1000 & "m"
            End If
        End If
    End Function

    <CommandMethod("sect")> Public Sub SectionMark()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim blkbubble As BlockReference
        Dim blkscl As New Clsutilities
        Dim blksecend As BlockReference
        Dim opclo As String
        Dim pko As New PromptKeywordOptions("Open or Closed? ")
        pko.Keywords.Add("Open")
        pko.Keywords.Add("Closed")

        Dim pkr As PromptResult = ed.GetKeywords(pko)
        Dim blockinsert As New ClsBlock_insert
        Dim blkname As String
        Dim objid As ObjectId
        Dim blkbubend As BlockReference

        Try


            Using tr As Transaction = db.TransactionManager.StartTransaction


                opclo = LCase(Strings.Left(pkr.StringResult, 1))

                If LCase(Thisdrawing.GetVariable("UserS1")) = "inds" Then
                    ClsBlock_insert.BLKSFFACT = 1
                Else
                    ClsBlock_insert.BLKSFFACT = 2 / 3
                End If

                blkname = "hel_sym_2p" & opclo & "-2023"
                If check_routines.Blkchk(blkname) = True Then

                Else
                    If blockinsert.Loadblock(blkname & ".dwg") = False Then
                        Exit Sub
                    End If
                End If

                objid = blockinsert.InsertBlockWithJig(blkname, False, "Section Bubble Insertion Point: ", False, blkscl.Getscale)
                If objid = Nothing Then Exit Sub
                tr.Commit()
            End Using
            Using tr = db.TransactionManager.StartTransaction

                blkbubble = tr.GetObject(objid, OpenMode.ForWrite)
                Clsblock_jig.BasePt = blkbubble.Position
                blkbubble.LayerId = check_routines.Layercheck("~-TSYM3")

                blkname = "hel_sym_secend-2023"
                If check_routines.Blkchk(blkname) = True Then

                Else
                    If blockinsert.Loadblock(blkname & ".dwg") = False Then
                        Exit Sub
                    End If
                End If

                objid = blockinsert.InsertBlockWithJig(blkname, False, "Section End Insertion Point: ", True, blkscl.Getscale, True)
                If objid = Nothing Then
                    blkbubble.Erase(True)
                    tr.Commit()
                    Exit Sub
                End If
                blksecend = tr.GetObject(objid, OpenMode.ForWrite)


                blkname = "hel_sym_secarr-2023"
                If check_routines.Blkchk(blkname) = True Then

                Else
                    If blockinsert.Loadblock(blkname & ".dwg") = False Then
                        Exit Sub
                    End If
                End If

                objid = blockinsert.InsertBlock(blkbubble.Position, blkname, blkbubble.ScaleFactors.X, blkbubble.ScaleFactors.Y, blkbubble.ScaleFactors.Z, blksecend.Rotation)
                blkbubend = tr.GetObject(objid, OpenMode.ForWrite)

                Dim defcol As DynamicBlockReferencePropertyCollection = blkbubend.DynamicBlockReferencePropertyCollection
                For Each dbpid As DynamicBlockReferenceProperty In defcol
                    If dbpid.PropertyName = "hel_section_line" Then
                        dbpid.Value = 15 * blkbubble.ScaleFactors.X
                    End If

                Next

                blksecend.LayerId = check_routines.Layercheck("~-TSYM3")
                blkbubend.LayerId = check_routines.Layercheck("~-TSYM3")

                Dim attcol As AttributeCollection = blkbubble.AttributeCollection
                Dim tst As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
                Dim tstyle As TextStyleTableRecord
                For Each attid As ObjectId In attcol
                    Dim attref As AttributeReference = tr.GetObject(attid, OpenMode.ForWrite)

                    If attref.Tag = "REF" Then
                        If LCase(Thisdrawing.GetVariable("UserS1")) = "inds" Then
                            tstyle = tr.GetObject(tst.Item("HEL40"), OpenMode.ForRead)
                            attref.LayerId = check_routines.Layercheck("~-TXT40")
                        Else
                            tstyle = tr.GetObject(tst.Item("HEL25"), OpenMode.ForRead)
                            attref.LayerId = check_routines.Layercheck("~-TXT25")
                        End If
                        attref.TextString = (ed.GetString("View Reference Letter? ")).StringResult
                        attref.TextStyleId = tstyle.ObjectId
                        attref.Height = tstyle.TextSize
                        attref.WidthFactor = 0.9

                    Else
                        If LCase(Thisdrawing.GetVariable("UserS1")) = "inds" Then
                            tstyle = tr.GetObject(tst.Item("HEL25"), OpenMode.ForRead)
                            attref.LayerId = check_routines.Layercheck("~-TXT25")
                        Else
                            tstyle = tr.GetObject(tst.Item("HEL15"), OpenMode.ForRead)
                            attref.LayerId = check_routines.Layercheck("~-TXT15")
                        End If
                        attref.TextString = (ed.GetString("View Drawing Reference Location? ")).StringResult
                        attref.TextStyleId = tstyle.ObjectId
                        attref.Height = tstyle.TextSize
                        attref.WidthFactor = 0.9
                    End If
                Next

                tr.Commit()
                tr.Dispose()
            End Using
            pko = New PromptKeywordOptions("Mirror Section? ")
            pko.Keywords.Add("Yes")
            pko.Keywords.Add("No")

            pkr = ed.GetKeywords(pko)
            Using tr = db.TransactionManager.StartTransaction
                If pkr.StringResult = "Yes" Then
                    blksecend = tr.GetObject(blksecend.ObjectId, OpenMode.ForWrite)
                    blkbubend = tr.GetObject(blkbubend.ObjectId, OpenMode.ForWrite)
                    blkbubble = tr.GetObject(blkbubble.ObjectId, OpenMode.ForWrite)
                    Dim mirscl As New Scale3d(blkbubble.ScaleFactors.X, blkbubble.ScaleFactors.Y * -1, blkbubble.ScaleFactors.Z)

                    blkbubend.ScaleFactors = mirscl
                    blksecend.ScaleFactors = mirscl
                End If

                tr.Commit()
                tr.Dispose()
            End Using

            Try
                Using tr = db.TransactionManager.StartTransaction
                    blksecend = tr.GetObject(blksecend.ObjectId, OpenMode.ForWrite)
                    If Thisdrawing.GetVariable("Users1") = "bldg" Then blksecend.Erase()
                    tr.Commit()
                    tr.Dispose()
                End Using
            Catch

            End Try

            ClsBlock_insert.BLKSFFACT = Nothing

        Catch ex As Exception

        End Try
    End Sub

    <CommandMethod("Detail")> Public Sub DetailBubble()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction
        Dim blkbubble As BlockReference
        Dim blkscl As New Clsutilities
        Dim objid As ObjectId
        Dim opclo As String
        Dim pko As New PromptKeywordOptions("Open or Closed? ")
        pko.Keywords.Add("Open")
        pko.Keywords.Add("Closed")
        Dim pkr As PromptResult = ed.GetKeywords(pko)

        opclo = LCase(Strings.Left(pkr.StringResult, 1))

        Dim blockinsert As New ClsBlock_insert
        Dim blkname As String
        Using tr
            If LCase(Thisdrawing.GetVariable("UserS1")) = "inds" Then
                ClsBlock_insert.BLKSFFACT = 1
            Else
                ClsBlock_insert.BLKSFFACT = 2 / 3
            End If

            blkname = "hel_sym_2p" & opclo & "-2023"
            If check_routines.Blkchk(blkname) = True Then

            Else
                If blockinsert.Loadblock(blkname & ".dwg") = False Then
                    Exit Sub
                End If
            End If

            objid = blockinsert.InsertBlockWithJig(blkname, False, "Detail Bubble Insertion Point: ", False, blkscl.Getscale, False)
            If objid = Nothing Then Exit Sub
            tr.Commit()
            tr.Dispose()
        End Using
        tr = db.TransactionManager.StartTransaction

        Using tr

            blkbubble = tr.GetObject(objid, OpenMode.ForWrite)
            Clsdetailjig.blkobj = blkbubble
            Clsblock_jig.BasePt = blkbubble.Position
            blkbubble.LayerId = check_routines.Layercheck("~-TSYM3")

            Dim attcol As AttributeCollection = blkbubble.AttributeCollection
            Dim tst As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
            Dim tstyle As TextStyleTableRecord
            For Each attid As ObjectId In attcol
                Dim attref As AttributeReference = tr.GetObject(attid, OpenMode.ForWrite)

                If attref.Tag = "REF" Then
                    If LCase(Thisdrawing.GetVariable("UserS1")) = "inds" Then
                        tstyle = tr.GetObject(tst.Item("HEL40"), OpenMode.ForRead)
                        attref.LayerId = check_routines.Layercheck("~-TXT40")
                    Else
                        tstyle = tr.GetObject(tst.Item("HEL25"), OpenMode.ForRead)
                        attref.LayerId = check_routines.Layercheck("~-TXT25")
                    End If
                    attref.TextString = (ed.GetString("View Reference Number? ")).StringResult
                    attref.TextStyleId = tstyle.ObjectId
                    attref.Height = tstyle.TextSize
                    attref.WidthFactor = 0.9
                Else
                    If LCase(Thisdrawing.GetVariable("UserS1")) = "inds" Then
                        tstyle = tr.GetObject(tst.Item("HEL25"), OpenMode.ForRead)
                        attref.LayerId = check_routines.Layercheck("~-TXT25")
                    Else
                        tstyle = tr.GetObject(tst.Item("HEL15"), OpenMode.ForRead)
                        attref.LayerId = check_routines.Layercheck("~-TXT15")
                    End If
                    attref.TextString = (ed.GetString("View Drawing Reference Location? ")).StringResult
                    attref.TextStyleId = tstyle.ObjectId
                    attref.Height = tstyle.TextSize
                    attref.WidthFactor = 0.9
                End If
            Next

            tr.Commit()
            tr.Dispose()
        End Using
        Dim SelPt As Autodesk.AutoCAD.EditorInput.PromptPointResult
        Dim drawjig As New Clsdetailjig
        SelPt = drawjig.StartJig()


        tr = db.TransactionManager.StartTransaction
        Dim myDB As Database = HostApplicationServices.WorkingDatabase
        Using tr
            Dim PLBOX As Polyline = Nothing
            Dim detailline As Line = Nothing
            Dim pline As Polyline3d
            Dim plinepts As New Point3dCollection
            Dim cntrpt As Point3d = Clsdetailjig.ppr.Value
            Dim detailcircle As Circle = Nothing
            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                Dim myBT As BlockTable = myDB.BlockTableId.GetObject(OpenMode.ForRead)
                Dim myBTR As BlockTableRecord
                If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
                    myBTR = myBT(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)
                Else
                    myBTR = myBT(BlockTableRecord.PaperSpace).GetObject(OpenMode.ForWrite)
                End If
                Try
                    If LCase(Clsdetailjig.clsdjpr.StringResult) = "circle" Or Clsdetailjig.clsdjpr.StringResult = Nothing Then
                        detailline = New Line(cntrpt, SelPt.Value)
                        detailcircle = New Circle(cntrpt, Vector3d.ZAxis, detailline.Length)
                        plinepts.Add(Clsdetailjig.line1.StartPoint)
                        plinepts.Add(Clsdetailjig.line1.EndPoint)
                        plinepts.Add(Clsdetailjig.line3.EndPoint)
                        pline = New Polyline3d(Poly3dType.SimplePoly, plinepts, False)
                        detailcircle.LayerId = check_routines.Layercheck("~-TSYM2")
                        pline.LayerId = check_routines.Layercheck("~-TSYM2")
                        myBTR.AppendEntity(pline)
                        myBTR.AppendEntity(detailcircle)
                    Else
                        PLBOX = Clsdetailjig.pl
                        detailline = New Line(cntrpt, SelPt.Value)
                        plinepts.Add(Clsdetailjig.line1.StartPoint)
                        plinepts.Add(Clsdetailjig.line1.EndPoint)
                        plinepts.Add(Clsdetailjig.line3.EndPoint)
                        pline = New Polyline3d(Poly3dType.SimplePoly, plinepts, False) With {
                        .LayerId = check_routines.Layercheck("~-TSYM2")
                        }
                        PLBOX.LayerId = check_routines.Layercheck("~-TSYM2")
                        PLBOX.ColorIndex = 256
                        myBTR.AppendEntity(pline)
                        myBTR.AppendEntity(PLBOX)

                    End If
                Catch
                    'tr = db.TransactionManager.StartTransaction
                    blkbubble = tr.GetObject(objid, OpenMode.ForWrite)
                    blkbubble.Erase(True)
                    tr.Commit()
                    Exit Sub
                End Try
            End Using
            If LCase(Clsdetailjig.clsdjpr.StringResult) = "circle" Or Clsdetailjig.clsdjpr.StringResult = Nothing Then
                tr.AddNewlyCreatedDBObject(pline, True)
                tr.AddNewlyCreatedDBObject(detailcircle, True)
                tr.Commit()
            Else
                tr.AddNewlyCreatedDBObject(pline, True)
                tr.AddNewlyCreatedDBObject(PLBOX, True)
                'tr.AddNewlyCreatedDBObject(detailcircle, True)
                tr.Commit()

            End If
            tr.Dispose()
        End Using

        Clsdetailjig.myLine.Dispose()
        Clsdetailjig.line1.Dispose()
        Clsdetailjig.line2.Dispose()
        Clsdetailjig.line3.Dispose()
        Clsdetailjig.blkobj.Dispose()
        Clsdetailjig.pl.Dispose()

        ClsBlock_insert.BLKSFFACT = Nothing
nd:
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    <CommandMethod("Title")> Public Sub DrawingTitle()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction '= db.TransactionManager.StartTransaction

        Dim blockinsert As New ClsBlock_insert
        Dim blkname As String

        Dim blkscl As New Clsutilities

        ClsBlock_insert.BLKSFFACT = 1
        blkname = "hel_dmy_title-2023"
        If check_routines.Blkchk(blkname) = True Then

        Else
            If blockinsert.Loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If

        Dim dmyobjid As ObjectId = blockinsert.InsertBlockWithJig(blkname, False, "Drawing Title Insertion Point: ", False, blkscl.Getscale, False)
        'tr.Commit()
        If dmyobjid = Nothing Then Exit Sub
        tr = db.TransactionManager.StartTransaction

        blkname = "hel_sym_title-2023"
        If check_routines.Blkchk(blkname) = True Then

        Else
            If blockinsert.Loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If

        Dim dummytitle As BlockReference = tr.GetObject(dmyobjid, OpenMode.ForWrite)
        Dim objid As ObjectId = blockinsert.InsertBlock(dummytitle.Position, blkname, dummytitle.ScaleFactors.X, dummytitle.ScaleFactors.Y, dummytitle.ScaleFactors.Z, dummytitle.Rotation)
        dummytitle.Erase()

        Dim dwgtitle As BlockReference = tr.GetObject(objid, OpenMode.ForWrite)
        Clsdetailjig.blkobj = dwgtitle
        Clsblock_jig.BasePt = dwgtitle.Position
        dwgtitle.LayerId = check_routines.Layercheck("~-TSYM4")

        Dim attcol As AttributeCollection = dwgtitle.AttributeCollection
        Dim tst As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
        Dim tstyle As TextStyleTableRecord
        Dim extents As Extents3d
        Dim PSO As New PromptStringOptions("") With {
            .AllowSpaces = True
            }

        For Each attid As ObjectId In attcol
            Dim attref As AttributeReference = tr.GetObject(attid, OpenMode.ForWrite)

            If attref.Tag = "TITLE" Then
                If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
                    tstyle = tr.GetObject(tst.Item("HEL40"), OpenMode.ForRead)
                Else
                    tstyle = tr.GetObject(tst.Item("PHEL40"), OpenMode.ForRead)
                End If

                PSO.Message = ("Drawing Title? ")
                attref.TextString = (ed.GetString(PSO).StringResult)
                If attref.TextString = "" Then
                    attref.TextString = InputBox("Please enter a view title")
                End If
                attref.TextStyleId = tstyle.ObjectId
                attref.Height = tstyle.TextSize
                attref.WidthFactor = 0.9
                attref.LayerId = check_routines.Layercheck("~-TXT40")
                Try
                    extents = attref.GeometricExtents
                Catch ex As Exception
                    attref.TextString = "TO BE DETERMINED"
                    extents = attref.GeometricExtents
                End Try

            ElseIf attref.Tag = "SCALE" Then
                Try
                    tstyle = tr.GetObject(tst.Item(utilities.Gettxtstyle), OpenMode.ForRead)
                Catch
                    Dim writetocmd As New Clsutilities
                    writetocmd.WritetoCMD("Herold Text Style not found, using current text style")
                    tstyle = tr.GetObject(tst.Item(Thisdrawing.ActiveTextStyle.Name), OpenMode.ForRead)
                End Try
                PSO.Message = "View Scale? <" & GETSCL() & "> "
                attref.TextString = ed.GetString(PSO).StringResult
                If attref.TextString = "" Then
                    attref.TextString = GETSCL()
                End If
                attref.TextStyleId = tstyle.ObjectId
                Try
                    attref.Height = tstyle.TextSize
                    attref.WidthFactor = 0.9
                Catch

                End Try
                attref.LayerId = check_routines.Layercheck("~-TXT" & Right(utilities.Gettxtstyle, 2))
            ElseIf attref.Tag = "ALT." Then
                Try
                    tstyle = tr.GetObject(tst.Item(utilities.Gettxtstyle), OpenMode.ForRead)
                Catch
                    Dim writetocmd As New Clsutilities
                    writetocmd.WritetoCMD("Herold Text Style not found, using current text style")
                    tstyle = tr.GetObject(tst.Item(Thisdrawing.ActiveTextStyle.Name), OpenMode.ForRead)
                End Try
                PSO.Message = "Optional Alternate View Title? "
                attref.TextString = ed.GetString(PSO).StringResult
                attref.TextStyleId = tstyle.ObjectId
                Try
                    attref.Height = tstyle.TextSize
                    attref.WidthFactor = 0.9
                Catch

                End Try
                attref.LayerId = check_routines.Layercheck("~-TXT15")
            End If
        Next

        Dim defcol As DynamicBlockReferencePropertyCollection = dwgtitle.DynamicBlockReferencePropertyCollection
        For Each dbpid As DynamicBlockReferenceProperty In defcol
            If dbpid.PropertyName = "hel_title_line" Then
                dbpid.Value = extents.MaxPoint.X - extents.MinPoint.X
            End If

        Next


        tr.Commit()

    End Sub

    <CommandMethod("vTitle")> Public Sub ViewTitle()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction
        Dim blkscl As New clsutilities
        Dim blockinsert As New ClsBlock_insert
        Dim blkname As String

        ClsBlock_insert.BLKSFFACT = 1
        Dim opclo As String
        Dim pko As New PromptKeywordOptions("Open or Closed? ")
        pko.Keywords.Add("Open")
        pko.Keywords.Add("Closed")
        Dim pkr As PromptResult = ed.GetKeywords(pko)

        opclo = LCase(Strings.Left(pkr.StringResult, 1))

        blkname = "hel_dmy_2pt-2023"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.Loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If

        Dim dmyobjid As ObjectId = blockinsert.InsertBlockWithJig(blkname, False, "Drawing Title Insertion Point: ", False, blkscl.getscale, False)
        If dmyobjid = Nothing Then Exit Sub
        tr.Commit()
        tr = db.TransactionManager.StartTransaction

        blkname = "hel_sym_2p" & opclo & "t-2023"
        If check_routines.blkchk(blkname) = True Then

        Else
            If blockinsert.Loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If

        Dim dummytitle As BlockReference = tr.GetObject(dmyobjid, OpenMode.ForWrite)
        Dim objid As ObjectId = blockinsert.InsertBlock(dummytitle.Position, blkname, dummytitle.ScaleFactors.X, dummytitle.ScaleFactors.Y, dummytitle.ScaleFactors.Z, dummytitle.Rotation)
        dummytitle.Erase()

        Dim dwgtitle As BlockReference = tr.GetObject(objid, OpenMode.ForWrite)
        Clsdetailjig.blkobj = dwgtitle
        clsblock_jig.BasePt = dwgtitle.Position
        dwgtitle.LayerId = check_routines.layercheck("~-TSYM5")

        Dim attcol As AttributeCollection = dwgtitle.AttributeCollection
        Dim tst As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
        Dim tstyle As TextStyleTableRecord
        Dim extents As Extents3d
        Dim PSO As New PromptStringOptions("") With {
            .AllowSpaces = True
            }
        For Each attid As ObjectId In attcol
            Dim attref As AttributeReference = tr.GetObject(attid, OpenMode.ForWrite)

            If attref.Tag = "TITLE" Then
                If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
                    tstyle = tr.GetObject(tst.Item("HEL40"), OpenMode.ForRead)
                Else
                    tstyle = tr.GetObject(tst.Item("PHEL40"), OpenMode.ForRead)
                End If
                PSO.Message = "Drawing Title? "
                attref.TextString = (ed.GetString(PSO)).StringResult
                If attref.TextString = "" Then
                    attref.TextString = InputBox("Please enter a view title")
                End If
                attref.TextStyleId = tstyle.ObjectId
                attref.Height = tstyle.TextSize
                attref.WidthFactor = 0.9
                attref.LayerId = check_routines.layercheck("~-TXT40")
                Try
                    extents = attref.GeometricExtents
                Catch ex As Exception
                    attref.TextString = "TO BE DETERMINED"
                    extents = attref.GeometricExtents
                End Try
            ElseIf attref.Tag = "SCALE" Then
                Try
                    tstyle = tr.GetObject(tst.Item(utilities.gettxtstyle), OpenMode.ForRead)
                Catch
                    Dim writetocmd As New clsutilities
                    writetocmd.WritetoCMD("Herold Text Style not found, using current text style")
                    tstyle = tr.GetObject(tst.Item(Thisdrawing.ActiveTextStyle.Name), OpenMode.ForRead)
                End Try
                PSO.Message = "View Scale? <" & GETSCL() & "> "
                attref.TextString = ed.GetString(PSO).StringResult
                If attref.TextString = "" Then
                    attref.TextString = GETSCL()
                End If
                attref.TextStyleId = tstyle.ObjectId
                Try
                    attref.Height = tstyle.TextSize
                    attref.WidthFactor = 0.9
                Catch
                End Try
                attref.LayerId = check_routines.layercheck("~-TXT" & Right(utilities.gettxtstyle, 2))
            ElseIf attref.Tag = "ALT-ABOVE" Then
                Try
                    tstyle = tr.GetObject(tst.Item(utilities.gettxtstyle), OpenMode.ForRead)
                Catch
                    Dim writetocmd As New clsutilities
                    writetocmd.WritetoCMD("Herold Text Style not found, using current text style")
                    tstyle = tr.GetObject(tst.Item(Thisdrawing.ActiveTextStyle.Name), OpenMode.ForRead)
                End Try
                PSO.Message = "Optional Alternate View Title Above? "
                attref.TextString = ed.GetString(PSO).StringResult
                attref.TextStyleId = tstyle.ObjectId
                Try
                    attref.Height = tstyle.TextSize
                    attref.WidthFactor = 0.9
                Catch
                End Try
                attref.LayerId = check_routines.layercheck("~-TXT15")
            ElseIf attref.Tag = "ALT-BELOW" Then
                Try
                    tstyle = tr.GetObject(tst.Item(utilities.gettxtstyle), OpenMode.ForRead)
                Catch
                    Dim writetocmd As New clsutilities
                    writetocmd.WritetoCMD("Herold Text Style not found, using current text style")
                    tstyle = tr.GetObject(tst.Item(Thisdrawing.ActiveTextStyle.Name), OpenMode.ForRead)
                End Try
                PSO.Message = "Optional Alternate View Title Below? "
                attref.TextString = ed.GetString(PSO).StringResult
                attref.TextStyleId = tstyle.ObjectId
                Try
                    attref.Height = tstyle.TextSize
                    attref.WidthFactor = 0.9
                Catch
                End Try

                attref.LayerId = check_routines.layercheck("~-TXT15")

            ElseIf attref.Tag = "REF" Then
                Try
                    If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
                        tstyle = tr.GetObject(tst.Item("HEL40"), OpenMode.ForRead)
                    Else
                        tstyle = tr.GetObject(tst.Item("PHEL40"), OpenMode.ForRead)
                    End If
                Catch
                    Dim writetocmd As New clsutilities
                    writetocmd.WritetoCMD("Herold Text Style not found, using current text style")
                    tstyle = tr.GetObject(tst.Item(Thisdrawing.ActiveTextStyle.Name), OpenMode.ForRead)
                End Try
                attref.TextString = ed.GetString("View Reference? ").StringResult
                attref.TextStyleId = tstyle.ObjectId
                Try
                    attref.Height = tstyle.TextSize
                    attref.WidthFactor = 0.9
                Catch
                End Try
                attref.LayerId = check_routines.layercheck("~-TXT40")
            ElseIf attref.Tag = "DWGREF" Then
                Try
                    tstyle = tr.GetObject(tst.Item(utilities.gettxtstyle), OpenMode.ForRead)
                Catch
                    Dim writetocmd As New clsutilities
                    writetocmd.WritetoCMD("Herold Text Style not found, using current text style")
                    tstyle = tr.GetObject(tst.Item(Thisdrawing.ActiveTextStyle.Name), OpenMode.ForRead)
                End Try
                attref.TextString = ed.GetString("View Drawing Reference Location? ").StringResult
                attref.TextStyleId = tstyle.ObjectId
                Try
                    attref.Height = tstyle.TextSize
                    attref.WidthFactor = 0.9
                Catch
                End Try
                attref.LayerId = check_routines.layercheck("~-TXT" & Right(utilities.gettxtstyle, 2))
            End If
        Next

        Dim defcol As DynamicBlockReferencePropertyCollection = dwgtitle.DynamicBlockReferencePropertyCollection
        For Each dbpid As DynamicBlockReferenceProperty In defcol
            If dbpid.PropertyName = "hel_title_line" Then
                dbpid.Value = extents.MaxPoint.X - dwgtitle.Position.X
            End If

        Next


        tr.Commit()
    End Sub
End Class
