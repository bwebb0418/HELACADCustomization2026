Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry



Public Class ClsLeaders

    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property
    Function Getarrowid(ByVal newarrname As String)
        Dim arrobjid As ObjectId '= ObjectId.Null
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim oldarrname As String = Application.GetSystemVariable("DIMBLK")
        Dim tr As Transaction = db.TransactionManager.StartTransaction
        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)


        Application.SetSystemVariable("DIMBLK", newarrname)

        If oldarrname.Length <> 0 Then
            Application.SetSystemVariable("DIMBLK", oldarrname)
        Else
            Application.SetSystemVariable("DIMBLK", ".")
        End If


        If newarrname = "." Then
            arrobjid = ObjectId.Null
        Else
            arrobjid = bt(newarrname)
        End If
        tr.Commit()
        Return arrobjid
    End Function

    'Public Sub TwoclickArrows()
    '    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim db As Database = doc.Database
    '    Dim ed As Editor = doc.Editor
    '    Dim laycheck As New ClsCheckroutines
    '    Dim pkw As New PromptKeywordOptions("Arrow Type: ") With {
    '        .AllowArbitraryInput = False,
    '        .AppendKeywordsToMessage = True
    '        }
    '    pkw.Keywords.Add("Arrow")
    '    pkw.Keywords.Add("BlankDot")
    '    pkw.Keywords.Add("Dot")
    '    pkw.Keywords.Add("Integral")

    '    Dim pkr As PromptResult = ed.GetKeywords(pkw)
    '    Dim arrtyp As ObjectId = ObjectId.Null
    '    Dim OSMODE As Integer = Application.GetSystemVariable("OSMODE")

    '    If LCase(Strings.Left(pkr.StringResult, 1)) = "a" Then
    '        Const arrowname As String = "."
    '        arrtyp = getarrowid(arrowname)
    '        Application.SetSystemVariable("OSMODE", 512)
    '    ElseIf LCase(Strings.Left(pkr.StringResult, 1)) = "b" Then
    '        Const arrowname As String = "_DOTBLANK"
    '        arrtyp = getarrowid(arrowname)
    '        Application.SetSystemVariable("OSMODE", 4)
    '    ElseIf LCase(Strings.Left(pkr.StringResult, 1)) = "d" Then
    '        Const arrowname As String = "_DOT"
    '        arrtyp = getarrowid(arrowname)
    '        Application.SetSystemVariable("OSMODE", 55)
    '    ElseIf LCase(Strings.Left(pkr.StringResult, 1)) = "i" Then
    '        Const arrowname As String = "_INTEGRAL"
    '        arrtyp = getarrowid(arrowname)
    '        Application.SetSystemVariable("OSMODE", 512)
    '    Else
    '        Const arrowname As String = "."
    '        arrtyp = getarrowid(arrowname)
    '        Application.SetSystemVariable("OSMODE", 512)
    '    End If
    '    Using tr As Transaction = db.TransactionManager.StartTransaction

    '        Try

    '            Dim pso As New PromptSelectionOptions With {
    '            .MessageForAdding = "Select Text: ",
    '            .SingleOnly = True,
    '            .RejectObjectsOnLockedLayers = False
    '            }
    '            Dim selectfilter(0) As TypedValue
    '            selectfilter.SetValue(New TypedValue(DxfCode.Start, "TEXT,MTEXT"), 0)
    '            Dim psf As New SelectionFilter(selectfilter)


    '            Dim psr As PromptSelectionResult = ed.GetSelection(pso, psf)

    '            Dim so As SelectedObject = psr.Value(0)
    '            Dim entt As Entity
    '            Dim objs As New DBObjectCollection


    '            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
    '            Dim btr As BlockTableRecord

    '            If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
    '                btr = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '            Else
    '                btr = tr.GetObject(bt(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
    '            End If
    '            If psr.Value.Count = 1 Then
    '                Dim ent As Entity = tr.GetObject(so.ObjectId, OpenMode.ForRead)
    '                If TypeOf ent Is MText Then
    '                    Dim mt As MText = ent

    '                    Dim cb As New MTextFragmentCallback(Function(frag As MTextFragment, obj As Object)
    '                                                            'ed.WriteMessage("\nMtext Fragment {0}", frag.Text)
    '                                                            Return MTextFragmentCallbackStatus.Continue
    '                                                        End Function)
    '                    mt.ExplodeFragments(cb)

    '                    mt.Explode(objs)
    '                    entt = objs(0)
    '                Else
    '                    entt = ent
    '                End If

    '                Dim acleader As New Leader

    '                acleader.SetDatabaseDefaults()

    '                Dim ppo As New PromptPointOptions("Pick Leader Start Point: ") With {
    '                .AllowNone = False
    '                }

    '                Dim ppr As PromptPointResult = ed.GetPoint(ppo)
    '                acleader.AppendVertex(ppr.Value)

    '                Dim txt As DBText = entt
    '                Dim EXTENTS As Extents3d = txt.GeometricExtents

    '                Dim STARTPT As Point3d
    '                Dim turnpt As Point3d
    '                If ppr.Value(0) < (EXTENTS.MaxPoint(0) + EXTENTS.MinPoint(0)) / 2 Then
    '                    STARTPT = New Point3d(EXTENTS.MinPoint(0) - (txt.Height / 2),
    '                                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
    '                                                     EXTENTS.MinPoint(2))
    '                    turnpt = New Point3d(EXTENTS.MinPoint(0) - (txt.Height * 2.5),
    '                                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
    '                                                     EXTENTS.MinPoint(2))

    '                Else
    '                    STARTPT = New Point3d(EXTENTS.MaxPoint(0) + (txt.Height / 2),
    '                                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
    '                                                     EXTENTS.MinPoint(2))
    '                    turnpt = New Point3d(EXTENTS.MaxPoint(0) + (txt.Height * 2.5),
    '                                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
    '                                                     EXTENTS.MinPoint(2))
    '                End If

    '                acleader.AppendVertex(turnpt)
    '                acleader.AppendVertex(STARTPT)



    '                btr.AppendEntity(acleader)
    '                tr.AddNewlyCreatedDBObject(acleader, True)
    '                acleader.Dimldrblk = (arrtyp)


    '                acleader.LayerId = laycheck.Layercheck("~-TDIM")
    '            End If
    '            tr.Commit()
    '        Catch ex As Exception
    '            ed.WriteMessage(ex.Message)
    '            tr.Abort()

    '            Application.SetSystemVariable("osmode", OSMODE)
    '        Finally



    '            Application.SetSystemVariable("osmode", OSMODE)
    '        End Try
    '        tr.Dispose()
    '    End Using

    'End Sub

    'Public Sub ThreeClickArrow()
    '    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim db As Database = doc.Database
    '    Dim ed As Editor = doc.Editor
    '    Dim laycheck As New ClsCheckroutines

    '    Dim pkw As New PromptKeywordOptions("Arrow Type: ") With {
    '    .AllowArbitraryInput = False,
    '    .AppendKeywordsToMessage = True
    '    }
    '    pkw.Keywords.Add("Arrow")
    '    pkw.Keywords.Add("BlankDot")
    '    pkw.Keywords.Add("Dot")
    '    pkw.Keywords.Add("Integral")

    '    Dim pkr As PromptResult = ed.GetKeywords(pkw)
    '    Dim arrtyp As ObjectId = ObjectId.Null
    '    Dim OSMODE As Integer = Application.GetSystemVariable("osmode")
    '    If LCase(Strings.Left(pkr.StringResult, 1)) = "a" Then
    '        Const arrowname As String = "."
    '        arrtyp = getarrowid(arrowname)
    '        Application.SetSystemVariable("OSMODE", 512)
    '    ElseIf LCase(Strings.Left(pkr.StringResult, 1)) = "b" Then
    '        Const arrowname As String = "_DOTBLANK"
    '        arrtyp = getarrowid(arrowname)
    '        Application.SetSystemVariable("OSMODE", 4)
    '    ElseIf LCase(Strings.Left(pkr.StringResult, 1)) = "d" Then
    '        Const arrowname As String = "_DOT"
    '        arrtyp = getarrowid(arrowname)
    '        Application.SetSystemVariable("OSMODE", 55)
    '    ElseIf LCase(Strings.Left(pkr.StringResult, 1)) = "i" Then
    '        Const arrowname As String = "_INTEGRAL"
    '        arrtyp = getarrowid(arrowname)
    '        Application.SetSystemVariable("OSMODE", 512)
    '    Else
    '        Const arrowname As String = "."
    '        arrtyp = getarrowid(arrowname)
    '        Application.SetSystemVariable("OSMODE", 512)
    '    End If
    '    Using tr As Transaction = db.TransactionManager.StartTransaction
    '        Try

    '            Dim pso As New PromptSelectionOptions With {
    '            .MessageForAdding = "Select Text: ",
    '            .SingleOnly = True,
    '            .RejectObjectsOnLockedLayers = False
    '            }
    '            Dim selectfilter(0) As TypedValue
    '            selectfilter.SetValue(New TypedValue(DxfCode.Start, "TEXT,MTEXT"), 0)
    '            Dim psf As New SelectionFilter(selectfilter)


    '            Dim psr As PromptSelectionResult = ed.GetSelection(pso, psf)

    '            Dim so As SelectedObject = psr.Value(0)
    '            Dim objs As New DBObjectCollection
    '            Dim BT As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
    '            Dim btr As BlockTableRecord
    '            Dim entt As Entity


    '            If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
    '                btr = tr.GetObject(BT(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '            Else
    '                btr = tr.GetObject(BT(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
    '            End If

    '            If psr.Value.Count = 1 Then
    '                Dim ent As Entity = tr.GetObject(so.ObjectId, OpenMode.ForRead)
    '                If TypeOf ent Is MText Then
    '                    Dim mt As MText = ent
    '                    Dim cb As New MTextFragmentCallback(Function(frag As MTextFragment, obj As Object)
    '                                                            'ed.WriteMessage("\nMtext Fragment {0}", frag.Text)
    '                                                            Return MTextFragmentCallbackStatus.Continue
    '                                                        End Function)
    '                    mt.ExplodeFragments(cb)
    '                    mt.Explode(objs)
    '                    entt = objs(0)
    '                Else
    '                    entt = ent
    '                End If

    '                'Dim extents As Extents3d = entt.GeometricExtents
    '                Dim txt As DBText = entt
    '                Dim extents As Extents3d = txt.GeometricExtents
    '                Dim txtpt As New Point3d((extents.MaxPoint(0) + extents.MinPoint(0)) / 2,
    '                                                     extents.MinPoint(1) + (txt.Height / 2),
    '                                                     extents.MinPoint(2))

    '                Dim acleader As New Leader

    '                acleader.SetDatabaseDefaults()

    '                Dim ppj As New PromptPointOptions("Pick Leader Turn Point: ") With {
    '                .BasePoint = txtpt,
    '                .UseBasePoint = True,
    '                .UseDashedLine = True
    '                }
    '                Dim pprj As PromptPointResult = ed.GetPoint(ppj)
    '                Dim turnpt As New Point3d(pprj.Value(0),
    '                                                     extents.MinPoint(1) + (txt.Height / 2),
    '                                                     extents.MinPoint(2))


    '                Dim ppo As New PromptPointOptions("Pick Leader Start Point: ") With {
    '                .BasePoint = turnpt,
    '                .UseDashedLine = True,
    '                .UseBasePoint = True
    '                }
    '                Dim ppr As PromptPointResult = ed.GetPoint(ppo)

    '                acleader.AppendVertex(ppr.Value)


    '                '' ''Dim turnpt As Point3d = New Point3d(pprj.Value(0), _
    '                '' ''                                     extents.MinPoint(1) + (txt.Height / 2), _
    '                '' ''                                     extents.MinPoint(2))

    '                Dim startpt As Point3d

    '                If pprj.Value(0) < (extents.MaxPoint(0) + extents.MinPoint(0)) / 2 Then
    '                    startpt = New Point3d(extents.MinPoint(0) - (txt.Height / 2),
    '                                                     extents.MinPoint(1) + (txt.Height / 2),
    '                                                     extents.MinPoint(2))
    '                Else
    '                    startpt = New Point3d(extents.MaxPoint(0) + (txt.Height / 2),
    '                                                     extents.MinPoint(1) + (txt.Height / 2),
    '                                                     extents.MinPoint(2))
    '                End If


    '                acleader.AppendVertex(turnpt)
    '                acleader.AppendVertex(startpt)



    '                btr.AppendEntity(acleader)
    '                tr.AddNewlyCreatedDBObject(acleader, True)
    '                'acleader.Annotation = ent.ObjectId
    '                acleader.Dimldrblk = (arrtyp)

    '                acleader.LayerId = laycheck.Layercheck("~-TDIM")

    '            End If
    '            tr.Commit()
    '        Catch ex As Exception
    '            ed.WriteMessage(ex.Message)
    '            tr.Abort()

    '        Finally



    '            Application.SetSystemVariable("osmode", OSMODE)
    '        End Try
    '        tr.Dispose()
    '    End Using

    'End Sub


End Class
