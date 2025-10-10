Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Public Class ClsThreePTARROWJIG
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig
    Dim BasePt As Point3d
    Dim myPR As PromptPointResult
    Public Shared myLine As New Line()
    Public Shared ppr As PromptPointResult
    Public Shared blkobj As BlockReference
    Public Shared line1 As New Line
    Public Shared line3 As New Line
    Public Shared STARTPT As Point3d
    Public Shared EXTENTS As Extents3d
    Public Shared txt As DBText
    Public Shared sppr As PromptPointResult
    Private line2 As New Line
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property
    <CommandMethod("THREEPOINT")> Public Sub StartJig()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim laycheck As New ClsCheckroutines
        Dim pkw As New PromptKeywordOptions("Arrow Type: ")
        Dim acleader As New Leader
        Dim GETARR As New ClsLeaders

        pkw.AllowArbitraryInput = False
        pkw.AppendKeywordsToMessage = True
        pkw.Keywords.Add("Arrow")
        pkw.Keywords.Add("BlankDot")
        pkw.Keywords.Add("Dot")
        pkw.Keywords.Add("Integral")

        Dim pkr As PromptResult = ed.GetKeywords(pkw)
        Dim arrtyp As ObjectId = ObjectId.Null
        Dim OSMODE As Integer = Application.GetSystemVariable("OSMODE")

        If LCase(Strings.Left(pkr.StringResult, 1)) = "a" Then
            Const arrowname As String = "."
            arrtyp = GETARR.Getarrowid(arrowname)
            Application.SetSystemVariable("OSMODE", 512)
        ElseIf LCase(Strings.Left(pkr.StringResult, 1)) = "b" Then
            Const arrowname As String = "_DOTBLANK"
            arrtyp = GETARR.Getarrowid(arrowname)
            Application.SetSystemVariable("OSMODE", 4)
        ElseIf LCase(Strings.Left(pkr.StringResult, 1)) = "d" Then
            Const arrowname As String = "_DOT"
            arrtyp = GETARR.Getarrowid(arrowname)
            Application.SetSystemVariable("OSMODE", 55)
        ElseIf LCase(Strings.Left(pkr.StringResult, 1)) = "i" Then
            Const arrowname As String = "_INTEGRAL"
            arrtyp = GETARR.Getarrowid(arrowname)
            Application.SetSystemVariable("OSMODE", 512)
        Else
            Const arrowname As String = "."
            arrtyp = GETARR.Getarrowid(arrowname)
            Application.SetSystemVariable("OSMODE", 512)
        End If
        Using tr As Transaction = db.TransactionManager.StartTransaction
            Try
                line1.ColorIndex = 1
                line2.ColorIndex = 1
            Catch ex As Exception

            End Try
            Try

                Dim pso As New PromptSelectionOptions With {
                .MessageForAdding = "Select Text: ",
                .SingleOnly = True,
                .RejectObjectsOnLockedLayers = False
                }
                Dim selectfilter(0) As TypedValue
                selectfilter.SetValue(New TypedValue(DxfCode.Start, "TEXT,MTEXT"), 0)
                Dim psf As New SelectionFilter(selectfilter)


                Dim psr As PromptSelectionResult = ed.GetSelection(pso, psf)

                Dim so As SelectedObject = psr.Value(0)
                Dim entt As Entity
                Dim objs As New DBObjectCollection


                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord

                If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
                    btr = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Else
                    btr = tr.GetObject(bt(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
                End If
                If psr.Value.Count = 1 Then
                    Dim ent As Entity = tr.GetObject(so.ObjectId, OpenMode.ForRead)
                    If TypeOf ent Is MText Then
                        Dim mt As MText = ent

                        Dim cb As New MTextFragmentCallback(Function(frag As MTextFragment, obj As Object)
                                                                'ed.WriteMessage("\nMtext Fragment {0}", frag.Text)
                                                                Return MTextFragmentCallbackStatus.Continue
                                                            End Function)
                        mt.ExplodeFragments(cb)

                        mt.Explode(objs)
                        entt = objs(0)
                    Else
                        entt = ent
                    End If

                    acleader.SetDatabaseDefaults()

                    txt = entt
                    EXTENTS = txt.GeometricExtents
                End If

                Dim ppo As New PromptPointOptions("Pick Leader Start Point: ")

                sppr = ed.GetPoint(ppo)


                BasePt = New Point3d((EXTENTS.MaxPoint.X + EXTENTS.MinPoint.X) / 2,
                                      (EXTENTS.MaxPoint.Y + EXTENTS.MinPoint.Y) / 2,
                                      (EXTENTS.MaxPoint.Z + EXTENTS.MinPoint.Z) / 2)

                'line1.ColorIndex = 1
                'line2.ColorIndex = 1

                myPR = ed.Drag(Me)
                Do
                    Select Case myPR.Status
                        Case PromptStatus.OK
                            acleader.AppendVertex(line2.EndPoint)
                            acleader.AppendVertex(line1.EndPoint)
                            acleader.AppendVertex(line1.StartPoint)



                            btr.AppendEntity(acleader)
                            tr.AddNewlyCreatedDBObject(acleader, True)
                            acleader.Dimldrblk = (arrtyp)


                            acleader.LayerId = laycheck.Layercheck("~-TDIM")

                            Exit Do
                    End Select
                Loop While myPR.Status <> PromptStatus.Cancel
                tr.Commit()
            Catch
                tr.Abort()


            Finally




            End Try
            tr.Dispose()
        End Using
        Application.SetSystemVariable("osmode", OSMODE)
    End Sub
    Protected Overrides Function Sampler(ByVal prompts As _
    Autodesk.AutoCAD.EditorInput.JigPrompts) As _
    Autodesk.AutoCAD.EditorInput.SamplerStatus
        myPR = prompts.AcquirePoint("Select leader turn point:")


        If myPR.Value.IsEqualTo(BasePt) Then
            Return SamplerStatus.NoChange
        Else

            If myPR.Value.X < BasePt.X And myPR.Value.X < (EXTENTS.MinPoint.X - (txt.Height / 2)) Then
                line1.StartPoint = New Point3d(EXTENTS.MinPoint(0) - (txt.Height / 2),
                                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
                                                     EXTENTS.MinPoint(2))
                line1.EndPoint = New Point3d(myPR.Value.X,
                                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
                                                     EXTENTS.MinPoint(2))

                line2.StartPoint = line1.EndPoint
                line2.EndPoint = sppr.Value
            ElseIf myPR.Value.X > BasePt.X And myPR.Value.X > (EXTENTS.MaxPoint.X + (txt.Height / 2)) Then
                line1.StartPoint = New Point3d(EXTENTS.MaxPoint(0) + (txt.Height / 2),
                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
                                     EXTENTS.MinPoint(2))
                line1.EndPoint = New Point3d(myPR.Value.X,
                                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
                                                     EXTENTS.MinPoint(2))

                line2.StartPoint = line1.EndPoint
                line2.EndPoint = sppr.Value

            ElseIf myPR.Value.X < BasePt.X And myPR.Value.X > (EXTENTS.MinPoint.X - (txt.Height / 2)) Then
                line1.StartPoint = New Point3d(EXTENTS.MinPoint(0) - (txt.Height / 2),
                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
                                     EXTENTS.MinPoint(2))
                line1.EndPoint = New Point3d(EXTENTS.MinPoint(0) - (txt.Height),
                                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
                                                     EXTENTS.MinPoint(2))

                line2.StartPoint = line1.EndPoint
                line2.EndPoint = sppr.Value
            ElseIf myPR.Value.X > BasePt.X And myPR.Value.X < (EXTENTS.MaxPoint.X + (txt.Height / 2)) Then
                line1.StartPoint = New Point3d(EXTENTS.MaxPoint(0) + (txt.Height / 2),
                                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
                                                     EXTENTS.MinPoint(2))
                line1.EndPoint = New Point3d(EXTENTS.MaxPoint(0) + (txt.Height),
                                                     EXTENTS.MinPoint(1) + (txt.Height / 2),
                                                     EXTENTS.MinPoint(2))

                line2.StartPoint = line1.EndPoint
                line2.EndPoint = sppr.Value
            End If


            Return SamplerStatus.OK
        End If
    End Function
    Protected Overrides Function WorldDraw(ByVal draw As _
    Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean

        draw.Geometry.Draw(line1)
        draw.Geometry.Draw(line2)

        Return True
    End Function
    Protected Overrides Sub Finalize()
        myLine.Dispose()
        line1.Dispose()
        Try
            If blkobj.IsDisposed = False Then
            Else

                blkobj.Dispose()
            End If
        Catch

        End Try

        line3.Dispose()
        ppr = Nothing
        line2.Dispose()
        MyBase.Finalize()
    End Sub


End Class
