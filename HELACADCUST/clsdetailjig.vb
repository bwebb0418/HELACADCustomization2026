Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Public Class Clsdetailjig
    Inherits Autodesk.AutoCAD.EditorInput.DrawJig
    Dim BasePt As Point3d
    Dim myPR As PromptPointResult
    Public Shared myLine As Line
    Public Shared ppr As PromptPointResult
    Public Shared blkobj As BlockReference
    Public Shared line1 As Line
    Public Shared line3 As Line
    'Public Shared pointcol As New Point3dCollection
    Public Shared line2 As Line
    Public Shared clsdjpr As PromptResult
    Public Shared pl As Polyline
    Public Shared radius As Double


    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property
    Function StartJig() As PromptPointResult
        'Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim pdo As New PromptDoubleOptions("Input new radius for box corners: ")

        Dim pdr As PromptDoubleResult
        Dim pko As New PromptKeywordOptions("Circle Detail, or Box Detail? ")
        pko.Keywords.Add("Circle")
        pko.Keywords.Add("Box")
        pko.AppendKeywordsToMessage = True
        pko.AllowNone = True

        clsdjpr = ed.GetKeywords(pko)
        Try
            radius = 5 * blkobj.ScaleFactors.X

        Catch
            Dim blkscl As New Clsutilities
            radius = blkscl.Getscale * 5
        End Try
        Dim ppo As PromptPointOptions
radiuschange:
        If LCase(Strings.Left(clsdjpr.StringResult, 1)) = "c" Or clsdjpr.StringResult = Nothing Then
            ppo = New PromptPointOptions("Pick Detail Center: ")
        Else
            ppo = New PromptPointOptions("Pick closest corner to detail bubble: ")
            ppo.Keywords.Add("Radius")
            ppo.AppendKeywordsToMessage = True
        End If
        ppo.AllowNone = False
        ppr = ed.GetPoint(ppo)

        If ppr.StringResult <> Nothing Then
            If ppr.StringResult.Contains("R") Or ppr.StringResult.Contains("r") Then
                pdo.DefaultValue = radius
                pdr = ed.GetDouble(pdo)
                'If pdr.Value = "" Then
                '    radius = 8
                'Else
                radius = pdr.Value

                'End If

                GoTo radiuschange
            End If
        End If

        If ppr.Status <> PromptStatus.OK Then
            myPR = Nothing
            Return myPR
            Exit Function
        End If

        myLine = New Line
        line1 = New Line
        line2 = New Line
        line3 = New Line
        pl = New Polyline

        myLine.StartPoint = ppr.Value
        myLine.ColorIndex = 1
        line1.ColorIndex = 2
        line3.ColorIndex = 2
        pl.ColorIndex = 2
        ' detailcirc.ColorIndex = 2
        myPR = ed.Drag(Me)
        Do
            Select Case myPR.Status
                Case PromptStatus.OK
                    Return myPR
                    Exit Do
            End Select
        Loop While myPR.Status <> PromptStatus.Cancel
        Return myPR
        ' End Using
    End Function
    Protected Overrides Function Sampler(ByVal prompts As Autodesk.AutoCAD.EditorInput.JigPrompts) As _
    Autodesk.AutoCAD.EditorInput.SamplerStatus
        Dim line1endpoint As Point3d
        If LCase(clsdjpr.StringResult) = "circle" Or clsdjpr.StringResult = Nothing Then
            myPR = prompts.AcquirePoint("Pick detail circle diameter:")
            If myPR.Value.IsEqualTo(BasePt) Then
                Return SamplerStatus.NoChange
            Else
                BasePt = myPR.Value
                myLine.EndPoint = BasePt

                If myPR.Value.X < blkobj.Position.X Then

                    line1.StartPoint = blkobj.Position
                    line1.EndPoint = blkobj.Position
                    line1endpoint = New Point3d(blkobj.Position.X - (10 * blkobj.ScaleFactors.X), line1.StartPoint.Y, line1.StartPoint.Z)
                    line1.EndPoint = line1endpoint
                    line2.StartPoint = line1endpoint
                    line2.EndPoint = ppr.Value
                Else

                    line1.StartPoint = blkobj.Position
                    line1.EndPoint = blkobj.Position
                    line1endpoint = New Point3d(blkobj.Position.X + (10 * blkobj.ScaleFactors.X), line1.StartPoint.Y, line1.StartPoint.Z)
                    line1.EndPoint = line1endpoint
                    line2.StartPoint = line1endpoint
                    line2.EndPoint = ppr.Value
                End If

                Dim line3endpoint As New Point3d(line2.StartPoint.X + line2.Length - myLine.Length, line2.StartPoint.Y, line2.StartPoint.Z)
                line3.StartPoint = line2.StartPoint
                line3.EndPoint = line3endpoint
                line3.TransformBy(Matrix3d.Rotation(line2.Angle, Vector3d.ZAxis, line3.StartPoint))

                Return SamplerStatus.OK

            End If
        Else
            myPR = prompts.AcquirePoint("Pick opposite corner of detail box:")
            If myPR.Value.IsEqualTo(BasePt) Then
                Return SamplerStatus.NoChange
            Else
                BasePt = myPR.Value

                Dim radii As Double = (radius)

                If myPR.Value.X < blkobj.Position.X Then
                    line1.StartPoint = blkobj.Position
                    line1.EndPoint = blkobj.Position
                    line1endpoint = New Point3d(blkobj.Position.X - (10 * blkobj.ScaleFactors.X), line1.StartPoint.Y, line1.StartPoint.Z)
                    line1.EndPoint = line1endpoint
                    line2.StartPoint = line1endpoint
                    line2.EndPoint = ppr.Value
                Else
                    line1.StartPoint = blkobj.Position
                    line1.EndPoint = blkobj.Position
                    line1endpoint = New Point3d(blkobj.Position.X + (10 * blkobj.ScaleFactors.X), line1.StartPoint.Y, line1.StartPoint.Z)
                    line1.EndPoint = line1endpoint
                    line2.StartPoint = line1endpoint
                    line2.EndPoint = ppr.Value
                End If

                Dim startpt As New Point2d(line2.EndPoint.X, line2.EndPoint.Y)
                Dim endpt As New Point2d(BasePt.X, BasePt.Y)
                pl.Reset(False, 0)
                pl.Reset(False, 1)
                pl.Reset(False, 2)
                pl.Reset(False, 3)
                pl.Reset(False, 4)
                pl.Reset(False, 5)
                pl.Reset(False, 6)
                pl.Reset(False, 7)
                pl.Reset(False, 8)




                If line2.EndPoint.X > line1.EndPoint.X And line2.EndPoint.Y > line1.EndPoint.Y Then
                    If endpt.X - startpt.X < radii * 2 Or endpt.Y - startpt.Y < radii * 2 Then
                        radii = 1 * blkobj.ScaleFactors.X
                    Else
                        radii = radius
                    End If

                    If myPR.Value.X > ppr.Value.X And myPR.Value.Y > ppr.Value.Y Then
                        line2.EndPoint = New Point3d(ppr.Value.X + radii, ppr.Value.Y + radii, ppr.Value.Z)
                        pl.AddVertexAt(0, New Point2d(startpt.X + radii, startpt.Y), 0, 0, 0)
                        pl.AddVertexAt(1, New Point2d(endpt.X - radii, startpt.Y), 0.4142, 0, 0)

                        pl.AddVertexAt(2, New Point2d(endpt.X, startpt.Y + radii), 0, 0, 0)
                        pl.AddVertexAt(3, New Point2d(endpt.X, endpt.Y - radii), 0.4142, 0, 0)

                        pl.AddVertexAt(4, New Point2d(endpt.X - radii, endpt.Y), 0, 0, 0)
                        pl.AddVertexAt(5, New Point2d(startpt.X + radii, endpt.Y), 0.4142, 0, 0)

                        pl.AddVertexAt(6, New Point2d(startpt.X, endpt.Y - radii), 0, 0, 0)
                        pl.AddVertexAt(7, New Point2d(startpt.X, startpt.Y + radii), 0.4142, 0, 0)

                        pl.AddVertexAt(8, New Point2d(startpt.X + radii, startpt.Y), 0, 0, 0)
                    End If
                ElseIf line2.EndPoint.X > line1.EndPoint.X And line2.EndPoint.Y < line1.EndPoint.Y Then
                    If endpt.X - startpt.X < radii * 2 Or startpt.Y - endpt.Y < radii * 2 Then
                        radii = 1 * blkobj.ScaleFactors.X
                    Else
                        radii = radius
                    End If

                    If myPR.Value.X > ppr.Value.X And myPR.Value.Y < ppr.Value.Y Then
                        line2.EndPoint = New Point3d(ppr.Value.X + radii, ppr.Value.Y - radii, ppr.Value.Z)
                        pl.AddVertexAt(0, New Point2d(startpt.X + radii, startpt.Y), 0, 0, 0)
                        pl.AddVertexAt(1, New Point2d(endpt.X - radii, startpt.Y), -0.4142, 0, 0)

                        pl.AddVertexAt(2, New Point2d(endpt.X, startpt.Y - radii), 0, 0, 0)
                        pl.AddVertexAt(3, New Point2d(endpt.X, endpt.Y + radii), -0.4142, 0, 0)

                        pl.AddVertexAt(4, New Point2d(endpt.X - radii, endpt.Y), 0, 0, 0)
                        pl.AddVertexAt(5, New Point2d(startpt.X + radii, endpt.Y), -0.4142, 0, 0)

                        pl.AddVertexAt(6, New Point2d(startpt.X, endpt.Y + radii), 0, 0, 0)
                        pl.AddVertexAt(7, New Point2d(startpt.X, startpt.Y - radii), -0.4142, 0, 0)

                        pl.AddVertexAt(8, New Point2d(startpt.X + radii, startpt.Y), 0, 0, 0)
                    End If
                ElseIf line2.EndPoint.X < line1.EndPoint.X And line2.EndPoint.Y > line1.EndPoint.Y Then
                    If startpt.X - endpt.X < radii * 2 Or endpt.Y - startpt.Y < radii * 2 Then
                        radii = 1 * blkobj.ScaleFactors.X
                    Else
                        radii = radius
                    End If

                    If myPR.Value.X < ppr.Value.X And myPR.Value.Y > ppr.Value.Y Then
                        line2.EndPoint = New Point3d(ppr.Value.X - radii, ppr.Value.Y + radii, ppr.Value.Z)
                        pl.AddVertexAt(0, New Point2d(startpt.X - radii, startpt.Y), 0, 0, 0)
                        pl.AddVertexAt(1, New Point2d(endpt.X + radii, startpt.Y), -0.4142, 0, 0)

                        pl.AddVertexAt(2, New Point2d(endpt.X, startpt.Y + radii), 0, 0, 0)
                        pl.AddVertexAt(3, New Point2d(endpt.X, endpt.Y - radii), -0.4142, 0, 0)

                        pl.AddVertexAt(4, New Point2d(endpt.X + radii, endpt.Y), 0, 0, 0)
                        pl.AddVertexAt(5, New Point2d(startpt.X - radii, endpt.Y), -0.4142, 0, 0)

                        pl.AddVertexAt(6, New Point2d(startpt.X, endpt.Y - radii), 0, 0, 0)
                        pl.AddVertexAt(7, New Point2d(startpt.X, startpt.Y + radii), -0.4142, 0, 0)

                        pl.AddVertexAt(8, New Point2d(startpt.X - radii, startpt.Y), 0, 0, 0)
                    End If
                ElseIf line2.EndPoint.X < line1.EndPoint.X And line2.EndPoint.Y < line1.EndPoint.Y Then
                    If startpt.X - endpt.X < radii * 2 Or startpt.Y - endpt.Y < radii * 2 Then
                        radii = 1 * blkobj.ScaleFactors.X
                    Else
                        radii = radius
                    End If
                    If myPR.Value.X < ppr.Value.X And myPR.Value.Y < ppr.Value.Y Then
                        line2.EndPoint = New Point3d(ppr.Value.X - radii, ppr.Value.Y - radii, ppr.Value.Z)
                        pl.AddVertexAt(0, New Point2d(startpt.X - radii, startpt.Y), 0, 0, 0)
                        pl.AddVertexAt(1, New Point2d(endpt.X + radii, startpt.Y), 0.4142, 0, 0)

                        pl.AddVertexAt(2, New Point2d(endpt.X, startpt.Y - radii), 0, 0, 0)
                        pl.AddVertexAt(3, New Point2d(endpt.X, endpt.Y + radii), 0.4142, 0, 0)

                        pl.AddVertexAt(4, New Point2d(endpt.X + radii, endpt.Y), 0, 0, 0)
                        pl.AddVertexAt(5, New Point2d(startpt.X - radii, endpt.Y), 0.4142, 0, 0)

                        pl.AddVertexAt(6, New Point2d(startpt.X, endpt.Y + radii), 0, 0, 0)
                        pl.AddVertexAt(7, New Point2d(startpt.X, startpt.Y - radii), 0.4142, 0, 0)

                        pl.AddVertexAt(8, New Point2d(startpt.X - radii, startpt.Y), 0, 0, 0)
                    End If
                End If

                Dim line3endpoint As New Point3d(line2.StartPoint.X + line2.Length - (radii), line2.StartPoint.Y, line2.StartPoint.Z)
                line3.StartPoint = line2.StartPoint
                line3.EndPoint = line3endpoint
                line3.TransformBy(Matrix3d.Rotation(line2.Angle, Vector3d.ZAxis, line3.StartPoint))


                Return SamplerStatus.OK

            End If
        End If

    End Function
    Protected Overrides Function WorldDraw(ByVal draw As _
    Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean
        ' draw.Geometry.WorldLine(New Point3d(0, 0, 0), BasePt)
        ' draw.Geometry.Circle(BasePt, 1.5, Vector3d.ZAxis)
        If LCase(clsdjpr.StringResult) = "circle" Or clsdjpr.StringResult = Nothing Then
            draw.Geometry.Draw(line1)
            draw.Geometry.Draw(line3)
            draw.Geometry.Draw(New Circle(ppr.Value, Vector3d.ZAxis, myLine.Length))
        Else
            draw.Geometry.Draw(line1)
            draw.Geometry.Draw(line3)

            'draw.Geometry.Polygon(pointcol)
            draw.Geometry.Polyline(pl, 0, pl.NumberOfVertices - 1)
            'draw.Geometry.Polyline(pointcol, Vector3d.ZAxis, System.IntPtr.Add(0, 0))
            'draw.Geometry.Draw(New Circle(ppr.Value, Vector3d.ZAxis, myLine.Length))
        End If
        Return True
    End Function
    Protected Overrides Sub Finalize()
        myLine.Dispose()
        line1.Dispose()
        blkobj.Dispose()
        line3.Dispose()
        ppr = Nothing
        line2.Dispose()
        MyBase.Finalize()
    End Sub


    '<CommandMethod("plinetest")> Public Sub testpline()

    '    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim db As Database = doc.Database
    '    Dim tr As Transaction = db.TransactionManager.StartTransaction
    '    Dim ed As Editor = doc.Editor
    '    Dim startpt As PromptPointResult = ed.GetPoint("Get Start Point")
    '    Dim endpt As PromptPointResult = ed.GetPoint("Get End Point")
    '    Dim pointcol As New Point2dCollection



    '    pointcol.Add(New Point2d(startpt.Value.X, startpt.Value.Y))
    '    pointcol.Add(New Point2d(endpt.Value.X, startpt.Value.Y))
    '    pointcol.Add(New Point2d(endpt.Value.X, endpt.Value.Y))
    '    pointcol.Add(New Point2d(startpt.Value.X, endpt.Value.Y))

    '    'pline = New Polyline3d(Poly3dType.SimplePoly, pointcol, True)
    '    Dim pl As New Polyline
    '    pl.AddVertexAt(0, New Point2d(startpt.Value.X + 30, startpt.Value.Y), 0, 0, 0)
    '    pl.AddVertexAt(1, New Point2d(endpt.Value.X - 30, startpt.Value.Y), 0.4142, 0, 0)
    '    pl.AddVertexAt(2, New Point2d(endpt.Value.X, startpt.Value.Y + 30), 0, 0, 0)
    '    pl.AddVertexAt(3, New Point2d(endpt.Value.X, endpt.Value.Y - 30), 0.4142, 0, 0)
    '    pl.AddVertexAt(4, New Point2d(endpt.Value.X - 30, endpt.Value.Y), 0, 0, 0)
    '    pl.AddVertexAt(5, New Point2d(startpt.Value.X + 30, endpt.Value.Y), 0.4142, 0, 0)
    '    pl.AddVertexAt(6, New Point2d(startpt.Value.X, endpt.Value.Y - 30), 0, 0, 0)
    '    pl.AddVertexAt(7, New Point2d(startpt.Value.X, startpt.Value.Y + 30), 0.4142, 0, 0)
    '    pl.AddVertexAt(8, New Point2d(startpt.Value.X + 30, startpt.Value.Y), 0, 0, 0)
    '    Using tr
    '        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
    '        Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        Dim objid As ObjectId = btr.AppendEntity(pl)

    '        pl = tr.GetObject(objid, OpenMode.ForWrite)

    '        'pl.SetBulgeAt(1, Math.Tan(45 * 0.5))

    '        tr.AddNewlyCreatedDBObject(pl, True)
    '        tr.Commit()
    '    End Using

    'End Sub

End Class
