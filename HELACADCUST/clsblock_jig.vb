Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry

Public Class Clsblock_jig
    Inherits Autodesk.AutoCAD.EditorInput.EntityJig
    Public Shared BasePt As Point3d
    Public Shared dynam As Boolean
    Public Shared dynamprop1 As String
    Public Shared dynamprop2 As String
    'Dim myMatrix As Matrix3d
    Dim myBRef As Autodesk.AutoCAD.DatabaseServices.BlockReference
    Dim myOpts As Autodesk.AutoCAD.EditorInput.JigPromptPointOptions
    Private lastMousePosition As Point3d = BasePt
    Private rotateblk As Boolean
    'Private mref2defmap As Dictionary(Of AttributeReference, AttributeDefinition)



    Sub New(ByVal BlockIns As Point3d, ByVal BlockID As ObjectId, ByVal blkscl As Double)
        MyBase.New(New Autodesk.AutoCAD.DatabaseServices.BlockReference(BlockIns, BlockID))
        myBRef = Me.Entity

        Dim BLKSCALE As Double
        If ClsBlock_insert.BLKSFFACT = Nothing Then ClsBlock_insert.BLKSFFACT = 1
        BLKSCALE = blkscl
        Dim sclfactor As New Scale3d(BLKSCALE, BLKSCALE, BLKSCALE)
        myBRef.ScaleFactors = sclfactor

    End Sub

    Function BeginJig(ByRef multi As Boolean, ByRef Prmptstring As String, ByRef rotate As Boolean, Optional ByRef ortho As Boolean = False) As PromptPointResult
        If BasePt = Nothing Then
            BasePt = New Point3d(0, 0, 0)
        End If
        If myOpts Is Nothing Then
            myOpts = New Autodesk.AutoCAD.EditorInput.JigPromptPointOptions() With {
            .Message = vbCrLf & Prmptstring,
            .Cursor = Autodesk.AutoCAD.EditorInput.CursorType.Crosshair,
            .UseBasePoint = False,
            .UserInputControls = UserInputControls.Accept3dCoordinates
            }
            If ortho = True Then
                myOpts.UserInputControls = UserInputControls.GovernedByOrthoMode
            End If
        End If
        Dim ed As Autodesk.AutoCAD.EditorInput.Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim myPR As PromptResult
        rotateblk = rotate
        myPR = ed.Drag(Me)
        If multi = True Then
            Do
                Select Case myPR.Status
                    Case Autodesk.AutoCAD.EditorInput.PromptStatus.OK
                        Return myPR
                        Exit Do
                    Case Autodesk.AutoCAD.EditorInput.PromptStatus.None
                        Return myPR
                        Exit Do
                    Case Autodesk.AutoCAD.EditorInput.PromptStatus.Other
                        Return myPR
                        Exit Do
                End Select
            Loop While myPR.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel
        Else
            HELACADCustomization.ClsBlock_insert.rotation = myBRef.Rotation
            Return myPR
        End If
        Return Nothing
    End Function
    Protected Overrides Function Sampler(ByVal prompts As JigPrompts) As SamplerStatus
        Dim myPPR As PromptPointResult
        myPPR = prompts.AcquirePoint(myOpts)
        Dim curPos As Point3d
        curPos = myPPR.Value
        If (curPos.IsEqualTo(BasePt)) Then
            Return SamplerStatus.NoChange
        Else
            'myMatrix = Autodesk.AutoCAD.Geometry.Matrix3d.Displacement(BasePt.GetVectorTo(curPos))
            BasePt = curPos
            Return SamplerStatus.OK
        End If

        lastMousePosition = myPPR.Value

        Return SamplerStatus.OK
    End Function
    Protected Overrides Function Update() As Boolean
        myBRef.Position = BasePt
        'Return False
        'Else
        If rotateblk = True Then
            Dim oldPosition As New Point2d(myBRef.Position.X, myBRef.Position.Y)
            Dim newPosition As New Point2d(lastMousePosition.X, lastMousePosition.Y)
            Dim utils As New Clsutilities
            myBRef.Rotation = oldPosition.GetVectorTo(newPosition).Angle - utils.Degtorad(180) '- Math.PI / 4
        End If

        Return False

    End Function

    Public Function Jigattribs() As Dictionary(Of AttributeReference, AttributeDefinition)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Dim tr As Transaction = db.TransactionManager.StartTransaction

        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim btr As BlockTableRecord = tr.GetObject(bt(myBRef.Name), OpenMode.ForRead)
        If myBRef.Database = Nothing Then
            Dim modelspace As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            modelspace.AppendEntity(myBRef)
            tr.AddNewlyCreatedDBObject(myBRef, True)
        End If
        Dim dict As New Dictionary(Of AttributeReference, AttributeDefinition)
        Dim attcol As AttributeCollection = myBRef.AttributeCollection
        Dim tst As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)

        Jigattribs = Nothing

        If btr.HasAttributeDefinitions Then
            For Each objid As ObjectId In btr
                Dim obj As DBObject = tr.GetObject(objid, OpenMode.ForRead)
                If TypeOf obj Is AttributeDefinition Then
                    Dim attdef As AttributeDefinition = obj
                    Dim attref As New AttributeReference
                    attref.SetAttributeFromBlock(attdef, myBRef.BlockTransform)
                    myBRef.AttributeCollection.AppendAttribute(attref)
                    tr.AddNewlyCreatedDBObject(attref, True)
                    dict.Add(attref, attdef)

                End If
            Next

            Jigattribs = dict
        End If

    End Function

End Class
