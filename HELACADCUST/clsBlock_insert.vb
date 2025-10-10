Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry


Public Class ClsBlock_insert
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property
    Public Shared rotation As Double
    Public Shared BLKSFFACT As Double
    Public Function InsertBlockWithJig(ByVal BlockName As String, ByRef multi As Boolean, ByRef prmptstring As String, ByRef rotate As Boolean, ByRef blkscl As Double, Optional ByRef ortho As Boolean = False) As ObjectId

        Dim myDB As Database
        Dim BLKSCALE As Double
        If BLKSFFACT = Nothing Then BLKSFFACT = 1
        BLKSCALE = blkscl * BLKSFFACT

        myDB = HostApplicationServices.WorkingDatabase
        Dim myJig As clsblock_jig
        Using myTrans As Transaction = myDB.TransactionManager.StartTransaction
            Dim myBT As BlockTable = myDB.BlockTableId.GetObject(OpenMode.ForRead)
            If myBT.Has(BlockName) Then
                Dim myBTR As BlockTableRecord = myBT(BlockName).GetObject(OpenMode.ForRead)
                myJig = New clsblock_jig(New Point3d(0, 0, 0), myBTR.ObjectId, BLKSCALE)
            Else
                Exit Function
            End If

        End Using
        Dim myBlkID As ObjectId
        Dim SelPt As Autodesk.AutoCAD.EditorInput.PromptPointResult
        Do
            SelPt = myJig.BeginJig(multi, prmptstring, rotate, ortho)
            If SelPt IsNot Nothing Then
                Select Case SelPt.Status
                    Case Autodesk.AutoCAD.EditorInput.PromptStatus.OK

                        myBlkID = InsertBlock(SelPt.Value, BlockName, BLKSCALE, BLKSCALE, BLKSCALE, rotation)
                    Case Autodesk.AutoCAD.EditorInput.PromptStatus.Other
                        Exit Function
                End Select
            End If
            If multi = False Then SelPt = Nothing
            If SelPt Is Nothing Then Exit Do
        Loop While SelPt.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK
        Return myBlkID
    End Function

    Public Function InsertBlock(ByVal InsPt As Autodesk.AutoCAD.Geometry.Point3d,
ByVal BlockName As String, ByVal XScale As Double,
ByVal YScale As Double, ByVal ZScale As Double, ByVal rotate As Double) _
As Autodesk.AutoCAD.DatabaseServices.ObjectId
        Dim myBlockRef As BlockReference
        Dim myDB As Database = HostApplicationServices.WorkingDatabase
        Using myTrans As Transaction = myDB.TransactionManager.StartTransaction
            Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
                Dim myBT As BlockTable = myDB.BlockTableId.GetObject(OpenMode.ForRead)
                Dim myBTR As BlockTableRecord
                If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
                    myBTR = myBT(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)
                Else
                    myBTR = myBT(BlockTableRecord.PaperSpace).GetObject(OpenMode.ForWrite)
                End If
                'Insert the Block
                Dim myBlockDef As BlockTableRecord = myBT(BlockName).GetObject(OpenMode.ForRead)
                myBlockRef = New BlockReference(InsPt, myBT(BlockName)) With {
                .ScaleFactors = New Autodesk.AutoCAD.Geometry.Scale3d(XScale, YScale, ZScale),
                .Rotation = rotate
                }
                myBTR.AppendEntity(myBlockRef)
                myTrans.AddNewlyCreatedDBObject(myBlockRef, True)
                'Set the Attribute Value
                Dim myAttColl As Autodesk.AutoCAD.DatabaseServices.AttributeCollection
                Dim myEnt As Autodesk.AutoCAD.DatabaseServices.Entity
                myAttColl = myBlockRef.AttributeCollection
                For Each entID As ObjectId In myBlockDef
                    myEnt = entID.GetObject(OpenMode.ForWrite)
                    If TypeOf myEnt Is Autodesk.AutoCAD.DatabaseServices.AttributeDefinition Then
                        Dim myAttDef As Autodesk.AutoCAD.DatabaseServices.AttributeDefinition = myEnt
                        Dim myAttRef As New Autodesk.AutoCAD.DatabaseServices.AttributeReference
                        myAttRef.SetAttributeFromBlock(myAttDef,
                        myBlockRef.BlockTransform)
                        myAttColl.AppendAttribute(myAttRef)
                        myTrans.AddNewlyCreatedDBObject(myAttRef, True)
                    End If
                Next
                myTrans.Commit()
            End Using
        End Using
        rotation = Nothing
        Return myBlockRef.ObjectId
    End Function

    Public Function Loadblock(ByRef name As String) As Boolean
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        Dim blkpath As String

        Dim profile As IConfigurationSection = Application.UserConfigurationManager.OpenCurrentProfile

        Dim suppathstart As String = profile.OpenSubsection("General").ReadProperty("ACAD", String.Empty)
        If suppathstart = "" Then
            Loadblock = False
            Exit Function
        End If

        Dim suppath As String() = suppathstart.Split(New Char() {";"c})
        Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument

            Using xdb As New Database(False, True)
                For Each path In suppath
                    Try
                        xdb.ReadDwgFile(path & "\" & name, FileOpenMode.OpenForReadAndReadShare, True, Nothing)
                        blkpath = path & "\" & name
                        Using tr As Transaction = doc.Database.TransactionManager.StartTransaction
                            Dim blkname As String = SymbolUtilityServices.GetBlockNameFromInsertPathName(blkpath)
                            Dim objid As ObjectId = doc.Database.Insert(blkname, xdb, True)
                            If objid.IsNull Then
                                MsgBox("Block Not Found")
                                Loadblock = False
                            End If
                            tr.Commit()
                            Loadblock = True
                            Exit Function
                        End Using
                    Catch ex As Exception
                        'MsgBox("Block " & name & " Not Found")


                    Finally

                    End Try

                Next


            End Using
        End Using
        Loadblock = False
        Dim writetocmd = New clsutilities
        writetocmd.WritetoCMD("Block " & name & " Not Found")
    End Function
End Class
