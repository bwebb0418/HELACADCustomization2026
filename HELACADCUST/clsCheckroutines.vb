Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
''' <summary>
''' Class for checking and creating AutoCAD entities like layers, text styles, etc.
''' Provides validation and setup functions for drawing standards.
''' </summary>

Public Class ClsCheckroutines
    Public ReadOnly Property Thisdrawing() As AcadDocument
    ''' <summary>
    ''' Checks for a layer by name and creates it if it doesn't exist.
    ''' </summary>
    ''' <param name="layname">The name of the layer.</param>
    ''' <returns>The ObjectId of the layer.</returns>
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public Function Layercheck(ByVal layname As String) As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction
        Dim layers As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
        Dim ltr As LayerTableRecord
        Dim ltrid As ObjectId
        If layers.Has(layname) = False Then

            ltr = New LayerTableRecord With {
                .Name = layname
            }
            layers.UpgradeOpen()
            ltrid = layers.Add(ltr)
            tr.AddNewlyCreatedDBObject(ltr, True)

        Else
            ltrid = layers.Item(layname)
        End If
        tr.Commit()
        Return ltrid
    End Function

    Public Function Dimstcheck(ByVal dimstname As String) As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction

        Dim dimsty As DimStyleTable = tr.GetObject(db.DimStyleTableId, OpenMode.ForRead)

        Dim dimstid As ObjectId
        If dimsty.Has(dimstname) = True Then
            dimstid = dimsty.Item(dimstname)
        Else

        End If
        tr.Commit()
        Return dimstid
    End Function

    Public Function Laychk(ByVal layname As String) As Boolean
        Dim layers As AcadLayers
        Dim layer As AcadLayer

        layers = Thisdrawing.Layers
        For Each layer In layers
            If layer.Name = layname Then
                Laychk = True
                Exit Function
            End If
        Next
        Laychk = False
    End Function

    Public Function Textcheck(ByVal txt As String) As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction

        Dim text As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)

        Dim txsid As ObjectId
        If text.Has(txt) = False Then
            Dim VTD As New ClsVTD
            VTD.Text_Gen(txt)

        Else
            txsid = text.Item(txt)

        End If
        tr.Commit()
        Return txsid
    End Function

    Public Function Txtchk(ByVal txt As String) As Boolean
        Dim tstyles As AcadTextStyles
        Dim tstyle As AcadTextStyle

        tstyles = Thisdrawing.TextStyles
        For Each tstyle In tstyles
            If tstyle.Name = txt Then
                Txtchk = True
                Exit Function
            End If
        Next
        Txtchk = False
    End Function

    Public Function Dimchk(ByVal ddim As String) As Boolean
        Dim dstyles As AcadDimStyles
        Dim dstyle As AcadDimStyle

        dstyles = Thisdrawing.DimStyles
        For Each dstyle In dstyles
            If dstyle.Name = ddim Then
                Dimchk = True
                Exit Function
            End If
        Next
        Dimchk = False
    End Function

    Public Function Blkchk(ByVal blkname As String) As Boolean
        Dim blks As AcadBlocks
        Dim blk As AcadBlock

        blks = Thisdrawing.Blocks
        For Each blk In blks
            If blk.Name = blkname Then
                Blkchk = True
                Exit Function
            End If
        Next
        Blkchk = False
    End Function

    Function SetLayerCurrent(ByVal slayername As String)

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                         OpenMode.ForRead)

            layercheck(slayername)

            If acLyrTbl.Has(slayername) = True Then
                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(slayername)

                '' Save the changes
                acTrans.Commit()
            End If

            '' Dispose of the transaction
        End Using
        SetLayerCurrent = Nothing
    End Function


End Class
