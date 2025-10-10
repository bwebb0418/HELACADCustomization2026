Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime

Public Class ClsGridlines
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public check_routines As New ClsCheckroutines
    Public text_dimensions As New ClsVTD

    Public glval As String


    <CommandMethod("Gridlines")> Public Sub Gridline()
        Dim ans As String
        Dim stpt As Object
        Dim ndpt As Object
        Dim gline As AcadLine
        Dim existing As String
        Dim gltag As AcadBlockReference
        Dim gltagattribs
        Dim orthomode As Integer
        Dim glvali

        On Error GoTo nd
        stpt = Thisdrawing.Utility.GetPoint(, "Pick Start Point of Gridline: ")
        ndpt = Thisdrawing.Utility.GetPoint(stpt, "Pick End Point of Gridline: ")
        On Error GoTo 0

        gline = Thisdrawing.ModelSpace.AddLine(stpt, ndpt)

        If check_routines.Laychk("~-GRID") = False Then Thisdrawing.Layers.Add("~-GRID")

        gline.Layer = "~-GRID"
        Gl_tag(stpt, gline)

        gltag = Thisdrawing.ModelSpace.Item(Thisdrawing.ModelSpace.Count - 1)

        On Error Resume Next
        existing = Thisdrawing.Utility.GetString(False, "New or Existing Gridline? ")

        If LCase(Left(existing, 1)) = "e" Then
            If check_routines.Laychk("~-GRIDTXTEX") = False Then Thisdrawing.Layers.Add("~-GRIDTXTEX")
            gltag.Layer = "~-GRIDTXTEX"
        Else
            If check_routines.Laychk("~-GRIDTXT") = False Then Thisdrawing.Layers.Add("~-GRIDTXT")
            gltag.Layer = "~-GRIDTXT"
        End If

multi:
        On Error Resume Next
        ans = Thisdrawing.Utility.GetString(False, "Add a Parallel Gridline? Yes or No ")
        On Error GoTo 0
        If LCase(Left(ans, 1)) = "y" Then
            orthomode = Thisdrawing.GetVariable("orthomode")
            Thisdrawing.SetVariable("orthomode", 1)
            stpt = gltag.InsertionPoint
            gltag.Copy()
            gltag.Move(gltag.InsertionPoint, Thisdrawing.Utility.GetPoint(gltag.InsertionPoint, "New Gridline Position: "))
            gline.Copy()
            gline.Move(stpt, gltag.InsertionPoint)
            gltagattribs = gltag.GetAttributes
            gltagattribs(0).TextString = "#"
            On Error Resume Next
            If glval = "" Then
                glval = Thisdrawing.Utility.GetString(False, "Insert Gridline Number: ")

            Else
                Dim numlet, letnum, num, lett
                letnum = glval Like "[a-z][0-9]"
                If letnum = False Then letnum = glval Like "[A-Z][0-9]"
                numlet = glval Like "[0-99][a-z]"
                If numlet = False Then numlet = glval Like "[0-99][A-Z]"
                num = glval Like "[0-99]"
                lett = glval Like "[a-z]"
                If lett = False Then lett = glval Like "[A-Z]"
                If letnum = True Then
                    glval = Gl_letter(Strings.Left(glval, 1))
                ElseIf numlet = True Then
                    If Len(glval) = 3 Then
                        glval = Strings.Left(glval, 2) + 1
                    ElseIf Len(glval) = 2 Then
                        glval = Strings.Left(glval, 1) + 1
                    End If
                ElseIf num = True Then
                    glval += 1
                ElseIf lett = True Then
                    glval = Gl_letter(glval)
                End If
                glvali = Thisdrawing.Utility.GetString(False, "Insert Gridline Number: <" & glval & "> ")
                If glvali <> "" Then glval = glvali
            End If
            gltagattribs(0).TextString = glval
            Thisdrawing.SetVariable("orthomode", orthomode)
            ans = ""
        Else
            GoTo nd
        End If

        GoTo multi
nd:

    End Sub

    <CommandMethod("DRAWCIRC")> Sub Build_GL_Tag()

        '' '' '' '' Get the current document and database
        ' '' '' ''Dim acDoc As document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
        '' '' ''Dim acCurDb As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database

        '' '' ''If check_routines.txtchk("HEL40") = False Then text_dimensions.Text_Gen("HEL40")

        '' '' '' '' Start a transaction
        '' '' ''Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

        '' '' ''    '' Open the Block table for read
        '' '' ''    Dim acBlkTbl As BlockTable
        '' '' ''    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

        '' '' ''    '' Open the Block table record Model space for write
        '' '' ''    Dim acBlkTblRec As BlockTableRecord
        '' '' ''    acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
        '' '' ''                                    OpenMode.ForWrite)

        '' '' ''    '' Create a circle that is at 2,3 with a radius of 4.25
        '' '' ''    Dim acCirc As Circle = New Circle()
        '' '' ''    acCirc.SetDatabaseDefaults()
        '' '' ''    acCirc.Center = New Point3d(0, 0, 0)
        '' '' ''    acCirc.Radius = 6.5

        '' '' ''    Dim acatrb As AttributeDefinition = New AttributeDefinition
        '' '' ''    acatrb.TextString = "#"
        '' '' ''    acatrb.Tag = "GRNO"
        '' '' ''    acatrb.Prompt = "Grid No."
        '' '' ''    'acatrb.= New Point3d(0, 0, 0)



        '' '' ''    '' Add the new object to the block table record and the transaction
        '' '' ''    acBlkTblRec.AppendEntity(acCirc)
        '' '' ''    acTrans.AddNewlyCreatedDBObject(acCirc, True)

        '' '' ''    acBlkTblRec.AppendEntity(acatrb)
        '' '' ''    acTrans.AddNewlyCreatedDBObject(acatrb, True)

        '' '' ''    '' Save the new object to the database
        '' '' ''    acTrans.Commit()
        '' '' ''End Using



        Dim txtht As Double
        Dim txtst As String
        Dim attrib As AcadAttribute
        Dim inspoint(0 To 2) As Double ' As Object = inspoint(0) = 0
        Dim blkcirc As AcadCircle
        Dim block As AcadBlock

        Dim arry(1) As AcadObject

        If check_routines.Txtchk("HEL40") = False Then text_dimensions.Text_Gen("HEL40")

        txtht = 4
        txtst = "HEL40"
        'inspoint = New Point2d(0, 0)
        inspoint(0) = 0
        inspoint(1) = 0
        inspoint(2) = 0

        attrib = Thisdrawing.ModelSpace.AddAttribute(txtht, AcAttributeMode.acAttributeModeNormal, "Grid No.", inspoint, "GRNO", "#")
        attrib.Alignment = AcAlignment.acAlignmentMiddleCenter
        attrib.StyleName = txtst

        blkcirc = Thisdrawing.ModelSpace.AddCircle(attrib.TextAlignmentPoint, 6.5)

        blkcirc.Layer = "0"
        attrib.Layer = "0"

        arry(0) = attrib : arry(1) = blkcirc

        block = Thisdrawing.Blocks.Add(inspoint, "hel_grbub")
        Thisdrawing.CopyObjects(arry, block)

        attrib.Delete() : blkcirc.Delete()

    End Sub

    Function Gl_tag(ByVal stpt As Object, ByVal gridline As AcadLine)
        Dim attribs
        Dim SCL As Double
        Dim uni As Double
        Dim gltag As AcadBlockReference
        Dim intersect As Object

        Dim mvpt(0 To 2) As Double

        Dim Line As AcadLine

        SCL = Thisdrawing.GetVariable("userr3")
        uni = Thisdrawing.GetVariable("userr2")

        If SCL = 0 Then SCL = 1
        If uni = 0 Then uni = 1

        If check_routines.Blkchk("hel_grbub") = False Then Build_GL_Tag()

        mvpt(0) = stpt(0) + (6.5 * SCL * uni)
        mvpt(1) = stpt(1)
        mvpt(2) = stpt(2)

        Line = Thisdrawing.ModelSpace.AddLine(stpt, mvpt)
        Line.Rotate(stpt, gridline.Angle)
        Line.ScaleEntity(Line.EndPoint, 2)

        gltag = Thisdrawing.ModelSpace.InsertBlock(Line.StartPoint, "hel_grbub", SCL * uni, SCL * uni, SCL * uni, 0)
        Line.Delete()
        'Thisdrawing.SendCommand "ate l "
        attribs = gltag.GetAttributes
        On Error Resume Next
        glval = Thisdrawing.Utility.GetString(False, "Insert Gridline Number: ")
        attribs(0).TextString = glval
    End Function

    Public Function Gl_letter(ByVal Chara As String)

        If Asc(Chara) >= Asc("a") And Asc(Chara) < Asc("z") Then
            Gl_letter = Chr(Asc(Chara) + 1)
        ElseIf Asc(Chara) >= Asc("A") And Asc(Chara) < Asc("Z") Then
            Gl_letter = Chr(Asc(Chara) + 1)
        Else
            Gl_letter = glval
        End If
    End Function


End Class
