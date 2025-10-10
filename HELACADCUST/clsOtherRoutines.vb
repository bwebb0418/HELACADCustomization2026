Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Windows

Public Class ClsOtherRoutines
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public check_routines As New ClsCheckroutines
    Public utilities As New Clsutilities
    <CommandMethod("ARCH_Underlay")> Public Sub Arch_Underlay()
        'Dim blks As AcadBlocks
        Dim blkref As AcadBlockReference
        Dim htch As AcadHatch
        Dim mtext As AcadMText
        Dim text As AcadText
        Dim ldim As AcadDimension
        Dim rdim As AcadDimRotated
        Dim adim As AcadDimAligned
        Dim lead As AcadLeader
        Dim attref As AcadAttribute
        Dim solid As AcadSolid

        MsgBox("Converting to Architectural Underlay, this may take a few minutes", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
        On Error GoTo nd

        Thisdrawing.Utility.Prompt("Exploding blocks")
        For Each obj In Thisdrawing.ModelSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                blkref.Explode()
                blkref.Delete()
            End If
        Next obj
        Thisdrawing.Utility.Prompt("Exploding blocks")
        For Each obj In Thisdrawing.ModelSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                blkref.Explode()
                blkref.Delete()
            End If
        Next obj
        Thisdrawing.Utility.Prompt("Exploding blocks")
        For Each obj In Thisdrawing.ModelSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                blkref.Explode()
                blkref.Delete()
            End If
        Next obj
        Thisdrawing.Utility.Prompt("Exploding blocks")
        For Each obj In Thisdrawing.ModelSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                blkref.Explode()
                blkref.Delete()
            End If
        Next obj
        Thisdrawing.Utility.Prompt("Exploding blocks")
        For Each obj In Thisdrawing.ModelSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                blkref.Explode()
                blkref.Delete()
            End If
        Next obj
        Thisdrawing.Utility.Prompt("Deleting hatches, text, dimensions, and leaders")
        For Each obj In Thisdrawing.ModelSpace
            If TypeOf obj Is AcadHatch Then
                htch = obj
                htch.Delete()
            ElseIf TypeOf obj Is AcadMText Then
                mtext = obj
                mtext.Delete()
            ElseIf TypeOf obj Is AcadText Then
                text = obj
                text.Delete()
            ElseIf TypeOf obj Is AcadDimension Then
                ldim = obj
                ldim.Delete()
            ElseIf TypeOf obj Is AcadDimRotated Then
                rdim = obj
                rdim.Delete()
            ElseIf TypeOf obj Is AcadDimAligned Then
                adim = obj
                adim.Delete()
            ElseIf TypeOf obj Is AcadAttribute Then
                attref = obj
                attref.Delete()
            ElseIf TypeOf obj Is AcadLeader Then
                lead = obj
                lead.Delete()
            ElseIf TypeOf obj Is AcadSolid Then
                solid = obj
                solid.Delete()
            End If
        Next obj

        Thisdrawing.Utility.Prompt("Setting objects to Arch layer")

        If check_routines.Laychk("S-ARCH") = False Then Thisdrawing.Layers.Add("S-ARCH")
        Thisdrawing.Layers.Item("S-ARCH").color = 254

        For Each obj In Thisdrawing.ModelSpace
            obj.layer = "S-ARCH"
            obj.color = AcColor.acByLayer
            'obj.Linetype = "bylayer"
            obj.Lineweight = ACAD_LWEIGHT.acLnWtByLayer
        Next

        Thisdrawing.Utility.Prompt("Purging drawing")
        Dim bindpurge As Clsbind_purge
        bindpurge = New Clsbind_purge
        bindpurge.Purgedwg()
        bindpurge.Purgedwg()
        bindpurge.Purgedwg()
        Thisdrawing.Regen(AcRegenType.acAllViewports)

        Thisdrawing.SendCommand("Flatten" & vbCr)
        Thisdrawing.SendCommand("all" & vbCr & vbCr & vbCr)


        MsgBox("Conversion to Architectural Underlay Sucessful", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
        Exit Sub
nd:
        MsgBox("Conversion to Architectural Underlay Unsucessful", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Error!")
    End Sub

    <CommandMethod("Boundingbox")> Public Sub Bounding_box()
        Dim obj As Object
        Dim maxpoint As Object = Nothing
        Dim minpoint As Object = Nothing
        Dim point
        Dim maxpoint2 As Object = Nothing
        Dim minpoint2 As Object = Nothing
        Dim vertlist(0 To 7) As Double
        Dim pline As AcadLWPolyline
        Dim userr1, userr2, userr3
        Dim space
        Dim MTOBJ As AcadMText
        Dim DTOBJ As AcadText
        Dim SSOBJS As AcadSelectionSet
        Dim i As Integer
        Dim filterdata(0) As Object
        Dim filtertype(0) As Int16

        For Each SSOBJS In Thisdrawing.SelectionSets
            If SSOBJS.Name = "BBobjs" Then
                SSOBJS.Delete()
                Exit For
            End If
        Next SSOBJS

        SSOBJS = Thisdrawing.SelectionSets.Add("BBobjs")

        'On Error GoTo err

        filtertype(0) = 0 : filterdata(0) = "Text,MText"
        On Error GoTo 0

        Thisdrawing.Utility.Prompt("Select Text to Box: ")
        SSOBJS.SelectOnScreen(filtertype, filterdata)


        ' ''Thisdrawing.Utility.Prompt("Select Text to Combine: ")
        ' ''SSOBJS.SelectOnScreen()

        On Error GoTo nd

        For i = 0 To SSOBJS.Count - 1

            obj = SSOBJS(i)

            If TypeOf obj Is AcadMText Then
                MTOBJ = obj
                MTOBJ.GetBoundingBox(minpoint, maxpoint)
                MTOBJ.AttachmentPoint = AcAttachmentPoint.acAttachmentPointMiddleCenter
                MTOBJ.GetBoundingBox(minpoint2, maxpoint2)
                MTOBJ.Move(minpoint2, minpoint)
            ElseIf TypeOf obj Is AcadText Then
                DTOBJ = obj
                DTOBJ.GetBoundingBox(minpoint, maxpoint)
                DTOBJ.Alignment = AcAlignment.acAlignmentMiddleCenter
                DTOBJ.GetBoundingBox(minpoint2, maxpoint2)
                DTOBJ.Move(minpoint2, minpoint)
            Else
                obj.GetBoundingBox(minpoint, maxpoint)
            End If

            If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
                space = Thisdrawing.ModelSpace
                userr2 = Thisdrawing.GetVariable("userr2")
            Else
                space = Thisdrawing.PaperSpace
                userr2 = 1
            End If

            userr1 = Thisdrawing.GetVariable("userr1")
            'userr2 = Thisdrawing.GetVariable("userr2")
            userr3 = Thisdrawing.GetVariable("userr3")

            vertlist(0) = minpoint(0) : vertlist(1) = minpoint(1)
            vertlist(2) = maxpoint(0) : vertlist(3) = minpoint(1)
            vertlist(4) = maxpoint(0) : vertlist(5) = maxpoint(1)
            vertlist(6) = minpoint(0) : vertlist(7) = maxpoint(1)

            pline = space.AddLightWeightPolyline(vertlist)

            pline.Closed = True
            If check_routines.Laychk("~-L18") = False Then Thisdrawing.Layers.Add("~-L18")
            pline.Layer = "~-L18"
            If TypeOf obj Is AcadText Then
                userr1 = obj.Height
            ElseIf TypeOf obj Is AcadMText Then
                userr1 = obj.Height
            End If
            pline.Offset(2 * ((userr1 / userr2 / userr3) * userr2 * userr3) / 3)
            pline.Offset(((userr1 / userr2 / userr3) * userr2 * userr3) / 3)
            pline.Delete()

        Next i

nd:
    End Sub

    <CommandMethod("TextCombine")> Public Sub Text_combine()
        Dim SSOBJS As AcadSelectionSet
        Dim filterdata(0) As Object
        Dim filtertype(0) As Int16
        Dim space
        Dim inspoint
        Dim secpoint
        Dim nmtobj As AcadMText
        Dim i As Integer
        Dim ntext = ""
        Dim width As Double
        Dim txtlay
        Dim txtht
        Dim dtext As AcadText
        Dim Line As AcadLine
        Dim Line1 As AcadLine
        Dim dtarray(0 To 1) As AcadEntity
        Dim hightext As AcadMText
        Dim j As Integer
        Dim refindex As Integer
        Dim comparetext As AcadMText
        Dim removearray(0 To 2) As AcadEntity
        Dim addarray(0 To 1) As AcadEntity
        Dim Sname As String
        Dim theight As Double
        Dim lname As String


        For Each SSOBJS In Thisdrawing.SelectionSets
            If SSOBJS.Name = "textobjs" Then
                SSOBJS.Delete()
                Exit For
            End If
        Next SSOBJS

        SSOBJS = Thisdrawing.SelectionSets.Add("textobjs")

        'On Error GoTo err

        filtertype(0) = 0 : filterdata(0) = "Text,MText"

        Thisdrawing.Utility.Prompt("Select Text to Combine: ")
        SSOBJS.SelectOnScreen(filtertype, filterdata)
        If SSOBJS.Count < 2 Then Exit Sub
        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If
        On Error GoTo err
        inspoint = utilities.Getpoint("Select New Text Insertion Point")
        secpoint = utilities.Getendpoint(inspoint, "Select Text Box Width")
        width = secpoint(0) - inspoint(0)

        j = 0

        Do

            For i = 0 To SSOBJS.Count - 1
                If TypeOf SSOBJS(i) Is AcadText Then
                    dtext = SSOBJS(i)
                    Line = space.AddLine(secpoint, inspoint)
                    SSOBJS.Select(AcSelect.acSelectionSetLast)
                    nmtobj = space.AddMText(dtext.InsertionPoint, width, dtext.TextString)
                    SSOBJS.Select(AcSelect.acSelectionSetLast)
                    nmtobj.StyleName = dtext.StyleName
                    nmtobj.Height = dtext.Height
                    nmtobj.Layer = dtext.Layer
                    '        nmtobj.AttachmentPoint = dtext.Alignment
                    dtarray(0) = dtext
                    dtarray(1) = Line
                    SSOBJS.RemoveItems(dtarray)
                    dtext.Delete()
                    Line.Delete()
                End If
            Next i

            j += 1

        Loop Until j = 3

        Sname = SSOBJS(0).StyleName
        theight = SSOBJS(0).Height
        lname = SSOBJS(0).Layer

        Line = space.AddLine(secpoint, inspoint)
        Line1 = space.AddLine(secpoint, inspoint)
        addarray(0) = Line
        addarray(1) = Line1

        Do
            If TypeOf SSOBJS(0) Is AcadMText Then
                hightext = SSOBJS(0)
            Else
                hightext = SSOBJS(1)
            End If

            j = 0

            Do
                i = 0
                For i = 0 To SSOBJS.Count - 1
                    comparetext = SSOBJS(i)
                    If hightext.InsertionPoint(1) < comparetext.InsertionPoint(1) Then
                        hightext = SSOBJS(i)
                        refindex = i
                    ElseIf hightext.InsertionPoint(1) = comparetext.InsertionPoint(1) Then
                        If comparetext.InsertionPoint(0) < hightext.InsertionPoint(0) Then
                            hightext = SSOBJS(i)
                            refindex = i
                        End If
                    End If

                Next i
                j += 1
            Loop Until j = SSOBJS.Count

            If ntext = "" Then
                ntext = hightext.TextString
            Else
                ntext = ntext & "\P" & hightext.TextString
            End If

            SSOBJS.AddItems(addarray)
            removearray(0) = hightext
            removearray(1) = Line
            removearray(2) = Line1
            SSOBJS.RemoveItems(removearray)
            hightext.Delete()


        Loop Until SSOBJS.Count = 0

        nmtobj = space.AddMText(inspoint, width, ntext)

        nmtobj.StyleName = Sname
        nmtobj.Height = theight
        nmtobj.Layer = lname

        Line.Delete()
        Line1.Delete()


        On Error Resume Next


err:
    End Sub

    <CommandMethod("TextAlign")> Public Sub Text_align()
        Dim pso1 As New PromptSelectionOptions
        Dim psr1 As PromptSelectionResult
        Dim psr2 As PromptSelectionResult
        Dim ent As Entity
        Dim txt As DBText
        Dim mtxt As MText
        Dim alignpoint As Point3d
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction

        pso1.AllowDuplicates = False
        pso1.SingleOnly = True
        pso1.RejectObjectsOnLockedLayers = False
        pso1.MessageForAdding = "Pick text to align to: "

        Dim selectfilter(0) As TypedValue
        selectfilter.SetValue(New TypedValue(DxfCode.Start, "TEXT,MTEXT"), 0)
        Dim psf As New SelectionFilter(selectfilter)

        psr1 = ed.GetSelection(pso1, psf)

        If psr1.Status <> PromptStatus.OK Then Exit Sub

        Dim pso2 As New PromptSelectionOptions With {
        .AllowDuplicates = False,
        .MessageForAdding = "Pick text to align: ",
        .MessageForRemoval = "Pick text to remove from align command: ",
        .SingleOnly = False
        }

        psr2 = ed.GetSelection(pso2, psf)

        If psr2.Status <> PromptStatus.OK Then Exit Sub

        tr = db.TransactionManager.StartTransaction
        Dim so1 As SelectedObject = psr1.Value(0)

        ent = tr.GetObject(so1.ObjectId, OpenMode.ForRead)
        If TypeOf ent Is DBText Then
            txt = ent
            alignpoint = txt.Position
        ElseIf TypeOf ent Is MText Then
            mtxt = ent
            alignpoint = mtxt.Location
        Else
            Exit Sub
        End If
        'ent = Nothing
        Try
            Dim SS As Autodesk.AutoCAD.EditorInput.SelectionSet = psr2.Value
            For Each so2 As SelectedObject In SS
                ent = tr.GetObject(so2.ObjectId, OpenMode.ForWrite)
                If TypeOf ent Is DBText Then
                    txt = ent
                    txt.Position = New Point3d(alignpoint.X, txt.Position.Y, txt.Position.Z)

                ElseIf TypeOf ent Is MText Then
                    mtxt = ent
                    mtxt.Location = New Point3d(alignpoint.X, mtxt.Location.Y, mtxt.Location.Z)
                Else

                End If
            Next
            tr.Commit()
        Catch ex As Exception
            tr.Abort()
        End Try

    End Sub

    Public Sub Oldlayerconvert()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction

        Using layers As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
            Dim ltr As LayerTableRecord
            Dim ltrid As ObjectId

            For Each ltrid In layers
                ltr = tr.GetObject(ltrid, OpenMode.ForRead)
                If Strings.Left(ltr.Name, 3) = "S-L" Or Strings.Left(ltr.Name, 3) = "S-T" Then
                    ltr.UpgradeOpen()
                    ltr.Name = "~-" & Strings.Right(ltr.Name, (Strings.Len(ltr.Name)) - 2)
                ElseIf Strings.Left(ltr.Name, 7) = "S-SHADE" Then
                    ltr.UpgradeOpen()
                    ltr.Name = "~-" & Strings.Right(ltr.Name, (Strings.Len(ltr.Name)) - 2)
                ElseIf Strings.Left(ltr.Name, 6) = "S-GRID" Then
                    ltr.UpgradeOpen()
                    ltr.Name = "~-" & Strings.Right(ltr.Name, (Strings.Len(ltr.Name)) - 2)
                ElseIf Strings.Left(ltr.Name, 6) = "S-XREF" Then
                    ltr.UpgradeOpen()
                    ltr.Name = "~-" & Strings.Right(ltr.Name, (Strings.Len(ltr.Name)) - 2)
                ElseIf Strings.Left(ltr.Name, 6) = "S-VPORT" Then
                    ltr.UpgradeOpen()
                    ltr.Name = "~-" & Strings.Right(ltr.Name, (Strings.Len(ltr.Name)) - 2)
                End If
            Next
            tr.Commit()
        End Using

    End Sub
    Function get_path(FName As String) As String
        Dim profile As IConfigurationSection = Application.UserConfigurationManager.OpenCurrentProfile
        Dim suppathstart As String = profile.OpenSubsection("General").ReadProperty("ACAD", String.Empty)
        Dim suppath As String() = suppathstart.Split(New Char() {";"c})
        Dim Fpath As String

        Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument


            For Each path In suppath
                Try
                    If System.IO.Directory.GetFiles(path, FName, IO.SearchOption.TopDirectoryOnly).Length > 0 Then
                        Fpath = path & "\" & FName
                        Return Fpath
                    Else

                    End If
                Catch ex As Exception



                Finally

                End Try

            Next path



        End Using
        Return Nothing
    End Function
End Class
