Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows

Public Class ClsIssues
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public blk As AcadBlockReference
    Dim check_routines As New ClsCheckroutines
    Dim utilities As New Clsutilities
    Public dte As String
    Dim vportlay As AcadLayer


    <CommandMethod("IFC")> Public Sub IFC()

        Dim blkref As AcadBlockReference
        Dim revnumber As String = ""
        Dim obj As Object
        Dim dynprops
        Dim onestamp As Boolean
        Dim attribs
        Dim attrib As AcadAttributeReference
        'Set ps = Thisdrawing.ActiveLayout
        onestamp = False
        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If LCase(blkref.EffectiveName) = "hel_stamps-2026" Then
                    dynprops = blkref.GetDynamicBlockProperties
                    If onestamp = True Then
                        blkref.Delete()
                        Exit For
                    End If
                    For I = LBound(dynprops) To UBound(dynprops)
                        If LCase(dynprops(I).PropertyName) = "visibility" Then
                            dynprops(I).Value = "Issued for Construction"
                            onestamp = True
                        End If
                    Next
                End If
            End If
        Next obj


        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If Left(LCase(blkref.EffectiveName), 9) = "hel_2026_" Then
                    attribs = blkref.GetAttributes
                    For I = LBound(attribs) To UBound(attribs)
                        If LCase(Thisdrawing.GetVariable("users1")) = "inds" Then
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    attrib.TextString = "0"
                                    revnumber = attrib.TextString
                                ElseIf attrib.TextString Like "[A-Z]" Then
                                    attrib.TextString = "0"
                                    revnumber = attrib.TextString
                                Else
                                    attrib.TextString = Int(attrib.TextString) + 1
                                    revnumber = attrib.TextString
                                End If
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    attribs(I + 2).TextString = "CONSTRUCTION"
                                    Exit For
                                End If
                            End If
                        Else
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    attrib.TextString = "0"
                                    revnumber = attrib.TextString
                                Else
                                    attrib.TextString = Int(attrib.TextString) + 1
                                    revnumber = attrib.TextString
                                End If
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    attribs(I + 2).TextString = "CONSTRUCTION"
                                    Exit For
                                End If
                            End If
                        End If
                    Next I
                End If
            End If
        Next obj


        If check_routines.Laychk("~-VPORT") = False Then

            vportlay = Thisdrawing.Layers.Add("~-VPORT")
            vportlay.Plottable = False
        Else
            Thisdrawing.Layers.Item("~-VPORT").Freeze = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").LayerOn = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").Lock = vbFalse
        End If
        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadPViewport Then
                obj.Layer = "~-vport"
            End If
        Next obj

        Exit Sub

        Return



    End Sub

    <CommandMethod("IFRS")> Public Sub IFRS1()
        IFRS("")
    End Sub


    Public Sub IFRS(Optional ByRef reviewvalue As String = "")

        Dim blkref As BlockReference
        Dim revnumber As String = ""
        Dim obj As Object
        Dim attribs
        Dim attrib As AcadAttributeReference
        Dim inspoint As Point3d
        Dim rotation
        Dim blkname
        Dim blkscl As New Clsutilities
        Dim intLast As Integer
        Dim intNext As Integer
        Dim blockinsert As New ClsBlock_insert
        Dim blkid As ObjectId
        Dim check_routines As New ClsCheckroutines

        rotation = 0

        blkname = "hel_stamps-2026"
        If check_routines.Blkchk(blkname) = True Then

        Else
            If blockinsert.Loadblock(blkname & ".dwg") = False Then
                Exit Sub
            End If
        End If
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Try
            blkid = blockinsert.InsertBlockWithJig(blkname, False, "Select Stamp Insertion Point: ", False, blkscl.Getscale)
        Catch
            Exit Sub
        End Try
        Dim tr As Transaction = doc.Database.TransactionManager.StartTransaction
        Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
            Try
                blkref = tr.GetObject(blkid, OpenMode.ForWrite)
            Catch
                tr.Abort()
                Exit Sub
            End Try
            Dim defcol As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
            For Each dbpid As DynamicBlockReferenceProperty In defcol

                If dbpid.PropertyName = "Visibility" Then
                    Try
                        dbpid.Value = "Not For Construction"
                    Catch ex As Exception

                    End Try
                End If
            Next
            blkref.LayerId = check_routines.Layercheck("0")
        End Using
        tr.Commit()

        inspoint = New Point3d(blkref.Position.X, blkref.Position.Y + (30 * blkscl.Getscale), blkref.Position.Z)


        blkid = blockinsert.InsertBlock(inspoint, blkname, blkscl.Getscale, blkscl.Getscale, blkscl.Getscale, rotation)

        tr = doc.Database.TransactionManager.StartTransaction
        Using doclock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument

            blkref = tr.GetObject(blkid, OpenMode.ForWrite)
            Dim defcol As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
            For Each dbpid As DynamicBlockReferenceProperty In defcol

                If dbpid.PropertyName = "Visibility" Then
                    Try
                        dbpid.Value = "Issued for Review"
                    Catch ex As Exception

                    End Try
                End If
            Next
            blkref.LayerId = check_routines.Layercheck("0")
        End Using
        tr.Commit()

        'blk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

        'dynprop = blk.GetDynamicBlockProperties
        'For I = LBound(dynprop) To UBound(dynprop)
        '    If dynprop(I).PropertyName = "Visibility" Then
        '        dynprop(I).Value = "Issued for Review"
        '    End If
        'Next I

        'blk.Layer = "0"
        'blk.Update()

        'blk = blk.Copy
        'inspoint(1) = inspoint(1) + (30 * blkscl)
        'blk.Move(blk.InsertionPoint, inspoint)

        'dynprop = blk.GetDynamicBlockProperties
        'For I = LBound(dynprop) To UBound(dynprop)
        '    If dynprop(I).PropertyName = "Visibility" Then
        '        dynprop(I).Value = "Not For Construction"
        '    End If
        'Next I

        'blk.Layer = "0"
        'blk.Update()
        'blkref.Dispose()
        Dim blkref2 As AcadBlockReference

        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref2 = obj
                If Left(LCase(blkref2.EffectiveName), 9) = "hel_2026_" Then
                    attribs = blkref2.GetAttributes
                    For I = LBound(attribs) To UBound(attribs)
                        If LCase(Thisdrawing.GetVariable("users1")) = "inds" Then
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    intLast = 64
                                Else
                                    intLast = Asc(attrib.TextString)
                                End If
                                If intLast = 90 Then intLast = 64 'If letter "z" 96+1= letter "a"
                                intNext = intLast + 1
                                If intNext = 73 Or intNext = 79 Then intNext += 1
                                attrib.TextString = Chr(intNext)
                                revnumber = attrib.TextString
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    Dim ed As Editor = doc.Editor
                                    If reviewvalue = "" Then
                                        Dim pso As New PromptStringOptions("Review Progress: ") With {
                                        .AllowSpaces = True
                                        }
                                        Dim psr As PromptResult = ed.GetString(pso)
                                        If psr.StringResult = "" Then
                                            attribs(I + 2).TextString = "CLIENT REVIEW"
                                        Else
                                            attribs(I + 2).TextString = psr.StringResult & " CLIENT REVIEW"
                                        End If
                                    Else
                                        attribs(I + 2).TextString = reviewvalue & " CLIENT REVIEW"
                                    End If
                                    Exit For
                                End If
                            End If
                        Else
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    attrib.TextString = "0"
                                    revnumber = attrib.TextString
                                Else
                                    attrib.TextString = Int(attrib.TextString) + 1
                                    revnumber = attrib.TextString
                                End If
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    Dim ed As Editor = doc.Editor
                                    If reviewvalue = "" Then
                                        Dim pso As New PromptStringOptions("Review Progress: ") With {
                                        .AllowSpaces = True
                                        }
                                        Dim psr As PromptResult = ed.GetString(pso)
                                        If psr.StringResult = "" Then
                                            attribs(I + 2).TextString = "CLIENT REVIEW"
                                        Else
                                            attribs(I + 2).TextString = psr.StringResult & " CLIENT REVIEW"
                                        End If
                                    Else
                                        attribs(I + 2).TextString = reviewvalue & " CLIENT REVIEW"
                                    End If
                                    Exit For
                                End If
                            End If
                        End If
                    Next I
                End If
            End If
        Next obj


        If check_routines.Laychk("~-VPORT") = False Then
            Dim vportlay As AcadLayer
            vportlay = Thisdrawing.Layers.Add("~-VPORT")
            vportlay.Plottable = False
        Else
            Thisdrawing.Layers.Item("~-VPORT").Freeze = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").LayerOn = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").Lock = vbFalse
        End If
        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadPViewport Then
                obj.Layer = "~-vport"
            End If
        Next obj

        Exit Sub

        Return



    End Sub

    <CommandMethod("IFR")> Public Sub IFR1()
        IFR("")
    End Sub

    Public Sub IFR(Optional ByRef reviewvalue As String = "")

        Dim blkref As AcadBlockReference
        Dim revnumber As String = ""
        Dim attribs
        Dim attrib As AcadAttributeReference
        Dim intLast As Integer
        Dim intNext As Integer
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim dynprops

        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If LCase(blkref.EffectiveName) = "hel_stamps-2026" Then
                    dynprops = blkref.GetDynamicBlockProperties
                    '' ''If onestamp = True Then
                    '' ''    blkref.Delete()
                    '' ''    Exit For
                    '' ''End If
                    For I = LBound(dynprops) To UBound(dynprops)
                        If LCase(dynprops(I).PropertyName) = "visibility" Then
                            If dynprops(I).Value = "Not For Construction" Then

                            Else
                                dynprops(I).Value = "Issued for Review"
                                'onestamp = True
                            End If
                        End If
                    Next
                End If
            End If
        Next obj

        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If Left(LCase(blkref.EffectiveName), 9) = "hel_2026_" Then
                    attribs = blkref.GetAttributes
                    For I = LBound(attribs) To UBound(attribs)
                        If LCase(Thisdrawing.GetVariable("users1")) = "inds" Then
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    intLast = 64
                                Else
                                    intLast = Asc(attrib.TextString)
                                End If
                                If intLast = 90 Then intLast = 64 'If letter "z" 96+1= letter "a"
                                intNext = intLast + 1
                                If intNext = 73 Or intNext = 79 Then intNext += 1
                                attrib.TextString = Chr(intNext)
                                revnumber = attrib.TextString
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    Dim ed As Editor = doc.Editor
                                    If reviewvalue = "" Then
                                        Dim pso As New PromptStringOptions("Review Progress: ") With {
                                        .AllowSpaces = True
                                        }
                                        Dim psr As PromptResult = ed.GetString(pso)
                                        If psr.StringResult = "" Then
                                            attribs(I + 2).TextString = "CLIENT REVIEW"
                                        Else
                                            attribs(I + 2).TextString = psr.StringResult & " CLIENT REVIEW"
                                        End If
                                    Else
                                        attribs(I + 2).TextString = reviewvalue & " CLIENT REVIEW"
                                    End If
                                    Exit For
                                End If
                            End If
                        Else
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    attrib.TextString = "0"
                                    revnumber = attrib.TextString
                                Else
                                    attrib.TextString = Int(attrib.TextString) + 1
                                    revnumber = attrib.TextString
                                End If
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    Dim ed As Editor = doc.Editor
                                    If reviewvalue = "" Then
                                        Dim pso As New PromptStringOptions("Review Progress: ") With {
                                        .AllowSpaces = True
                                        }
                                        Dim psr As PromptResult = ed.GetString(pso)
                                        If psr.StringResult = "" Then
                                            attribs(I + 2).TextString = "CLIENT REVIEW"
                                        Else
                                            attribs(I + 2).TextString = psr.StringResult & " CLIENT REVIEW"
                                        End If
                                    Else
                                        attribs(I + 2).TextString = reviewvalue & " CLIENT REVIEW"
                                    End If
                                    Exit For
                                End If
                            End If
                        End If
                    Next I
                End If
            End If
        Next obj


        If check_routines.Laychk("~-VPORT") = False Then
            Dim vportlay As AcadLayer
            vportlay = Thisdrawing.Layers.Add("~-VPORT")
            vportlay.Plottable = False
        Else
            Thisdrawing.Layers.Item("~-VPORT").Freeze = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").LayerOn = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").Lock = vbFalse
        End If
        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadPViewport Then
                obj.Layer = "~-vport"
            End If
        Next obj

        Exit Sub

        Return



    End Sub

    <CommandMethod("IFT")> Public Sub Ift()

        Dim blkref As AcadBlockReference
        Dim revnumber As String = ""
        Dim obj As Object
        Dim dynprops
        'Dim onestamp As Boolean
        Dim attribs
        Dim intlast
        Dim intnext
        Dim attrib As AcadAttributeReference

        'onestamp = False
        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If LCase(blkref.EffectiveName) = "hel_stamps-2026" Then
                    dynprops = blkref.GetDynamicBlockProperties

                    For I = LBound(dynprops) To UBound(dynprops)
                        If LCase(dynprops(I).PropertyName) = "visibility" Then
                            If dynprops(I).Value = "Not For Construction" Then

                            Else
                                dynprops(I).Value = "Issued for Tender"

                            End If
                        End If
                    Next
                End If
            End If
        Next obj


        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If Left(LCase(blkref.EffectiveName), 9) = "hel_2026_" Then
                    attribs = blkref.GetAttributes
                    For I = LBound(attribs) To UBound(attribs)
                        If LCase(Thisdrawing.GetVariable("users1")) = "inds" Then
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    intlast = 64
                                Else
                                    intlast = Asc(attrib.TextString)
                                End If
                                If intlast = 90 Then intlast = 64 'If letter "z" 96+1= letter "a"
                                intnext = intlast + 1
                                If intnext = 73 Or intnext = 79 Then intnext += 1
                                attrib.TextString = Chr(intnext)
                                revnumber = attrib.TextString
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    attribs(I + 2).TextString = "TENDER"
                                    Exit For
                                End If
                            End If
                        Else
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    attrib.TextString = "0"
                                    revnumber = attrib.TextString
                                Else
                                    attrib.TextString = Int(attrib.TextString) + 1
                                    revnumber = attrib.TextString
                                End If
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    attribs(I + 2).TextString = "TENDER"
                                    Exit For
                                End If
                            End If
                        End If
                    Next I
                End If
            End If
        Next obj


        If check_routines.Laychk("~-VPORT") = False Then
            Dim vportlay As AcadLayer
            vportlay = Thisdrawing.Layers.Add("~-VPORT")
            vportlay.Plottable = False
        Else
            Thisdrawing.Layers.Item("~-VPORT").Freeze = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").LayerOn = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").Lock = vbFalse
        End If
        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadPViewport Then
                obj.Layer = "~-vport"
            End If
        Next obj

        Exit Sub

        Return

    End Sub

    <CommandMethod("IFI")> Public Sub Ifi()

        Dim blkref As AcadBlockReference
        Dim revnumber As String = ""
        Dim obj As Object
        Dim dynprops

        Dim attribs
        Dim attrib As AcadAttributeReference
        Dim intLast As Integer
        Dim intNext As Integer
        'Set ps = Thisdrawing.ActiveLayout

        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If LCase(blkref.EffectiveName) = "hel_stamps-2026" Then
                    dynprops = blkref.GetDynamicBlockProperties
                    '' ''If onestamp = True Then
                    '' ''    blkref.Delete()
                    '' ''    Exit For
                    '' ''End If
                    For I = LBound(dynprops) To UBound(dynprops)
                        If LCase(dynprops(I).PropertyName) = "visibility" Then
                            If dynprops(I).Value = "Not For Construction" Then

                            Else
                                dynprops(I).Value = "Issued for Information"
                                'onestamp = True
                            End If
                        End If
                    Next
                End If
            End If
        Next obj


        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If Left(LCase(blkref.EffectiveName), 9) = "hel_2026_" Then
                    attribs = blkref.GetAttributes
                    For I = LBound(attribs) To UBound(attribs)
                        If LCase(Thisdrawing.GetVariable("users1")) = "inds" Then
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    attrib.TextString = "0"
                                    revnumber = attrib.TextString
                                ElseIf attrib.TextString Like "[A-Z]" Then
                                    If attrib.TextString = "" Then
                                        intLast = 64
                                    Else
                                        intLast = Asc(attrib.TextString)
                                    End If
                                    If intLast = 90 Then intLast = 64 'If letter "z" 96+1= letter "a"
                                    intNext = intLast + 1
                                    If intNext = 73 Or intNext = 79 Then intNext += 1
                                    attrib.TextString = Chr(intNext)
                                    revnumber = attrib.TextString
                                Else
                                    attrib.TextString = Int(attrib.TextString) + 1
                                    revnumber = attrib.TextString
                                End If
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    attribs(I + 2).TextString = "INFORMATION"
                                    Exit For
                                End If
                            End If
                        Else
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    attrib.TextString = "0"
                                    revnumber = attrib.TextString
                                Else
                                    attrib.TextString = Int(attrib.TextString) + 1
                                    revnumber = attrib.TextString
                                End If
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    attribs(I + 2).TextString = "INFORMATION"
                                    Exit For
                                End If
                            End If
                        End If
                    Next I
                End If
            End If
        Next obj


        If check_routines.Laychk("~-VPORT") = False Then
            Dim vportlay As AcadLayer
            vportlay = Thisdrawing.Layers.Add("~-VPORT")
            vportlay.Plottable = False
        Else
            Thisdrawing.Layers.Item("~-VPORT").Freeze = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").LayerOn = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").Lock = vbFalse
        End If
        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadPViewport Then
                obj.Layer = "~-vport"
            End If
        Next obj

        Exit Sub

        Return

    End Sub
    <CommandMethod("IFA")> Public Sub IFA()

        Dim blkref As AcadBlockReference
        Dim revnumber As String = ""
        Dim obj As Object
        Dim attribs
        Dim attrib As AcadAttributeReference
        Dim intLast As Integer
        Dim intNext As Integer
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim dynprops

        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If LCase(blkref.EffectiveName) = "hel_stamps-2026" Then
                    dynprops = blkref.GetDynamicBlockProperties
                    '' ''If onestamp = True Then
                    '' ''    blkref.Delete()
                    '' ''    Exit For
                    '' ''End If
                    For I = LBound(dynprops) To UBound(dynprops)
                        If LCase(dynprops(I).PropertyName) = "visibility" Then
                            If dynprops(I).Value = "Not For Construction" Then

                            Else
                                dynprops(I).Value = "Issued for Approval"
                                'onestamp = True
                            End If
                        End If
                    Next
                End If
            End If
        Next obj

        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If Left(LCase(blkref.EffectiveName), 10) = "hel_border" Or Left(LCase(blkref.EffectiveName), 14) = "hel_industrial" Then
                    attribs = blkref.GetAttributes
                    For I = LBound(attribs) To UBound(attribs)
                        If LCase(Thisdrawing.GetVariable("users1")) = "inds" Then
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    intLast = 64
                                Else
                                    intLast = Asc(attrib.TextString)
                                End If
                                If intLast = 90 Then intLast = 64 'If letter "z" 96+1= letter "a"
                                intNext = intLast + 1
                                If intNext = 73 Or intNext = 79 Then intNext += 1
                                attrib.TextString = Chr(intNext)
                                revnumber = attrib.TextString
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    attribs(I + 2).TextString = "Approval"
                                    Exit For
                                End If
                            End If
                        Else
                            If LCase(attribs(I).TagString) = "rev" Or LCase(attribs(I).TagString) = "r#" Then
                                attrib = attribs(I)
                                If attrib.TextString = "" Then
                                    attrib.TextString = "0"
                                    revnumber = attrib.TextString
                                Else
                                    attrib.TextString = Int(attrib.TextString) + 1
                                    revnumber = attrib.TextString
                                End If
                            ElseIf LCase(attribs(I).TagString) Like "r#" Or LCase(attribs(I).TagString) Like "r1#" Then
                                If LCase(attribs(I).TextString) = "" Then
                                    attribs(I).TextString = revnumber
                                    If dte = "" Then dte = Format(Now, "yyyy.MM.dd")
                                    attribs(I + 1).TextString = dte
                                    attribs(I + 2).TextString = "Approval"
                                    Exit For
                                End If
                            End If
                        End If
                    Next I
                End If
            End If
        Next obj


        If check_routines.Laychk("~-VPORT") = False Then
            Dim vportlay As AcadLayer
            vportlay = Thisdrawing.Layers.Add("~-VPORT")
            vportlay.Plottable = False
        Else
            Thisdrawing.Layers.Item("~-VPORT").Freeze = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").LayerOn = vbFalse
            Thisdrawing.Layers.Item("~-VPORT").Lock = vbFalse
        End If
        For Each obj In Thisdrawing.PaperSpace
            If TypeOf obj Is AcadPViewport Then
                obj.Layer = "~-vport"
            End If
        Next obj

        Exit Sub

        Return



    End Sub

End Class
