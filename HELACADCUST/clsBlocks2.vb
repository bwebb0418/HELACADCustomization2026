Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows

Public Class ClsBlocks2
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public blk As AcadBlockReference
    Public dynprop As Object
    Public ATTRIBS
    Public existing As String
    Public utilities As New clsutilities
    Public check_routines As New ClsCheckroutines

    <CommandMethod("titleblock")> Sub Showtblkmenu()
        Dim frmtblk As New FrmTitleblock
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        frmtblk.ShowDialog()
        If frmtblk.DialogResult = System.Windows.Forms.DialogResult.OK Then
            ed.WriteMessage("Titleblock insertion complete")
        Else
            ed.WriteMessage("Titleblock insertion error")
        End If
    End Sub

    <CommandMethod("WBTag")> Sub Wood_Beam_tag()
        Dim space
        Dim inspoint, inspoint2
        Dim secpoint, secpoint2
        Dim ATTRIBS
        Dim dynprop
        Dim rotation
        Dim blkname As String


        Dim blkscl

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If
        On Error GoTo nd
        inspoint = utilities.Getpoint("Select Beam Insert Point")
        secpoint = utilities.Getendpoint(inspoint, "Select Beam End Point")
        rotation = utilities.Getrotation(inspoint, secpoint)
        blkscl = 1
        If blkscl = 0 Then blkscl = 1
        If utilities.Radtodeg(rotation) > 105 And utilities.Radtodeg(rotation) < 285 Then
            inspoint2 = inspoint
            secpoint2 = secpoint
            inspoint = secpoint2
            secpoint = inspoint2
            rotation = utilities.Getrotation(inspoint, secpoint)
        End If

        If Thisdrawing.GetVariable("Userr3") = 1 Then
            blkname = "hel_wood_beam-2026"
        Else
            blkname = "hel_wood_beam_i-2026"
        End If

        If check_routines.Blkchk(blkname) = True Then

        Else
            blkname = $"{blkname}.dwg"
        End If

        blk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

        existing = Thisdrawing.Utility.GetString(False, "New or Existing Beam? ")

        Dim lname As String

        ATTRIBS = blk.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            If LCase(Strings.Left(existing, 1)) = "e" Then
                lname = "~-TXTEX"
                ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                If check_routines.Laychk(lname) = False Then
                    Thisdrawing.Layers.Add(lname)
                End If
                ATTRIBS(i).Layer = lname
            Else
                ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                    Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
                End If
                ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
                On Error GoTo 0
            End If
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Beam Size? ")
        Next i

        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Beam Length" Then
                Dim line As AcadLine
                line = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
                dynprop(i).Value = line.Length
                line.Delete()
            End If
        Next i

        On Error Resume Next

        If LCase(Strings.Left(existing, 1)) = "e" Then
            lname = "S-WBMEX"
        Else
            lname = "S-WBM"
        End If

        If check_routines.Laychk(lname) = False Then Thisdrawing.Layers.Add(lname)
        blk.Layer = lname

        blk.Update()

nd:
    End Sub
    <CommandMethod("GRTag")> Sub Girt_tag()
        Dim space
        Dim inspoint, inspoint2
        Dim secpoint, secpoint2
        Dim ATTRIBS
        Dim dynprop
        Dim rotation
        Dim blkname As String
        Dim blkscl

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If
        On Error GoTo nd
        inspoint = utilities.Getpoint("Select Girder Truss Insert Point")
        secpoint = utilities.Getendpoint(inspoint, "Select Girder Truss End Point")
        rotation = utilities.Getrotation(inspoint, secpoint)
        blkscl = 1
        If blkscl = 0 Then blkscl = 1
        If utilities.Radtodeg(rotation) > 105 And utilities.Radtodeg(rotation) < 285 Then
            inspoint2 = inspoint
            secpoint2 = secpoint
            inspoint = secpoint2
            secpoint = inspoint2
            rotation = utilities.Getrotation(inspoint, secpoint)
        End If

        If Thisdrawing.GetVariable("Userr3") = 1 Then
            blkname = "hel_girt_tag-2026"
        Else
            blkname = "hel_girt_tag_i-2026"
        End If

        If check_routines.Blkchk(blkname) = True Then

        Else
            blkname &= ".dwg"
        End If

        existing = Thisdrawing.Utility.GetString(False, "New or Existing Girder? ")

        Dim lname As String

        blk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

        ATTRIBS = blk.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            If LCase(Left(existing, 1)) = "e" Then
                lname = "~-TXTEX"
                ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                If check_routines.Laychk(lname) = False Then
                    Thisdrawing.Layers.Add(lname)
                End If
                ATTRIBS(i).Layer = lname
            Else
                ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                    Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
                End If
                ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
                On Error GoTo 0
            End If
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Girt Size? ")
        Next i

        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Beam Length" Then
                Dim line As AcadLine
                line = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
                dynprop(i).Value = line.Length
                line.Delete()
            End If
        Next i

        On Error Resume Next

        If LCase(Left(existing, 1)) = "e" Then
            lname = "S-WTRUSSEX"
        Else
            lname = "S-WTRUSS"
        End If

        If check_routines.Laychk(lname) = False Then Thisdrawing.Layers.Add(lname)
        blk.Layer = lname

        blk.Update()
        blk.Update()

nd:
    End Sub


    <CommandMethod("breakline")> Sub Break_Line()
        Dim space
        Dim inspoint, inspoint2, secpoint
        Dim dynprop
        Dim rotation
        Dim blkname
        Dim blkscl

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
            blkscl = utilities.Getscale
        Else
            space = Thisdrawing.PaperSpace
            blkscl = Thisdrawing.GetVariable("Userr3")
        End If
        On Error GoTo nd
        inspoint = utilities.Getpoint("Select the Start of the Break Line: ")
        secpoint = utilities.Getendpoint(inspoint, "Select the End of the Break Line: ")
        rotation = utilities.Getrotation(inspoint, secpoint)
        If blkscl = 0 Then blkscl = 1

        If check_routines.Blkchk("hel_brklin-2026") = True Then
            blkname = "hel_brklin-2026"
        Else
            blkname = "hel_brklin-2026.dwg"
        End If

        blk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Break Length" Then
                Dim line As AcadLine
                line = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
                dynprop(i).Value = line.Length
                line.Delete()
            End If
        Next i

        If check_routines.Laychk("~-TSYM1") = False Then Thisdrawing.Layers.Add("~-TSYM1")
        blk.Layer = "~-TSYM1"
        blk.Update()
nd:
    End Sub

    <CommandMethod("ShearW")> Sub Shearwall()
        Dim blk As AcadBlock
        Dim blks As AcadBlocks
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim line As AcadLine
        Dim uni As Double
        Dim scl As Double
        Dim size As String
        Dim width As String
        Dim inspoint, secpoint, inspoint2, secpoint2
        Dim holddowns As String
        Dim rot, blkscl
        Dim prop As AcadDynamicBlockReferenceProperty
        Dim Sw As String
        Dim dynprop

        uni = Thisdrawing.GetVariable("userr3")
        If uni <> 1 Then
            uni = 1 / 25
            width = "4 or 6: "
        Else
            width = "100 or 150: "
        End If

        scl = Thisdrawing.GetVariable("userr2")

        blks = Thisdrawing.Blocks

sizer:
        On Error GoTo nd
        size = Thisdrawing.Utility.GetString(False, "Shearwall width? " & width)

        If size <> "4" And size <> "6" And size <> "100" And size <> "150" Then GoTo sizer

        If size = "100" Then
            size = 4
        ElseIf size = "150" Then
            size = 6
        End If
        On Error GoTo 0
        If check_routines.Blkchk("hel_shearwall" & size & "-2026") = True Then
            blkname = "hel_shearwall" & size & "-2026"
        Else
            blkname = "hel_shearwall" & size & "-2026.dwg"
        End If
        On Error GoTo nd
        inspoint = utilities.Getpoint("Select Shearwall Insertion Point: ")
        secpoint = utilities.Getendpoint(inspoint, "select Shearwall End Point: ")
        rot = utilities.Getrotation(inspoint, secpoint)
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, uni, uni, uni, rot)

        If utilities.Radtodeg(rot) > 105 And utilities.Radtodeg(rot) < 285 Then
            inspoint2 = inspoint
            secpoint2 = secpoint
            inspoint = secpoint2
            secpoint = inspoint2
            rot = utilities.Getrotation(inspoint, secpoint)
        End If

        If Right(blkname, 3) = "dwg" Then
            blkref.Delete()
            ChangeLTS()
            blkname = Left(blkname, Len(blkname) - 4)
            blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, uni, uni, uni, rot)
        End If

        holddowns = Thisdrawing.Utility.GetString(False, "Does this Shearwall Require Hold Downs? Yes or No?")
        If LCase(Left(holddowns, 1)) = "n" Then
            holddowns = "Without Hold Downs"
        Else
            holddowns = "With Hold Downs"
        End If

        dynprop = blkref.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Distance" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, secpoint)
            End If
            If dynprop(i).PropertyName = "Visibility" Then
                prop = dynprop(i)
                prop.Value = holddowns
            End If
        Next i

        'Inserts the Shearwall Tag

        If check_routines.Blkchk("hel_shearwall_tag-2026") = True Then
            blkname = "hel_shearwall_tag-2026"
        Else
            blkname = "hel_shearwall_tag-2026.dwg"
        End If

        line = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
        line.ScaleEntity(inspoint, 0.5)
        line.Rotate(line.EndPoint, utilities.Degtorad(-90))
        line.ScaleEntity(line.EndPoint, (5 * utilities.Getscale) / line.Length)

        blkscl = utilities.Getscale * Thisdrawing.GetVariable("Userr1")
        If blkscl = 0 Then blkscl = 1
        blkref = Thisdrawing.ModelSpace.InsertBlock(line.StartPoint, blkname, blkscl, blkscl, blkscl, rot)

        line.Delete()

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
            End If
            Sw = Thisdrawing.Utility.GetString(False, "Shearwall Number? ")
            If Sw = "" Then Sw = 1
            ATTRIBS(i).TextString = "SW" & Sw
            On Error GoTo 0
        Next i
nd:
    End Sub

    Function ChangeLTS() As Boolean
        Dim blk As AcadBlock
        Dim blks As AcadBlocks
        Dim line As AcadLine
        Dim uni As Double
        Dim scl As Double
        Dim size As String
        Dim obj As Object

        blks = Thisdrawing.Blocks

        uni = Thisdrawing.GetVariable("userr3")

        scl = Thisdrawing.GetVariable("userr2")

        For Each blk In blks
            If Strings.Left(blk.Name, 19) = "hel_shearwall4-2026" Or Strings.Left(blk.Name, 19) = "hel_shearwall6-2026" Then
                size = Strings.Left(blk.Name, 14)
                size = Strings.Right(size, 1)
                If uni = "1" Then
                    If scl = 50 Then
                        For Each obj In blk
                            If TypeOf obj Is AcadLine Then
                                line = obj
                                If size = 6 Then line.LinetypeScale = 0.75
                                If size = 4 Then line.LinetypeScale = 0.5
                            End If
                        Next obj
                    ElseIf scl = 100 Then
                        For Each obj In blk
                            If TypeOf obj Is AcadLine Then
                                line = obj
                                If size = 6 Then line.LinetypeScale = 0.375
                                If size = 4 Then line.LinetypeScale = 0.25
                            End If
                        Next obj
                    Else
                        For Each obj In blk
                            If TypeOf obj Is AcadLine Then
                                line = obj
                                line.LinetypeScale = 1000
                            End If
                        Next obj
                    End If
                Else
                    If scl = 48 Then
                        For Each obj In blk
                            If TypeOf obj Is AcadLine Then
                                line = obj
                                If size = 6 Then line.LinetypeScale = 0.832
                                If size = 4 Then line.LinetypeScale = 0.555
                            End If
                        Next obj
                    ElseIf scl = 64 Then
                        For Each obj In blk
                            If TypeOf obj Is AcadLine Then
                                line = obj
                                If size = 6 Then line.LinetypeScale = 0.624
                                If size = 4 Then line.LinetypeScale = 0.416
                            End If
                        Next obj
                    ElseIf scl = 96 Then
                        For Each obj In blk
                            If TypeOf obj Is AcadLine Then
                                line = obj
                                If size = 6 Then line.LinetypeScale = 0.416
                                If size = 4 Then line.LinetypeScale = 0.277
                            End If
                        Next obj
                    Else
                        For Each obj In blk
                            If TypeOf obj Is AcadLine Then
                                line = obj
                                line.LinetypeScale = 10000000
                            End If
                        Next obj
                    End If

                End If
            End If
        Next blk


        ChangeLTS = True

    End Function

    <CommandMethod("Joistlayout")> Sub Joist_layout()
        Dim space, jostins, joistsec, fillins, fillsec
        Dim ATTRIBS, dynprop, rotation, blkname
        Dim blkscl, blkref
        Dim osmode As Integer
        Dim joistline As AcadLine
        Dim fillline As AcadLine
        Dim inspoint, joistins, fillins2, fillsec2, joistins2, joistsec2
        Dim material As String
        Dim prop As AcadDynamicBlockReferenceProperty


        material = Thisdrawing.Utility.GetString(False, "Joist Material? ")

        osmode = Thisdrawing.GetVariable("osmode")

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If
        On Error GoTo nd
        Thisdrawing.SetVariable("osmode", 512)
        joistins = utilities.Getpoint("Select 1st Point of Joist Span: ")
        Thisdrawing.SetVariable("osmode", 128)
        joistsec = utilities.Getendpoint(joistins, "Select 2nd Point of Joist Span: ")

        joistins(2) = 0 : joistsec(2) = 0

        joistline = space.AddLine(joistins, joistsec)

        On Error Resume Next
        Thisdrawing.Linetypes.Load("Phantom", "acad.lin")
        On Error GoTo 0

        joistline.Linetype = "phantom"
        joistline.color = ACAD_COLOR.acRed

        Thisdrawing.SetVariable("osmode", 512)
        fillins = utilities.Getpoint("Select 1st Point of Joist Spacing: ")
        Thisdrawing.SetVariable("osmode", 128)
        fillsec = utilities.Getendpoint(fillins, "Select 2nd Point of Joist spacing: ")

        fillsec(2) = 0 : fillins(2) = 0

        fillline = space.AddLine(fillins, fillsec)

        inspoint = joistline.IntersectWith(fillline, AcExtendOption.acExtendBoth)
        blkscl = utilities.Getscale
        rotation = utilities.Getrotation(fillins, fillsec)

        If utilities.Radtodeg(rotation) > 105 And utilities.Radtodeg(rotation) < 285 Then
            fillins2 = fillins
            fillsec2 = fillsec
            fillins = fillsec2
            fillsec = fillins2
            rotation = utilities.Getrotation(fillins, fillsec)
        End If

        If utilities.Radtodeg(joistline.Angle) > 195 Then
            joistins2 = joistins
            joistsec2 = joistsec
            joistins = joistsec2
            joistsec = joistins2
        End If

        If utilities.Radtodeg(joistline.Angle) <= 15 Then
            joistins2 = joistins
            joistsec2 = joistsec
            joistins = joistsec2
            joistsec = joistins2
        End If

        If Thisdrawing.GetVariable("Userr3") = 1 Then
            If check_routines.Blkchk("hel_" & material & "_joist-2026") = True Then
                blkname = "hel_" & material & "_joist-2026"
            Else
                blkname = "hel_" & material & "_joist-2026.dwg"
            End If
        Else
            If check_routines.Blkchk("hel_" & material & "_joist_i-2026") = True Then
                blkname = "hel_" & material & "_joist_i-2026"
            Else
                blkname = "hel_" & material & "_joist_i-2026.dwg"
            End If
        End If


        blkref = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

        dynprop = blkref.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Span1" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, joistsec)
            End If
            If dynprop(i).PropertyName = "Span2" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, joistins)
            End If
            If dynprop(i).PropertyName = "Fill1" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, fillsec)
            End If
            If dynprop(i).PropertyName = "Fill2" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, fillins)
            End If
        Next i

        fillline.Delete()
        joistline.Delete()

        existing = Thisdrawing.Utility.GetString(False, "New or Existing Joist? ")

        Dim lname As String
        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle

            If LCase(Left(existing, 1)) = "e" Then
                lname = "~-TXTEX"
                ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                If check_routines.Laychk(lname) = False Then
                    Thisdrawing.Layers.Add(lname)
                End If
                ATTRIBS(i).Layer = lname
            Else
                ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                    Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
                End If
                ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
                On Error GoTo 0
            End If
        Next i

        On Error Resume Next

        If material = "wood" Then
            If LCase(Left(existing, 1)) = "e" Then
                lname = "S-WJSTEX"
            Else
                lname = "S-WJST"
            End If
        ElseIf material = "steel" Then
            If LCase(Left(existing, 1)) = "e" Then
                lname = "S-SJSTEX"
            Else
                lname = "S-SJST"
            End If
        Else
            lname = "0"
        End If

        If check_routines.Laychk(lname) = False Then Thisdrawing.Layers.Add(lname)
        blkref.Layer = lname

        Thisdrawing.SendCommand("attedit" & vbCr & "l" & vbCr)

        On Error Resume Next
        Thisdrawing.Linetypes.Item("Phantom").Delete()

nd:

        Thisdrawing.SetVariable("osmode", osmode)
    End Sub

    <CommandMethod("Post")> Sub Post()

        Dim blks As AcadBlocks
        Dim blkref As AcadBlockReference
        Dim blkname

        Dim uni As Double
        Dim scl As Double
        Dim size As String
        Dim inspoint

        Dim prop As AcadDynamicBlockReferenceProperty
        Dim ATTRIBS

        Dim width As String
        Try
            uni = Thisdrawing.GetVariable("userr3")
            'If uni <> 1 Then uni = 1 / 25

            If uni <> 1 Then
                uni = 1
                width = "4,6,8,10 or 12: "
            Else
                width = "100,150,200,250 or 300: "
            End If

            scl = Thisdrawing.GetVariable("userr2")

sizer:

            size = Thisdrawing.Utility.GetString(False, "Post size? " & width)

            If size <> "4" And size <> "6" And size <> "100" And size <> "150" And size <> "8" And
            size <> "10" And size <> "200" And size <> "250" And size <> "12" And size <> "300" Then GoTo sizer

            If size > 20 Then
                size /= 25
            ElseIf size < 20 Then
                size &= "_i"
            End If

            blks = Thisdrawing.Blocks

            If check_routines.Blkchk("hel_post" & size & "-2026") = True Then
                blkname = "hel_post" & size & "-2026"
            Else
                blkname = "hel_post" & size & "-2026.dwg"
            End If

            inspoint = utilities.Getpoint("Select Post Insertion Point: ")

            blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, uni, uni, uni, 0)

            Dim posttype As String
            Dim criptype As String
            Dim viz As Boolean
            posttype = Thisdrawing.Utility.GetString(False, "Post Type? ")
            criptype = Thisdrawing.Utility.GetString(False, "Cripple Type? ")
            If posttype = "" And criptype = "" Then viz = False


            ATTRIBS = blkref.GetAttributes
            For i = LBound(ATTRIBS) To UBound(ATTRIBS)
                Try
                    ATTRIBS(i).StyleName = utilities.Gettxtstyle
                    ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                    If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                        Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
                    End If
                    ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
                    If ATTRIBS(i).TagString = "PTAG" Then
                        ATTRIBS(i).TextString = posttype
                    Else
                        ATTRIBS(i).TextString = criptype
                    End If
                Catch

                End Try

            Next i

            dynprop = blkref.GetDynamicBlockProperties
            For i = LBound(dynprop) To UBound(dynprop)
                Try
                    If dynprop(i).PropertyName = "Distance" Then
                        prop = dynprop(i)

                        If Right(size, 2) = "_i" Then size = Left(size, Len(size) - 2)
                        If size > (5 * Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height) Then
                            prop.Value = size + (2 * Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height)
                        Else
                            prop.Value = (5 * Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height)
                        End If
                    ElseIf dynprop(i).propertyname = "Visibility1" Then
                        prop = dynprop(i)
                        If viz = True Then
                            prop.Value = "True"
                        Else
                            prop.Value = "False"
                        End If
                    End If
                Catch

                End Try
            Next i


        Catch

        End Try

    End Sub

    Function Quick_Steel_Beam_tag()
        Dim space
        Dim inspoint, inspoint2
        Dim secpoint, secpoint2
        Dim ATTRIBS
        Dim dynprop
        Dim rotation
        Dim blkname As String
        Dim blkscl
        Dim userinput As String
        Dim blkref As AcadBlockReference

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If
        On Error GoTo nd
        inspoint = utilities.Getpoint("Select Beam Insert Point: ")
        secpoint = utilities.Getendpoint(inspoint, "Select Beam End Point: ")
        rotation = utilities.Getrotation(inspoint, secpoint)
        blkscl = 1
        If blkscl = 0 Then blkscl = 1
        If utilities.Radtodeg(rotation) > 105 And utilities.Radtodeg(rotation) < 285 Then
            inspoint2 = inspoint
            secpoint2 = secpoint
            inspoint = secpoint2
            secpoint = inspoint2
            rotation = utilities.Getrotation(inspoint, secpoint)
        End If

        If Thisdrawing.GetVariable("Userr3") = 1 Then
            blkname = "hel_steel_beam-2026"
        Else
            blkname = "hel_steel_beam_i-2026"
        End If

        If check_routines.Blkchk(blkname) = True Then

        Else
            blkname &= ".dwg"
        End If

        blk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

        Dim lname As String

        ATTRIBS = blk.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
            On Error GoTo 0
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Beam Size? ")
        Next i

        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Beam Length" Then
                Dim line As AcadLine
                line = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
                dynprop(i).Value = line.Length
                line.Delete()
            End If
        Next i

        On Error Resume Next

        lname = "S-SBM"

        If check_routines.Laychk(lname) = False Then Thisdrawing.Layers.Add(lname)
        blk.Layer = lname

        blk.Update()
        blk.Update()

nd:
    End Function

    Function Steel_Beam_tag()
        Dim space
        Dim inspoint, inspoint2
        Dim secpoint, secpoint2
        Dim ATTRIBS
        Dim dynprop
        Dim rotation
        Dim blkname As String
        Dim blkscl
        Dim userinput As String
        Dim line1 As AcadLine
        Dim line2 As AcadLine
        Dim blkref As AcadBlockReference

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If
        On Error GoTo nd
        inspoint = utilities.Getpoint("Select Beam Insert Point: ")
        secpoint = utilities.Getendpoint(inspoint, "Select Beam End Point: ")
        rotation = utilities.Getrotation(inspoint, secpoint)
        blkscl = 1
        If blkscl = 0 Then blkscl = 1
        If utilities.Radtodeg(rotation) > 105 And utilities.Radtodeg(rotation) < 285 Then
            inspoint2 = inspoint
            secpoint2 = secpoint
            inspoint = secpoint2
            secpoint = inspoint2
            rotation = utilities.Getrotation(inspoint, secpoint)
        End If

        If Thisdrawing.GetVariable("Userr3") = 1 Then
            blkname = "hel_steel_beam-2026"
        Else
            blkname = "hel_steel_beam_i-2026"
        End If

        If check_routines.Blkchk(blkname) = True Then

        Else
            blkname &= ".dwg"
        End If

        blk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

        existing = Thisdrawing.Utility.GetString(False, "New or Existing Beam? ")

        Dim lname As String


        ATTRIBS = blk.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            If LCase(Left(existing, 1)) = "e" Then
                lname = "~-TXTEX"
                ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                If check_routines.Laychk(lname) = False Then
                    Thisdrawing.Layers.Add(lname)
                End If
                ATTRIBS(i).Layer = lname
            Else
                ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                    Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
                End If
                ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
                On Error GoTo 0
            End If
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Beam Size? ")
        Next i

        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Beam Length" Then
                Dim line As AcadLine
                line = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
                dynprop(i).Value = line.Length
                line.Delete()
            End If
        Next i

        On Error Resume Next

        If LCase(Left(existing, 1)) = "e" Then
            lname = "S-SBMEX"
        Else
            lname = "S-SBM"
        End If

        If check_routines.Laychk(lname) = False Then Thisdrawing.Layers.Add(lname)
        blk.Layer = lname

        blk.Update()
        blk.Update()

        On Error Resume Next
        userinput = Thisdrawing.Utility.GetString(False, "Is This a Moment Connection? Yes or No? ")
        If LCase(Left(userinput, 1)) = "n" Then
            Exit Function
        End If

        line1 = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
        line2 = Thisdrawing.ModelSpace.AddLine(secpoint, inspoint)

        line1.ScaleEntity(line1.StartPoint, (250 * Thisdrawing.GetVariable("Userr3")) / line1.Length)
        line2.ScaleEntity(line2.StartPoint, (250 * Thisdrawing.GetVariable("Userr3")) / line2.Length)
        line1.Visible = False
        line2.Visible = False

        If check_routines.Blkchk("hel_moment_tag-2026") = True Then
            blkname = "hel_moment_tag-2026"
        Else
            blkname = "hel_moment_tag-2026.dwg"
        End If

        blkscl = utilities.Getscale * Thisdrawing.GetVariable("userr1")
        If blkscl = 0 Then blkscl = 1

        blkref = Thisdrawing.ModelSpace.InsertBlock(line1.EndPoint, blkname, blkscl, blkscl, blkscl, line1.Angle)

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
            If ATTRIBS(i).Tag = "FORCE" Then
                ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Moment Value: ") & "kN"
            End If
            On Error GoTo 0
        Next i

        blkref = blkref.Copy
        blkref.Move(line1.EndPoint, line2.EndPoint)


        line1.Delete()
        line2.Delete()
nd:
    End Function

    Function Steel_Beam_Rec()
        Dim space
        Dim inspoint, inspoint2
        Dim secpoint, secpoint2
        Dim ATTRIBS
        Dim dynprop
        Dim rotation
        Dim blkname
        Dim blkscl
        Dim userinput As String
        Dim line1 As AcadLine
        Dim line2 As AcadLine
        Dim blkref As AcadBlockReference
        Dim depth, shape

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If
        On Error GoTo nd
        inspoint = utilities.Getpoint("Select Beam Insert Point: ")
        secpoint = utilities.Getendpoint(inspoint, "Select Beam End Point: ")
        rotation = utilities.Getrotation(inspoint, secpoint)
        blkscl = Thisdrawing.GetVariable("Userr3")
        If blkscl = 0 Then blkscl = 1

        If utilities.Radtodeg(rotation) > 105 And utilities.Radtodeg(rotation) < 285 Then
            inspoint2 = inspoint
            secpoint2 = secpoint
            inspoint = secpoint2
            secpoint = inspoint2
            rotation = utilities.Getrotation(inspoint, secpoint)
        End If

        If check_routines.Blkchk("hel_steel_beam_dbl-2026") = True Then
            blkname = "hel_steel_beam_dbl-2026"
        Else
            blkname = "hel_steel_beam_dbl-2026.dwg"
        End If

        blk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

shape:
        shape = Thisdrawing.Utility.GetString(False, "Beam Shape: ")
        shape = UCase(shape)
        If shape <> "S" And shape <> "W" And shape <> "H" And shape <> "HSS" And shape <> "L" Then
            MsgBox("Please Input Shape in the following format:" & vbCr & "S, W, H, HSS or L", vbOKOnly, "Error")
            GoTo shape
        End If
        On Error GoTo nd
        depth = CDbl(Thisdrawing.Utility.GetReal("Specify Beam Depth: "))

        existing = Thisdrawing.Utility.GetString(False, "New or Existing Beam? ")

        Dim lname As String

        ATTRIBS = blk.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            If LCase(Left(existing, 1)) = "e" Then
                lname = "~-TXTEX"
                ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                If check_routines.Laychk(lname) = False Then
                    Thisdrawing.Layers.Add(lname)
                End If
                ATTRIBS(i).Layer = lname
            Else
                ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                    Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
                End If
                ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
                On Error GoTo 0
            End If
        Next i

        dynprop = blk.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Beam Length" Then
                Dim line As AcadLine
                line = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
                dynprop(i).Value = line.Length
                line.Delete()
            End If
            If dynprop(i).PropertyName = "Distance" Then
                dynprop(i).Value = depth
            End If
        Next i

        On Error Resume Next

        If LCase(Left(existing, 1)) = "e" Then
            lname = "S-SBMEX"
        Else
            lname = "S-SBM"
        End If

        If check_routines.Laychk(lname) = False Then Thisdrawing.Layers.Add(lname)
        blk.Layer = lname

        blk.Update()
        blk.Update()

        On Error Resume Next
        userinput = Thisdrawing.Utility.GetString(False, "Is This a Moment Connection? Yes or No? ")
        If LCase(Left(userinput, 1)) = "n" Then
            Exit Function
        End If

        line1 = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
        line2 = Thisdrawing.ModelSpace.AddLine(secpoint, inspoint)

        line1.ScaleEntity(line1.StartPoint, (250 * blkscl) / line1.Length)
        line2.ScaleEntity(line2.StartPoint, (250 * blkscl) / line2.Length)
        line1.Visible = False
        line2.Visible = False

        If check_routines.Blkchk("hel_moment_tag-2026") = True Then
            blkname = "hel_moment_tag-2026"
        Else
            blkname = "hel_moment_tag-2026.dwg"
        End If

        blkscl = utilities.Getscale * Thisdrawing.GetVariable("userr1")
        If blkscl = 0 Then blkscl = 1

        blkref = Thisdrawing.ModelSpace.InsertBlock(line1.EndPoint, blkname, blkscl, blkscl, blkscl, line1.Angle)

        dynprop = blkref.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Distance" Then
                dynprop(i).Value = depth
            End If
        Next i


        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
            If ATTRIBS(i).Tag = "FORCE" Then
                ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Moment Value: ") & "kN"
            End If
            On Error GoTo 0
        Next i

        blkref = blkref.Copy
        blkref.Move(line1.EndPoint, line2.EndPoint)

nd:

        line1.Delete()
        line2.Delete()
    End Function

    <CommandMethod("SBTag")> Sub Steelbeam()
        On Error Resume Next
        If LCase(Left(Thisdrawing.Utility.GetString(False, "Single Line Beam or Outline of Beam [S/O]? "), 1)) <> "o" Then
            Call Steel_Beam_tag()
            Exit Sub
        Else
            Call Steel_Beam_Rec()
        End If
    End Sub
    <CommandMethod("QSBeam")> Sub Quicksteelbeam()
        Call Quick_Steel_Beam_tag()
    End Sub

    <CommandMethod("recTag")> Sub Rec_tag()
        Dim inspoint
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim blkscl
        Dim maxpt = Nothing
        Dim minpt = Nothing
        Dim ATTRIBS

        blkscl = utilities.Getscale
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.Getpoint("Select Tag Insertion Point: ")

        If check_routines.Blkchk("hel_rec_tag-2026") = True Then
            blkname = "hel_rec_tag-2026"
        Else
            blkname = "hel_rec_tag-2026.dwg"
        End If
        On Error GoTo nd
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Tag Value: ")
            ATTRIBS(i).GetBoundingBox(minpt, maxpt)
            On Error GoTo 0
        Next i

        maxpt(1) = minpt(1)

        dynprop = blkref.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Distance" Then
                dynprop(i).Value = utilities.Getlength(minpt, maxpt) / 2
            ElseIf dynprop(i).PropertyName = "Distance1" Then
                dynprop(i).Value = utilities.Getlength(minpt, maxpt) / 2
            End If
        Next i

nd:

    End Sub

    <CommandMethod("ElevDTag")> Sub Elev_diff()
        Dim inspoint
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim blkscl
        Dim maxpt = Nothing
        Dim minpt = Nothing
        Dim ATTRIBS

        blkscl = utilities.Getscale
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.Getpoint("Select Tag Insertion Point: ")

        If check_routines.Blkchk("hel_elvdiff-2026") = True Then
            blkname = "hel_elvdiff-2026"
        Else
            blkname = "hel_elvdiff-2026.dwg"
        End If
        On Error GoTo nd
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        dynprop = blkref.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Visibility" Then
                If Thisdrawing.GetVariable("users1") = "inds" Then
                    dynprop(i).Value = "Industrial"
                Else
                    dynprop(i).Value = "Buildings"
                End If
            End If
        Next i

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
            ATTRIBS(i).TextString = Thisdrawing.Utility.GetString(False, "Input Elevation Difference: ")
            ATTRIBS(i).GetBoundingBox(minpt, maxpt)
            On Error GoTo 0
        Next i

        maxpt(1) = minpt(1)


nd:
    End Sub

    <CommandMethod("drainbearing")> Sub Drain_bearing()
        Dim inspoint
        Dim blkref As AcadBlockReference
        Dim blkname
        Dim blkscl
        Dim options
        Dim dynprop

        options = Thisdrawing.Utility.GetString(False)

        blkscl = Thisdrawing.GetVariable("Userr3")
        If blkscl = 0 Then blkscl = 1

        inspoint = utilities.Getpoint("Select Bottom Corner of Footing: ")

        If check_routines.Blkchk("hel_drain_bearing-2026") = True Then
            blkname = "hel_drain_bearing-2026"
        Else
            blkname = "hel_drain_bearing-2026.dwg"
        End If
        On Error GoTo nd
        blkref = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

        dynprop = blkref.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Visibility" Then
                If options = "y" Then
                    dynprop(i).Value = "Drain"
                Else
                    dynprop(i).Value = "No Drain"
                End If
            End If
        Next i
nd:
    End Sub

    <CommandMethod("spanlayout")> Sub Span_layout()
        Dim space, Spanins, Spansec, fillins, fillsec
        Dim ATTRIBS, dynprop, rotation, blkname
        Dim blkscl, blkref
        Dim osmode As Integer
        Dim Spanline As AcadLine
        Dim fillline As AcadLine
        Dim inspoint, fillsec2, spanins2, fillins2, spansec2
        Dim loc As String = ""
        Dim material As String = ""
        Dim objtype As String = ""
        Dim prop As AcadDynamicBlockReferenceProperty
        Dim existing As String = ""
        On Error GoTo nd1
pickobj:
        objtype = Thisdrawing.Utility.GetString(False, "Joist or Rebar? ")
        If LCase(Left(objtype, 1)) <> "j" And LCase(Left(objtype, 1)) <> "r" Then
            GoTo pickobj
        ElseIf LCase(Left(objtype, 1)) = "j" Then
            GoTo joistmat
        End If

location:
        loc = Thisdrawing.Utility.GetString(False, "Top or Bottom Mat? ")
        If LCase(Left(loc, 1)) <> "t" And LCase(Left(loc, 1)) <> "b" Then
            GoTo location
        Else
            GoTo ins
        End If

joistmat:
        material = Thisdrawing.Utility.GetString(False, "Steel or Wood joists? ")
        existing = Thisdrawing.Utility.GetString(False, "New or Existing Joist? ")
        If LCase(Left(material, 1)) <> "s" And LCase(Left(material, 1)) <> "w" Then GoTo joistmat

ins:

        osmode = Thisdrawing.GetVariable("osmode")

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If
        On Error GoTo nd1
        Thisdrawing.SetVariable("osmode", 512)
        Spanins = utilities.Getpoint("Select 1st Point of Span: ")
        Thisdrawing.SetVariable("osmode", 128)
        Spansec = utilities.Getendpoint(Spanins, "Select 2nd Point of Span: ")
        Spanline = space.AddLine(Spanins, Spansec)

        On Error Resume Next
        Thisdrawing.Linetypes.Load("Phantom", "acad.lin")
        On Error GoTo 0

        Spanline.Linetype = "phantom"
        Spanline.color = ACAD_COLOR.acRed
        On Error GoTo eraseline
        Thisdrawing.SetVariable("osmode", 512)
        fillins = utilities.Getpoint("Select 1st Point of Extents: ")
        Thisdrawing.SetVariable("osmode", 128)
        fillsec = utilities.Getendpoint(fillins, "Select 2nd Point of Extents: ")
        fillline = space.AddLine(fillins, fillsec)
        GoTo good
eraseline:

        Spanline.Delete()
        GoTo nd1

good:
        inspoint = Spanline.IntersectWith(fillline, AcExtendOption.acExtendBoth)
        blkscl = utilities.Getscale
        rotation = utilities.Getrotation(fillins, fillsec)

        If utilities.Radtodeg(rotation) > 105 And utilities.Radtodeg(rotation) < 285 Then
            fillins2 = fillins
            fillsec2 = fillsec
            fillins = fillsec2
            fillsec = fillins2
            rotation = utilities.Getrotation(fillins, fillsec)
        End If

        If utilities.Radtodeg(Spanline.Angle) > 195 Then
            spanins2 = Spanins
            spansec2 = Spansec
            Spanins = spansec2
            Spansec = spanins2
        End If

        If utilities.Radtodeg(Spanline.Angle) <= 15 Then
            spanins2 = Spanins
            spansec2 = Spansec
            Spanins = spansec2
            Spansec = spanins2
        End If

        If check_routines.Blkchk("hel_span_layout-2026") = True Then
            blkname = "hel_span_layout-2026"
        Else
            blkname = "hel_span_layout-2026.dwg"
        End If

        blkref = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

        dynprop = blkref.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Span1" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, Spansec)
            End If
            If dynprop(i).PropertyName = "Span2" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, Spanins)
            End If
            If dynprop(i).PropertyName = "Fill1" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, fillsec)
            End If
            If dynprop(i).PropertyName = "Fill2" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, fillins)
            End If
            If dynprop(i).PropertyName = "Visibility" Then
                prop = dynprop(i)
                If LCase(Left(objtype, 1)) = "j" Then
                    prop.Value = "Joist"
                Else
                    prop.Value = "Rebar"
                End If
            End If
        Next i

        fillline.Delete()
        Spanline.Delete()

        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle
            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
            On Error GoTo 0
        Next i

        Thisdrawing.SendCommand("attedit ")
        Thisdrawing.SendCommand("l ")

        On Error Resume Next
        Thisdrawing.Linetypes.Item("Phantom").Delete()

        If LCase(Left(objtype, 1)) = "r" Then
            GoTo rebar_layers
        ElseIf LCase(Left(objtype, 1)) = "j" Then
            GoTo Joist_layers
        End If

Joist_layers:



        If LCase(Left(material, 1)) = "w" Then
            If LCase(Left(existing, 1)) = "e" Then
                blkref.Layer = "S-WJSTEX"
            Else
                blkref.Layer = "S-WJST"
            End If
        ElseIf LCase(Left(material, 1)) = "s" Then
            If LCase(Left(existing, 1)) = "e" Then
                blkref.Layer = "S-SJSTEX"
            Else
                blkref.Layer = "S-SJST"
            End If
        Else
            blkref.Layer = "0"
        End If

        GoTo nd1

rebar_layers:
        If LCase(Left(loc, 1)) = "t" Then
            If check_routines.Laychk("S-RCONT") = False Then Thisdrawing.Layers.Add("S-RCONT")
            blkref.Layer = "S-RCONT"
            GoTo nd1
        Else
            If check_routines.Laychk("S-RBOT") = False Then Thisdrawing.Layers.Add("S-RBOT")
            blkref.Layer = "S-RBOT"
            GoTo nd1
        End If

nd1:
        Thisdrawing.SetVariable("osmode", osmode)
    End Sub



    <CommandMethod("rebarlayout")> Sub Rebar_layout()
        Dim space, jostins, joistsec, fillins, fillsec
        Dim ATTRIBS, dynprop, rotation, blkname
        Dim blkscl, blkref
        Dim osmode As Integer
        Dim joistline As AcadLine
        Dim fillline As AcadLine
        Dim inspoint, joistins, fillsec2, fillins2, joistsec2, joistins2
        Dim material As String
        Dim prop As AcadDynamicBlockReferenceProperty


        osmode = Thisdrawing.GetVariable("osmode")

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
        Else
            space = Thisdrawing.PaperSpace
        End If
        On Error GoTo nd
        Thisdrawing.SetVariable("osmode", 512)
        joistins = utilities.Getpoint("Select 1st Point of Rebar Extents: ")
        Thisdrawing.SetVariable("osmode", 128)
        joistsec = utilities.Getendpoint(joistins, "Select 2nd Point of rebar Extents: ")

        joistins(2) = 0 : joistsec(2) = 0

        joistline = space.AddLine(joistins, joistsec)

        On Error Resume Next
        Thisdrawing.Linetypes.Load("Phantom", "acad.lin")
        On Error GoTo 0

        joistline.Linetype = "phantom"
        joistline.color = ACAD_COLOR.acRed

        Thisdrawing.SetVariable("osmode", 512)
        fillins = utilities.Getpoint("Select 1st Point of Rebar Spacing: ")
        Thisdrawing.SetVariable("osmode", 128)
        fillsec = utilities.Getendpoint(fillins, "Select 2nd Point of Rebar spacing: ")

        fillsec(2) = 0 : fillins(2) = 0

        fillline = space.AddLine(fillins, fillsec)

        inspoint = joistline.IntersectWith(fillline, AcExtendOption.acExtendBoth)
        blkscl = utilities.Getscale
        rotation = utilities.Getrotation(fillins, fillsec)

        If utilities.Radtodeg(rotation) > 105 And utilities.Radtodeg(rotation) < 285 Then
            fillins2 = fillins
            fillsec2 = fillsec
            fillins = fillsec2
            fillsec = fillins2
            rotation = utilities.Getrotation(fillins, fillsec)
        End If

        If utilities.Radtodeg(joistline.Angle) > 195 Then
            joistins2 = joistins
            joistsec2 = joistsec
            joistins = joistsec2
            joistsec = joistins2
        End If

        If utilities.Radtodeg(joistline.Angle) <= 15 Then
            joistins2 = joistins
            joistsec2 = joistsec
            joistins = joistsec2
            joistsec = joistins2
        End If

        If Thisdrawing.GetVariable("Userr3") = 1 Then
            If check_routines.Blkchk("hel_rebarspc-2026") = True Then
                blkname = "hel_rebarspc-2026"
            Else
                blkname = "hel_rebarspc-2026.dwg"
            End If
        Else
            If check_routines.Blkchk("hel_rebarspc_i") = True Then
                blkname = "hel_rebarspc_i"
            Else
                blkname = "hel_rebarspc_i.dwg"
            End If
        End If


        blkref = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, rotation)

        dynprop = blkref.GetDynamicBlockProperties
        For i = LBound(dynprop) To UBound(dynprop)
            If dynprop(i).PropertyName = "Span1" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, joistsec)
            End If
            If dynprop(i).PropertyName = "Span2" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, joistins)
            End If
            If dynprop(i).PropertyName = "Fill1" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, fillsec)
            End If
            If dynprop(i).PropertyName = "Fill2" Then
                prop = dynprop(i)
                prop.Value = utilities.Getlength(inspoint, fillins)
            End If
        Next i

        fillline.Delete()
        joistline.Delete()

        Dim lname As String
        ATTRIBS = blkref.GetAttributes
        For i = LBound(ATTRIBS) To UBound(ATTRIBS)
            On Error Resume Next
            ATTRIBS(i).StyleName = utilities.Gettxtstyle

            ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
            If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
            End If
            ATTRIBS(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
            On Error GoTo 0

        Next i

        On Error Resume Next


        lname = "S-RCONT"


        If check_routines.Laychk(lname) = False Then Thisdrawing.Layers.Add(lname)
        blkref.Layer = lname

        Thisdrawing.SendCommand("attedit" & vbCr & "l" & vbCr)

        On Error Resume Next
        Thisdrawing.Linetypes.Item("Phantom").Delete()

nd:
        Thisdrawing.SetVariable("osmode", osmode)
    End Sub

    <CommandMethod("ScaleBar")> Public Sub InsertScaleBar()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction
        Dim blockname As String
        Dim check_routines As New ClsCheckroutines
        Dim blockinsert As New ClsBlock_insert

        Using tr
            If Application.GetSystemVariable("MEASUREMENT") <> 1 Then
                ed.WriteMessage("Currently this routine only applies to Metric scales")
                Exit Sub
            End If

            Dim dimscl

            If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
                dimscl = 1
            Else
                dimscl = 1 / Application.GetSystemVariable("Userr2")
            End If

            If dimscl = 0 Then dimscl = 1

            blockname = "hel_scalebar_1-" & Format(Application.GetSystemVariable("Userr2"), "000") & "-2026"

            If check_routines.Blkchk(blockname) = True Then

            Else
                If blockinsert.Loadblock(blockname & ".dwg") = False Then
                    ed.WriteMessage("Scale Bar for current scale does not exist, contact system administrator")
                    Exit Sub
                End If
            End If
            Try
                Dim objid As ObjectId = blockinsert.InsertBlockWithJig(blockname, False, "Scalebar insertion point: ", False, dimscl)
                tr.Commit()
            Catch
                ed.WriteMessage("Scale Bar for current scale does not exist, contact system administrator")
                tr.Abort()

            End Try

        End Using
        tr.Dispose()
    End Sub

    <CommandMethod("NorthArrow")> Public Sub Northarrow()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction
        Dim blockname As String
        Dim check_routines As New ClsCheckroutines
        Dim blockinsert As New ClsBlock_insert
        Dim blkscl As New Clsutilities

        Using tr
            Try
                blockname = "hel_northarrow-2026"

                If check_routines.Blkchk(blockname) = True Then

                Else
                    blockinsert.Loadblock(blockname & ".dwg")
                End If

                Dim pko As New PromptKeywordOptions("Drawing size? ") With {
                    .AllowNone = False,
                    .AllowArbitraryInput = False
                }
                pko.Keywords.Add("Full Size")
                pko.Keywords.Add("Half Size")
                pko.Keywords.Add("Key Plan")
                Dim pr As PromptResult = ed.GetKeywords(pko)

                If LCase(Strings.Left(pr.StringResult, 1)) = "f" Then
                    ClsBlock_insert.BLKSFFACT = 1
                ElseIf LCase(Strings.Left(pr.StringResult, 1)) = "h" Then
                    ClsBlock_insert.BLKSFFACT = 0.75
                Else
                    ClsBlock_insert.BLKSFFACT = 0.5
                End If
                blockinsert.InsertBlockWithJig(blockname, False, "North Arrow Insertion point? ", False, blkscl.Getscale)
                tr.Commit()
            Catch ex As Exception
                tr.Abort()
            End Try
            tr.Dispose()

        End Using
        ClsBlock_insert.BLKSFFACT = Nothing
    End Sub

    <CommandMethod("clmark")> Sub Clmark()

        Dim blkname As String = "hel_clmark-2026"
        Dim blkscl As New Clsutilities
        'Dim prop As AcadDynamicBlockReferenceProperty
        Dim blockinsert As New ClsBlock_insert
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction
        Try
            Using tr

                If check_routines.Blkchk(blkname) = True Then

                Else
                    If blockinsert.Loadblock(blkname & ".dwg") = False Then
                        Exit Sub
                    End If
                End If

                Dim objid As ObjectId = blockinsert.InsertBlockWithJig(blkname, False, "Centerline Insertion Point: ", False, blkscl.Getscale)
                If objid = Nothing Then Exit Sub

                tr.Commit()

                tr = db.TransactionManager.StartTransaction

                Dim clmarkblk As BlockReference = tr.GetObject(objid, OpenMode.ForWrite)

                clmarkblk.LayerId = check_routines.Layercheck("~-TSYM3")

                Dim attcol As AttributeCollection = clmarkblk.AttributeCollection

                Dim pso As New PromptStringOptions("First Line of Centerline Mark text: ") With {
                    .AllowSpaces = True
                }

                Dim pss1 As PromptResult = ed.GetString(pso)
                Dim pss2 As PromptResult
                If pss1.Status = PromptStatus.OK Then
                    pso.Message = "Second line of Centerline Mark text: "
                    pss2 = ed.GetString(pso)
                Else
                    clmarkblk.Erase()
                    tr.Commit()
                    Exit Sub
                End If

                For Each attid As ObjectId In attcol
                    Dim attref As AttributeReference = tr.GetObject(attid, OpenMode.ForWrite)

                    If attref.Tag = "COL." And pss2.StringResult = "" Then
                        attref.TextString = pss1.StringResult
                    ElseIf attref.Tag = "COL." And pss2.StringResult <> "" And pss1.StringResult = "" Then
                        attref.TextString = pss2.StringResult
                        Exit For
                    ElseIf attref.Tag = "CL2-1" And pss2.StringResult <> "" Then
                        attref.TextString = pss1.StringResult
                    ElseIf attref.Tag = "CL2-2" And pss2.StringResult <> "" Then
                        attref.TextString = pss2.StringResult
                    End If
                    attref.LayerId = check_routines.Layercheck("~-TXT" & Strings.Right(utilities.Gettxtstyle, 2))
                Next

                Dim defcol As DynamicBlockReferencePropertyCollection = clmarkblk.DynamicBlockReferencePropertyCollection
                For Each dbpid As DynamicBlockReferenceProperty In defcol

                    If dbpid.PropertyName = "Visibility" Then
                        If pss2.StringResult = "" Or pss1.StringResult = "" Then
                            dbpid.Value = "1-LINE"

                        Else
                            dbpid.Value = "2-LINE"
                        End If
                    ElseIf dbpid.PropertyName = "Angle" Then

                    End If

                Next

                tr.Commit()

                tr = db.TransactionManager.StartTransaction
                clmarkblk = tr.GetObject(objid, OpenMode.ForWrite)
                Dim pao As New PromptAngleOptions("Select a point along the Centerline: ") With {
                    .BasePoint = clmarkblk.Position,
                    .UseBasePoint = True,
                    .UseDashedLine = True
                }

                Dim par As PromptDoubleResult = ed.GetAngle(pao)

                If utilities.Radtodeg(par.Value) > 105 And utilities.Radtodeg(par.Value) < 285 Then

                    clmarkblk.TransformBy(Matrix3d.Rotation(par.Value + utilities.Degtorad(270), Vector3d.ZAxis, clmarkblk.Position))
                Else
                    clmarkblk.TransformBy(Matrix3d.Rotation(par.Value + utilities.Degtorad(90), Vector3d.ZAxis, clmarkblk.Position))
                End If


                tr.Commit()
            End Using
        Catch ex As Exception

        End Try
    End Sub

    <CommandMethod("eclmark")> Sub E_clmark()
        Dim inspoint As Object = Nothing
        Dim INTPOINT As Object
        Dim blkname
        Dim blkscl
        'Dim prop As AcadDynamicBlockReferenceProperty
        Dim obj As Object = Nothing
        'Dim osmode
        Dim line1 As AcadLine
        Dim line2 As AcadLine
        Dim line3 As AcadLine
        Dim rot As Double
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction
        Dim ATTRIBS
        Dim dynprop
        'Dim rotation


        Try

            Thisdrawing.Utility.GetEntity(obj, inspoint, "Select Centerline: ")

        Catch
            Call Clmark()

        End Try

        Using tr
            Try
                If check_routines.Laychk("~-LCEN") = False Then
                    Thisdrawing.Layers.Add("~-LCEN")
                    Thisdrawing.Layers.Item("~-LCEN").Linetype = "CENTER"
                End If
                obj.Layer = "~-LCEN"
                If TypeOf obj Is AcadLine Then
                    line1 = Thisdrawing.ModelSpace.AddLine(inspoint, obj.StartPoint)
                    line2 = Thisdrawing.ModelSpace.AddLine(inspoint, obj.EndPoint)
                    If line1.Length < line2.Length Then
                        inspoint = obj.StartPoint
                    Else
                        inspoint = obj.EndPoint
                    End If
                    line1.Delete()
                    line2.Delete()

                ElseIf TypeOf obj Is AcadPolyline Or TypeOf obj Is AcadLWPolyline Then

                ElseIf TypeOf obj Is AcadCircle Or TypeOf obj Is AcadArc Then
                    line1 = Thisdrawing.ModelSpace.AddLine(obj.Center, inspoint)
                    rot = line1.Angle
                    INTPOINT = line1.IntersectWith(obj, AcExtendOption.acExtendBoth)
                    inspoint(0) = INTPOINT(0)
                    inspoint(1) = INTPOINT(1)
                    inspoint(2) = INTPOINT(2)
                    line2 = Thisdrawing.ModelSpace.AddLine(inspoint, line1.EndPoint)
                    inspoint(0) = INTPOINT(3)
                    inspoint(1) = INTPOINT(4)
                    inspoint(2) = INTPOINT(5)
                    line3 = Thisdrawing.ModelSpace.AddLine(inspoint, line1.EndPoint)
                    If line2.Length < line3.Length Then
                        inspoint = line2.StartPoint
                    Else
                        inspoint = line3.StartPoint
                    End If
                    line1.Delete()
                    line2.Delete()
                    line3.Delete()
                Else


                End If

                If check_routines.Blkchk("hel_clmark-2026") = True Then
                    blkname = "hel_clmark-2026"
                Else
                    blkname = "hel_clmark-2026.dwg"
                End If

                blkscl = utilities.Getscale
                If blkscl = 0 Then blkscl = 1

                blk = Thisdrawing.ModelSpace.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

                Dim pso As New PromptStringOptions("First Line of Centerline Mark text: ") With {
                       .AllowSpaces = True
                }
                Dim pss1 As PromptResult = ed.GetString(pso)
                Dim pss2 As PromptResult
                If pss1.Status = PromptStatus.OK Then
                    pso.Message = "Second line of Centerline Mark text: "
                    pss2 = ed.GetString(pso)
                Else
                    blk.Delete()
                    Exit Sub
                End If

                dynprop = blk.GetDynamicBlockProperties
                ATTRIBS = blk.GetAttributes
                For i = LBound(ATTRIBS) To UBound(ATTRIBS)
                    Try
                        ATTRIBS(i).StyleName = utilities.Gettxtstyle
                        ATTRIBS(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                        If check_routines.Laychk("~-TXT" & Strings.Right(utilities.Gettxtstyle, 2)) = False Then
                            Thisdrawing.Layers.Add("~-TXT" & Strings.Right(utilities.Gettxtstyle, 2))
                        End If
                        ATTRIBS(i).Layer = "~-TXT" & Strings.Right(utilities.Gettxtstyle, 2)
                    Catch
                    End Try
                Next i


                If check_routines.Laychk("~-TSYM3") = False Then Thisdrawing.Layers.Add("~-TSYM3")
                blk.Layer = "~-TSYM3"
                blk.Update()

                For i = LBound(dynprop) To UBound(dynprop)
                    If dynprop(i).PropertyName = "Visibility" Then
                        If pss1.StringResult = "" Or pss2.StringResult = "" Then
                            dynprop(i).Value = "1-LINE"
                            For j = 0 To UBound(ATTRIBS)
                                If ATTRIBS(j).TagString = "COL." Then
                                    ATTRIBS(j).TextString = pss1.StringResult & pss2.StringResult
                                End If
                            Next j
                        Else
                            dynprop(i).Value = "2-LINE"
                            For j = 0 To UBound(ATTRIBS)
                                If ATTRIBS(j).TagString = "CL2-1" Then
                                    ATTRIBS(j).TextString = pss1.StringResult
                                ElseIf ATTRIBS(j).TagString = "CL2-2" Then
                                    ATTRIBS(j).TextString = pss2.StringResult
                                End If
                            Next j
                        End If
                    End If
                Next i

                Dim angle As Double

                If rot <> 0 Then
                    angle = rot
                Else
                    angle = obj.angle + utilities.Degtorad(90)
                End If
                If angle = Nothing Then angle = utilities.Degtorad(0)
                If utilities.Radtodeg(angle) > 105 And utilities.Radtodeg(angle) < 285 Then
                    angle += utilities.Degtorad(180)
                End If
                blk.Rotation = angle
                tr.Commit()
            Catch
                tr.Abort()
                Exit Try


            End Try
        End Using
        tr.Dispose()
    End Sub
End Class
