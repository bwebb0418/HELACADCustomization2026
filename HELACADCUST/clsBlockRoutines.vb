Imports System.Runtime
Imports System.Diagnostics
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows

Public Class ClsBlockRoutines
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public utilities As New Clsutilities
    Public check_routines As New ClsCheckroutines
    Function Blockupdate()

        Dim blkref As AcadBlockReference
        Dim blkref2 As AcadBlockReference
        Dim dynprops, ATTRIBS
        Dim i As Integer
        Dim j As Integer
        Dim blkscl
        Dim obj As Object
        Dim dynprops2, attribs2
        Dim attrib As AcadAttributeReference
        Dim dynprop As AcadDynamicBlockReferenceProperty
        Dim minpoint As Object = Nothing
        Dim maxpoint As Object = Nothing

        For Each obj In Thisdrawing.ModelSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If blkref.XScaleFactor <> 1 And blkref.XScaleFactor <> 10 / 254 Then
                    blkscl = utilities.Getscale
                Else
                    blkscl = blkref.XScaleFactor
                End If
                If blkref.IsDynamicBlock And Left(blkref.EffectiveName, 13) <> "hel_shearwall-2023" And blkref.EffectiveName <> "hel_sym_secarr-2023" Then
                    blkref2 = Thisdrawing.ModelSpace.InsertBlock(blkref.InsertionPoint, blkref.EffectiveName, blkscl, blkscl, blkscl, blkref.Rotation)
                    dynprops = blkref.GetDynamicBlockProperties
                    dynprops2 = blkref2.GetDynamicBlockProperties
                    For i = LBound(dynprops) To UBound(dynprops)
                        If dynprops(i).PropertyName <> "Origin" Then dynprops2(i).Value = dynprops(i).Value
                    Next i
                    ATTRIBS = blkref.GetAttributes
                    attribs2 = blkref2.GetAttributes
                    For i = LBound(ATTRIBS) To UBound(ATTRIBS)
                        On Error Resume Next
                        attribs2(i).StyleName = utilities.Gettxtstyle
                        attribs2(i).Height = Thisdrawing.TextStyles.Item(utilities.Gettxtstyle).Height
                        If check_routines.Laychk("~-TXT" & Right(utilities.Gettxtstyle, 2)) = False Then
                            Thisdrawing.Layers.Add("~-TXT" & Right(utilities.Gettxtstyle, 2))
                        End If
                        attribs2(i).Layer = "~-TXT" & Right(utilities.Gettxtstyle, 2)
                        On Error GoTo 0
                        attribs2(i).TextString = ATTRIBS(i).TextString
                    Next i
                    blkref2.Layer = blkref.Layer
                    blkref.Delete()

                End If
            End If
        Next obj


        For Each obj In Thisdrawing.ModelSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj
                If blkref.IsDynamicBlock And blkref.EffectiveName = "hel_rec_tag-2023" Then
                    dynprops = blkref.GetDynamicBlockProperties
                    For i = 0 To UBound(dynprops)
                        If dynprops(i).PropertyName = "Distance" Or dynprops(i).PropertyName = "Distance1" Then
                            dynprop = dynprops(i)
                            ATTRIBS = blkref.GetAttributes
                            For j = 0 To UBound(ATTRIBS)
                                If ATTRIBS(j).TagString = "RECTAG" Then
                                    attrib = ATTRIBS(j)
                                    attrib.GetBoundingBox(minpoint, maxpoint)
                                    maxpoint(1) = blkref.InsertionPoint(1)
                                    minpoint(1) = blkref.InsertionPoint(1)
                                    dynprop.Value = (maxpoint(0) - blkref.InsertionPoint(0))
                                End If
                            Next j
                        End If
                    Next i
                End If
            End If
        Next obj


    End Function

    Function Attribsfix()
        Dim ATTRIBS
        Dim BLOCK As AcadBlockReference
        Dim obj As AcadObject
        Dim i As Integer
        Dim txtlay As String = ""

        With Thisdrawing
            If .GetVariable("UserR1") = "2" Then txtlay = "~-TXT20"
            If .GetVariable("UserR1") = "2.5" Then txtlay = "~-TXT25"
            If .GetVariable("UserR1") = "3" Then txtlay = "~-TXT30"
        End With

        For Each obj In Thisdrawing.ModelSpace
            If TypeOf obj Is AcadBlockReference Then
                BLOCK = obj
                If BLOCK.EffectiveName = "hel_steel_beam_i-2023" Or BLOCK.EffectiveName = "hel_steel_beam-2023" Or
                    BLOCK.EffectiveName = "hel_wood_beam_i-2023" Or BLOCK.EffectiveName = "hel_wood_beam-2023" Or
                    BLOCK.EffectiveName = "hel_girt_tag_i-2023" Or BLOCK.EffectiveName = "hel_girt_tag-2023" Then
                    ATTRIBS = BLOCK.GetAttributes
                    For i = 0 To UBound(ATTRIBS)
                        ATTRIBS(i).Height = Thisdrawing.ActiveTextStyle.Height
                        ATTRIBS(i).StyleName = Thisdrawing.ActiveTextStyle.Name
                        check_routines.Laychk(txtlay)
                        ATTRIBS(i).Layer = txtlay
                    Next i
                End If
            End If
        Next obj
        Attribsfix = Nothing
    End Function

    Function Attfix()

        Dim blkref As AcadBlockReference
        Dim attribref
        Dim i
        Dim obj As Object

        For Each obj In Thisdrawing.ModelSpace
            If TypeOf obj Is AcadBlockReference Then
                blkref = obj

                attribref = blkref.GetAttributes
                For i = 0 To UBound(attribref)
                    attribref(i).TextString = attribref(i).TextString
                Next i
            End If
        Next obj

        Attfix = Nothing

    End Function


    Sub Block_Replace()
        Dim blknew As AcadBlockReference
        Dim blkold As AcadBlockReference = Nothing
        Dim blkrpl As AcadBlockReference
        Dim ssobjs As AcadSelectionSet
        Dim point(0 To 2) As Object
        Dim i As Integer
        Dim objs As Object = Nothing
        Dim filterdata(0) As Object
        Dim filtertype(0) As Integer
        Dim movept
        Dim startpt
        Dim BLKTEMP As AcadBlockReference
        Dim METHOD

        For Each ssobjs In Thisdrawing.SelectionSets
            If ssobjs.Name = "blockreplace" Then
                ssobjs.Delete()
                Exit For
            End If
        Next ssobjs

        ssobjs = Thisdrawing.SelectionSets.Add("blockreplace")

        On Error GoTo err
picknewblock:
        Thisdrawing.Utility.GetEntity(objs, point, "Pick New Block: ")

        If TypeOf objs Is AcadBlockReference Then
            blknew = objs
        Else
            GoTo picknewblock
        End If

        filtertype(0) = 0 : filterdata(0) = "Insert"
        On Error GoTo 0

        METHOD = Thisdrawing.Utility.GetString(1, "Name OR Pick? ")

        If LCase(METHOD) = "p" Then
            Thisdrawing.Utility.Prompt("Select Blocks to Replace: ")
            ssobjs.SelectOnScreen(filtertype, filterdata)

            For i = 0 To ssobjs.Count - 1
                blkold = ssobjs(i)
                With Thisdrawing.ModelSpace
                    BLKTEMP = .InsertBlock(blkold.InsertionPoint, blknew.EffectiveName, blkold.XScaleFactor, blkold.YScaleFactor, blkold.ZScaleFactor, blkold.Rotation)
                    BLKTEMP.Layer = blkold.Layer
                End With
            Next i

            For i = 0 To ssobjs.Count - 1
                ssobjs(i).Delete()
            Next i

        Else

            Thisdrawing.Utility.GetEntity(blkold, point, "Select Blocks to Replace: ")
            Dim blkname
            Dim ATTRIBSNEW
            Dim ATTRIBSOLD
            blkname = blkold.EffectiveName
            For Each objs In Thisdrawing.ModelSpace
                If TypeOf objs Is AcadBlockReference Then
                    blkrpl = objs
                    If blkrpl.EffectiveName = blkname Then
                        BLKTEMP = Thisdrawing.ModelSpace.InsertBlock(blkrpl.InsertionPoint, blknew.EffectiveName, blkrpl.XScaleFactor, blkrpl.YScaleFactor, blkrpl.ZScaleFactor, blkrpl.Rotation)
                        BLKTEMP.Layer = blkrpl.Layer
                        ATTRIBSOLD = blkrpl.GetAttributes
                        If UBound(ATTRIBSOLD) <= 0 Then ATTRIBSOLD = blknew.GetAttributes
                        ATTRIBSNEW = BLKTEMP.GetAttributes
                        For i = 0 To UBound(ATTRIBSOLD)
                            ATTRIBSNEW(i).TextString = ATTRIBSOLD(i).TextString
                        Next i
                        blkrpl.Delete()
                    End If
                End If
            Next objs
        End If

err:
    End Sub

End Class
