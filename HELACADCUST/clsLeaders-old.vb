''Imports Autodesk.AutoCAD.Interop
''Imports Autodesk.AutoCAD.Interop.Common
''Imports Autodesk.AutoCAD.ApplicationServices
''Imports Autodesk.AutoCAD.Runtime
''Imports Autodesk.AutoCAD.EditorInput
''Imports Autodesk.AutoCAD.DatabaseServices



''Public Class clsLeadersold


''    Public check_routines As clsCheckroutines = New clsCheckroutines

''    Public ReadOnly Property Thisdrawing() As AcadDocument
''        Get
''            Return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
''        End Get
''    End Property


''    Dim linecount As Integer

''    <CommandMethod("twopoint")> Public Sub tptarr()
''        Dim txt As AcadText
''        Dim mtxt As AcadMText
''        Dim acadobj As AcadObject
''        Dim leader(0 To 8) As Double
''        Dim leadSTPT As Object
''        'Dim leadTRNPT As Variant
''        'Dim leadNDPT As Variant
''        Dim acadlead As AcadLeader
''        Dim txtmax As Object
''        Dim txtmin As Object
''        Dim OSMODE As Integer
''        Dim MT2TOBJs As AcadText
''        Dim space
''        Dim strendtype
''        Dim endtype As AcDimArrowheadType
''        Dim point

''        'On Error GoTo err

''        strendtype = Thisdrawing.Utility.GetString(False, "Arrow End Type: Arrow, Blank Dot, Dot, Integral ")

''        If LCase(Left(strendtype, 1)) = "a" Then
''            endtype = AcDimArrowheadType.acArrowClosed
''        ElseIf LCase(Left(strendtype, 1)) = "b" Then
''            endtype = AcDimArrowheadType.acArrowDotBlank
''        ElseIf LCase(Left(strendtype, 1)) = "d" Then
''            endtype = AcDimArrowheadType.acArrowDot
''        ElseIf LCase(Left(strendtype, 1)) = "i" Then
''            endtype = AcDimArrowheadType.acArrowIntegral
''        Else
''            endtype = AcDimArrowheadType.acArrowClosed
''        End If

''        If Thisdrawing.ActiveLayout.Name = "Model" Then
''            space = Thisdrawing.ModelSpace
''        Else
''            space = Thisdrawing.PaperSpace
''        End If

''        If check_routines.Laychk("~-TDIM") = False Then Thisdrawing.Layers.Add("~-TDIM")

''        OSMODE = Thisdrawing.GetVariable("OSMODE")
''        If endtype = AcDimArrowheadType.acArrowDotBlank Then
''            Thisdrawing.SetVariable("OSMODE", 4)
''        ElseIf endtype = AcDimArrowheadType.acArrowDot Then
''            Thisdrawing.SetVariable("osmode", 55)
''        Else
''            Thisdrawing.SetVariable("OSMODE", 512)
''        End If
''clear:
''        'On Error GoTo err
''        Thisdrawing.Utility.GetEntity(acadobj, point, "Pick Text")
''        'If TypeOf acadobj Is AcadText Then
''        'Set txt = acadobj
''        If TypeOf acadobj Is AcadMText Then
''            mtxt = acadobj
''            GoTo mtext
''        ElseIf TypeOf acadobj Is AcadText Then
''            txt = acadobj
''        Else
''            Exit Sub
''        End If
''        'On Error Resume Next
''        txt.GetBoundingBox(txtmin, txtmax)
''        leadSTPT = Thisdrawing.Utility.GetPoint(, "Pick Leader Start Location")
''        If leadSTPT(0) < txtmin(0) Or leadSTPT(0) < (txtmin(0) + txtmax(0)) / 2 Then
''            leader(6) = txtmin(0) - (txtmax(1) - txtmin(1)) / 2
''            leader(7) = txtmin(1) + (txtmax(1) - txtmin(1)) / 2
''            leader(8) = txtmin(2)
''            leader(3) = leader(6) - (txtmax(1) - txtmin(1))
''            leader(4) = txtmin(1) + (txtmax(1) - txtmin(1)) / 2
''            leader(5) = txtmin(2)
''        ElseIf leadSTPT(0) > txtmax(0) Or leadSTPT(0) > (txtmin(0) + txtmax(0)) / 2 Then
''            leader(6) = txtmax(0) + (txtmax(1) - txtmin(1)) / 2
''            leader(7) = txtmin(1) + (txtmax(1) - txtmin(1)) / 2
''            leader(8) = txtmin(2)
''            leader(3) = leader(6) + (txtmax(1) - txtmin(1))
''            leader(4) = txtmin(1) + (txtmax(1) - txtmin(1)) / 2
''            leader(5) = txtmin(2)
''        End If
''        leader(0) = leadSTPT(0) : leader(1) = leadSTPT(1) : leader(2) = leadSTPT(2)
''        'leader(3) = leadTRNPT(0): leader(4) = leadTRNPT(1): leader(5) = leadTRNPT(2)
''        'leader(6) = leadNDPT(0): leader(7) = leadNDPT(1): leader(8) = leadNDPT(2)
''        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
''            acadlead = Thisdrawing.ModelSpace.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)

''        Else
''            acadlead = Thisdrawing.PaperSpace.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)
''        End If
''        ' ''acadlead = space.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)
''        ' ''If LCase(space.Name) <> LCase("*Model_Space") Then
''        ' ''acadlead.ScaleFactor = 1
''        ' ''End If
''        acadlead.Layer = "~-TDIM"
''        acadlead.ArrowheadType = endtype
''        If endtype = AcDimArrowheadType.acArrowDotBlank Then
''            acadlead.ArrowheadSize = 4 * Thisdrawing.GetVariable("userr3")
''        End If
''        Thisdrawing.SetVariable("OSMODE", OSMODE)
''        Exit Sub

''mtext:
''        mtxt.GetBoundingBox(txtmin, txtmax)

''        leadSTPT = Thisdrawing.Utility.GetPoint(, "Pick Leader Start Location")
''        'leadSTPT = utils.getpoint("Pick Leader Start Location")

''        If leadSTPT(0) < txtmin(0) Or leadSTPT(0) < (txtmin(0) + txtmax(0)) / 2 Then
''            leader(6) = txtmin(0) - (mtxt.Height / 2)
''            leader(7) = txtmax(1) - (mtxt.Height / 2)
''            leader(8) = txtmin(2)
''            leader(3) = leader(6) - mtxt.Height
''            leader(4) = txtmax(1) - (mtxt.Height / 2)
''            leader(5) = txtmin(2)
''        ElseIf leadSTPT(0) > txtmax(0) Or leadSTPT(0) > (txtmin(0) + txtmax(0)) / 2 Then
''            'Dim i As Integer
''            'Dim lastobj As AcadText
''            'I = UBound(MT2TOBJS)'sets the refernce location on the right side to the Bottom line
''            'i = 0 'LBound(MT2TOBJS) 'sets the refernce location on the right side to the top line
''            'lastobj = MT2TOBJs
''            'lastobj.GetBoundingBox(txtmin, txtmax)
''            'endofline(mtxt)
''            'MT2TOBJs = endofline(mtxt) 'GetMTextAsText(mtxt)
''            MT2TOBJs.GetBoundingBox(txtmin, txtmax)
''            leader(6) = txtmax(0) + (mtxt.Height / 2)
''            'leader(7) = txtmax(1) - (mtxt.Height / 2)
''            leader(7) = mtxt.InsertionPoint(1) - (mtxt.Height / 2)
''            leader(8) = txtmin(2)
''            leader(3) = leader(6) + mtxt.Height
''            leader(4) = mtxt.InsertionPoint(1) - (mtxt.Height / 2)
''            'leader(4) = txtmax(1) - (mtxt.Height / 2)
''            leader(5) = txtmin(2)
''            MT2TOBJs.Delete()
''        End If




''        leader(0) = leadSTPT(0) : leader(1) = leadSTPT(1) : leader(2) = leadSTPT(2)
''        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
''            acadlead = Thisdrawing.ModelSpace.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)

''        Else
''            acadlead = Thisdrawing.PaperSpace.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)
''        End If


''        ' ''acadlead = space.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)
''        acadlead.Layer = "~-TDIM"
''        acadlead.ArrowheadType = endtype
''        If endtype = AcDimArrowheadType.acArrowDotBlank Then
''            acadlead.ArrowheadSize = 4 * Thisdrawing.GetVariable("userr3")
''        End If
''        ' ''If LCase(space.Name) <> LCase("*Model_Space") Then
''        ' ''    acadlead.ScaleFactor = 1
''        ' ''End If

''err:
''        Thisdrawing.SetVariable("OSMODE", OSMODE)
''    End Sub

''    Public Function GetMTextAsText(ByVal aMT As AcadMText) As AcadText


''        'find the end of the first line
''        Dim mtstring As String
''        Dim stlen As Integer
''        mtstring = aMT.TextString
''        stlen = (mtstring.IndexOf("\P"))
''        'Dim newtxt As AcadText

''        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
''            If stlen <> Nothing Then
''                GetMTextAsText = Thisdrawing.ModelSpace.AddText(Strings.Left(mtstring, stlen), aMT.InsertionPoint, aMT.Height)
''            Else
''                GetMTextAsText = Thisdrawing.ModelSpace.AddText(aMT.TextString, aMT.InsertionPoint, aMT.Height)
''            End If
''        Else
''            If stlen <> Nothing Then
''                GetMTextAsText = Thisdrawing.PaperSpace.AddText(Strings.Left(mtstring, stlen), aMT.InsertionPoint, aMT.Height)
''            Else
''                GetMTextAsText = Thisdrawing.PaperSpace.AddText(aMT.TextString, aMT.InsertionPoint, aMT.Height)
''            End If
''        End If


''        ' '' ''Set aBlk = Thisdrawing.ObjectIdToObject(aMT.OwnerID)
''        '' ''If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
''        '' ''    iTC = Thisdrawing.ModelSpace.Count
''        '' ''Else
''        '' ''    iTC = Thisdrawing.PaperSpace.Count
''        '' ''End If
''        ' '' ''If (Not UCase(aBlk.Name) Like "**MODEL_SPACE*") Then
''        ' '' ''Set VOBJS(0) = aMT
''        ' '' ''Thisdrawing.CopyObjects VOBJS, Thisdrawing.ModelSpace
''        ' '' ''Else

''        '' ''MsgBox(iTC)
''        ' '' ''End If
''        ' '' ''On Error GoTo 0
''        '' ''aMT.Copy()
''        '' ''MsgBox(Thisdrawing.ModelSpace.Count)
''        '' ''Thisdrawing.SendCommand("explode ")
''        '' ''Thisdrawing.SendCommand("L ")
''        '' ''MsgBox(Thisdrawing.ModelSpace.Count)
''        '' ''If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
''        '' ''    spacecount = Thisdrawing.ModelSpace.Count
''        '' ''Else
''        '' ''    spacecount = Thisdrawing.PaperSpace.Count
''        '' ''End If
''        '' ''lCT = spacecount - iTC - 1
''        '' ''MsgBox(lCT)
''        '' ''ReDim VOBJ(0 To lCT)
''        '' ''lT = 0
''        '' ''For lCT = iTC To spacecount - 1
''        '' ''    If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
''        '' ''        VOBJ(lT) = Thisdrawing.ModelSpace.Item(lCT)
''        '' ''    Else
''        '' ''        VOBJ(lT) = Thisdrawing.PaperSpace.Item(lCT)
''        '' ''    End If

''        '' ''    lT = lT + 1
''        '' ''Next
''        '' ''MsgBox(UBound(VOBJ))
''        '' ''GetMTextAsText = VOBJ

''        '' ''Dim i As Integer
''        '' ''For i = 0 To UBound(VOBJ)
''        '' ''    'VOBJ(i).Visible = False
''        '' ''Next i

''    End Function

''    <CommandMethod("GETend")> Public Sub endofline()
''        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
''        Dim db As Database = doc.Database
''        Dim ed As Editor = doc.Editor
''        'Dim objid
''        Dim tr As Transaction = db.TransactionManager.StartTransaction
''        Dim pso As PromptSelectionOptions = New PromptSelectionOptions

''        pso.MessageForAdding = "\nSelect Text: "
''        pso.AllowDuplicates = False
''        pso.AllowSubSelections = True
''        pso.RejectObjectsFromNonCurrentSpace = True
''        pso.RejectObjectsOnLockedLayers = False
''        pso.SingleOnly = True

''        Dim psr As PromptSelectionResult = ed.GetSelection(pso)
''        If psr.Status = PromptStatus.OK Then

''            Dim so As SelectedObject = psr.Value(0)
''            With tr

''                'objid = aMT.ObjectID
''                Dim obj As DBObjectCollection = New DBObjectCollection
''                Dim ent As Entity = .GetObject(so.ObjectId, OpenMode.ForRead)
''                Dim mt As MText = ent
''                Dim cb As MTextFragmentCallback = New MTextFragmentCallback(Function(frag As MTextFragment, objid As Object)
''                                                                                ed.WriteMessage("Textfound {0}", frag.Text)
''                                                                                Return MTextFragmentCallbackStatus.Continue
''                                                                            End Function) '(Function(frag As MTextFragment, txtobj As Object) (ed.WriteMessage("Textfound {0}", frag.Text))(MTextFragmentCallbackStatus.Continue)) return
''                mt.ExplodeFragments(cb)


''            End With
''        End If
''    End Sub

''    <CommandMethod("threepoint")> Public Sub thptarr()
''        Dim txt As AcadText
''        Dim mtxt As AcadMText
''        Dim acadobj As AcadObject
''        Dim leader(0 To 8) As Double
''        Dim leadSTPT As Object
''        Dim leadTRPT As Object
''        'Dim leadTRNPT As Variant
''        'Dim leadNDPT As Variant
''        Dim acadlead As AcadLeader
''        Dim txtmax As Object
''        Dim txtmin As Object
''        Dim OSMODE As Integer
''        Dim MT2TOBJS As AcadText
''        Dim space
''        Dim strendtype As String
''        Dim endtype As AcDimArrowheadType
''        Dim point

''        'On Error GoTo err
''        strendtype = Thisdrawing.Utility.GetString(False, "Arrow End Type: Arrow, Blank Dot, Dot, Integral ")

''        If LCase(Left(strendtype, 1)) = "a" Then
''            endtype = AcDimArrowheadType.acArrowOblique
''        ElseIf LCase(Left(strendtype, 1)) = "b" Then
''            endtype = AcDimArrowheadType.acArrowDotBlank
''        ElseIf LCase(Left(strendtype, 1)) = "d" Then
''            endtype = AcDimArrowheadType.acArrowDot
''        ElseIf LCase(Left(strendtype, 1)) = "i" Then
''            endtype = AcDimArrowheadType.acArrowIntegral
''        Else
''            endtype = AcDimArrowheadType.acArrowClosed
''        End If

''        If Thisdrawing.ActiveLayout.Name = "Model" Then
''            space = Thisdrawing.ModelSpace
''        Else
''            space = Thisdrawing.PaperSpace
''        End If

''        If check_routines.Laychk("~-TDIM") = False Then Thisdrawing.Layers.Add("~-TDIM")

''        OSMODE = Thisdrawing.GetVariable("OSMODE")
''        If endtype = AcDimArrowheadType.acArrowDotBlank Then
''            Thisdrawing.SetVariable("OSMODE", 4)
''        ElseIf endtype = AcDimArrowheadType.acArrowDot Then
''            Thisdrawing.SetVariable("osmode", 55)
''        Else
''            Thisdrawing.SetVariable("OSMODE", 512)
''        End If
''clear:
''        On Error GoTo err
''        Thisdrawing.Utility.GetEntity(acadobj, point, "Pick Text")
''        'If TypeOf acadobj Is AcadText Then
''        'Set txt = acadobj
''        If TypeOf acadobj Is AcadMText Then
''            mtxt = acadobj
''            GoTo mtext
''        ElseIf TypeOf acadobj Is AcadText Then
''            txt = acadobj
''        Else
''            Exit Sub
''        End If
''        'On Error Resume Next
''        txt.GetBoundingBox(txtmin, txtmax)
''        leadTRPT = Thisdrawing.Utility.GetPoint(, "Pick Leader Turn Point")
''        leadSTPT = Thisdrawing.Utility.GetPoint(, "Pick Leader Start Location")
''        If leadSTPT(0) < txtmin(0) Or leadSTPT(0) < (txtmin(0) + txtmax(0)) / 2 Then
''            leader(6) = txtmin(0) - (txtmax(1) - txtmin(1)) / 2
''            leader(7) = txtmin(1) + (txtmax(1) - txtmin(1)) / 2
''            leader(8) = txtmin(2)
''            If leadTRPT(0) > txtmin(0) And leadTRPT(0) < txtmax(0) Then
''                leader(3) = leader(6) - (txtmax(1) - txtmin(1))
''            Else
''                leader(3) = leadTRPT(0)
''            End If
''            leader(4) = txtmin(1) + (txtmax(1) - txtmin(1)) / 2
''            leader(5) = txtmin(2)
''        ElseIf leadSTPT(0) > txtmax(0) Or leadSTPT(0) > (txtmin(0) + txtmax(0)) / 2 Then
''            leader(6) = txtmax(0) + (txtmax(1) - txtmin(1)) / 2
''            leader(7) = txtmin(1) + (txtmax(1) - txtmin(1)) / 2
''            leader(8) = txtmin(2)
''            If leadTRPT(0) > txtmin(0) And leadTRPT(0) < txtmax(0) Then
''                leader(3) = leader(6) + (txtmax(1) - txtmin(1))
''            Else
''                leader(3) = leadTRPT(0)
''            End If
''            leader(4) = txtmin(1) + (txtmax(1) - txtmin(1)) / 2
''            leader(5) = txtmin(2)
''        End If
''        leader(0) = leadSTPT(0) : leader(1) = leadSTPT(1) : leader(2) = leadSTPT(2)
''        'leader(3) = leadTRNPT(0): leader(4) = leadTRNPT(1): leader(5) = leadTRNPT(2)
''        'leader(6) = leadNDPT(0): leader(7) = leadNDPT(1): leader(8) = leadNDPT(2)
''        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
''            acadlead = Thisdrawing.ModelSpace.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)

''        Else
''            acadlead = Thisdrawing.PaperSpace.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)
''        End If
''        ' ''acadlead = space.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)
''        ' ''If LCase(space.Name) <> LCase("*Model_Space") Then
''        ' ''acadlead.ScaleFactor = 1
''        ' ''End If
''        acadlead.Layer = "~-TDIM"
''        acadlead.ArrowheadType = endtype
''        If endtype = AcDimArrowheadType.acArrowDotBlank Then
''            acadlead.ArrowheadSize = 4 * Thisdrawing.GetVariable("userr3")
''        End If
''        Thisdrawing.SetVariable("OSMODE", OSMODE)
''        Exit Sub

''mtext:
''        mtxt.GetBoundingBox(txtmin, txtmax)

''        leadTRPT = Thisdrawing.Utility.GetPoint(, "Pick Leader Turn Point")
''        leadSTPT = Thisdrawing.Utility.GetPoint(, "Pick Leader Start Location")
''        If leadSTPT(0) < txtmin(0) Or leadSTPT(0) < (txtmin(0) + txtmax(0)) / 2 Then
''            leader(6) = txtmin(0) - (mtxt.Height / 2)
''            leader(7) = txtmax(1) - (mtxt.Height / 2)
''            leader(8) = txtmin(2)
''            leader(3) = leadTRPT(0)
''            If leadTRPT(0) > txtmin(0) And leadTRPT(0) < txtmax(0) Then
''                leader(3) = leader(6) - (txtmax(1) - txtmin(1))
''            Else
''                leader(3) = leadTRPT(0)
''            End If
''            leader(4) = txtmax(1) - (mtxt.Height / 2)
''            leader(5) = txtmin(2)
''        ElseIf leadSTPT(0) > txtmax(0) Or leadSTPT(0) > (txtmin(0) + txtmax(0)) / 2 Then
''            'Dim i As Integer
''            'Dim lastobj As AcadText
''            'I = UBound(MT2TOBJS)'sets the refernce location on the right side to the Bottom line
''            'i = LBound(MT2TOBJS) 'sets the refernce location on the right side to the top line
''            'lastobj = MT2TOBJS(i)
''            'lastobj.GetBoundingBox(txtmin, txtmax)
''            MT2TOBJS = GetMTextAsText(mtxt)
''            MT2TOBJS.GetBoundingBox(txtmin, txtmax)
''            leader(6) = txtmax(0) + (mtxt.Height / 2)
''            leader(7) = mtxt.InsertionPoint(1) - (mtxt.Height / 2)
''            'leader(7) = txtmax(1) - (mtxt.Height / 2)
''            leader(8) = txtmin(2)
''            If leadTRPT(0) > txtmin(0) And leadTRPT(0) < txtmax(0) Then
''                leader(3) = leader(6) + (txtmax(1) - txtmin(1))
''            Else
''                leader(3) = leadTRPT(0)
''            End If
''            leader(4) = mtxt.InsertionPoint(1) - (mtxt.Height / 2)
''            'leader(4) = txtmax(1) - (mtxt.Height / 2)
''            leader(5) = txtmin(2)
''            MT2TOBJS.Delete()
''        End If

''        'For i = 0 To UBound(MT2TOBJS)
''        '    MT2TOBJS(i).Delete()
''        'Next i

''        leader(0) = leadSTPT(0) : leader(1) = leadSTPT(1) : leader(2) = leadSTPT(2)
''        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
''            acadlead = Thisdrawing.ModelSpace.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)
''        Else
''            acadlead = Thisdrawing.PaperSpace.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)
''        End If
''        ' ''acadlead = space.AddLeader(leader, Nothing, AcDimArrowheadType.acArrowClosed)
''        ' ''If LCase(space.Name) <> LCase("*Model_Space") Then
''        ' ''acadlead.ScaleFactor = 1
''        ' ''End If
''        acadlead.Layer = "~-TDIM"
''        acadlead.ArrowheadType = endtype
''        If endtype = AcDimArrowheadType.acArrowDotBlank Then
''            acadlead.ArrowheadSize = 4 * Thisdrawing.GetVariable("userr3")
''        End If


''err:
''        Thisdrawing.SetVariable("OSMODE", OSMODE)
''    End Sub

''End Class
