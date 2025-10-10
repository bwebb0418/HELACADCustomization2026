Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Windows
Imports System.Runtime
Imports System.Diagnostics

Public Class Clsbind_purge
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public Function IsXrefResolved(ByVal xrBlockName As String)
        'If the XREf is not resolved (unloaded or not found),
        'then attempting to reference the XRefDatabase property
        'of the corresponding Block object for the Xref will
        'raise an error with the description "No database".
        On Error Resume Next
        Dim objBlock As AcadDatabase
        objBlock = Thisdrawing.Blocks(CType(xrBlockName, Index)).XRefDatabase
        If Err.Number = 0 Then
            IsXrefResolved = True
        Else
            IsXrefResolved = False
        End If

    End Function


    Function Preplotchk()
        'chkfonts()
        Preplotchk = Nothing
    End Function

    Function Chkfonts()
        Dim tstyle As AcadTextStyle
        Dim fnt As String
        Dim bld As Boolean
        Dim ita As Boolean
        Dim cha As Long
        Dim paf As Long
        On Error GoTo err

        For Each tstyle In Thisdrawing.TextStyles
            If Left(tstyle.Name, 9) = "HEL-Title" Then
                GoTo nextfnt
            End If
            If Thisdrawing.GetVariable("useri1") <> 7 Then GoTo logochk
            If Thisdrawing.GetVariable("users1") = "arch" Then
                If Left(tstyle.Name, 4) = "HEL2" Or Left(tstyle.Name, 5) = "PHEL2" Then
                    tstyle.fontFile = "Archquik.shx"
                ElseIf Left(tstyle.Name, 4) = "HEL3" Or Left(tstyle.Name, 5) = "PHEL3" Then
                    tstyle.fontFile = "Archstyl.shx"
                ElseIf Left(tstyle.Name, 4) = "HEL4" Or Left(tstyle.Name, 5) = "PHEL4" Then
                    tstyle.fontFile = "Archtitl.shx"
                ElseIf Left(tstyle.Name, 4) = "HEL4" Or Left(tstyle.Name, 5) = "PHEL4" Then
                    tstyle.fontFile = "Archtitl.shx"
                End If
            Else
                If tstyle.Name = "HEL_sym" Then
                    tstyle.fontFile = "ROMANS.shx"
                ElseIf Left(tstyle.Name, 3) = "HEL" Then
                    tstyle.fontFile = "ROMANS.shx"
                ElseIf Left(tstyle.Name, 4) = "PHEL" Then
                    tstyle.fontFile = "ROMANS.shx"
                End If
            End If
logochk:
            If tstyle.Name = "LOGO" Or tstyle.Name = "LOGO-HEL" Then
                tstyle.SetFont("Eras Bold ITC", False, False, 0, 0)
            ElseIf tstyle.Name = "HEL-Logo" Then
                tstyle.SetFont("Eras Bold ITC", False, False, 0, 0)
            ElseIf tstyle.Name = "LOGO-HEL2" Then
                tstyle.fontFile = "romans.shx"
            ElseIf LCase(tstyle.Name = "hel-logo2") Then
                tstyle.fontFile = "romans.shx"
            ElseIf tstyle.Name = "STYLE1" Or tstyle.Name = "STAMP-HEL" Then
                tstyle.SetFont("Swis721 blkoul bt", False, False, 0, 0)
            End If
nextfnt:

        Next tstyle

        Thisdrawing.Regen(AcRegenType.acAllViewports)

        Exit Function

err:

        MsgBox("Contact System Administrator" & vbCr & "Font Check Error - Missing Font", vbCritical + vbOKOnly, "Error!")
    End Function


    Function Reloadxref()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        Reloadxref = Nothing

        Using tr As Transaction = db.TransactionManager.StartTransaction

            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord
            Dim blkid As ObjectId
            Dim enumr As IEnumerator = bt.GetEnumerator
            ed.WriteMessage(vbLf)
            Dim lxrefids As New ObjectIdCollection
            Try
                Do While enumr.MoveNext
                    blkid = enumr.Current
                    btr = tr.GetObject(blkid, OpenMode.ForWrite)
                    If btr.IsFromExternalReference Then
                        If btr.IsUnloaded = False Then
                            If Not lxrefids.Contains(blkid) Then lxrefids.Add(blkid)
                            'ed.WriteMessage("Xref {0} has be reloaded." & vbLf, btr.Name)
                        End If
                    End If
                Loop
            Catch
                tr.Abort()
                Exit Function
            End Try

            If lxrefids.Count <> 0 Then
                db.ReloadXrefs(lxrefids)
                tr.Commit()
            Else
                ed.WriteMessage("Error reloading Xref's.  Please perform a manual reload")
                tr.Abort()
            End If
        End Using
    End Function

    Function Bindxref()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        Dim btr As BlockTableRecord
        Dim blkid As ObjectId
        Dim bt As BlockTable
        Dim enumr As IEnumerator
        ed.WriteMessage(vbLf)
        Dim lxrefids As New ObjectIdCollection
        Dim uxrefids As New ObjectIdCollection
        Bindxref = Nothing
        Using tr As Transaction = db.TransactionManager.StartTransaction

            bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            enumr = bt.GetEnumerator
            Do While enumr.MoveNext
                blkid = enumr.Current
                btr = tr.GetObject(blkid, OpenMode.ForWrite)
                If btr.IsFromExternalReference Then
                    If btr.IsUnloaded = False Then
                        'If Not lxrefids.Contains(blkid) Then lxrefids.Add(blkid)
                    Else
                        Dim msgboxres As MsgBoxResult
                        msgboxres = MsgBox("Xref " & btr.Name & " is currently unloaded." & vbCr & "If it is not loaded it will not be bound!" & vbCr &
                                           "Do you wish to load it now?", MsgBoxStyle.YesNo, "Unloaded Xref")
                        If msgboxres = MsgBoxResult.Yes Then
                            If Not uxrefids.Contains(blkid) Then uxrefids.Add(blkid)
                            '    If Not lxrefids.Contains(blkid) Then lxrefids.Add(blkid)
                        Else
                            db.DetachXref(blkid)
                            ed.WriteMessage("Xref {0} has been detached." & vbLf, btr.Name)
                        End If
                    End If
                End If
            Loop
            If uxrefids.Count <> 0 Then
                Try
                    db.ReloadXrefs(uxrefids)
                Catch
                    ed.WriteMessage("Error reloading Xref's.  Please perform a manual reload")
                End Try
                'db.ResolveXrefs(True, False)
            End If
            tr.Commit()
        End Using
        Using tr = db.TransactionManager.StartTransaction
            ed.WriteMessage(vbLf)
            bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            enumr = bt.GetEnumerator
            Do While enumr.MoveNext
                blkid = enumr.Current
                btr = tr.GetObject(blkid, OpenMode.ForWrite)
                If btr.IsFromExternalReference Then
                    If btr.IsUnloaded = False Then
                        If Not lxrefids.Contains(blkid) Then
                            lxrefids.Clear()
                            lxrefids.Add(blkid)
                            If lxrefids.Count <> 0 Then
                                Try
                                    db.BindXrefs(lxrefids, False)
                                Catch ex As Exception
                                    ed.WriteMessage("There was an error binding Xref {0}. Please manually bind this xref." & vbLf, btr.Name)
                                End Try
                            End If
                        End If
                    End If
                End If
            Loop

            tr.Commit()
        End Using
    End Function
    Function Purgedwg()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        Dim objids As New ObjectIdCollection
        Dim dobjids As New ObjectIdCollection
        Dim symtrec As SymbolTableRecord
        'Dim obj As DBObject
        ed.WriteMessage(vbLf)

        Purgedwg = Nothing

        Using tr As Transaction = doc.TransactionManager.StartTransaction

            Dim blkt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)

            For Each objid As ObjectId In blkt
                objids.Add(objid)
            Next

            db.Purge(objids)

            For Each pobjid As ObjectId In objids
                symtrec = tr.GetObject(pobjid, OpenMode.ForWrite)
                Try
                    symtrec.Erase(True)
                Catch ex As Exception
                    doc.Editor.WriteMessage("Block Object {0} could not be purged!" & vbLf, pobjid)
                End Try
            Next
            tr.Commit()

        End Using

        Using tr = doc.TransactionManager.StartTransaction
            objids.Clear()

            Dim lryt As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
            Dim tst As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
            Dim dst As DimStyleTable = tr.GetObject(db.DimStyleTableId, OpenMode.ForRead)
            Dim lnt As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)

            For Each objid As ObjectId In lryt
                objids.Add(objid)
            Next

            For Each objid As ObjectId In tst
                objids.Add(objid)
            Next

            For Each objid As ObjectId In dst
                objids.Add(objid)
            Next

            For Each objid As ObjectId In lnt
                objids.Add(objid)
            Next

            db.Purge(objids)

            For Each pobjid As ObjectId In objids
                symtrec = tr.GetObject(pobjid, OpenMode.ForWrite)
                Try
                    symtrec.Erase(True)
                Catch ex As Exception
                    doc.Editor.WriteMessage("Style Object {0} could not be purged!", pobjid)
                End Try
            Next

            tr.Commit()
        End Using


    End Function
    <CommandMethod("BINDPURGE")> Sub BindandPurge()
        Bindxref()
        Purgedwg()
        Purgedwg()
        Purgedwg()
    End Sub

End Class
