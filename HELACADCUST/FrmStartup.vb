Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports System.Diagnostics
Imports System.Reflection


Public Class FrmStartup

    Dim startupfrm As FrmStartup
    Dim xd As ClsXdata
    ReadOnly checker As New ClsCheckroutines
    Dim uni As String
    Dim scl As Integer
    Dim sclchk As String
    Dim txt As String
    Dim discp As String
    Dim dimst As String
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Private Sub AddScales()
        With Me.CboSCL.Items
            'Imperial Inches
            If rdoINS.Checked = True Then
                .Clear()
                .Add("1/16")
                .Add("1/8")
                .Add("3/16")
                .Add("1/4")
                .Add("3/8")
                .Add("1/2")
                .Add("3/4")
                .Add("1")
                .Add("1 1/2")
                .Add("3")
                .Add("6")
                .Add("12")
                'lblunit.Caption = " = 1'-0"""
                'Imperial Feet
            ElseIf rdoFTS.Checked = True Then
                .Clear()
                .Add("10")
                .Add("20")
                .Add("30")
                .Add("40")
                .Add("50")
                .Add("60")
                'lblunit.Caption = "Under Construction"
                'Metric Millimeters
            ElseIf rdoMMS.Checked = True Then
                .Clear()
                .Add("1")
                .Add("2")
                .Add("5")
                .Add("10")
                .Add("20")
                .Add("25")
                .Add("50")
                .Add("75")
                .Add("100")
                .Add("125")
                .Add("200")
                .Add("500")
                'lblunit.Caption = "Under Construction"
                'Metric Meters
            ElseIf rdoMTS.Checked = True Then
                .Clear()
                .Add("100")
                .Add("200")
                .Add("250")
                .Add("300")
                .Add("400")
                .Add("500")
                .Add("1000")
                .Add("2000")
                .Add("2500")
                .Add("3000")
                .Add("4000")
                .Add("5000")
                .Add("7500")
                .Add("10000")
                .Add("12500")
                .Add("20000")
                .Add("50000")
                'lblunit.Caption = "Under Construction"
                'Everything else
            Else
                '.Clear()
                'MsgBox("Contact AutoCAD Systems Manager", vbCritical + vbAbortRetryIgnore, "Error!")
                '.Text = "ERROR!"
            End If
        End With
        tboCSC.Text = ""
    End Sub


    Private Sub RdoMMS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoMMS.CheckedChanged
        AddScales()
    End Sub

    Private Sub RdoINS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoINS.CheckedChanged
        AddScales()
    End Sub

    Private Sub RdoMTS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoMTS.CheckedChanged
        AddScales()
    End Sub

    Private Sub RdoFTS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoFTS.CheckedChanged
        AddScales()
    End Sub

    Private Sub ButSTR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butSTR.Click

        startupfrm = Me

        xd = New ClsXdata

        If rdoMMS.Checked = True Then
            uni = "mms"
        ElseIf rdoFTS.Checked = True Then
            uni = "fts"
        ElseIf rdoMTS.Checked = True Then
            uni = "mts"
        ElseIf rdoINS.Checked = True Then
            uni = "ins"
        Else
            MsgBox("You must choose your units!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly _
       , "Error!")

            Exit Sub
        End If

        If rdoARW.Checked = True Then
            dimst = "arrws"
        ElseIf rdoTCK.Checked = True Then
            dimst = "ticks"
        Else
            MsgBox("You must choose a dimenstion style!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly _
       , "Error!")
            Exit Sub
        End If

        If rdo20.Checked = True Then
            txt = "2.0"
        ElseIf rdo25.Checked = True Then
            txt = "2.5"
        Else
            MsgBox("You must choose a text size!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly _
                   , "Error!")
            Exit Sub

        End If

        If rdoARC.Checked = True Then
            discp = "arch"
        ElseIf rdoBDE.Checked = True Then
            discp = "blen"
        ElseIf rdoIND.Checked = True Then
            discp = "Inds"
        ElseIf rdoBDG.Checked = True Then
            discp = "bldg"
        Else
            MsgBox("You must choose a discipline!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error!")
            Exit Sub
        End If

        If tboCSC.Text = "" Then
            MsgBox("You must choose a scale!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error!")
            Exit Sub
        Else
            sclchk = tboCSC.Text.Trim()
            Dim result As Boolean = Integer.TryParse(sclchk, scl)
            If result Then
                scl = tboCSC.Text
            Else
                MsgBox("You must choose a scale!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error!")
                Exit Sub
            End If
        End If
        startupfrm.Hide()
        xd.Compilexdata(uni, scl, txt, discp, dimst)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        Dwgboundary()
    End Sub

    Private Sub ButEND_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butEND.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Abort
        startupfrm = Me
        startupfrm.Hide()

        'xd.CanceledStartup()

    End Sub

    Private Sub CboSCL_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboSCL.MouseHover

    End Sub

    Private Sub CboSCL_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboSCL.SelectedIndexChanged
        ' ''If Me.CboSCL.SelectedText = "" Then
        ' ''    tboCSC.Text = ""
        ' ''    Exit Sub
        ' ''End If
        'converts Cboscale.selectedvalue to a number that can be used for the scale factor
        With Me.CboSCL
            If .SelectedItem.ToString = "1/16" Then tboCSC.Text = 12 * 16
            If .SelectedItem.ToString = "1/8" Then tboCSC.Text = 12 * 8
            If .SelectedItem.ToString = "3/16" Then tboCSC.Text = 12 * 16 / 3
            If .SelectedItem.ToString = "1/4" Then tboCSC.Text = 12 * 4
            If .SelectedItem.ToString = "3/8" Then tboCSC.Text = 12 * 8 / 3
            If .SelectedItem.ToString = "1/2" Then tboCSC.Text = 12 * 2
            If .SelectedItem.ToString = "3/4" Then tboCSC.Text = 12 * 4 / 3
            If .SelectedItem.ToString = "1" Then tboCSC.Text = 12
            If .SelectedItem.ToString = "1 1/2" Then tboCSC.Text = 12 / 1.5
            If .SelectedItem.ToString = "3" Then tboCSC.Text = 12 / 3
            If .SelectedItem.ToString = "6" Then tboCSC.Text = 12 / 6
            If .SelectedItem.ToString = "12" Then tboCSC.Text = 1
            If rdoMMS.Checked = True Or rdoMTS.Checked = True Then tboCSC.Text = .SelectedItem.ToString
            If rdoFTS.Checked = True Then tboCSC.Text = .SelectedItem.ToString
        End With
    End Sub

    Private Sub RdoARC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoARC.CheckedChanged
        rdoTCK.Checked = True
        rdoINS.Checked = True
        rdo20.Checked = True
    End Sub

    Private Sub RdoBDE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoBDE.CheckedChanged
        rdoTCK.Checked = True
        rdoINS.Checked = True
        rdo20.Checked = True
    End Sub

    Private Sub RdoBDG_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoBDG.CheckedChanged
        rdoTCK.Checked = True
        rdoMMS.Checked = True
        rdo20.Checked = True
    End Sub

    Private Sub RdoIND_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoIND.CheckedChanged
        rdoARW.Checked = True
        rdoINS.Checked = True
        rdo25.Checked = True
    End Sub

    Function Dwgboundary()
        Dwgboundary = Nothing
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim pline As Polyline
        Dim check_routines As New ClsCheckroutines
        Dim tr As Transaction = doc.TransactionManager.StartTransaction
        Dim Getscl As New Clsutilities
        Dim layer As LayerTableRecord = tr.GetObject(check_routines.Layercheck("defpoints"), OpenMode.ForWrite)
        Dim laycol As Autodesk.AutoCAD.Colors.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Colors.ColorMethod.ByColor, 7)
        layer.IsPlottable = False
        layer.Color = laycol ' FromColorIndex(Colors.ColorMethod.ByColor, 7)
        tr.Commit()
        tr = doc.TransactionManager.StartTransaction
        Dim blktbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim blktblrec As BlockTableRecord = tr.GetObject(blktbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

        db.Limmin = New Point2d(0, 0)
        db.Limmax = New Point2d(787.4 * Getscl.Getscale, 571.5 * Getscl.Getscale)

        Dim point1 As New Point2d(db.Limmin.X, db.Limmin.Y)
        Dim point2 As New Point2d(db.Limmax.X, db.Limmin.Y)
        Dim point3 As New Point2d(db.Limmax.X, db.Limmax.Y)
        Dim point4 As New Point2d(db.Limmin.X, db.Limmax.Y)
        pline = New Polyline
        pline.SetDatabaseDefaults()
        pline.AddVertexAt(0, point1, 0, 0, 0)
        pline.AddVertexAt(1, point2, 0, 0, 0)
        pline.AddVertexAt(2, point3, 0, 0, 0)
        pline.AddVertexAt(3, point4, 0, 0, 0)
        pline.Closed = True

        pline.Layer = layer.Name
        blktblrec.AppendEntity(pline)
        tr.AddNewlyCreatedDBObject(pline, True)
        If Application.GetSystemVariable("TILEMODE") = 0 Then Application.SetSystemVariable("TILEMODE", 1)
        Dim view As New ViewTableRecord
        Dim point5 As New Point2d(db.Limmax.X / 2, db.Limmax.Y / 2)
        view.CenterPoint = point5
        view.Height = (db.Limmax.X - db.Limmin.X)
        view.Width = (db.Limmax.Y - db.Limmin.Y)
        ed.SetCurrentView(view)

        Application.SetSystemVariable("TILEMODE", 0)
        ed.SwitchToPaperSpace()

        'ed.SwitchToModelSpace()
        tr.Commit()
        tr = doc.TransactionManager.StartTransaction

        Dim layoutdict As DBDictionary = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead)

        For Each dictent As DBDictionaryEntry In layoutdict
            Dim layoutname As String = dictent.Key
            If layoutname = "Layout2" Then
                LayoutManager.Current.DeleteLayout(layoutname)
            End If

        Next
        ed.Regen()
        tr.Commit()
        'Rename Paperspace Tab
        tr = doc.TransactionManager.StartTransaction
        blktbl = db.BlockTableId.GetObject(OpenMode.ForRead)
        'blktblrec = tr.GetObject(blktbl(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
        Dim symtblenum As SymbolTableEnumerator = blktbl.GetEnumerator

        While symtblenum.MoveNext
            blktblrec = symtblenum.Current.GetObject(OpenMode.ForRead)
            If blktblrec.IsLayout Then
                If blktblrec.Name = "*Paper_Space" Then
                    Dim layoutmgr As LayoutManager = LayoutManager.Current
                    Dim layoutobj As Layout = tr.GetObject(blktblrec.LayoutId, OpenMode.ForRead)
                    If layoutobj.LayoutName = "Layout1" Then
                        layoutmgr.RenameLayout(layoutobj.LayoutName, "Master Layout")
                    ElseIf layoutobj.LayoutName = "Layout2" Then
                        layoutmgr.DeleteLayout(layoutobj.LayoutName)
                    End If
                End If
            End If
        End While
        ed.Regen()
        tr.Commit()
        'Set Limits and Zoom
        'db.Plimmin = New Point2d(0, 0)
        'db.Plimmax = New Point2d(914.4 * getscl.getobjectscale, 609.6 * getscl.getobjectscale)

        Dim view2 As ViewTableRecord = ed.GetCurrentView
        Dim point6 As New Point2d(914.4 * Getscl.Getobjectscale / 2, 609.6 * Getscl.Getobjectscale / 2)
        view2.CenterPoint = point6
        view2.Height = (609.6 * Getscl.Getobjectscale)
        view2.Width = (914.4 * Getscl.Getobjectscale)

        ed.SetCurrentView(view2)
        'Delete predefined viewports
        tr = doc.TransactionManager.StartTransaction
        blktbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        blktblrec = tr.GetObject(blktbl(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
        Dim blktblenum As BlockTableRecordEnumerator = blktblrec.GetEnumerator

        'While blktblenum.MoveNext
        '    Dim objid As ObjectId = blktblenum.Current
        '    Dim dbobj As DBObject = tr.GetObject(objid, OpenMode.ForWrite, False, True)
        '    'If TypeOf dbobj Is Viewport Then
        '    '    dbobj.Erase()
        '    'End If
        'End While
        tr.Commit()


        Application.SetSystemVariable("TILEMODE", 1)

    End Function

    ' Private Sub FrmStartup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '     Label3.Text = "Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build ' & "." & My.Application.Info.Version.Revision
    ' End Sub

    Private Sub FrmStartup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim now As DateTime = DateTime.Now
        Dim version As String = now.ToString("yy") & "." & now.ToString("MM") & "." & now.ToString("dd") & ".0"  ' Adjust revision as needed
        Label3.Text = "Version " & version
    End Sub

    Private Sub btnClient_Click(sender As Object, e As EventArgs) Handles btnClient.Click
        Dim MBoxResult As MsgBoxResult
        MsgBox("Warning, choosing Client Specific disables Herold customization" & vbLf & "Do you wish to continue?", vbYesNo, "Warning!")

        xd = New ClsXdata

        If MBoxResult = vbYes Then
            xd.Compilexdata("Client", "Client", "Client", "Client", "Client")
            Me.DialogResult = System.Windows.Forms.DialogResult.Abort
            Exit Sub
        End If
    End Sub
End Class