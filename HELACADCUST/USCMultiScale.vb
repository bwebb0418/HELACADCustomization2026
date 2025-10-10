Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors

Public Class USCMultiScale

    Private Sub Cmdconvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdconvert.Click
        Dim result As MsgBoxResult
        result = MsgBox("Changing to a MultiScale Drawing cannot be undone!" & vbCr & "Are you sure you wish to do this?",
                        vbYesNo & vbCritical & vbDefaultButton2, "Warning!")
        If result = vbNo Then
            Exit Sub
        End If

        'cmdconvert.Enabled = False

        Dim xdata As New ClsXdata
        Dim units As String
        Dim scale As String
        Dim txt As String
        Dim discp As String
        Dim dimst As String
        'Dim syst As String

        units = xdata.Readxdata("units")
        scale = xdata.Readxdata("scale")
        txt = xdata.Readxdata("text")
        discp = xdata.Readxdata("discp")
        dimst = xdata.Readxdata("dimst")
        'syst = xdata.readxdata("systm")

        units = Strings.Right((units), Strings.Len(units) - 6)
        scale = Strings.Right((scale), Strings.Len(scale) - 6)
        txt = Strings.Right((txt), Strings.Len(txt) - 6)
        discp = Strings.Right((discp), Strings.Len(discp) - 6)
        dimst = Strings.Right((dimst), Strings.Len(dimst) - 6)
        'syst = Strings.Right((syst), Strings.Len(syst) - 6)


        xdata.Compilexdata(units, scale, txt, discp, dimst, "psms")



        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction

        Dim dimsty As DimStyleTable = tr.GetObject(db.DimStyleTableId, OpenMode.ForRead)
        Dim dimstyrec As DimStyleTableRecord

        For Each dimstyid As ObjectId In dimsty
            dimstyrec = tr.GetObject(dimstyid, OpenMode.ForWrite)
            If dimstyrec.Name.StartsWith("HEL") Then
                dimstyrec.Name = dimstyrec.Name & "-" & scale
            End If
        Next

        Dim txtsty As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
        Dim txtstyrec As TextStyleTableRecord

        For Each txtstyid As ObjectId In txtsty
            txtstyrec = tr.GetObject(txtstyid, OpenMode.ForWrite)
            If txtstyrec.Name.StartsWith("HEL") And Not txtstyrec.Name.Contains("sym") And Not txtstyrec.Name.Contains("Dim") And Not txtstyrec.Name.Contains("Weld") Then
                txtstyrec.Name = txtstyrec.Name & "-" & scale
            End If
        Next
        tr.Commit()


        'End If

        'If dimstid.IsNull = True Then
        '    Str = New DimStyleTableRecord
        'Else
        '    Str = tr.GetObject(dimstid, OpenMode.ForWrite)
        'End If

        rdo11.Enabled = True
    End Sub

    Private Sub USCMultiScale_Disposed(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.Windows.PalettePersistEventArgs) Handles Me.Disposed
        e.ConfigurationSection.WriteProperty("MultiScale", 25)
    End Sub

    'Private Sub USCMultiScale_Load(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.Windows.PalettePersistEventArgs) Handles Me.Load
    '    Dim a As Double = CType(e.ConfigurationSection.ReadProperty("MultiScale", 25), Double)
    'End Sub

    Function Addscale(ByVal scale As String)

        Dim xdata As New ClsXdata
        Dim units As String
        Dim txt As String
        Dim discp As String
        Dim dimst As String
        ' Dim syst As String

        units = xdata.Readxdata("units")
        txt = xdata.Readxdata("text")
        discp = xdata.Readxdata("discp")
        dimst = xdata.Readxdata("dimst")
        'syst = xdata.readxdata("systm")

        units = Strings.Right((units), Strings.Len(units) - 6)
        'scale = Strings.Right((scale), Strings.Len(scale) - 6)
        txt = Strings.Right((txt), Strings.Len(txt) - 6)
        discp = Strings.Right((discp), Strings.Len(discp) - 6)
        dimst = Strings.Right((dimst), Strings.Len(dimst) - 6)
        'syst = Strings.Right((syst), Strings.Len(syst) - 6)


        xdata.Compilexdata(units, scale, txt, discp, dimst, "psms")
        xdata.Readxdata()
        Addscale = Nothing

    End Function

    Private Sub Rdo11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdo11.CheckedChanged
        Dim vtd As New ClsVTD
        rdo11.Checked = False
        rdo11.BackColor = Drawing.Color.DarkGray
        'Addscale(1)
        'vtd.SetSystemVariables()
        'vtd.Text_Gen(, 1)
        'vtd.Dim_Gen(1)


    End Sub
End Class
