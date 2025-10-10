Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry

Public Class frmlumber
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        With Me.CboDLSIZE.Items
            .Add("38x38 / 2x2")
            .Add("38x64 / 2x3")
            .Add("38x89 / 2x4")
            .Add("38x140 / 2x6")
            .Add("38x184 / 2x8")
            .Add("38x235 / 2x10")
            .Add("38x286 / 2x12")
            .Add("64x64 / 3x3")
            .Add("64x89 / 3x4")
            .Add("64x140 / 3x6")
            .Add("64x184 / 3x8")
            .Add("64x235 / 3x10")
            .Add("64x286 / 3x12")
            .Add("89x89 / 4x4")
            .Add("89x140 / 4x6")
            .Add("89x184 / 4x8")
            .Add("89x235 / 4x10")
            .Add("89x286 / 4x12")
            .Add("140x140 / 6x6")
            .Add("140x184 / 6x8")
            .Add("140x235 / 6x10")
            .Add("140x286 / 6x12")
            .Add("140x336 / 6x14")
            .Add("140x387 / 6x16")
            .Add("184x184 / 8x8")
            .Add("184x235 / 8x10")
            .Add("184x286 / 8x12")
            .Add("184x336 / 8x14")
            .Add("184x387 / 8x16")
            .Add("235x235 / 10x10")
            .Add("235x286 / 10x12")
            .Add("235x336 / 10x14")
            .Add("235x387 / 10x16")
            .Add("286x286 / 12x12")
            .Add("286x336 / 12x14")
            .Add("286x387 / 12x16")
        End With

        With Me.CboRSSIZE.Items
            .Add("51x51 / 2x2")
            .Add("51x76 / 2x3")
            .Add("51x102 / 2x4")
            .Add("51x152 / 2x6")
            .Add("51x203 / 2x8")
            .Add("51x254 / 2x10")
            .Add("51x305 / 2x12")
            .Add("76x76 / 3x3")
            .Add("76x102 / 3x4")
            .Add("76x152 / 3x6")
            .Add("76x203 / 3x8")
            .Add("76x254 / 3x10")
            .Add("76x305 / 3x12")
            .Add("102x102 / 4x4")
            .Add("102x152 / 4x6")
            .Add("102x203 / 4x8")
            .Add("102x254 / 4x10")
            .Add("102x305 / 4x12")
            .Add("152x152 / 6x6")
            .Add("152x203 / 6x8")
            .Add("152x254 / 6x10")
            .Add("152x305 / 6x12")
            .Add("152x356 / 6x14")
            .Add("152x406 / 6x16")
            .Add("203x203 / 8x8")
            .Add("203x254 / 8x10")
            .Add("203x305 / 8x12")
            .Add("203x356 / 8x14")
            .Add("203x406 / 8x16")
            .Add("254x254 / 10x10")
            .Add("254x305 / 10x12")
            .Add("254x356 / 10x14")
            .Add("254x406 / 10x16")
            .Add("305x305 / 12x12")
            .Add("305x356 / 12x14")
            .Add("305x406 / 12x16")
        End With

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Function Mettoimp(ByVal size)
        Mettoimp = ""
        Select Case size
            Case 38
                Mettoimp = 1.5
            Case 51
                Mettoimp = 2
            Case 64
                Mettoimp = 2.5
            Case 76
                Mettoimp = 3
            Case 89
                Mettoimp = 3.5
            Case 102
                Mettoimp = 4
            Case 140
                Mettoimp = 5.5
            Case 152
                Mettoimp = 6
            Case 184
                Mettoimp = 7.25
            Case 203
                Mettoimp = 8
            Case 235
                Mettoimp = 9.25
            Case 254
                Mettoimp = 10
            Case 286
                Mettoimp = 11.25
            Case 305
                Mettoimp = 12
            Case 337
                Mettoimp = 13.25
            Case 356
                Mettoimp = 14
            Case 387
                Mettoimp = 15.25
            Case 406
                Mettoimp = 16
        End Select
        Return Mettoimp
    End Function
    '' ''<CommandMethod("Timber")> Public Sub showtimbermenu()
    '' ''    Me.Show()
    '' ''End Sub
End Class