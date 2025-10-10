Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Windows
Public Class Clscodetest
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    Public xdata As New ClsXdata

    'All code in this module is for testing/debugging purposes only

    <CommandMethod("UV")> Public Sub UV()
        'Reads the User Variables Stored in the Drawing.
        Dim retval As String = ""
        Dim I As Object
        With Thisdrawing
            For I = 1 To 5
                retval = retval & "UserI" & (I) & " = " & .GetVariable("UserI" & (I)) & vbCr
            Next I
            For I = 1 To 5
                retval = retval & "UserR" & (I) & " = " & .GetVariable("UserR" & (I)) & vbCr
            Next I
            For I = 1 To 5
                retval = retval & "UserS" & (I) & " = " & .GetVariable("UserS" & (I)) & vbCr
            Next I

        End With
        MsgBox(retval)
    End Sub


    <CommandMethod("CXD")> Public Sub CXD()
        'This function WILL Clear any XData Set by the HEL Setup Routine
        Dim DataType(1) As Int16
        Dim data(1) As Object
        DataType(0) = 1001 : data(0) = "HEL-Setup"
        DataType(1) = 1000 : data(1) = "0"
        xdata.Writexdata(DataType, data)
    End Sub

    <CommandMethod("xd")> Public Sub Xd()
        'Quick function to read the xdata through the XDtoLSP Override
        xdata.Readxdata()
    End Sub


    <CommandMethod("laycount")> Public Sub Laycount()
        MsgBox(Thisdrawing.Layers.Count)
    End Sub

    Public Function DeleteObjectFromBlock(ByVal ent As AcadEntity) As Long

        Dim blkDef As AcadBlock

        blkDef = Thisdrawing.ObjectIdToObject(ent.OwnerID)
        ent.Delete()
        DeleteObjectFromBlock = blkDef.Count

    End Function

    Function Find_dict()
        On Error Resume Next
        Dim dicts As AcadDictionaries
        Dim dict As Object
        dicts = Thisdrawing.Dictionaries
        For Each dict In dicts
            'Debug.Print(dict.Name)
        Next

    End Function

    Function TOOLBAR_OFF()
        TOOLBAR_OFF = Nothing

        Dim HELtb As AcadToolbar
        Dim HELmnu As AcadMenuGroup

        HELmnu = Thisdrawing.Application.MenuGroups(CType("HEROLDM", Index))
        For Each HELtb In HELmnu.Toolbars
            HELtb.Visible = False
        Next HELtb
    End Function

    Function TOOLBAR_ON()
        TOOLBAR_ON = Nothing
        Dim HELtb As AcadToolbar
        Dim HELmnu As AcadMenuGroup

        HELmnu = Thisdrawing.Application.MenuGroups(CType("HEROLDM", Index))
        For Each HELtb In HELmnu.Toolbars
            HELtb.Visible = True
        Next HELtb
    End Function

    <CommandMethod("cuv")> Public Sub Cuv()
        Dim I As Integer
        With Thisdrawing
            For I = 1 To 5
                .SetVariable("UserI" & (I), vbEmpty)
            Next I
            For I = 1 To 5
                .SetVariable("UserR" & (I), vbEmpty)
            Next I
            For I = 1 To 5
                .SetVariable("UserS" & (I), "")
            Next I
        End With
    End Sub


    <CommandMethod("getpoint")> Public Sub Getpoints()
        On Error GoTo nd
        Dim points As Object
        Dim I, J, K As Integer
        I = 0
        J = 1
        K = 2
        Do
            points = Thisdrawing.Utility.GetPoint(, "Pick point to map")
            Thisdrawing.Utility.Prompt("PT(" & I & ") = ;" & points(0) & ":PT(" & J & ") = ;" & points(1) & ":PT(" & K & ") = ;" & points(2))
            I += 3
            J += 3
            K += 3
        Loop

nd:
    End Sub



End Class
