Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports MgdAcApplication = Autodesk.AutoCAD.ApplicationServices.Application
''' <summary>
''' Utility class providing common functions for AutoCAD operations.
''' Includes user input, calculations, and system variable handling.
''' </summary>

    ''' <summary>
    ''' Gets the active AutoCAD document.
    ''' </summary>
    ''' <returns>The active AcadDocument.</returns>
Public Class Clsutilities
    Function WritetoCMD(TEXTLINE As String)
        Thisdrawing.Utility.Prompt(TEXTLINE & vbCrLf)
        WritetoCMD = Nothing
    End Function
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Try
                Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
            Catch ex As Exception
                Thisdrawing = Nothing
            End Try
        End Get
    End Property

    Function Radtodeg(ByVal RADROT)
        Dim pi
        pi = 4 * Math.Atan(1)
        Radtodeg = RADROT * (180 / pi)
    End Function

    Function Degtorad(ByVal degrot)
        Dim pi
        pi = 4 * Math.Atan(1)
        Degtorad = (degrot * pi) / 180
    End Function

    Function Getpoint(ByVal Prompt As String)
        On Error GoTo errhandler
        Getpoint = Thisdrawing.Utility.GetPoint(, Prompt)
        Exit Function

errhandler:


    End Function

    Function Getendpoint(ByVal inspoint As Object, ByVal Prompt As String)
        On Error GoTo errhandler
        Getendpoint = Thisdrawing.Utility.GetPoint(inspoint, Prompt)
        Exit Function

errhandler:


    ''' <summary>
    ''' Determines the appropriate text style based on active space and user variables.
    ''' </summary>
    ''' <returns>The text style name.</returns>
    End Function

    Function Getrotation(ByVal inspoint As Object, ByVal secpoint As Object)
        Dim line As AcadLine
        line = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
        Getrotation = line.Angle
        line.Delete()
    End Function

    Function Gettxtstyle()
        Try
            If Thisdrawing.ActiveSpace = AcActiveSpace.acPaperSpace Then
                Gettxtstyle = "PHEL" & (Application.GetSystemVariable("Userr1") * 10)
            Else
                Gettxtstyle = "HEL" & (Application.GetSystemVariable("Userr1") * 10)
            End If
            Dim check As New ClsCheckroutines
            If check.Txtchk(Gettxtstyle) = False Then
                Dim writetocmd As New Clsutilities
                writetocmd.WritetoCMD("Herold Text Style not found, using current text style")
                Gettxtstyle = Thisdrawing.ActiveTextStyle.Name
            End If
        Catch
            Dim writetocmd As New Clsutilities
            writetocmd.WritetoCMD("Herold Text Style not found, using current text style")
            Gettxtstyle = Thisdrawing.ActiveTextStyle.Name
        End Try
    End Function

    Function Getscale() As Double
        Dim units
        Dim dimscl


        units = Application.GetSystemVariable("Userr3")
        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            dimscl = Application.GetSystemVariable("Userr2")
        Else
            dimscl = 1
        End If

        If units = 0 Then units = 1
        If dimscl = 0 Then dimscl = 1

        Getscale = units * dimscl

    End Function

    Function Getobjectscale() As Double
        Dim units

        units = Application.GetSystemVariable("Userr3")

        If units = 0 Then units = 1

        Getobjectscale = units

    End Function

    Function Getlength(ByVal inspoint, ByVal secpoint)
        Dim line As AcadLine
        line = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
        Getlength = line.Length
        line.Delete()

    End Function

    '' ''Function layerexport()
    '' ''    Dim xlapp As Excel.Application
    '' ''    Dim xlbook As Excel.Workbook
    '' ''    Dim xlsheet As Excel.Worksheet
    '' ''    Dim layers As AcadLayers
    '' ''    Dim layer As AcadLayer

    '' ''    xlbook = GetObject("M:\custom\development\r2007\Layers.xls")
    '' ''    xlapp = xlbook.Parent
    '' ''    layers = Thisdrawing.layers
    '' ''    xlsheet = xlbook.Sheets("sheet1")

    '' ''    I = 0
    '' ''    For Each layer In layers
    '' ''        I = I + 1
    '' ''        xlsheet.Cells(I, 1) = layer.Name
    '' ''        xlsheet.Cells(I, 2) = layer.Linetype
    '' ''        xlsheet.Cells(I, 3) = layer.color
    '' ''    Next layer

    '' ''    xlbook.save()

    '' ''End Function

    <CommandMethod("bounding")> Public Sub Boundingbox()
        Dim obj As AcadObject = Nothing
        Dim point(0 To 2) As Double
        Dim maxpoint As Object = Nothing
        Dim minpoint As Object = Nothing

        Thisdrawing.Utility.GetEntity(obj, point, "Pick Object")
        obj.getboundingbox(minpoint, maxpoint)
        'MsgBox(point(0) & "-" & point(1))
        'MsgBox(minpoint(0) & "-" & minpoint(1))
        'MsgBox(maxpoint(0) & "-" & maxpoint(1))

        Thisdrawing.ModelSpace.AddLine(minpoint, maxpoint)
    End Sub
    <CommandMethod("dwgline")> Sub Linerot()
        Dim line As AcadLine
        Dim inspoint As Object
        Dim secpoint As Object

        inspoint = (Thisdrawing.Utility.GetPoint(, "Pick First Point"))
        secpoint = Thisdrawing.Utility.GetPoint(, "Pick Second Point")
        line = Thisdrawing.ModelSpace.AddLine(inspoint, secpoint)
    End Sub
    <CommandMethod("profilename")> Public Sub Profilename()
        Dim prefs As AcadPreferences = Application.Preferences
        MsgBox(prefs.Profiles.ActiveProfile)
    End Sub

    <CommandMethod("tm")> Sub Tilemode()
        If Thisdrawing.ActiveLayout.Name = "Model" Then
            Thisdrawing.SetVariable("tilemode", 0)
        Else
            Thisdrawing.SetVariable("tilemode", 1)
        End If
    End Sub

    Function Loadlinetype(ByVal linetype As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Const filename As String = "Acad.lin"
        Using tr As Transaction = db.TransactionManager.StartTransaction
            Try
                Dim path As String = HostApplicationServices.Current.FindFile(filename, db, FindFileHint.Default)
                db.LoadLineTypeFile(linetype, path)
            Catch ex As Exception
                'ed.WriteMessage(linetype & " : " & ex.Message)

                'If ex.ErrorStatus = ErrorStatus.DuplicateRecordName Then

                'End If
            End Try

            tr.Commit()
        End Using
        Loadlinetype = Nothing

    End Function

    Public Function Serialnumber()

        Serialnumber = Application.GetSystemVariable("_PKSER")

        'MsgBox(product)

    End Function
End Class
