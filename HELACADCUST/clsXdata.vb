Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Windows

Public Class ClsXdata
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property


    Dim I As Integer
    Dim xdlay As AcadLayer

    Sub Writexdata(ByVal datatype As Object, ByVal data As Object)

        For Me.I = LBound(datatype) To UBound(datatype)

            If data(Me.I) = "" Then Exit For
        Next

        xdlay = Thisdrawing.Layers.Item(0)
        xdlay.SetXData(datatype, data)
        If UBound(datatype) > 2 Then readxdata()
    End Sub

    Public Sub Compilexdata(ByVal units As String, ByVal scale As String, ByVal txt As String, ByVal discp As String,
                            ByVal dimst As String, Optional ByVal syst As String = "psss")
        xdlay = Thisdrawing.Layers(0)
        Dim datatype(0 To 7) As Int16
        Dim data(0 To 7) As Object

        datatype(0) = 1001 : data(0) = "HEL-Setup"
        datatype(1) = 1000 : data(1) = "Txtht=" & txt
        datatype(2) = 1000 : data(2) = "Cursc=" & scale
        datatype(3) = 1000 : data(3) = "Systm=" & syst
        datatype(4) = 1000 : data(4) = "Units=" & units
        datatype(5) = 1000 : data(5) = "Discp=" & discp
        datatype(6) = 1000 : data(6) = "Scale=" & scale
        datatype(7) = 1000 : data(7) = "Dimst=" & dimst

        Writexdata(datatype, data)

        setuservariables(units, scale, txt, discp, dimst, syst)

    End Sub

    Public Sub Setuservariables(ByVal units As String, ByVal scale As String, ByVal txt As String,
                                ByVal discp As String, ByVal dimst As String, ByVal syst As String)
        Dim uni As Double
        Dim txtht As Double
        Dim scl As Double
        Dim result As Boolean

        If LCase(txt) = "client" Then
            '' ''txtht = 0
            Exit Sub
        Else
            '' ''txtht = CDec(txt)
            result = Double.TryParse(CDbl(txt), txtht)
            If result Then 'MsgBox("Decimal")
                Thisdrawing.SetVariable("Userr1", txtht)
            End If
        End If

        If LCase(dimst) = "client" Then
            Exit Sub
        Else

            Thisdrawing.SetVariable("Users2", dimst)
        End If

        result = Double.TryParse(CDbl(scale), scl)
        If result Then 'MsgBox("Decimal")
            Thisdrawing.SetVariable("userr2", scl)
        End If


        If units = "mts" Then
            uni = 1
        ElseIf units = "mms" Then
            uni = 1
        Else
            uni = 10 / 254
        End If

        Thisdrawing.SetVariable("userr3", uni)

        If LCase(discp) = "arch" Then
            Thisdrawing.SetVariable("Users3", "AR_")
        ElseIf lcase(discp) = "blen" Then
            Thisdrawing.SetVariable("Users3", "BE_")
        ElseIf lcase(discp) = "inds" Then
            Thisdrawing.SetVariable("Users3", "ST_")
        ElseIf lcase(discp) = "bldg" Then
            Thisdrawing.SetVariable("Users3", "ST_")
        End If

        Thisdrawing.SetVariable("Users1", discp)
        Thisdrawing.SetVariable("Users4", units)
        Thisdrawing.SetVariable("Users5", syst)
    End Sub

    <CommandMethod("TestXdata")> Public Sub TestXdata()
        compilexdata("Meters", "50", "2.5", "Indus", "Ticks", "psss")
    End Sub

    Function Readxdata(Optional ByVal strng As String = "None")
        Dim units As String = ""
        Dim scale As String = ""
        Dim txt As String = ""
        Dim discp As String = ""
        Dim dimst As String = ""
        Dim syst As String = ""
        Dim datatype(0 To 7) As Int16
        Dim data(0 To 7) As Object
        Dim modemacro As String = ""

        Readxdata = Nothing
        Try
            xdlay = Thisdrawing.Layers.Item(0)

            xdlay.GetXData("HEL-Setup", datatype, data)
            If LCase(data(1)) = "client" Then
                Thisdrawing.SetVariable("UserS3", "client")
                Exit Function
            End If

            For Me.I = LBound(datatype) To UBound(datatype)
                'If datatype(I) = "" Then datatype(I) = "blank"
                If Me.I <> 6 And Me.I <> 8 Then
                    modemacro = modemacro & data(Me.I) & " ; "
                End If
                If Me.I = 1 Then
                    txt = Strings.Right(data(Me.I), Strings.Len(data(Me.I)) - 6)
                    If LCase(strng) = "text" Then
                        Readxdata = data(Me.I)
                    End If
                ElseIf Me.I = 2 Then
                    scale = Strings.Right(data(Me.I), Strings.Len(data(Me.I)) - 6)
                    If LCase(strng) = "scale" Then
                        Readxdata = data(Me.I)
                    End If
                ElseIf Me.I = 3 Then
                    syst = Strings.Right(data(Me.I), Strings.Len(data(Me.I)) - 6)
                    If LCase(strng) = "systm" Then
                        Readxdata = data(Me.I)
                    End If
                ElseIf Me.I = 4 Then
                    units = Strings.Right(data(Me.I), Strings.Len(data(Me.I)) - 6)
                    If LCase(strng) = "units" Then
                        Readxdata = data(Me.I)
                    End If
                ElseIf Me.I = 5 Then
                    discp = Strings.Right(data(Me.I), Strings.Len(data(Me.I)) - 6)
                    If LCase(strng) = "discp" Then
                        Readxdata = data(Me.I)
                    End If
                ElseIf Me.I = 6 Then
                    'discp = Strings.Right(data(I), Strings.Len(data(I)) - 6)
                    If LCase(strng) = "bscal" Then
                        Readxdata = data(Me.I)
                    End If
                ElseIf Me.I = 7 Then
                    dimst = Strings.Right(data(Me.I), Strings.Len(data(Me.I)) - 6)
                    If LCase(strng) = "dimst" Then
                        Readxdata = data(Me.I)
                    End If
                End If
            Next
            Thisdrawing.SetVariable("modemacro", modemacro)
            Setuservariables(units, scale, txt, discp, dimst, syst)
        Catch

        End Try


    End Function

    Function Check_xdata()
        Dim datatype(0 To 7) As Int16
        Dim data(0 To 7) As Object
        Check_xdata = False
        Dim I As Integer
        Try

            xdlay = Thisdrawing.Layers.Item(0)


            xdlay.GetXData("HEL-Setup", datatype, data)
            If datatype Is Nothing Then
                Check_xdata = False
            Else
                For I = LBound(datatype) To UBound(datatype)
                    Check_xdata = True
                    Exit For
                Next


            End If
        Catch
            Check_xdata = False
        End Try

    End Function

    Function Check_client() As Boolean
        If LCase(Thisdrawing.GetVariable("userS3")) = "client" Or Thisdrawing.GetVariable("userS3") = "" Then
            Return True
        End If
        Return False
    End Function
End Class
