'Imports System.Runtime
'Imports System.Diagnostics
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
'Imports MS.Internal.WindowsRuntime.Windows
'Imports System.IO
'Imports Autodesk.AutoCAD.Customization



''' <summary>
''' Main customization class for HELACAD AutoCAD plugin.
''' Implements IExtensionApplication to integrate with AutoCAD's lifecycle.
''' Handles events, commands, and UI for structural engineering workflows.
''' </summary>

Public Class Customization



    Implements Autodesk.AutoCAD.Runtime.IExtensionApplication

    Public WithEvents Thisdrawing As AcadDocument = DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)

    Public WithEvents Docs As DocumentCollection = Application.DocumentManager

    Public layerfrm As Frmlayergen

    Public startupfrm As FrmStartup

    Public bindpurge As Clsbind_purge
    Public Shared curlayer As String

    Public blockroutines As ClsBlockRoutines

    Friend Shared m_ps As PaletteSet = Nothing

    Dim checkxdata As ClsXdata
    Dim Strtuproutine As ClsVTD
    Dim attswitch As Boolean = False
    Dim clientswitch As Boolean
    ReadOnly WritetoCMD As New Clsutilities


    '<CommandMethod("HELVER")> Sub HELVERSION()
    '    Dim splash As New spsHEL
    '    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim ed As Editor = doc.Editor
    '    splash.Show()
    '    ed.WriteMessage(vbLf & "HEL Customization Loaded" & vbLf & "Version {0}.{1}.{2}.{3}" & vbLf, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)
    '    If splash.DialogResult = System.Windows.Forms.DialogResult.OK Then
    '        splash.Dispose()
    '        splash = Nothing
    '    End If

    'End Sub


    Private Sub Docs_DocumentActivated(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.ApplicationServices.DocumentCollectionEventArgs) Handles Docs.DocumentActivated
        Try

            If Application.DocumentManager.MdiActiveDocument IsNot Nothing Then
                Thisdrawing = DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)

            End If
        Catch

        Finally
            If License_Check() = True Then
                ' WritetoCMD.WritetoCMD("License is valid")
            Else

            End If
        End Try
    End Sub

    Private Sub Thisdrawing_Activate() Handles Thisdrawing.Activate

        checkxdata = New ClsXdata
        If checkxdata.Check_xdata = True Then
            checkxdata.Readxdata()
            'layerfrm = New frmlayergen
            'layerfrm.newlayerconvert()
            'blockroutines.attfix()
        ElseIf checkxdata.Check_xdata = False Then
            'DrawingStartup()
        End If


    End Sub

    Private Sub Thisdrawing_BeginCommand(ByVal CommandName As String) Handles Thisdrawing.BeginCommand
        'Dim hkey, key, shl
        Dim retval As MsgBoxResult
        Dim path = Thisdrawing.Path
        Dim userpath = Environ("Appdata") & "\Autodesk\AutoCAD 2023\R24.2\enu\Plotters\Plot Styles"
        Dim objshell
        Dim objLink
        'Dim objscr
        Dim chkrtine As New ClsCheckroutines
        'Dim profile
        Dim prefs As AcadPreferences = Application.Preferences
        Dim dimst As String = ""
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim check_routines As New ClsCheckroutines

        Dim serial As New Clsutilities

        '        If serial.Serialnumber = "396-48853632" Then Exit Sub
        'MsgBox(CommandName)

        'COMMANDS THAT ACCESS THE REGISTRY
        If CommandName = "QSAVE" Or CommandName = "SAVE" Or CommandName = "SAVEAS" Then
            Dim profile2 As IConfigurationSection = Application.UserConfigurationManager.OpenGlobalSection
            profile2.OpenSubsection("FixedProfile").OpenSubsection("General").WriteProperty("ACETMOVEBAK", "H:\BACKUP DRAWING FILES")
        ElseIf CommandName = "NEW" Then
            Dim profile2 As IConfigurationSection = Application.UserConfigurationManager.OpenGlobalSection
            profile2.OpenSubsection("FixedProfile").OpenSubsection("General").WriteProperty("ACETMOVEBAK", "H:\BACKUP DRAWING FILES")

        ElseIf CommandName = "REVCLOUD" Then

            Dim profile2 As IConfigurationSection = Application.UserConfigurationManager.OpenCurrentProfile

            profile2.OpenSubsection("Variables").WriteProperty("*REVCLOUDMAXARCLENGTH", 7 * Thisdrawing.GetVariable("USERR3"))
            profile2.OpenSubsection("Variables").WriteProperty("*REVCLOUDMINARCLENGTH", 5 * Thisdrawing.GetVariable("USERR3"))

            doc.Database.Clayer = check_routines.Layercheck("~-LREV")
            'Thisdrawing.ActiveLayer = Thisdrawing.Layers("~-LREV")
        ElseIf CommandName = "PLOT" Then

            'Thisdrawing.SendCommand("TEXTTOFRONT " & "A ")
            ' ***The following code will set the printer paths to the proper directories to ensure that
            ' ***everyone is using the same PC3, PMP and CTB files all the time
            Dim profile2 As IConfigurationSection = Application.UserConfigurationManager.OpenCurrentProfile

            'bindpurge = New clsbind_purge
            'bindpurge.preplotchk()

            'For I = 0 To Len(Thisdrawing.Path)
            '    If Left(Right(Thisdrawing.Path, I), 12) = "04S Drawings" Then
            '        path = Left(Thisdrawing.Path, Len(Thisdrawing.Path) - I + 12)
            '        Exit For
            '    Else
            '        path = Thisdrawing.Path
            '    End If
            'Next I

            If Thisdrawing.GetVariable("Dwgtitled") = 0 Then
                retval = vbNo
            ElseIf Thisdrawing.GetVariable("useri1") <> 7 And Thisdrawing.GetVariable("useri1") <> 6 Then
                retval = MsgBox("Does this drawing require a Client Specific Pentable?" _
                    & vbCr & "Client Specific Pentables must be stored in the folder:" & vbCr &
                    path, vbYesNo, "CTB Directory Selection")
                Thisdrawing.SetVariable("useri1", retval)
            Else
                retval = Thisdrawing.GetVariable("userI1")
            End If
            Adjusttitlelines()
            Dim OBJ As Object

            If chkrtine.Laychk("~-VPORT") = False Then
                Dim vportlay As AcadLayer
                vportlay = Thisdrawing.Layers.Add("~-VPORT")
                vportlay.Plottable = False
            Else
                Thisdrawing.Layers.Item("~-VPORT").Freeze = False
                Thisdrawing.Layers.Item("~-VPORT").LayerOn = True
                Thisdrawing.Layers.Item("~-VPORT").Lock = False
            End If
            For Each OBJ In Thisdrawing.PaperSpace
                If TypeOf OBJ Is AcadPViewport Then
                    OBJ.Layer = "~-vport"
                End If
            Next OBJ

            'folder setup
            Try
                For Each Filefound As String In IO.Directory.GetFiles(userpath, "*.lnk")
                    My.Computer.FileSystem.DeleteFile(Filefound)
                Next

                For Each Filefound As String In IO.Directory.GetFiles(userpath, "*.ctb")
                    My.Computer.FileSystem.DeleteFile(Filefound)
                Next

                For Each Filefound As String In IO.Directory.GetFiles(userpath, "*.stb")
                    My.Computer.FileSystem.DeleteFile(Filefound)
                Next

            Catch ex As Exception

            Finally
            End Try


            If retval = vbNo Then

                Try

                    objshell = CreateObject("WScript.Shell")
                    objLink = objshell.CreateShortcut(userpath & "\HELCTB.lnk")
                    'objLink.Arguments = Environ("ALLUSERSPROFILE") & "\Autodesk\ApplicationPlugins\helacad2023.bundle\Contents\Resources\Plotters\CTB\"
                    objLink.TargetPath = Environ("ALLUSERSPROFILE") & "\Autodesk\ApplicationPlugins\helacad2023.bundle\Contents\Resources\Plotters\CTB\"
                    objLink.Save
                Catch ex As Exception

                End Try

            Else

                Try

                    objshell = CreateObject("WScript.Shell")
                    objLink = objshell.CreateShortcut(userpath & "\CustomPentable.lnk")
                    'objLink.Arguments = Environ("ALLUSERSPROFILE") & "\Autodesk\ApplicationPlugins\helacad2023.bundle\Contents\Resources\Plotters\CTB\"
                    objLink.TargetPath = path
                    objLink.Save
                    objLink = Nothing
                    objLink = objshell.CreateShortcut(userpath & "\HELCTB.lnk")
                    objLink.TargetPath = Environ("ALLUSERSPROFILE") & "\Autodesk\ApplicationPlugins\helacad2023.bundle\Contents\Resources\Plotters\CTB\"
                    objLink.Save
                    objLink = Nothing
                    objshell = Nothing
                Catch ex As Exception

                End Try

            End If


        If prefs.Files.PrinterStyleSheetPath <> userpath Then
                If path = "Null" Then WritetoCMD.WritetoCMD("Setting path to user profile")
                prefs.Files.PrinterStyleSheetPath = userpath
                objshell = CreateObject("WScript.Shell")
                objshell.SendKeys("{ESC}")
                Try
                    objshell.run(Environ("ALLUSERSPROFILE") & "\Autodesk\ApplicationPlugins\helacad2023.bundle\Contents\Resources\Plotters\plotscript2.vbs")
                Catch ex As Exception
                    WritetoCMD.WritetoCMD("Local script error, trying network")
                    objshell.run("m:\structural\r2023\Plotters\plotscript2.vbs")
                End Try

            End If


        End If


        If CommandName = "UNDO" Then
            Exit Sub


        ElseIf CommandName = "MLEADER" Then
            Try
                Dim txt As Double
                Dim txtlay As String = ""

                txt = Thisdrawing.GetVariable("UserR1") 'Default Text Height as Selected by User
                If txt = "2" Then
                    txtlay = "~-TXT20"
                ElseIf txt = "2.5" Then
                    txtlay = "~-TXT25"
                End If
                check_routines.SetLayerCurrent(txtlay)
                Dim ml_gen As New ClsVTD
                ml_gen.ML_Gen()
            Catch ex As Exception

            End Try
        ElseIf Left(CommandName, 3) = "DIM" Or CommandName.Contains("DIM_FOR_GRIP") Then
            If curlayer = "" Then curlayer = Thisdrawing.ActiveLayer.Name
            Try
                Thisdrawing.ActiveLayer = Thisdrawing.Layers.Item("~-TDIM")
                If Thisdrawing.GetVariable("Users2") = "arrws" Then
                    If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
                        dimst = "HEL-Arrows"
                    Else
                        dimst = "PHEL-Arrows"
                    End If
                ElseIf Thisdrawing.GetVariable("Users2") = "ticks" Then
                    If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
                        dimst = "HEL-Ticks"
                    Else
                        dimst = "PHEL-Ticks"
                    End If
                Else
                    dimst = Thisdrawing.ActiveDimStyle.Name
                End If
                Thisdrawing.ActiveDimStyle = Thisdrawing.DimStyles.Item(dimst)
            Catch
                dimst = Thisdrawing.ActiveDimStyle.Name

            Finally
                Thisdrawing.ActiveDimStyle = Thisdrawing.DimStyles.Item(dimst)
            End Try

        End If

    End Sub
    <CommandMethod("typetag")> Public Sub Runtypeform()
        Dim typetagfrm As New FrmTypetag
        typetagfrm.Show()
    End Sub


    <CommandMethod("laygen")> Sub LayerGen()
        If LayerGenerator() = True Then
            WritetoCMD.WritetoCMD("Layer Generation Completed")
        Else
            WritetoCMD.WritetoCMD("Layer Generation Not Completed")
        End If
    End Sub

    Public Function LayerGenerator() As Boolean
        checkxdata = New ClsXdata
        If checkxdata.Check_xdata = False Then
            LayerGenerator = False
            Exit Function
        End If

        layerfrm = New Frmlayergen
        layerfrm.ShowDialog()
        If layerfrm.DialogResult = System.Windows.Forms.DialogResult.OK Then
            LayerGenerator = True
        Else
            LayerGenerator = False
        End If
    End Function

    <CommandMethod("ChangeDWGStartup")> Sub ChangeDrawingStartup()

        checkxdata = New ClsXdata
        Strtuproutine = New ClsVTD
        If checkxdata.Check_xdata = True Then
            Dim MSGBOXRES As MsgBoxResult = MsgBox("Do you want to rerun the drawing setup?", vbYesNo, "Change Drawing Setup?")
            If MSGBOXRES = vbYes Then
                GoTo rerun
            End If
            checkxdata.Readxdata()
            layerfrm = New Frmlayergen
            'layerfrm.Newlayerconvert()
            blockroutines = New ClsBlockRoutines
            blockroutines.Attfix()
        ElseIf checkxdata.Check_xdata = False Then
rerun:
            startupfrm = New FrmStartup
            startupfrm.ShowDialog()
            If startupfrm.DialogResult = System.Windows.Forms.DialogResult.OK Then
                LayerGenerator()
                Strtuproutine.Text_Gen()
                If Strtuproutine.SetSystemVariables = False Then
                    Strtuproutine.SetSystemVariables()
                End If
                Strtuproutine.Dim_Gen()
                Thisdrawing.Utility.Prompt("Drawing Startup Complete" & vbCrLf)
            ElseIf startupfrm.DialogResult = System.Windows.Forms.DialogResult.Abort Then
                Thisdrawing.Utility.Prompt("Drawing Startup Aborted" & vbCrLf)
            ElseIf startupfrm.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
                Thisdrawing.Utility.Prompt("Draw ing Startup Aborted" & vbCrLf)
            End If
    ''' <summary>
    ''' Runs the full drawing startup process, including layer/text/dimension generation.
    ''' Prompts user for customization usage and sets up the drawing.
    ''' </summary>
        End If


        Layoutswitch()
    End Sub

    <CommandMethod("DWGStartup")> Sub DrawingStartup()

        checkxdata = New ClsXdata
        Strtuproutine = New ClsVTD
        If checkxdata.Check_xdata = True Then
            checkxdata.Readxdata()
            layerfrm = New Frmlayergen
            'layerfrm.Newlayerconvert()
            blockroutines = New ClsBlockRoutines
            blockroutines.Attfix()
            Exit Sub

        Else
            Dim response As MsgBoxResult
            response = MsgBox("Do you want to use the Herold customization?", vbYesNo, "Use Herold")
            If response = vbNo Then
                Dim Dummy_XData As New ClsXdata
                Dummy_XData.Compilexdata("Client", "Client", "Client", "Client", "Client")
                WritetoCMD.WritetoCMD("Herold Customization disabled for this drawing.")
                Exit Sub
            End If

            startupfrm = New FrmStartup
            startupfrm.ShowDialog()
        End If

        If startupfrm.DialogResult = System.Windows.Forms.DialogResult.OK Then
            If LayerGenerator() = True Then
                WritetoCMD.WritetoCMD("Layer Generation Complete")
            Else
                WritetoCMD.WritetoCMD("Layer Generation not Complete")
            End If

            If Strtuproutine.Text_Gen() = True Then
                WritetoCMD.WritetoCMD("Text Generation Complete")
            Else
                WritetoCMD.WritetoCMD("Text Generation not Complete")
                WritetoCMD.WritetoCMD("Drawing Startup Not Complete")
                Exit Sub
            End If

            If Strtuproutine.SetSystemVariables = True Then
                WritetoCMD.WritetoCMD("Setting System Variables Complete")
            Else
                WritetoCMD.WritetoCMD("Setting System Variables Not Complete")
                WritetoCMD.WritetoCMD("Drawing Startup Not Complete")
                Exit Sub
            End If
            If Strtuproutine.Dim_Gen() = True Then
                WritetoCMD.WritetoCMD("Dimension Generation Complete")
            Else
                WritetoCMD.WritetoCMD("Dimension Generation Not Complete")
                WritetoCMD.WritetoCMD("Drawing Startup Not Complete")
                Exit Sub
            End If
            If Strtuproutine.ML_Gen = True Then
                WritetoCMD.WritetoCMD("MLeader Generation Complete")
            Else
                WritetoCMD.WritetoCMD("MLeader Generation Not Complete")
                WritetoCMD.WritetoCMD("Drawing Startup Not Complete")
                Exit Sub
            End If
            WritetoCMD.WritetoCMD("Drawing Startup Complete")
        ElseIf startupfrm.DialogResult = System.Windows.Forms.DialogResult.Abort Then
            Thisdrawing.Utility.Prompt("Drawing Startup Aborted" & vbCrLf)
        ElseIf startupfrm.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
            Thisdrawing.Utility.Prompt("Drawing Startup Aborted" & vbCrLf)

    ''' <summary>
    ''' Initializes the customization on AutoCAD startup.
    ''' Loads menus, sets trusted paths, and displays version info.
    ''' </summary>
        End If


        Layoutswitch()
    End Sub

    Public Sub Initialize() Implements Autodesk.AutoCAD.Runtime.IExtensionApplication.Initialize
        LoadHELMenus()

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        doc.SendStringToExecute("Cleanscreenon ", True, False, False)
        doc.SendStringToExecute("Cleanscreenoff ", True, False, False)
        Dim ed As Editor = doc.Editor
        ed.WriteMessage(vbLf & "HEL Customization Loaded" & vbLf & "Version {0}.{1}.{2}" & vbLf, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build) ', My.Application.Info.Version.Revision)
        Application.SetSystemVariable("trustedpaths", "C:\ProgramData\AutoDesk\ApplicationPlugins\helacad2023.bundle\Contents...;M:\Structural\r2023\...")
    End Sub


    Public Sub Terminate() Implements Autodesk.AutoCAD.Runtime.IExtensionApplication.Terminate

    End Sub

    Private Sub Thisdrawing_BeginSave(ByVal FileName As String) Handles Thisdrawing.BeginSave

    End Sub

    Private Sub Thisdrawing_Deactivate() Handles Thisdrawing.Deactivate

    End Sub

    Private Sub Thisdrawing_EndCommand(ByVal CommandName As String) Handles Thisdrawing.EndCommand
        Dim Client As New ClsXdata
        If Client.Check_client = True Then


        ElseIf CommandName = "XREF" Or CommandName = "SAVE" Or CommandName = "QSAVE" Then
            Dim xref As AcadBlock
            Dim blkref As AcadBlockReference
            Dim handle
            On Error Resume Next
            For Each xref In Thisdrawing.Blocks
                If xref.IsXRef Then
                    handle = xref.Name
                    For Each OBJ In Thisdrawing.ModelSpace
                        If TypeOf OBJ Is AcadBlockReference Then
                            blkref = OBJ
                            If blkref.Name = handle Then
                                blkref.Layer = "~-XREF"
                            End If
                        End If
                    Next
                End If
            Next
            'End If

        ElseIf CommandName = "ATTEDIT" Or CommandName = "EATTEDIT" Or CommandName = "DDEDIT" Or CommandName = "FIND" Then
            'Adjusttitlelines()
            attswitch = True
        ElseIf CommandName = "MTEXT" Or CommandName = "TEXT" Or Strings.Left(CommandName, 3) = "DIM" Or CommandName.Contains("DIM_FOR_GRIP") Then
            If curlayer <> "" Then
                Dim checkroutines As New ClsCheckroutines
                checkroutines.SetLayerCurrent(curlayer)
                curlayer = ""
            End If
        ElseIf CommandName = "LAYOUT_CONTROL" Or CommandName = "TM" Then
            'Thisdrawing.Regen(AcRegenType.acAllViewports)
            'Thisdrawing.SendCommand("Regenall ")
        ElseIf CommandName = "SAVEAS" Then
            Thisdrawing.SetVariable("USERI1", 0)
        End If


    End Sub

    Private Sub Thisdrawing_LayoutSwitched(ByVal LayoutName As String) Handles Thisdrawing.LayoutSwitched
        Layoutswitch()
    End Sub

    <CommandMethod("FontUpdate")> Public Sub FontUpdate()

        Dim blkref As AcadBlockReference
        Dim attrib As AcadAttributeReference
        Dim attribs

        Dim j As Integer
        Dim maxpoint(0 To 2) As Double
        Dim minpoint(0 To 2) As Double
        Dim obj As Object



        Dim Client As New ClsXdata
        If Client.Check_client = True Then Exit Sub

        Try
            Dim textgen As New ClsVTD
            Dim WritetoCMD As New Clsutilities
            If textgen.Text_Gen("ALL") = True Then
                WritetoCMD.WritetoCMD("Updated textstyle fonts")
            End If

        Catch ex As Exception

        End Try
        Try

            For Each obj In Thisdrawing.ModelSpace
                If TypeOf obj Is AcadBlockReference Then
                    blkref = obj
                    If UCase(Left(blkref.EffectiveName, 3)) = "HEL" Then
                        attribs = blkref.GetAttributes
                        For j = 0 To UBound(attribs)
                            attrib = attribs(j)
                            attrib.ScaleFactor = 0.9
                        Next j
                    End If
                End If
            Next obj
        Catch

        End Try

        Try
            For Each obj In Thisdrawing.PaperSpace
                If TypeOf obj Is AcadBlockReference Then
                    blkref = obj
                    If UCase(Left(blkref.EffectiveName, 3)) = "HEL" Then
                        attribs = blkref.GetAttributes
                        For j = 0 To UBound(attribs)
                            attrib = attribs(j)
                            attrib.ScaleFactor = 0.9
                        Next j
                    End If
                End If
            Next obj
        Catch

        End Try

        Call Adjusttitlelines()

    End Sub
    <CommandMethod("Titlelines")> Public Sub Adjusttitlelines()

        Dim blkref As AcadBlockReference
        Dim attrib As AcadAttributeReference
        Dim dynprop As AcadDynamicBlockReferenceProperty
        Dim dynprops, attribs
        Dim I As Integer
        Dim j As Integer
        Dim maxpoint(0 To 2) As Double
        Dim minpoint(0 To 2) As Double
        Dim obj As Object
        Dim rot As Double
        rot = 0

        Dim Client As New ClsXdata
        If Client.Check_client = True Then Exit Sub

        Try

            For Each obj In Thisdrawing.ModelSpace
                If TypeOf obj Is AcadBlockReference Then
                    blkref = obj
                    If blkref.IsDynamicBlock Then
                        If blkref.Rotation <> 0 Then
                            rot = blkref.Rotation
                            blkref.Rotation = 0
                        End If
                        dynprops = blkref.GetDynamicBlockProperties
                        For I = 0 To UBound(dynprops)
                            If dynprops(I).PropertyName = "hel_title_line" Then
                                dynprop = dynprops(I)
                                attribs = blkref.GetAttributes
                                For j = 0 To UBound(attribs)
                                    If attribs(j).TagString = "TITLE" Then
                                        attrib = attribs(j)
                                        attrib.GetBoundingBox(minpoint, maxpoint)
                                        maxpoint(1) = blkref.InsertionPoint(1)
                                        minpoint(1) = blkref.InsertionPoint(1)
                                        dynprop.Value = maxpoint(0) - blkref.InsertionPoint(0)
                                    End If
                                Next j
                            End If
                        Next I
                        If rot <> 0 Then
                            blkref.Rotation = rot
                            rot = 0
                        End If
                    End If
                End If
            Next obj
        Catch
        End Try

        Try


            For Each obj In Thisdrawing.PaperSpace
                If TypeOf obj Is AcadBlockReference Then
                    blkref = obj
                    If blkref.IsDynamicBlock Then
                        If blkref.Rotation <> 0 Then
                            rot = blkref.Rotation
                            blkref.Rotation = 0
                        End If
                        dynprops = blkref.GetDynamicBlockProperties
                        For I = 0 To UBound(dynprops)
                            If dynprops(I).PropertyName = "hel_title_line" Then
                                dynprop = dynprops(I)
                                attribs = blkref.GetAttributes
                                For j = 0 To UBound(attribs)
                                    If attribs(j).TagString = "TITLE" Then
                                        attrib = attribs(j)
                                        attrib.GetBoundingBox(minpoint, maxpoint)
                                        maxpoint(1) = blkref.InsertionPoint(1)
                                        minpoint(1) = blkref.InsertionPoint(1)
                                        dynprop.Value = maxpoint(0) - blkref.InsertionPoint(0)
                                    End If
                                Next j
                            End If
                        Next I
                        If rot <> 0 Then
                            blkref.Rotation = rot
                            rot = 0
                        End If
                    End If
                End If
            Next obj
        Catch
        End Try


    End Sub

    Public Sub Layoutswitch()
        Dim Client As New ClsXdata
        If Client.Check_client = True Then
            clientswitch = True
        ElseIf clientswitch = True Then

        Else
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Dim tr As Transaction = db.TransactionManager.StartTransaction
            Dim scl As Integer
            Dim uni As Double
            Dim dimstychk As ClsCheckroutines
            Using tr
                scl = Application.GetSystemVariable("Userr2")
                uni = Application.GetSystemVariable("userr3")
                If scl = 0 Then scl = 1
                If uni = 0 Then uni = 1

                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
                Dim dimsty As DimStyleTable = tr.GetObject(db.DimStyleTableId, OpenMode.ForRead)

                Dim dimstid As ObjectId
                If LCase(Strings.Left(btr.Name, 6)) = "*paper" Then
                    If uni = 1 Then
                        Application.SetSystemVariable("ltscale", 10)
                    Else
                        Application.SetSystemVariable("ltscale", 0.375)
                    End If
                    Try
                        If Application.GetSystemVariable("Userr1") = "2" Then
                            Application.SetSystemVariable("Textstyle", "PHEL20")
                        ElseIf Application.GetSystemVariable("Userr1") = "2.5" Then
                            Application.SetSystemVariable("Textstyle", "PHEL25")
                        End If
                    Catch ex As Exception

                    End Try

                    Try
                        dimstychk = New ClsCheckroutines
                        If Application.GetSystemVariable("Users2") = "arrws" Then
                            dimstid = dimstychk.Dimstcheck("PHEL-Arrows")
                            db.Dimstyle = dimstid
                            db.SetDimstyleData(tr.GetObject(dimstid, OpenMode.ForRead))
                        Else
                            dimstid = dimstychk.Dimstcheck("PHEL-Ticks")
                            db.Dimstyle = dimstid
                            db.SetDimstyleData(tr.GetObject(dimstid, OpenMode.ForRead))
                        End If

                    Catch ex As Exception

                    End Try


                Else
                    If uni = 1 Then
                        Application.SetSystemVariable("ltscale", 10 * scl)
                    Else
                        Application.SetSystemVariable("ltscale", 0.375 * scl)
                    End If
                    Try
                        If Application.GetSystemVariable("Userr1") = "2" Then
                            Application.SetSystemVariable("Textstyle", "HEL20")
                        ElseIf Application.GetSystemVariable("Userr1") = "2.5" Then
                            Application.SetSystemVariable("Textstyle", "HEL25")
                        End If
                    Catch ex As Exception

                    End Try


                    Try
                        dimstychk = New ClsCheckroutines
                        If Application.GetSystemVariable("Users2") = "arrws" Then
                            dimstid = dimstychk.Dimstcheck("HEL-Arrows")
                            db.Dimstyle = dimstid
                            db.SetDimstyleData(tr.GetObject(dimstid, OpenMode.ForRead))
                        Else
                            dimstid = dimstychk.Dimstcheck("HEL-Ticks")
                            db.Dimstyle = dimstid
                            db.SetDimstyleData(tr.GetObject(dimstid, OpenMode.ForRead))
                        End If
                    Catch ex As Exception

                    End Try

                    'ed.Regen()
                End If

                tr.Commit()

            End Using
        End If
    End Sub

    <CommandMethod("loadpalette")> Sub LoadPalettes()

        Dim woodpalette As uscWOOD = New uscWOOD
        'Dim multiscale As USCMultiScale = New USCMultiScale
        Dim drawingissue As uscIssues = New uscIssues
        Dim masonrypalette As uscMasonry = New uscMasonry
        If m_ps = Nothing Then

            m_ps = New Autodesk.AutoCAD.Windows.PaletteSet("Custom HEL Routines", New Guid("{ECBFEC73-9FE4-4aa2-8E4B-3068E94A2BFB}"))
            m_ps.Add("Wood Routines", woodpalette)
            m_ps.Add("Masonry Units", masonrypalette)
            m_ps.Add("Drawing Issues", drawingissue)
            'm_ps.Add("MultiScale", multiscale)
            m_ps.Style = PaletteSetStyles.ShowAutoHideButton Or PaletteSetStyles.ShowCloseButton
            m_ps.Visible = True

        Else
            m_ps.Visible = True

        End If

    End Sub


    <CommandMethod("Timber")> Sub LoadTimberPalette()
        If m_ps = Nothing Then
            LoadPalettes()

        Else
            m_ps.Visible = True

        End If
        For i As Integer = 0 To m_ps.Count - 1
            If m_ps.Item(i).Name = "Wood Routines" Then m_ps.Activate(i)
        Next
    End Sub


    <CommandMethod("Issues")> Sub LoadIssuesPalette()
        If m_ps = Nothing Then
            LoadPalettes()
        Else
            m_ps.Visible = True

        End If
        For i As Integer = 0 To m_ps.Count - 1
            If m_ps.Item(i).Name = "Drawing Issues" Then m_ps.Activate(i)
        Next
    End Sub

    <CommandMethod("Masonry")> Sub LoadMasonryPalette()
        If m_ps = Nothing Then
            LoadPalettes()
        Else
            m_ps.Visible = True

        End If
        For i As Integer = 0 To m_ps.Count - 1
            If m_ps.Item(i).Name = "Masonry Units" Then m_ps.Activate(i)
        Next
    End Sub


    Private Sub Thisdrawing_ObjectModified([Object] As Object) Handles Thisdrawing.ObjectModified
        Dim Client As New ClsXdata
        If Client.Check_client = True Then

        ElseIf attswitch = True Then
            On Error GoTo nd
            If TypeOf [Object] Is AcadAttributeReference Then
                Dim ent As AcadEntity = [Object]
                Dim attribref As AcadAttributeReference = ent
                Dim oacad As Autodesk.AutoCAD.Interop.AcadApplication = Thisdrawing.Application
                Dim obj As Object = oacad.ActiveDocument.ObjectIdToObject(attribref.OwnerID)
                If TypeOf obj Is AcadBlockReference Then
                    Dim blkref As AcadBlockReference = obj


                    Dim attrib As AcadAttributeReference
                    Dim dynprop As AcadDynamicBlockReferenceProperty
                    Dim dynprops, attribs
                    Dim I As Integer
                    Dim j As Integer
                    Dim maxpoint(0 To 2) As Double
                    Dim minpoint(0 To 2) As Double
                    Dim rot = Nothing
                    If LCase(blkref.EffectiveName) = "hel_sym_title" Or LCase(blkref.EffectiveName) = "hel_sym_2pct" Or LCase(blkref.EffectiveName) = "hel_sym_2pot" Then
                        If blkref.IsDynamicBlock Then
                            If blkref.Rotation <> 0 Then
                                rot = blkref.Rotation
                                blkref.Rotation = 0
                            End If
                            dynprops = blkref.GetDynamicBlockProperties
                            For I = 0 To UBound(dynprops)
                                If dynprops(I).PropertyName = "hel_title_line" Then
                                    dynprop = dynprops(I)
                                    attribs = blkref.GetAttributes
                                    For j = 0 To UBound(attribs)
                                        If attribs(j).TagString = "TITLE" Then
                                            attrib = attribs(j)
                                            attrib.GetBoundingBox(minpoint, maxpoint)
                                            maxpoint(1) = blkref.InsertionPoint(1)
                                            minpoint(1) = blkref.InsertionPoint(1)
                                            dynprop.Value = maxpoint(0) - blkref.InsertionPoint(0)
                                        End If
                                    Next j
                                End If
                            Next I
                            If rot <> 0 Then
                                blkref.Rotation = rot
                                rot = 0
                            End If
                        End If
                    End If
                End If



            End If
            attswitch = False
        End If


nd:
    End Sub

    <CommandMethod("LoadHELMenus")> Sub LoadHELMenus()
        Dim Client As New ClsXdata
        If Client.Check_client = True Then Exit Sub

        Dim mg As AcadMenuGroup
        Try
            mg = Application.MenuGroups.item("STEELBUNDLE")
        Catch
            Application.MenuGroups.load("SteelBundlex64.cuix", False)
        End Try
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        ed.WriteMessage("SteelPlus Menu Loaded" & vbLf)
        Try
            mg = Application.MenuGroups.item("HEL-ACAD-2023")

        Catch
            Application.MenuGroups.load("HEL-ACAD-2023.cuix", False)
        End Try
        ed.WriteMessage("HEL-ACAD-2023 Menu Loaded")
    End Sub

    Public Function License_Check() As Boolean
        License_Check = True
        'Dim Dte As Date = Now
        'If (Dte.Year) > "2025" Then
        'MsgBox("Herold AutoCAD Customizations have expired" & vbLf & "AutoCAD will now close" & vbLf & "Please remove the exipred customization files to open AutoCAD" & vbLf & "Update as soon as possible", vbOKOnly + vbCritical, "Warning!")
        'Application.Quit()
        'ElseIf Dte.Year = "2024" Then
        'If Dte.Month = 12 And Dte.Day > 15 Then
        'MsgBox("Herold AutoCAD Customizations will exipre December 31, 2024" & vbLf & "Please update as soon as possible", vbOKOnly + vbCritical, "Warning!")
        'ElseIf (Dte.Month) > 12 Or (Dte.Month) = 13 Then
        'MsgBox("Herold AutoCAD Customizations have expired" & vbLf & "AutoCAD will now close" & vbLf & "Please remove the exipred customization files to open AutoCAD" & vbLf & "Update as soon as possible", vbOKOnly + vbCritical, "Warning!")
        'Application.Quit()
        'End If
        'End If
    End Function

    Private Sub Thisdrawing_BeginLisp(FirstLine As String) Handles Thisdrawing.BeginLisp
        If FirstLine = "(C:WELD)" Then
            Try
                Dim scl As Integer = Thisdrawing.GetVariable("UserR2") '_DWSC
                Dim uni As Double = Thisdrawing.GetVariable("UserR3") 'SF
                If Thisdrawing.GetVariable("userr3") = 0 Then Thisdrawing.SetVariable("Userr3", 1)
                If uni = 1 Then
                    Thisdrawing.SetVariable("dimscale", scl * 25.4)
                Else

                End If
            Catch

            End Try

        End If
    End Sub

    Private Sub Thisdrawing_EndLisp() Handles Thisdrawing.EndLisp
        Try
            Dim scl As Integer = Thisdrawing.GetVariable("UserR2") '_DWSC
            Thisdrawing.SetVariable("dimscale", scl)
        Catch
        End Try

    End Sub
End Class
