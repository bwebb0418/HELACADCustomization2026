Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Windows
Imports Microsoft.VisualBasic.Interaction
Imports System.Data.Sql
Imports System.IO.File

Public Class FrmTitleblock
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property
    Public check_routines As New clsCheckroutines

    'Dim address As String
    'Dim Projname As String
    'Dim Clientname As String
    Dim space
    Dim blkscl As Integer
    ReadOnly inspoint(0 To 2) As Double
    Dim path As String
    'Dim prname1 As String
    'Dim prnamefinal1 As String
    'Dim prnamefinal2 As String
    'Dim prnamefinal3 As String



    Private Sub FrmTitleblock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim path As String


        'tboDate.Text = Format(Now, "yyyy.MM.dd")




        'If Thisdrawing.GetVariable("Dwgtitled") <> 0 Then
        '    path = Thisdrawing.Path
        '    path = Strings.Right(path, Strings.Len(path) - 12)
        '    path = Strings.Left(path, 8)
        '    tboPN.Text = path
        '    path = Strings.Left(path, 4) & Strings.Right(path, 3)
        '    If path.Contains("-") Then
        '        path = tboPN.Text
        '        path = Strings.Left(path, 3) & Strings.Right(path, 4)
        '    End If

        '    'If readtbfile(tboPN.Text) = False Then
        '    '    If sql_read(path) <> False Then
        '    '        tboCLI.Text = UCase(Clientname)
        '    '        tboPL.Text = UCase(address)
        '    '        If Strings.Len(Projname) <= 22 Then
        '    '            tboPT1.Text = Projname
        '    '        Else
        '    '            Dim pname As String() = Projname.Split(New Char() {" "c})
        '    '            For Each word In pname
        '    '                prname1 = prname1 & word & " "
        '    '                If Strings.Len(prname1) < 22 Then
        '    '                    prnamefinal1 = prname1
        '    '                ElseIf Strings.Len(prname1) < 44 Then
        '    '                    prnamefinal2 = Strings.Right(prname1, Strings.Len(prname1) - Strings.Len(prnamefinal1))
        '    '                Else
        '    '                    prnamefinal3 = Strings.Right(prname1, Strings.Len(prname1) - Strings.Len(prnamefinal1) - Strings.Len(prnamefinal2))
        '    '                End If
        '    '            Next
        '    '            tboPT1.Text = Trim(UCase(prnamefinal1))
        '    '            tboPT2.Text = Trim(UCase(prnamefinal2))
        '    '            tboPT3.Text = Trim(UCase(prnamefinal3))
        '    '        End If
        '    '    End If
        '    'End If
        '    prname1 = Nothing
        '    prnamefinal1 = Nothing
        '    prnamefinal2 = Nothing
        '    prnamefinal3 = Nothing
        'End If


        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then
            space = Thisdrawing.ModelSpace
            blkscl = Thisdrawing.GetVariable("userr2")
        Else
            space = Thisdrawing.PaperSpace
            blkscl = 1
        End If
        Try
            Loadtblocks("HEL_2023_85x11")
            Cbosize.Items.Add("HEL_2023_85x11")
        Catch
            Thisdrawing.Utility.Prompt("Error Loading 8.5x11 Titleblock")
        End Try
        Try
            Loadtblocks("HEL_2023_11x17")
            Cbosize.Items.Add("HEL_2023_11x17")
        Catch
            Thisdrawing.Utility.Prompt("Error Loading 11x17 Titleblock")
        End Try
        Try
            Loadtblocks("HEL_2023_22x17H")
            Cbosize.Items.Add("HEL_2023_22x17H")
        Catch
            Thisdrawing.Utility.Prompt("Error Loading 18x24 Titleblock")
        End Try
        Try
            Loadtblocks("HEL_2023_24x36")
            Cbosize.Items.Add("HEL_2023_24x36")
        Catch
            Thisdrawing.Utility.Prompt("Error Loading 24x36 Titleblock")
        End Try
        Try
            Loadtblocks("HEL_2023_Ind_11x17")
            Cbosize.Items.Add("HEL_2023_Ind_11x17")
        Catch
            Thisdrawing.Utility.Prompt("Error Loading Industrial 11x17 Titleblock")
        End Try
        Try
            Loadtblocks("HEL_2023_Ind_24x36")
            Cbosize.Items.Add("HEL_2023_Ind_24x36")
        Catch
            Thisdrawing.Utility.Prompt("Error Loading Industrial 24x36 Titleblock")
        End Try


    End Sub

    Private Sub Loadtblocks(ByRef blkname As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = db.TransactionManager.StartTransaction
        Dim tblk As AcadBlockReference
        Dim blockinsert As New ClsBlock_insert
        Using tr
            inspoint(0) = 0 : inspoint(1) = 0 : inspoint(2) = 0
            Try
                If check_routines.Blkchk(blkname) = True Then

                Else
                    If blockinsert.Loadblock(blkname & ".dwg") = False Then
                        Exit Sub
                    End If
                End If

                tblk = space.InsertBlock(inspoint, blkname, 1, 1, 1, 0)
                tblk.Delete()
                tr.Commit()
            Catch
                tr.Abort()
            End Try
            tr.Dispose()
        End Using
    End Sub

    Private Sub CmdGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGo.Click
        Try


            Dim units
            Dim dynprops
            Dim i As Integer
            Dim tblk As AcadBlockReference
            Dim blkname As Object
            Dim size
            Dim prop As AcadDynamicBlockReferenceProperty = Nothing

            If Cbosize.SelectedItem = Nothing Then
                Cbosize.Text = "HEL_2023_24x36"
            End If

            If path = Nothing And Not Cbosize.SelectedItem.ToString.Contains("HEL") Then

                'If LCase(Right(path, 3)) <> "dwg" Then GoTo strt
                If check_routines.Blkchk("HEL_Border") = True Then
                    blkname = "HEL_Border"
                Else
                    blkname = "HEL_Border.dwg"
                End If

                tblk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)
                dynprops = tblk.GetDynamicBlockProperties

                For i = 0 To UBound(dynprops)
                    If dynprops(i).PropertyName = "Border" Then
                        prop = dynprops(i)
                        If Cbosize.SelectedItem = Nothing Then
                            size = "Buildings D - Vertical"
                            prop.Value = Strings.Trim(size)
                        Else
                            prop.Value = Cbosize.SelectedItem.ToString
                        End If
                    End If
                Next i

                units = Thisdrawing.GetVariable("Userr3")

                If units = 1 And prop.Value <> "ISO A1" Then
                    tblk.XScaleFactor = 25.4
                    tblk.YScaleFactor = 25.4
                    tblk.ZScaleFactor = 25.4
                ElseIf units = 0 And prop.Value = "ISO A1" Then
                    tblk.XScaleFactor = 10 / 254
                    tblk.YScaleFactor = 10 / 254
                    tblk.ZScaleFactor = 10 / 254
                End If

            ElseIf path = Nothing And Cbosize.SelectedItem.ToString.Contains("HEL") Then

                blkname = Cbosize.SelectedItem.ToString

                If check_routines.Blkchk(blkname) = True Then
                    blkname = Cbosize.SelectedItem.ToString
                Else
                    blkname = Cbosize.SelectedItem.ToString & ".dwg"
                End If

                tblk = space.InsertBlock(inspoint, blkname, blkscl, blkscl, blkscl, 0)

                units = Thisdrawing.GetVariable("Userr3")

                If units = 1 Then
                    tblk.XScaleFactor = 25.4
                    tblk.YScaleFactor = 25.4
                    tblk.ZScaleFactor = 25.4

                End If

            Else
                tblk = space.InsertBlock(inspoint, Cbosize.SelectedItem.ToString, blkscl, blkscl, blkscl, 0)
            End If

            'attedit:
            '            tblk.Update()
            '            'Thisdrawing.SendCommand("attedit l ")
            '            Dim attribs
            '            Dim attrib As AcadAttributeReference

            '            attribs = tblk.GetAttributes()
            '            For i = 0 To UBound(attribs)
            '                attrib = attribs(i)
            '                If LCase(attrib.TagString) = "date" Then
            '                    attrib.TextString = tboDate.Text
            '                ElseIf LCase(attrib.TagString) = "proj-num" Then
            '                    attrib.TextString = tboPN.Text
            '                ElseIf LCase(attrib.TagString) = "draft" Then
            '                    attrib.TextString = tboDWGB.Text
            '                ElseIf LCase(attrib.TagString) = "draftrev" Then
            '                    attrib.TextString = tboCHKB.Text
            '                ElseIf LCase(attrib.TagString) = "designed" Then
            '                    attrib.TextString = tboDESB.Text
            '                ElseIf LCase(attrib.TagString) = "designrev" Then
            '                    attrib.TextString = tboREVB.Text
            '                ElseIf LCase(attrib.TagString) = "project-title" Then 'And attrib.MTextAttribute = True Then
            '                    attrib.TextString = tboPT1.Text & "\P" & tboPT2.Text & "\P" & tboPT3.Text
            '                ElseIf LCase(attrib.TagString) = "project-title-1" Then
            '                    attrib.TextString = tboPT1.Text
            '                ElseIf LCase(attrib.TagString) = "project-title-2" Then
            '                    attrib.TextString = tboPT2.Text
            '                ElseIf LCase(attrib.TagString) = "project-title-3" Then
            '                    attrib.TextString = tboPT3.Text
            '                ElseIf LCase(attrib.TagString) = "title" Then 'And attrib.MTextAttribute = True Then
            '                    attrib.TextString = tboDT1.Text & "\P" & tboDT2.Text & "\P" & tboDT3.Text & "\P" & tboDT4.Text & "\P" & tboDT5.Text
            '                ElseIf LCase(attrib.TagString) = "title-1" Then
            '                    attrib.TextString = tboDT1.Text
            '                ElseIf LCase(attrib.TagString) = "title-2" Then
            '                    attrib.TextString = tboDT2.Text
            '                ElseIf LCase(attrib.TagString) = "title-3" Then
            '                    attrib.TextString = tboDT3.Text
            '                ElseIf LCase(attrib.TagString) = "title-4" Then
            '                    attrib.TextString = tboDT4.Text
            '                ElseIf LCase(attrib.TagString) = "title-5" Then
            '                    attrib.TextString = tboDT5.Text
            '                ElseIf LCase(attrib.TagString) = "client" Then
            '                    attrib.TextString = tboCLI.Text
            '                ElseIf LCase(attrib.TagString) = "location" Then
            '                    attrib.TextString = tboPL.Text
            '                ElseIf LCase(attrib.TagString) = "dwgnum" Then
            '                    attrib.TextString = tbosht.Text
            '                End If
            '            Next i
            Me.Hide()
            path = Nothing
            '            'If readtbfile(tboPN.Text) = False Then writetbfile()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Catch
            Me.DialogResult = System.Windows.Forms.DialogResult.Abort
        End Try

    End Sub

    Private Sub CmdCAN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCAN.Click
        Me.Hide()
    End Sub

    Private Sub Cbosize_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cbosize.SelectedIndexChanged
        Dim oShell As Shell32.Shell
        Dim oFolder As Shell32.Folder
        'Dim oItems 'As Shell32.FolderItems
        Dim item 'As Shell32.FolderItem


        If Cbosize.SelectedItem.ToString = "Browse..." Then
strt:

            oShell = New Shell32.Shell ' ActiveX interface to shell32.dll
            oFolder = oShell.BrowseForFolder(0, "", 512, "M:\Structural\r2023\Blocks\Client Titleblocks\")
            If oFolder Is Nothing Then
                Exit Sub
            End If

            Dim fso
            Dim Folder
            Dim Files

            fso = CreateObject("scripting.filesystemobject")
            path = "M:\Structural\r2023\Blocks\Client Titleblocks\" + oFolder.Title
            Folder = fso.GetFolder(path)
            Files = Folder.Files

            Dim I As Integer
            For I = 0 To Cbosize.Items.Count - 1
                If Cbosize.Items(I) = "Browse..." Then
                    Cbosize.Items.Remove(Cbosize.Items(I))

                End If
            Next

            For Each item In Files
                If LCase(Strings.Right(item.Name, 3)) = "dwg" Then
                    Cbosize.Items.Add(item.Name)
                    'frmtblk2.Label1 = Folder & "\"
                End If
            Next item
        End If
        Cbosize.Items.Add("Browse...")
    End Sub

    Private Sub FrmTitleblock_Activated(sender As Object, e As EventArgs) Handles Me.Activated

    End Sub

    'Private Sub TboScale_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles tboScale.Enter
    '    Dim RETVAL As MsgBoxResult

    '    If LCase(tboScale.Text) <> "as shown" Then Exit Sub
    '    RETVAL = MsgBox("Do you want to change from 'AS SHOWN'?", vbYesNo + MsgBoxStyle.Exclamation)
    '    If RETVAL = MsgBoxResult.No Then Exit Sub
    '    'Dim getscale As clsTitles_Tags = New clsTitles_Tags
    '    tboScale.Text = clsTitles_Tags.GETSCL

    'End Sub

    'Private Sub TboScale_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tboScale.GotFocus

    'End Sub


    'Private Sub TboScale_lostfocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tboScale.LostFocus
    '    tboScale.Text = UCase(tboScale.Text)
    'End Sub

    '    Public Function sql_read(ByVal path As String)
    '        Dim SQLStr As String
    '        Dim Connstr As String
    '        Dim SQLConn As New SqlClient.SqlConnection() 'The SQL Connection
    '        Dim SQLCmd As New SqlClient.SqlCommand() 'The SQL Command
    '        Dim SQLdr As SqlClient.SqlDataReader        'The Local Data Store
    '        Dim client As String

    '        On Error GoTo nd
    '        'Connstr = "PROVIDER=SQLOLEDB.1;"

    '        'Connect to the Pubs database on the local server.
    '        Connstr = "DATA SOURCE=Herold-ACC;INITIAL CATALOG=vision;"

    '        'Use an integrated login.
    '        'strConn = strConn & " INTEGRATED SECURITY=sspi;"
    '        Connstr = Connstr & "UID=sa;pwd=Pa$$w0rd;"

    '        SQLStr = "SELECT name, clientid FROM pr where left(wbs1,7) = '" & path & "'"


    '        SQLConn.ConnectionString = Connstr 'Set the Connection String
    '        SQLConn.Open() 'Open the connection


    '        SQLCmd.Connection = SQLConn 'Sets the Connection to use with the SQL Command
    '        SQLCmd.CommandText = SQLStr 'Sets the SQL String
    '        SQLdr = SQLCmd.ExecuteReader 'Gets Data

    '        Do While SQLdr.Read() 'While Data is Present        

    '            'MsgBox(SQLdr("Address1")) 'Show data in a Message Box
    '            'MsgBox(SQLdr("City")) 'Show data in a Message Box
    '            'MsgBox(SQLdr("State")) 'Show data in a Message Box
    '            'MsgBox(SQLdr("ZIP")) 'Show data in a Message Box
    '            client = SQLdr("clientid")
    '            Projname = SQLdr("Name")
    '            SQLdr.NextResult() 'Move to the Next Record
    '        Loop
    '        SQLdr.Close() 'Close the SQLDataReader        

    '        SQLConn.Close() 'Close the connection

    '        SQLStr = "SELECT name FROM CL where clientid = '" & client & "'"

    '        SQLConn.Open() 'Open the connection


    '        SQLCmd.Connection = SQLConn 'Sets the Connection to use with the SQL Command
    '        SQLCmd.CommandText = SQLStr 'Sets the SQL String
    '        SQLdr = SQLCmd.ExecuteReader 'Gets Data

    '        Do While SQLdr.Read() 'While Data is Present        

    '            Clientname = SQLdr("name")

    '            SQLdr.NextResult() 'Move to the Next Record
    '        Loop
    '        SQLdr.Close() 'Close the SQLDataReader        

    '        SQLConn.Close() 'Close the connection

    '        SQLStr = "SELECT Address1, City, state,ZIP, phone, fax FROM claddress where clientid ='" & client & "'"

    '        SQLConn.Open() 'Open the connection

    '        SQLCmd.Connection = SQLConn 'Sets the Connection to use with the SQL Command
    '        SQLCmd.CommandText = SQLStr 'Sets the SQL String
    '        SQLdr = SQLCmd.ExecuteReader 'Gets Data

    '        Do While SQLdr.Read() 'While Data is Present        

    '            address = SQLdr("Address1") & " " & SQLdr("City") & " " & SQLdr("State") & " " & SQLdr("Zip")
    '            'address = (SQLdr("City")) 'Show data in a Message Box
    '            'MsgBox(SQLdr("State")) 'Show data in a Message Box
    '            'MsgBox(SQLdr("ZIP")) 'Show data in a Message Box

    '            SQLdr.NextResult() 'Move to the Next Record
    '        Loop
    '        SQLdr.Close() 'Close the SQLDataReader        

    '        SQLConn.Close() 'Close the connection
    '        sql_read = True
    '        Exit Function
    'nd:
    '        sql_read = False
    '    End Function



    'Sub writetbfile()
    '    Dim projectline As String
    '    Dim INFOPATH As String
    '    For I = 0 To Len(Thisdrawing.Path)
    '        If Strings.Left(Strings.Right(Thisdrawing.Path, I), 12) = "04S Drawings" Then
    '            INFOPATH = Strings.Left(Thisdrawing.Path, Strings.Len(Thisdrawing.Path) - I + 12)
    '        End If
    '    Next I

    '    If IO.File.Exists(INFOPATH & "\titleblockinfo.csv") = False Then
    '        Using file As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(INFOPATH & "\titleblockinfo.csv", True)
    '            projectline = "Projectnumber,ProjectTitle#1,ProjectTitle#2,ProjectTitle#3,Client,DrawnBy,CheckedBy,DesignedBy,ReviewedBy,ProjectLocation" & vbCr

    '            file.WriteLine(projectline)
    '            file.Close()
    '        End Using
    '    End If
    '    Using file As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(INFOPATH & "\titleblockinfo.csv", True)
    '        Dim pn, pt1, pt2, pt3, cli, dwgb, chkb, desb, revb, pl

    '        pn = tboPN.Text
    '        pt1 = tboPT1.Text
    '        pt2 = tboPT2.Text
    '        pt3 = tboPT3.Text
    '        cli = tboCLI.Text
    '        dwgb = tboDWGB.Text
    '        chkb = tboCHKB.Text
    '        desb = tboDESB.Text
    '        revb = tboREVB.Text
    '        pl = tboPL.Text
    '        pn = Strings.Replace(pn, ",", "`")
    '        pt1 = Strings.Replace(pt1, ",", "`")
    '        pt2 = Strings.Replace(pt2, ",", "`")
    '        pt3 = Strings.Replace(pt3, ",", "`")
    '        cli = Strings.Replace(cli, ",", "`")
    '        dwgb = Strings.Replace(dwgb, ",", "`")
    '        chkb = Strings.Replace(chkb, ",", "`")
    '        desb = Strings.Replace(desb, ",", "`")
    '        revb = Strings.Replace(revb, ",", "`")
    '        pl = Strings.Replace(pl, ",", "`")


    '        projectline = pn & "," & pt1 & "," & pt2 & "," & pt3 & "," & cli & "," & dwgb & "," & chkb & "," & desb & "," & revb & "," & pl & vbCr


    '        file.WriteLine(projectline)
    '        file.Close()
    '    End Using

    'End Sub

    'Function readtbfile(ByVal projnum As String) As Boolean
    '    Dim fh As Integer, tmp
    '    Dim INFOPATH As String
    '    For I = 0 To Len(Thisdrawing.Path)
    '        If Strings.Left(Strings.Right(Thisdrawing.Path, I), 12) = "04S Drawings" Then
    '            INFOPATH = Strings.Left(Thisdrawing.Path, Strings.Len(Thisdrawing.Path) - I + 12)
    '        End If
    '    Next I

    '    If IO.File.Exists(INFOPATH & "\titleblockinfo.csv") = True Then


    '        'Try 'On Error Resume Next
    '        fh = Microsoft.VisualBasic.FreeFile
    '        Try
    '            Microsoft.VisualBasic.FileOpen(fh, INFOPATH & "\titleblockinfo.csv", Microsoft.VisualBasic.OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

    '            Do While Not EOF(fh)
    '                tmp = Microsoft.VisualBasic.FileSystem.LineInput(fh)
    '                Dim tblockentries As String() = tmp.Split(New Char() {","c})
    '                If tblockentries(0) = projnum Then
    '                    tboPN.Text = Strings.Replace(tblockentries(0), "`", ",")
    '                    tboPT1.Text = Strings.Replace(tblockentries(1), "`", ",")
    '                    tboPT2.Text = Strings.Replace(tblockentries(2), "`", ",")
    '                    tboPT3.Text = Strings.Replace(tblockentries(3), "`", ",")
    '                    tboCLI.Text = Strings.Replace(tblockentries(4), "`", ",")
    '                    tboDWGB.Text = Strings.Replace(tblockentries(5), "`", ",")
    '                    tboCHKB.Text = Strings.Replace(tblockentries(6), "`", ",")
    '                    tboDESB.Text = Strings.Replace(tblockentries(7), "`", ",")
    '                    tboREVB.Text = Strings.Replace(tblockentries(8), "`", ",")
    '                    tboPL.Text = Strings.Replace(tblockentries(9), "`", ",")
    '                    Return True
    '                End If

    '            Loop
    '            FileClose(fh)

    '        Catch ex As Exception
    '            Return False

    '        End Try
    '    Else
    '        Return False
    '    End If
    '    Return True
    'End Function
End Class