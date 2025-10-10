Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
'Imports System.Data.OleDb

Public Class ClsVTD


    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property
    Public Function SetSystemVariables() As Boolean
        SetSystemVariables = False
        Try

            Dim scl As Long
            Dim pi As Double
            ''''Dim lisp As clsVLAX
            pi = 4 * Math.Atan(1)

            'this is to ensure that there is always a variable to pass on for items that require
            'a drawing scale for setting variables
            If Application.GetSystemVariable("userr2") = 0 Then Application.SetSystemVariable("userr2", 1)
            scl = Application.GetSystemVariable("UserR2")

            'This is to make sure that there is always a variable to pass on to Insbase
            Dim insbase(0 To 2) As Double
            insbase(0) = "0.0" : insbase(1) = "0.0" : insbase(2) = "0.0"
            ' ''If datum(0) = "" Then
            ' ''    datum = insbase
            ' ''End If
            'This is to set the snapunit variable
            Dim snapunit(0 To 1) As Double
            snapunit(0) = 0.5 : snapunit(1) = 0.5

            'This is to set the snapbase Variable
            Dim snapbase(0 To 1) As Double
            snapbase(0) = insbase(0) : insbase(1) = insbase(1)
            Dim ltscale As Double

            If Application.GetSystemVariable("UserR3") = "1" Then
                ltscale = 10 * scl
            Else
                ltscale = 0.375 * scl
            End If
            'Variables Specific to metric and imperial are broken out and set here
            '.Utility.Prompt ("System Variables")
            If Application.GetSystemVariable("UserR3") = "1" Then
                Application.SetSystemVariable("Lunits", 2)
                Application.SetSystemVariable("measurement", 1)
                'Application.SetSystemVariable("chamfera", 0)
                'Application.SetSystemVariable("chamferb", 0)
                'APPLICATION.SETSYSTEMVariable "chamferc", "20"
                'APPLICATION.SETSYSTEMVariable "chamferd", "20"
                Application.SetSystemVariable("cmlscale", 20)
                Application.SetSystemVariable("donutod", 25)
                Application.SetSystemVariable("hpscale", (scl / 3))
                Application.SetSystemVariable("insunits", 4)
                Application.SetSystemVariable("insunitsdefsource", 4)
                Application.SetSystemVariable("insunitsdeftarget", 4)
                Application.SetSystemVariable("ltscale", ltscale) 'LTScale variable uses Units from XData.XDtoLSP Module
                Application.SetSystemVariable("luprec", 3)
                Application.SetSystemVariable("lwunits", 1)
            Else
                Application.SetSystemVariable("lwunits", 0)
                Application.SetSystemVariable("luprec", 4)
                Application.SetSystemVariable("ltscale", ltscale) 'LTScale variable uses Units from XData.XDtoLSP Module
                Application.SetSystemVariable("insunits", 1)
                Application.SetSystemVariable("insunitsdeftarget", 1)
                Application.SetSystemVariable("insunitsdefsource", 1)
                Application.SetSystemVariable("hpscale", (scl / 3))
                Application.SetSystemVariable("donutod", 1.0#)
                Application.SetSystemVariable("cmlscale", 1.0#)
                'Application.SetSystemVariable("chamfera", 0)
                'Application.SetSystemVariable("chamferb", 0)
                'APPLICATION.SETSYSTEMVariable "chamferc", "0.75"
                'APPLICATION.SETSYSTEMVariable "chamferd", "0.75"
                Application.SetSystemVariable("Lunits", 4)
                Application.SetSystemVariable("measurement", 0)
            End If

            'APPLICATION.SETSYSTEMVariable "*_toolpalettepath", "m:\custom\industrial\toolpallets"
            'APPLICATION.SETSYSTEMVariable "_linfo"
            'APPLICATION.SETSYSTEMVariable "_pkser"
            'APPLICATION.SETSYSTEMVariable "_server"
            'APPLICATION.SETSYSTEMVariable "_vernum"
            Application.SetSystemVariable("acadlspasdoc", 1) 'Loads acad.lsp with every drawing opened
            'Application.SetSystemVariable("acadprefix", "m:\structural\r2026\...")
            'APPLICATION.SETSYSTEMVariable "acadver"
            'APPLICATION.SETSYSTEMVariable "acis15"
            'APPLICATION.SETSYSTEMVariable "acisoutver"
            'APPLICATION.SETSYSTEMVariable "adcstate"
            'APPLICATION.SETSYSTEMVariable "aflags"
            Application.SetSystemVariable("angbase", 0)
            Application.SetSystemVariable("angdir", 0)
            Application.SetSystemVariable("apbox", 1)
            Application.SetSystemVariable("aperture", 4)
            'APPLICATION.SETSYSTEMVariable "area"
            'APPLICATION.SETSYSTEMVariable "assiststate"
            Application.SetSystemVariable("attdia", 1)
            Application.SetSystemVariable("attmode", 1)
            Application.SetSystemVariable("attreq", 1)
            'APPLICATION.SETSYSTEMVariable "auditctl"
            Application.SetSystemVariable("aunits", 0)
            Application.SetSystemVariable("auprec", 2)
            'APPLICATION.SETSYSTEMVariable "autosnap", 63
            'APPLICATION.SETSYSTEMVariable "auxstat"
            'APPLICATION.SETSYSTEMVariable "axismode"
            'APPLICATION.SETSYSTEMVariable "axisunit"
            Application.SetSystemVariable("backgroundplot", 2)
            'APPLICATION.SETSYSTEMVariable "backz"
            Application.SetSystemVariable("bactioncolor", "7")
            Application.SetSystemVariable("bdependencyhighlight", 1)
            Application.SetSystemVariable("bgripobjcolor", "144")
            If Application.GetSystemVariable("bgripobjsize") <> 8 Then Application.SetSystemVariable("bgripobjsize", 8) 'this Variable causes a regen
            Application.SetSystemVariable("bindtype", 0)
            Application.SetSystemVariable("blipmode", 0)
            Application.SetSystemVariable("blockeditlock", 0)
            'APPLICATION.SETSYSTEMVariable "blockeditor", 0 Read Only Variable
            Application.SetSystemVariable("bparametercolor", "7")
            Application.SetSystemVariable("bparameterfont", "simplex.shx")
            If Application.GetSystemVariable("bparametersize") <> 12 Then Application.SetSystemVariable("bparametersize", 12) ' This variable causes a regen
            Application.SetSystemVariable("btmarkdisplay", 1)
            Application.SetSystemVariable("bvmode", 1)
            Application.SetSystemVariable("calcinput", 1)
            'APPLICATION.SETSYSTEMVariable "cdate"
            Application.SetSystemVariable("cecolor", "bylayer")
            Application.SetSystemVariable("celtscale", 1)
            Application.SetSystemVariable("celtype", "bylayer")
            Application.SetSystemVariable("celweight", -1)
            Application.SetSystemVariable("centermt", 0)
            'APPLICATION.SETSYSTEMVariable "chamfera" 'see above code
            'APPLICATION.SETSYSTEMVariable "chamferb" 'see above code
            'APPLICATION.SETSYSTEMVariable "chamferc" 'see above code
            'APPLICATION.SETSYSTEMVariable "chamferd" 'see above code
            Application.SetSystemVariable("chammode", 0)
            Application.SetSystemVariable("circlerad", 0)
            Application.SetSystemVariable("clayer", "0")
            'APPLICATION.SETSYSTEMVariable "cleanscreenstate", 0
            'APPLICATION.SETSYSTEMVariable "clistate" 'requries clarification
            'APPLICATION.SETSYSTEMVariable "cmdactive" 'this variable has no bearing on normal operation but is used for Coding
            Application.SetSystemVariable("cmddia", 1)
            Application.SetSystemVariable("cmdecho", 0)
            Application.SetSystemVariable("cmdinputhistorymax", 20)
            'APPLICATION.SETSYSTEMVariable "cmdnames" 'this variable has no bearing on normal operation but is used for Coding
            Application.SetSystemVariable("cmljust", 0)
            'APPLICATION.SETSYSTEMVariable "cmlscale" 'see above code
            'APPLICATION.SETSYSTEMVariable "cmlstyle" 'HEL should create a library of MLINE styles for R2006.
            Application.SetSystemVariable("compass", 0)
            Application.SetSystemVariable("coords", 2)
            Application.SetSystemVariable("cplotstyle", "bylayer")
            'APPLICATION.SETSYSTEMVariable "cprofile"
            'APPLICATION.SETSYSTEMVariable "cputicks" 'requires clarification
            Application.SetSystemVariable("crossingareacolor", 3)
            'APPLICATION.SETSYSTEMVariable "ctab"
            Application.SetSystemVariable("ctablestyle", "standard")
            'APPLICATION.SETSYSTEMVariable "currentprofile" 'requires clarification
            'APPLICATION.SETSYSTEMVariable "cursorsize", 5
            'APPLICATION.SETSYSTEMVariable "cvport"
            'APPLICATION.SETSYSTEMVariable "date"
            'APPLICATION.SETSYSTEMVariable "dbcstate", 0 Read-Only Variable
            'APPLICATION.SETSYSTEMVariable "dbglistall"
            'APPLICATION.SETSYSTEMVariable "dbmod"
            'APPLICATION.SETSYSTEMVariable "dctcust" 'HEROLD ENGINEERING Requires a customize dictionary - save it on the server!!
            Application.SetSystemVariable("dctmain", "enc") 'British English (ize)
            'APPLICATION.SETSYSTEMVariable "deflplstyle" 'Read-only Variable
            'APPLICATION.SETSYSTEMVariable "defplstyle" 'Read-only Variable
            Application.SetSystemVariable("delobj", 1)
            Application.SetSystemVariable("demandload", 3)
            'APPLICATION.SETSYSTEMVariable "diastat", 1 Read-only Variable (exit state of last dialogbox)
            'APPLICATION.SETSYSTEMVariable "dispsilh"
            'APPLICATION.SETSYSTEMVariable "distance"
            Application.SetSystemVariable("donutid", 0.0#)
            'APPLICATION.SETSYSTEMVariable "donutod" 'see code above
            Application.SetSystemVariable("dragmode", 2)
            Application.SetSystemVariable("dragp1", 10)
            Application.SetSystemVariable("dragp2", 25)
            'Application.SetSystemVariable("draworderctl", 3)
            'APPLICATION.SETSYSTEMVariable "drstate", 'variable not documented but currently displays "0" - clarification required
            Application.SetSystemVariable("dtexted", 0)
            Application.SetSystemVariable("dwgcheck", 3)
            'APPLICATION.SETSYSTEMVariable "dwgcodepage"
            'APPLICATION.SETSYSTEMVariable "dwgname"
            'APPLICATION.SETSYSTEMVariable "dwgprefix"
            'APPLICATION.SETSYSTEMVariable "dwgtitled"
            'APPLICATION.SETSYSTEMVariable "dwgwrite"
            'APPLICATION.SETSYSTEMVariable "dyndigrip", 31
            'APPLICATION.SETSYSTEMVariable "dyndivis", 2
            'APPLICATION.SETSYSTEMVariable "dynmode", 3
            'APPLICATION.SETSYSTEMVariable "dynpicoords", 1
            'APPLICATION.SETSYSTEMVariable "dynpiformat", 1
            'APPLICATION.SETSYSTEMVariable "dynpivis", 1
            'APPLICATION.SETSYSTEMVariable "dynprompt", 1
            'APPLICATION.SETSYSTEMVariable "dyntooltips", 1
            Application.SetSystemVariable("edgemode", 1)
            Application.SetSystemVariable("elevation", 0)
            'APPLICATION.SETSYSTEMVariable "enterprisemenu", '"show complete path name here with file name and extension"
            'APPLICATION.SETSYSTEMVariable "entexts", 1
            'APPLICATION.SETSYSTEMVariable "entmods"
            'Application.SetSystemVariable("epdfshx", 0)
            'APPLICATION.SETSYSTEMVariable "errno"
            'APPLICATION.SETSYSTEMVariable "*errno"
            Application.SetSystemVariable("expert", 1)
            Application.SetSystemVariable("explmode", 1)
            'APPLICATION.SETSYSTEMVariable "extmax"
            'APPLICATION.SETSYSTEMVariable "extmin"
            Application.SetSystemVariable("extnames", 1)
            'APPLICATION.SETSYSTEMVariable "facetratio"
            'APPLICATION.SETSYSTEMVariable "facetres"
            'APPLICATION.SETSYSTEMVariable "fflimit"
            If Application.GetSystemVariable("fielddisplay") = 0 Then Application.SetSystemVariable("fielddisplay", 1) 'This variable causes a regen
            Application.SetSystemVariable("fieldeval", 31)
            Application.SetSystemVariable("filedia", 1)
            Application.SetSystemVariable("filletrad", 0)
            Application.SetSystemVariable("fillmode", 1)
            'APPLICATION.SETSYSTEMVariable "flatland"
            Application.SetSystemVariable("fontalt", "romans.shx")
            Application.SetSystemVariable("fontmap", ".")
            'APPLICATION.SETSYSTEMVariable "force_paging", 'variable is not documented in R2006
            'APPLICATION.SETSYSTEMVariable "frontz"
            'APPLICATION.SETSYSTEMVariable "fullopen", 'Read-only variable
            Application.SetSystemVariable("fullplotpath", 0)
            'APPLICATION.SETSYSTEMVariable "gfang", 'Variable used with Gradient Fills
            'APPLICATION.SETSYSTEMVariable "gfclr1", 'Variable used with Gradient Fills
            'APPLICATION.SETSYSTEMVariable "gfclr2", 'Variable used with Gradient Fills
            'APPLICATION.SETSYSTEMVariable "gfclrlum", 'Variable used with Gradient Fills
            'APPLICATION.SETSYSTEMVariable "gfclrstate", 'Variable used with Gradient Fills
            'APPLICATION.SETSYSTEMVariable "gfname", 'Variable used with Gradient Fills
            'APPLICATION.SETSYSTEMVariable "gfshift", 'Variable used with Gradient Fills
            Application.SetSystemVariable("gridmode", 0)
            'APPLICATION.SETSYSTEMVariable "gridunit"
            'APPLICATION.SETSYSTEMVariable "gripblock", 1 'User Preference
            Application.SetSystemVariable("gripcolor", 160)
            Application.SetSystemVariable("gripdyncolor", 140)
            Application.SetSystemVariable("griphot", 1)
            Application.SetSystemVariable("griphover", 3)
            Application.SetSystemVariable("gripobjlimit", 400)
            'Application.SetSystemVariable("grips", 1)
            'Application.SetSystemVariable("gripsize", 3)
            Application.SetSystemVariable("griptips", 1)
            'APPLICATION.SETSYSTEMVariable "halogap", Used in 3D drawings and plotting
            'APPLICATION.SETSYSTEMVariable "handles",
            'APPLICATION.SETSYSTEMVariable "hideprecision", Used in 3D drawings and plotting
            'APPLICATION.SETSYSTEMVariable "hidetext", Used in 3D drawings and plotting
            Application.SetSystemVariable("highlight", 1)
            Application.SetSystemVariable("hpang", 0)
            'APPLICATION.SETSYSTEMVariable "hpassoc", 1 'user preference
            Application.SetSystemVariable("hpbound", 1)
            Application.SetSystemVariable("hpdouble", 0)
            Application.SetSystemVariable("hpdraworder", 1)
            Application.SetSystemVariable("hpgaptol", (0.5 * scl))
            Application.SetSystemVariable("hpinherit", 0)
            'APPLICATION.SETSYSTEMVariable "hpname", "ansi31" ' user preference
            Application.SetSystemVariable("hpobjwarning", 10000)
            Application.SetSystemVariable("hporigin", "0,0")
            Application.SetSystemVariable("hporiginmode", 0)
            'APPLICATION.SETSYSTEMVariable "hpscale" 'see above code
            Application.SetSystemVariable("hpseparate", 1)
            Application.SetSystemVariable("hpspace", 1.0#)
            Application.SetSystemVariable("hyperlinkbase", "")
            Application.SetSystemVariable("imagehlt", 0)
            Application.SetSystemVariable("indexctl", 3)
            'Application.SetSystemVariable("inetlocation", "http://heroldnan:8080")
            Application.SetSystemVariable("inputhistorymode", 15)
            'Application.SetSystemVariable("insbase", insbase)
            Application.SetSystemVariable("insname", "")
            'APPLICATION.SETSYSTEMVariable("insunits", 0) 'sets the insertion units to unitless
            Application.SetSystemVariable("insunitsdefsource", 0) 'see above code
            Application.SetSystemVariable("insunitsdeftarget", 0) 'see above code
            Application.SetSystemVariable("intelligentupdate", 20)
            Application.SetSystemVariable("intersectioncolor", 257)
            Application.SetSystemVariable("intersectiondisplay", 0)
            Application.SetSystemVariable("isavebak", 1)
            Application.SetSystemVariable("isavepercent", 0)
            Application.SetSystemVariable("isolines", 8)
            'APPLICATION.SETSYSTEMVariable "lastangle", 0 Read-Only Variable
            'APPLICATION.SETSYSTEMVariable "lastpoint", "" Read-Only Variable
            'APPLICATION.SETSYSTEMVariable "lastprompt", "" Read-Only Variable
            Application.SetSystemVariable("layerfilteralert", 2)
            Application.SetSystemVariable("layoutregenctl", 2)
            'APPLICATION.SETSYSTEMVariable "lazyload"
            'APPLICATION.SETSYSTEMVariable "lenslength"
            Application.SetSystemVariable("lightingunits", 0)
            Application.SetSystemVariable("limcheck", 0)
            'Set limits based on the drawing size in Model Space
            'this will be set from the drawing size listing in the start-up dialogue
            'APPLICATION.SETSYSTEMVariable "limmax"
            'APPLICATION.SETSYSTEMVariable "limmin"
            Application.SetSystemVariable("lispinit", 1)
            'APPLICATION.SETSYSTEMVariable "locale", "" Read-Only Variable
            'APPLICATION.SETSYSTEMVariable "localrootprefix" '"C:\\Documents and Settings\\dwyatt\\Local Settings\\Application Data\\Autodesk\\AutoCAD 2006\\R16.2\\enu\\"
            'this variable could be used to set things or send things to this folder on the local C: Drive
            Application.SetSystemVariable("logfilemode", 0)
            'APPLICATION.SETSYSTEMVariable "logfilename"
            'APPLICATION.SETSYSTEMVariable "logfilepath"
            'APPLICATION.SETSYSTEMVariable "loginname", 'displays users name - should be used for drawing creation to determine last editor
            'APPLICATION.SETSYSTEMVariable "longfname"
            'APPLICATION.SETSYSTEMvariable "Ltscale" see code above
            'APPLICATION.SETSYSTEMVariable "lunits" 'This variable to remain commented out in this section as per calculation above
            'APPLICATION.SETSYSTEMVariable "luprec", 'see code above
            Application.SetSystemVariable("lwdefault", 0)
            Application.SetSystemVariable("lwdisplay", 0)
            'APPLICATION.SETSYSTEMVariable "lwunits" 'see code above
            Application.SetSystemVariable("macrotrace", 0)
            Application.SetSystemVariable("maxactvp", 64)
            'APPLICATION.SETSYSTEMVariable "maxobjmem" 'Variable is no longer used
            Application.SetSystemVariable("maxsort", 600)
            'APPLICATION.SETSYSTEMVariable "mbuttonpan", 1 The value of this variable should be decided by the user
            'set measureinit based on measurement variable
            'APPLICATION.SETSYSTEMVariable "measureinit"
            'APPLICATION.SETSYSTEMVariable "measurement" 'This variable to remain commented out in this section as per calculation above
            Application.SetSystemVariable("menuctl", 1)
            Application.SetSystemVariable("menuecho", 0)
            'APPLICATION.SETSYSTEMVariable "menuname" 'Read-only Variable
            'APPLICATION.SETSYSTEMVariable "millisecs" 'Read-only Variable
            Application.SetSystemVariable("mirrtext", 0)
            'APPLICATION.SETSYSTEMVariable("modemacro", xd2(6))
            'APPLICATION.SETSYSTEMVariable "msmstate", 0 'Turns on/off the Markup Manager Tool Palette
            Application.SetSystemVariable("msolescale", 0)
            Application.SetSystemVariable("mtexted", "internal")
            Application.SetSystemVariable("mtextfixed", 0)
            Application.SetSystemVariable("mtjigstring", "ABC abc")
            Application.SetSystemVariable("NAVVCUBEDISPLAY", 1)
            'set this variable with help from VBA
            'APPLICATION.SETSYSTEMVariable "mydocumentsprefix" Read-Only Variable
            '******************************************************************************************************
            'If APPLICATION.GETSYSTEMVARIABLE("nfwstate") = "0" Then
            '   APPLICATION.SETSYSTEMVariable "nfwstate", 0#
            'Else
            '   APPLICATION.SETSYSTEMVariable "nfwstate", 1
            'End If
            'APPLICATION.SETSYSTEMVariable "nfwstate" 'see code above - attempting to get user to understand NEW commands and Topics
            '******************************************************************************************************
            'APPLICATION.SETSYSTEMVariable "nodename" ' Read-Only Variable
            Application.SetSystemVariable("nomutt", 0) 'This variable reduces dependancy on Command Line.
            Application.SetSystemVariable("obscuredcolor", 257)
            Application.SetSystemVariable("obscuredltype", 0)
            Application.SetSystemVariable("offsetdist", 0.0#)
            Application.SetSystemVariable("offsetgaptype", 0)
            Application.SetSystemVariable("oleframe", 2)
            Application.SetSystemVariable("olehide", 0)
            Application.SetSystemVariable("olequality", 3)
            Application.SetSystemVariable("olestartup", 0)
            'APPLICATION.SETSYSTEMVariable "opmstate", 2 'Variable is undocumented but should work, clarification is required ''Read-only Variable
            'APPLICATION.SETSYSTEMVariable "orthomode", 1 'This is a USER PREFERENCE
            'APPLICATION.SETSYSTEMVariable "osmode", 4133
            Application.SetSystemVariable("osnapcoord", 2)
            Application.SetSystemVariable("osnaphatch", 0)
            'APPLICATION.SETSYSTEMVariable "osnapnodlegacy", 0 'Variable is undocumented but should work, clarification is required
            Application.SetSystemVariable("osnapoverride", 0)
            Application.SetSystemVariable("osnapz", 0)
            Application.SetSystemVariable("paletteopaque", 0)
            Application.SetSystemVariable("paperupdate", 0)
            Application.SetSystemVariable("pdfshx", 0)
            'APPLICATION.SETSYSTEMVariable "pdmode", 0
            'If APPLICATION.GETSYSTEMVARIABLE("pdsize") <> 0 Then APPLICATION.SETSYSTEMVariable "pdsize", 0  'this variable causes a regen
            Application.SetSystemVariable("peditaccept", 1)
            Application.SetSystemVariable("pellipse", 0)
            'APPLICATION.SETSYSTEMVariable "perimeter" read-only variable
            'APPLICATION.SETSYSTEMVariable "pfacevmax" read-only variable
            'APPLICATION.SETSYSTEMVariable "phandle"
            Application.SetSystemVariable("pickadd", 1)
            Application.SetSystemVariable("pickauto", 1)
            'APPLICATION.SETSYSTEMVariable "pickbox", 3 'This is a USER PREFERENCE
            Application.SetSystemVariable("pickdrag", 0)
            Application.SetSystemVariable("pickfirst", 1)
            'Application.SetSystemVariable("pickstyle", 3)
            'APPLICATION.SETSYSTEMVariable "platform" read-only variable
            Application.SetSystemVariable("plinegen", 0)
            Application.SetSystemVariable("plinetype", 2)
            Application.SetSystemVariable("plinewid", 0)
            'APPLICATION.SETSYSTEMVariable "plotid" read-only variable
            Application.SetSystemVariable("plotoffset", 1)
            Application.SetSystemVariable("plotrotmode", 1)
            'APPLICATION.SETSYSTEMVariable "plotter" obsolete
            Application.SetSystemVariable("plquiet", 1)
            Application.SetSystemVariable("polaraddang", "")
            'APPLICATION.SETSYSTEMVariable "polarang", 90 * pi / 180
            Application.SetSystemVariable("polardist", 0)
            'APPLICATION.SETSYSTEMVariable "polarmode", 4 'This is a USER PREFERENCE
            Application.SetSystemVariable("polysides", 4)
            '******************************************************************************************************
            'APPLICATION.SETSYSTEMVariable "popups", "1" 'don't know how to get this to work but as long as it stays as "1" we are ok
            '******************************************************************************************************
            'APPLICATION.SETSYSTEMVariable "previeweffect", 2
            'APPLICATION.SETSYSTEMVariable "previewfilter", 7
            'APPLICATION.SETSYSTEMVariable "product" read-only variable
            'APPLICATION.SETSYSTEMVariable "program" read-only variable
            'APPLICATION.SETSYSTEMVariable "projectname"
            Application.SetSystemVariable("projmode", 1)
            Application.SetSystemVariable("proxygraphics", 1)
            Application.SetSystemVariable("proxynotice", 0)
            Application.SetSystemVariable("proxyshow", 1)
            'Application.SetSystemVariable("proxywebsearch", 1) 'this may not work very well...
            Application.SetSystemVariable("psltscale", 1)
            'APPLICATION.SETSYSTEMVariable "psprolog"
            'APPLICATION.SETSYSTEMVariable "psquality"
            '******************************************************************************************************
            'APPLICATION.SETSYSTEMVariable "pstylemode", 1 'Must stay as 1 for CTB files to work.
            '******************************************************************************************************
            Application.SetSystemVariable("pstylepolicy", 1) 'Must stay as 1 for CTB files to work.
            '******************************************************************************************************
            Application.SetSystemVariable("psvpscale", (1 / scl)) 'Variable defines the 1/XP viewport scale factor
            'APPLICATION.SETSYSTEMVariable "pucbase"
            'APPLICATION.SETSYSTEMVariable "qaflags" 'Variable is undocumented and is VERY confusing
            'APPLICATION.SETSYSTEMVariable "qauslock" 'VAriable is undocumented and is VERY confusing
            'APPLICATION.SETSYSTEMVariable "qcstate", 1 'read only variable stating whether the quick calc is on or off
            Application.SetSystemVariable("qtextmode", 0)
            'APPLICATION.SETSYSTEMVariable "queuedregenmax" 'Variable is undocumented
            'Application.SetSystemVariable("rasterdpi", 300)
            'Application.SetSystemVariable("rasterpreview", 1)
            Application.SetSystemVariable("re-init", 0)
            Application.SetSystemVariable("recoverymode", 2)
            'APPLICATION.SETSYSTEMVariable "refeditname"
            'APPLICATION.SETSYSTEMVariable "regenmode", 1'user preference
            Application.SetSystemVariable("rememberfolders", 1)
            Application.SetSystemVariable("reporterror", 0)
            'APPLICATION.SETSYSTEMVariable "riaspect" 'Variable is obsolete after R13
            'APPLICATION.SETSYSTEMVariable "ribackg" 'Variable is obsolete after R13
            'APPLICATION.SETSYSTEMVariable "riedge" 'Variable is obsolete after R13
            'APPLICATION.SETSYSTEMVariable "rigamut" 'Variable is obsolete after R13
            'APPLICATION.SETSYSTEMVariable "rigrey" 'Variable is obsolete after R13
            'APPLICATION.SETSYSTEMVariable "rithresh" 'Variable is obsolete after R13
            'APPLICATION.SETSYSTEMVariable "roamablerootprefix"  'Vairable is read-only but displays the following:
            'ROAMABLEROOTPREFIX = "C:\Documents and Settings\dwyatt\Application "
            'Data\Autodesk\AutoCAD 2006\R16.2\enu\" (read only)
            Application.SetSystemVariable("rtdisplay", 1)
            'APPLICATION.SETSYSTEMVariable "savefile" 'Current Autosave File Name
            'APPLICATION.SETSYSTEMVariable "savefilepath" 'Current Path of Autosave File Name
            'APPLICATION.SETSYSTEMVariable "saveimages" 'Vairable is obsolete after R13
            'APPLICATION.SETSYSTEMVariable "savename" 'File Name and Directory of DWG once it has been saved
            'APPLICATION.SETSYSTEMVariable "savetime", 5
            'APPLICATION.SETSYSTEMVariable "screenboxes", 0
            'APPLICATION.SETSYSTEMVariable "screenmode", 3
            'APPLICATION.SETSYSTEMVariable "screensize"
            Application.SetSystemVariable("sdi", 0)
            Application.SetSystemVariable("selectionarea", 1)
            'APPLICATION.SETSYSTEMVariable "selectionareaopacity", 25 'User Variable
            'APPLICATION.SETSYSTEMVariable "selectionpreview", 3  This is a user preference do not uncomment bwebb 05jan06
            Application.SetSystemVariable("shadedge", 3)
            Application.SetSystemVariable("shadedif", 70)
            'APPLICATION.SETSYSTEMVariable "shortcutmenu", 11 This should be left up to the indivdual user
            Application.SetSystemVariable("showlayerusage", 0)
            'APPLICATION.SETSYSTEMVariable "shpname" 'Variable is read-only
            Application.SetSystemVariable("sigwarn", 0)
            Application.SetSystemVariable("sketchinc", 0.1)
            Application.SetSystemVariable("skpoly", 1)
            Application.SetSystemVariable("snapang", 0)
            'Application.SetSystemVariable("snapbase", snapbase)
            Application.SetSystemVariable("snapisopair", 0)
            Application.SetSystemVariable("snapmode", 0)
            Application.SetSystemVariable("snapstyl", 0)
            Application.SetSystemVariable("snaptype", 0)
            'Application.SetSystemVariable("snapunit", snapunit)
            Application.SetSystemVariable("solidcheck", 1)
            Application.SetSystemVariable("sortents", 112)
            'APPLICATION.SETSYSTEMVariable "spaceswitch"
            Application.SetSystemVariable("splframe", 0)
            Application.SetSystemVariable("splinesegs", 8)
            Application.SetSystemVariable("splinetype", 6)
            'APPLICATION.SETSYSTEMVariable "ssfound"
            Application.SetSystemVariable("sslocate", 1)
            Application.SetSystemVariable("ssmautoopen", 1)
            'APPLICATION.SETSYSTEMVariable "sspolltime", 60 'don't know how this should work
            Application.SetSystemVariable("ssmsheetstatus", 2)
            'APPLICATION.SETSYSTEMVariable "ssmstate", 1 'don't know how this should work
            Application.SetSystemVariable("standardsviolation", 2)
            'APPLICATION.SETSYSTEMVariable "startup", 0 'User Preference
            'APPLICATION.SETSYSTEMVariable "startuptoday" 'Variable is obsolete
            Application.SetSystemVariable("surftab1", 6)
            Application.SetSystemVariable("surftab2", 6)
            Application.SetSystemVariable("surftype", 6)
            Application.SetSystemVariable("surfu", 6)
            Application.SetSystemVariable("surfv", 6)
            'APPLICATION.SETSYSTEMVariable "syscodepage"
            Application.SetSystemVariable("tableindicator", 1)
            Application.SetSystemVariable("tabmode", 0)
            'APPLICATION.SETSYSTEMVariable "target"
            'APPLICATION.SETSYSTEMVariable "tbcustomize" 'this would be a great system variable for the CUI that we want to lock down
            'APPLICATION.SETSYSTEMVariable "tdcreate" 'Variable could be used to identify Edit Time etc.
            'APPLICATION.SETSYSTEMVariable "tdindwg" 'Variable could be used to identify Edit Time etc.
            'APPLICATION.SETSYSTEMVariable "tducreate" 'Variable could be used to identify Edit Time etc.
            'APPLICATION.SETSYSTEMVariable "tdupdate" 'Variable could be used to identify Edit Time etc.
            'APPLICATION.SETSYSTEMVariable "tdusrtimer" 'Variable could be used to identify Edit Time etc.
            'APPLICATION.SETSYSTEMVariable "tduupdate" 'Variable could be used to identify Edit Time etc.
            Application.SetSystemVariable("tempoverrides", 1)
            'APPLICATION.SETSYSTEMVariable "tempprefix", 'Variable to state path of temp. files. Should be saved to local drive.
            Application.SetSystemVariable("texteval", 0)
            Application.SetSystemVariable("textfill", 1)
            Application.SetSystemVariable("textqlty", 50)
            'APPLICATION.SETSYSTEMVariable "textsize" 'should be used as per the startup
            'APPLICATION.SETSYSTEMVariable "textstyle" 'should be used as per the startup
            'APPLICATION.SETSYSTEMVariable "thickness"
            'APPLICATION.SETSYSTEMVariable "tilemode", 1
            Application.SetSystemVariable("tooltipmerge", 0)
            Application.SetSystemVariable("tooltips", 1)
            'APPLICATION.SETSYSTEMVariable "tpstate", 1 'not sure here
            'APPLICATION.SETSYSTEMVariable "tracewid"
            'APPLICATION.SETSYSTEMVariable "trackpath"
            Application.SetSystemVariable("trayicons", 1)
            Application.SetSystemVariable("traynotify", 1)
            Application.SetSystemVariable("traytimeout", 0)
            If Application.GetSystemVariable("treedepth") <> 3020 Then Application.SetSystemVariable("treedepth", 3020) 'This variable causes a regen
            Application.SetSystemVariable("treemax", 10000000)
            Application.SetSystemVariable("trimmode", 1)
            Application.SetSystemVariable("trustedpaths", "C:\ProgramData\AutoDesk\ApplicationPlugins\helacad2026.bundle\Contents...;M:\Structural\r2026\...")
            Application.SetSystemVariable("tspacefac", 1)
            Application.SetSystemVariable("tspacetype", 1)
            Application.SetSystemVariable("tstackalign", 1)
            Application.SetSystemVariable("tstacksize", 75)
            'APPLICATION.SETSYSTEMVariable "ucsaxisang", 90 'don't know about this one
            'APPLICATION.SETSYSTEMVariable "ucsbase", "world" 'don't know about this one
            Application.SetSystemVariable("ucsfollow", 0)
            Application.SetSystemVariable("ucsicon", 1)
            'APPLICATION.SETSYSTEMVariable "ucsname"
            'APPLICATION.SETSYSTEMVariable "ucsorg"
            Application.SetSystemVariable("ucsortho", 1)
            Application.SetSystemVariable("ucsview", 1)
            Application.SetSystemVariable("ucsvp", 1)
            'APPLICATION.SETSYSTEMVariable "ucsxdir", "1, 0, 0"
            'APPLICATION.SETSYSTEMVariable "ucsydir", 0, 1, 0
            'APPLICATION.SETSYSTEMVariable "undoctl", 21 'don't know about this one
            'APPLICATION.SETSYSTEMVariable "undomarks", 0 'don't know about this one
            'APPLICATION.SETSYSTEMVariable "undoondisk"
            Application.SetSystemVariable("unitmode", 0)
            Application.SetSystemVariable("updatethumbnail", 15)

            '**** User Variables used by this VBA Project DO NOT USE
            'APPLICATION.SETSYSTEMVariable "useri1"
            'APPLICATION.SETSYSTEMVariable "useri2"
            'APPLICATION.SETSYSTEMVariable "useri3"
            'APPLICATION.SETSYSTEMVariable "useri4"
            'APPLICATION.SETSYSTEMVariable "useri5"
            'APPLICATION.SETSYSTEMVariable "userr1"
            'APPLICATION.SETSYSTEMVariable "userr2"
            'APPLICATION.SETSYSTEMVariable "userr3"
            'APPLICATION.SETSYSTEMVariable "userr4"
            'APPLICATION.SETSYSTEMVariable "userr5"
            'APPLICATION.SETSYSTEMVariable "users1"
            'APPLICATION.SETSYSTEMVariable "users2"
            'APPLICATION.SETSYSTEMVariable "users3"
            'APPLICATION.SETSYSTEMVariable "users4"
            'APPLICATION.SETSYSTEMVariable "users5"
            ' ****

            'APPLICATION.SETSYSTEMVariable "viewctr"
            'APPLICATION.SETSYSTEMVariable "viewdir"
            'APPLICATION.SETSYSTEMVariable "viewmode", 0 'don't know about this one
            'APPLICATION.SETSYSTEMVariable "viewsize"
            'APPLICATION.SETSYSTEMVariable "viewtwist"
            Application.SetSystemVariable("visretain", 1)
            'APPLICATION.SETSYSTEMVariable "vpmaximizedstate", 0 'don't know about this one
            'APPLICATION.SETSYSTEMVariable "vsmax"
            'APPLICATION.SETSYSTEMVariable "msmin"
            'APPLICATION.SETSYSTEMVariable "vtduration"
            'APPLICATION.SETSYSTEMVariable "vtenable"
            'APPLICATION.SETSYSTEMVariable "vtfps"
            Application.SetSystemVariable("whiparc", 1)
            'APPLICATION.SETSYSTEMVariable "whipthread"
            'APPLICATION.SETSYSTEMVariable "windowareccolor", 5 'don't know about this one
            Application.SetSystemVariable("wmfbkgnd", 0)
            Application.SetSystemVariable("wmfforegnd", 0)
            'APPLICATION.SETSYSTEMVariable "worlducs", 1 'don't know about this one
            Application.SetSystemVariable("worldview", 1)
            'APPLICATION.SETSYSTEMVariable "writestat"
            'APPLICATION.SETSYSTEMVariable "wscurrent"
            Application.SetSystemVariable("xclipframe", 0)
            Application.SetSystemVariable("xedit", 1)
            Application.SetSystemVariable("xfadectl", 50)
            Application.SetSystemVariable("xloadctl", 2)
            'APPLICATION.SETSYSTEMVariable "xloadpath"
            'APPLICATION.SETSYSTEMVariable "xrefctl"
            Application.SetSystemVariable("xrefnotify", 2)
            Application.SetSystemVariable("xreftype", 1)
            Application.SetSystemVariable("zoomfactor", 60)

            'lisp = New clsVLAX
            'lispAPPLICATION.SETSYSTEMLispSymbol("_dwsc", APPLICATION.GETSYSTEMVARIABLE("dimscale"))
            Application.SetSystemVariable("insunits", 0)

            SetSystemVariables = True
        Catch ex As Exception


        Finally

        End Try

    End Function

    Private Function TStyleCreate(TStyle As String, THeight As Double, Optional Ms As Boolean = False) As Boolean
        Dim scl As Double
        Dim uni As Double = Thisdrawing.GetVariable("UserR3")
        Dim fnt As String = "Arial"
        Dim txt As AcadTextStyle
        Dim WritetoCMD As New Clsutilities


        Try
            If Ms = True Then
                scl = Thisdrawing.GetVariable("UserR2")
            Else
                scl = 1
            End If
            If Loadfont(TStyle, fnt) = True Then
                txt = Thisdrawing.TextStyles.Item(TStyle)
                With txt
                    .Height = (THeight * scl * uni)
                    .ObliqueAngle = "0.0"
                    .Width = ".9"
                End With
            Else

            End If
        Catch Ex As Exception
            WritetoCMD.WritetoCMD("Failure to  generate test style " & TStyle & " Exception " & Ex.ToString)
        End Try
        txt = Nothing
        scl = Nothing
        uni = Nothing

        Return True
    End Function

    Public Function Text_Gen(Optional txtht As String = "ALL") As Boolean
        Text_Gen = False

        Try
            'This function generates all of the text styles

            If txtht = "ALL" Then
                txtht = "HEL15"
                TStyleCreate(txtht, 1.5, True)
                txtht = "HEL20"
                TStyleCreate(txtht, 2.0, True)
                txtht = "HEL25"
                TStyleCreate(txtht, 2.5, True)
                txtht = "HEL30"
                TStyleCreate(txtht, 3.0, True)
                txtht = "HEL35"
                TStyleCreate(txtht, 3.5, True)
                txtht = "HEL40"
                TStyleCreate(txtht, 4.0, True)
                txtht = "HEL50"
                TStyleCreate(txtht, 5.0, True)
                txtht = "PHEL15"
                TStyleCreate(txtht, 1.5)
                txtht = "PHEL20"
                TStyleCreate(txtht, 2.0)
                txtht = "PHEL25"
                TStyleCreate(txtht, 2.5)
                txtht = "PHEL30"
                TStyleCreate(txtht, 3.0)
                txtht = "PHEL35"
                TStyleCreate(txtht, 3.5)
                txtht = "PHEL40"
                TStyleCreate(txtht, 4.0)
                txtht = "PHEL50"
                TStyleCreate(txtht, 5.0)
                txtht = "HELWeld"
                TStyleCreate(txtht, 0.0)
                txtht = "HEL_sym"
                TStyleCreate(txtht, 0.0)
                txtht = "HELDim"
                TStyleCreate(txtht, 0.0)
                txtht = "HEL-Title3"
                TStyleCreate(txtht, 0.0)
                txtht = "HEL-Title2"
                TStyleCreate(txtht, 0.0)
                txtht = "HEL-Title1"
                TStyleCreate(txtht, 0.0)
            ElseIf txtht = "CLIENT" Then
                Exit Function
            ElseIf txtht = "HEL15" Then
                TStyleCreate(txtht, 1.5, True)
            ElseIf txtht = "HEL20" Then
                TStyleCreate(txtht, 2.0, True)
            ElseIf txtht = "HEL25" Then
                TStyleCreate(txtht, 2.5, True)
            ElseIf txtht = "HEL30" Then
                TStyleCreate(txtht, 3.0, True)
            ElseIf txtht = "HEL35" Then
                TStyleCreate(txtht, 3.5, True)
            ElseIf txtht = "HEL40" Then
                TStyleCreate(txtht, 4.0, True)
            ElseIf txtht = "HEL50" Then
                TStyleCreate(txtht, 5.0, True)
            ElseIf txtht = "PHEL15" Then
                TStyleCreate(txtht, 1.5)
            ElseIf txtht = "PHEL20" Then
                TStyleCreate(txtht, 2.0)
            ElseIf txtht = "PHEL25" Then
                TStyleCreate(txtht, 2.5)
            ElseIf txtht = "PHEL30" Then
                TStyleCreate(txtht, 3.0)
            ElseIf txtht = "PHEL35" Then
                TStyleCreate(txtht, 3.5)
            ElseIf txtht = "PHEL40" Then
                TStyleCreate(txtht, 4.0)
            ElseIf txtht = "PHEL50" Then
                TStyleCreate(txtht, 5.0)
            ElseIf txtht = "HELWeld" Then
                TStyleCreate(txtht, 0.0)
            ElseIf txtht = "HEL_sym" Then
                txtht = "HEL_sym"
                TStyleCreate(txtht, 0.0)
            ElseIf txtht = "HELDim" Then
                txtht = "HELDim"
                TStyleCreate(txtht, 0.0)
            End If


            'Sets current Text Style
            With Thisdrawing
                If LCase(.GetVariable("Ctab")) = "model" Then
                    If .GetVariable("UserR1") = "2" Then .SetVariable("Textstyle", "HEL20")
                    If .GetVariable("UserR1") = "2.5" Then .SetVariable("Textstyle", "HEL25")
                    If .GetVariable("UserR1") = "3" Then .SetVariable("Textstyle", "HEL30")
                Else
                    If .GetVariable("UserR1") = "2" Then .SetVariable("Textstyle", "PHEL20")
                    If .GetVariable("UserR1") = "2.5" Then .SetVariable("Textstyle", "PHEL25")
                    If .GetVariable("UserR1") = "3" Then .SetVariable("Textstyle", "PHEL30")
                End If


            End With
            Text_Gen = True

        Catch


        End Try

    End Function

    Public Function Dim_Gen() As Boolean
        'This function sets all of the dimension styles

        Dim Client As New ClsXdata
        If Client.Check_client = True Then Return False


        Dim_Gen = False

        Try
            Dim dms As String

            Dim blksty As String
            Dim blkovr As Integer
            Dim txt As Double
            Dim scl As Integer
            Dim uni As Double



            If Thisdrawing.GetVariable("users1") = "arch" _
            Or Thisdrawing.GetVariable("users1") = "blen" Then
                blksty = "_archtick"
                blkovr = 0
            Else
                blksty = "_Oblique"
                blkovr = 1
            End If

            If Thisdrawing.GetVariable("userr1") = 0 Then Thisdrawing.SetVariable("Userr1", 2.5)

            txt = Thisdrawing.GetVariable("UserR1") 'Default Text Height as Selected by User
            scl = Thisdrawing.GetVariable("UserR2") '_DWSC

            If Thisdrawing.GetVariable("userr3") = 0 Then Thisdrawing.SetVariable("Userr3", 1)

            uni = Thisdrawing.GetVariable("UserR3") 'SF
            dms = Thisdrawing.GetVariable("UserS2") 'Dimension Style Selected by User
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Dim tr As Transaction = db.TransactionManager.StartTransaction

            Dim st As DimStyleTable = tr.GetObject(db.DimStyleTableId, OpenMode.ForWrite)
            Dim str As DimStyleTableRecord
            Dim arrows As New ClsLeaders
            Dim tstyle As New ClsCheckroutines
            Dim blayercol As Color = Color.FromColorIndex(ColorMethod.ByLayer, 256)
            Dim textcol As Color = Color.FromColorIndex(ColorMethod.ByAci, 6)

            Dim dimstid As ObjectId
            Dim chkstyl As New ClsCheckroutines



            If dms = "ticks" Then
                dimstid = chkstyl.Dimstcheck("HEL-Ticks")
            Else
                dimstid = chkstyl.Dimstcheck("HEL-Arrows")
            End If

            If dimstid.IsNull = True Then
                str = New DimStyleTableRecord
            Else
                str = tr.GetObject(dimstid, OpenMode.ForWrite)
            End If

            If uni = "1" Then
                If dms = "ticks" Then
                    'Metric Ticks go here
                    str.Name = "HEL-Ticks"
                    str.Dimasz = 2.5
                    str.Dimblk = arrows.Getarrowid("_OBLIQUE") '(blksty)
                    str.Dimdle = 1.5875
                    str.Dimtsz = blkovr
                Else
                    'Metric Arrows go here
                    str.Name = "HEL-Arrows"
                    str.Dimasz = 2.5
                    str.Dimblk = arrows.Getarrowid(".")
                    str.Dimdle = 0
                    str.Dimtsz = 0

                End If
                'Standardized Metric Settings
                str.Dimadec = 0
                str.Dimatfit = 2
                str.Dimcen = 10
                str.Dimclrd = blayercol
                str.Dimclre = blayercol
                str.Dimclrt = textcol
                str.Dimdec = 0
                If txt = "2.5" Then
                    str.Dimdli = 10
                ElseIf txt = "2" Then
                    str.Dimdli = 8
                End If
                str.Dimdsep = "."
                str.Dimexe = 1
                str.Dimexo = 1.5
                str.Dimfrac = 2
                str.Dimgap = 1
                str.Dimjust = 0
                str.Dimldrblk = arrows.Getarrowid(".") 'This setting should draw an arrowhead, remember to change the LISP File!!
                str.Dimlunit = 2
                str.Dimscale = scl
                str.Dimtih = 0
                str.Dimtix = 0
                str.Dimtmove = 0
                str.Dimtofl = 1
                str.Dimtoh = 0
                str.Dimtvp = 1
                str.Dimtxsty = tstyle.Textcheck("HELDIM")
                str.Dimtxt = txt
                str.Dimupt = 0
            ElseIf uni <> "1" Then

                If dms = "ticks" Then
                    'Imperial Ticks go here
                    str.Name = ("HEL-Ticks")
                    str.Dimasz = 0.09735
                    str.Dimblk = arrows.Getarrowid("_OBLIQUE")
                    str.Dimdle = 0.0625
                    str.Dimtsz = 0
                    str.Dimgap = 0.0625


                Else
                    'Imperial Arrows go here
                    str.Name = "HEL-Arrows"
                    str.Dimasz = 0.09375
                    str.Dimblk = arrows.Getarrowid(".")
                    str.Dimdle = 0
                    str.Dimtsz = 0
                    str.Dimgap = 0.0625


                End If
                str.Dimadec = 0
                str.Dimatfit = 2
                str.Dimazin = 1
                str.Dimcen = 0.375
                str.Dimclrd = blayercol
                str.Dimclre = blayercol
                str.Dimclrt = textcol
                str.Dimdec = 4
                If txt = "2.5" Then
                    str.Dimdli = 0.375
                ElseIf txt = "2" Then
                    str.Dimdli = 0.3125
                End If
                str.Dimdsep = "."
                str.Dimexe = 0.0625
                str.Dimexo = 0.09375
                str.Dimfrac = 2
                str.Dimjust = 0
                str.Dimldrblk = arrows.Getarrowid(".") 'This setting should draw an arrowhead, remember to change the LISP File!!
                str.Dimlfac = 1
                str.Dimlim = 0
                str.Dimlunit = 4
                str.Dimrnd = 0
                str.Dimscale = scl
                str.Dimtfac = 0.75
                str.Dimtih = 0
                str.Dimtix = 0
                str.Dimtmove = 0
                str.Dimtofl = 1
                str.Dimtoh = 0
                str.Dimtvp = 1
                str.Dimtxsty = tstyle.Textcheck("HELDIM")
                str.Dimtxt = (txt * uni)
                str.Dimupt = 0
                str.Dimzin = 3

            End If

            Try
                If dimstid.IsNull = True Then
                    dimstid = st.Add(str)
                    tr.AddNewlyCreatedDBObject(str, True)
                    db.Dimstyle = dimstid
                    tr.Commit()
                Else
                    db.Dimstyle = dimstid
                    tr.Commit()
                End If
            Catch
                tr.Abort()
            End Try
            tr = db.TransactionManager.StartTransaction
            st = tr.GetObject(db.DimStyleTableId, OpenMode.ForWrite)
            If dms = "ticks" Then
                dimstid = chkstyl.Dimstcheck("PHEL-Ticks")
            Else
                dimstid = chkstyl.Dimstcheck("PHEL-Arrows")
            End If

            If dimstid.IsNull = True Then
                str = New DimStyleTableRecord
            Else
                str = tr.GetObject(dimstid, OpenMode.ForWrite)
            End If

            If uni = "1" Then
                If dms = "ticks" Then
                    'Metric Ticks go here
                    str.Name = "PHEL-Ticks"
                    str.Dimasz = 2.5
                    str.Dimblk = arrows.Getarrowid("_OBLIQUE") '(blksty)
                    str.Dimdle = 1.5875
                    str.Dimtsz = blkovr
                Else
                    'Metric Arrows go here
                    str.Name = "PHEL-Arrows"
                    str.Dimasz = 2.5
                    str.Dimblk = arrows.Getarrowid(".")
                    str.Dimdle = 0
                    str.Dimtsz = 0

                End If
                'Standardized Metric Settings
                str.Dimadec = 0
                str.Dimatfit = 2
                str.Dimcen = 10
                str.Dimclrd = blayercol
                str.Dimclre = blayercol
                str.Dimclrt = textcol
                str.Dimdec = 0
                If txt = "2.5" Then
                    str.Dimdli = 10
                ElseIf txt = "2" Then
                    str.Dimdli = 8
                End If
                str.Dimdsep = "."
                str.Dimexe = 1
                str.Dimexo = 1.5
                str.Dimfrac = 2
                str.Dimgap = 1
                str.Dimjust = 0
                str.Dimldrblk = arrows.Getarrowid(".") 'This setting should draw an arrowhead, remember to change the LISP File!!
                str.Dimlunit = 2
                str.Dimscale = 1
                str.Dimlfac = scl
                str.Dimtih = 0
                str.Dimtix = 0
                str.Dimtmove = 0
                str.Dimtofl = 1
                str.Dimtoh = 0
                str.Dimtvp = 1
                str.Dimtxsty = tstyle.Textcheck("HELDIM")
                str.Dimtxt = txt
                str.Dimupt = 0
            ElseIf uni <> "1" Then

                If dms = "ticks" Then
                    'Imperial Ticks go here
                    str.Name = ("PHEL-Ticks")
                    str.Dimasz = 0.09735
                    str.Dimblk = arrows.Getarrowid("_OBLIQUE")
                    str.Dimdle = 0.0625
                    str.Dimtsz = 0
                    str.Dimgap = 0.0625


                Else
                    'Imperial Arrows go here
                    str.Name = "PHEL-Arrows"
                    str.Dimasz = 0.09375
                    str.Dimblk = arrows.Getarrowid(".")
                    str.Dimdle = 0
                    str.Dimtsz = 0
                    str.Dimgap = 0.0625


                End If
                'Standardized Imperial Settings
                str.Dimadec = 0
                str.Dimatfit = 2
                str.Dimazin = 1
                str.Dimcen = 0.375
                str.Dimclrd = blayercol
                str.Dimclre = blayercol
                str.Dimclrt = textcol
                str.Dimdec = 4
                If txt = "2.5" Then
                    str.Dimdli = 0.375
                ElseIf txt = "2" Then
                    str.Dimdli = 0.3125
                End If
                str.Dimdsep = "."
                str.Dimexe = 0.0625
                str.Dimexo = 0.09375

                str.Dimfrac = 2

                str.Dimjust = 0
                str.Dimldrblk = arrows.Getarrowid(".") 'This setting should draw an arrowhead, remember to change the LISP File!!
                str.Dimlfac = scl
                str.Dimlim = 0
                str.Dimlunit = 4
                str.Dimrnd = 0
                str.Dimscale = 1
                str.Dimtfac = 0.75
                str.Dimtih = 0
                str.Dimtix = 0
                str.Dimtmove = 0
                str.Dimtofl = 1
                str.Dimtoh = 0
                str.Dimtvp = 1
                str.Dimtxsty = tstyle.Textcheck("HELDIM")
                str.Dimtxt = (txt * uni)


                str.Dimupt = 0
                str.Dimzin = 3

            End If

            Try
                If dimstid.IsNull = True Then
                    dimstid = st.Add(str)
                    tr.AddNewlyCreatedDBObject(str, True)
                    db.Dimstyle = dimstid
                    tr.Commit()
                Else
                    db.Dimstyle = dimstid
                    tr.Commit()
                End If
            Catch
                tr.Abort()
            End Try
            Dim_Gen = True

        Catch

        End Try

    End Function


    Public Function ML_Gen() As Boolean
        'This function sets the Mleader style

        Dim Client As New ClsXdata
        If Client.Check_client = True Then Return False


        ML_Gen = False

        Try
            Dim dms As String
            Dim txt As Double
            Dim scl As Integer
            Dim uni As Double
            Dim textstyle As String = ""


            If Thisdrawing.GetVariable("userr1") = 0 Then Thisdrawing.SetVariable("Userr1", 2.5)

            txt = Thisdrawing.GetVariable("UserR1") 'Default Text Height as Selected by User

            If txt = "2" Then
                textstyle = "HEL20"
            ElseIf txt = "2.5" Then
                textstyle = "HEL25"
            End If

            scl = Thisdrawing.GetVariable("UserR2") '_DWSC

            If Thisdrawing.GetVariable("userr3") = 0 Then Thisdrawing.SetVariable("Userr3", 1)

            uni = Thisdrawing.GetVariable("UserR3") 'SF
            dms = Thisdrawing.GetVariable("UserS2") 'Dimension Style Selected by User

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Dim tr As Transaction = db.TransactionManager.StartTransaction

            Const Stylename As String = "HEL"
            Dim mlstyles As DBDictionary = tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForWrite)
            Dim tstyle As New ClsCheckroutines
            Dim mleaderstyle As ObjectId = ObjectId.Null


            If mlstyles.Contains(Stylename) Then
                mleaderstyle = mlstyles.GetAt(Stylename)
            Else

                Dim btable As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                If btable.Has("_DOT") = False Then
                    Application.SetSystemVariable("DIMBLK", "_DOT")
                End If
                Dim Newstyle As New MLeaderStyle With {
                .ArrowSize = 2.5 * scl * uni,
                .TextAlignAlwaysLeft = True,
                .TextStyleId = tstyle.Textcheck(textstyle),
                .TextHeight = txt * scl * uni,
                .LandingGap = .TextHeight / 2,
                .DoglegLength = .TextHeight / 2,
                .ExtendLeaderToText = True,
                .LeaderLineColor = Color.FromColorIndex(ColorMethod.ByPen, 1)
                }
                Newstyle.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader)
                Newstyle.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader)

                mleaderstyle = Newstyle.PostMLeaderStyleToDb(db, Stylename)
                tr.AddNewlyCreatedDBObject(Newstyle, True)

            End If

            db.MLeaderstyle = mleaderstyle

            tr.Commit()
            tr.Dispose()

            ML_Gen = True
        Catch ex As Exception

        End Try

    End Function
    Public Function Text(ByVal txtht As String)

        Dim txttype As String
        Dim txtlay As String

        Dim check_routines As New ClsCheckroutines
        Dim Client As New ClsXdata
        If Client.Check_client = True Then GoTo text_create

        HELACADCustomization.Customization.curlayer = Thisdrawing.ActiveLayer.Name
        txtlay = "~-TXT" & txtht

        txtht = "HEL" & txtht


        check_routines.SetLayerCurrent(txtlay)

        If Thisdrawing.ActiveSpace = AcActiveSpace.acModelSpace Then

        Else
            txtht = "P" & txtht
        End If

        If check_routines.Txtchk(txtht) = False Then
            If Text_Gen(txtht) = True Then
                Thisdrawing.ActiveTextStyle = Thisdrawing.TextStyles.Item(txtht)
            End If
        Else
            Thisdrawing.ActiveTextStyle = Thisdrawing.TextStyles.Item(txtht)
        End If
text_create:

        txttype = Thisdrawing.GetVariable("Users5")

        If txttype = "DT" Then
            Thisdrawing.SendCommand("dtext ")
        ElseIf txttype = "MT" Then
            Thisdrawing.SendCommand("mtext ")
        Else
            Thisdrawing.SendCommand("mtext ")
        End If
        Text = Nothing

    End Function

    <CommandMethod("20mm")> Public Sub Tzmmtext()
        Text("20")
    End Sub
    <CommandMethod("25mm")> Public Sub T5mm_text()
        Text("25")
    End Sub
    <CommandMethod("30mm")> Public Sub Thmm_text()
        Text("30")
    End Sub
    <CommandMethod("35mm")> Public Sub Th5mm_text()
        Text("35")
    End Sub
    <CommandMethod("40mm")> Public Sub Ftmm_text()
        Text("40")
    End Sub
    <CommandMethod("50mm")> Public Sub Ffmm_text()
        Text("50")
    End Sub
    <CommandMethod("chgtxttype")> Public Sub ChangeTextType()
        Dim txttype As String
        ''''Dim lisp As clsVLAX = New clsVLAX

        txttype = Thisdrawing.GetVariable("Users5")

        If txttype = "DT" Then
            Thisdrawing.SetVariable("Users5", "MT")
        ElseIf txttype = "MT" Then
            Thisdrawing.SetVariable("Users5", "DT")
        Else
            Thisdrawing.SetVariable("Users5", "MT")
        End If
    End Sub

    Public Function Loadfont(Newstyle As String, fonttoload As String) As Boolean
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        Dim db As Database = doc.Database

        Dim tm As Transaction = db.TransactionManager.StartTransaction()

        Dim ed As Editor = doc.Editor
        Try
            Using tm

                Dim st As TextStyleTable = CType(tm.GetObject(db.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)

                Dim str As New TextStyleTableRecord With {
                .Name = Newstyle
                }
                Try
                    st.Add(str)

                    str.Font = New Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(fonttoload, False, False, Nothing, Nothing)

                    tm.AddNewlyCreatedDBObject(str, True)

                    tm.Commit()
                Catch
                    For Each obj In st
                        Dim ESTR As TextStyleTableRecord
                        ESTR = tm.GetObject(obj, OpenMode.ForWrite, False)
                        'If Left(ESTR.Name, 3 = "HEL" Or Left(ESTR.Name, 4) = "PHEL" then
                        If ESTR.Name = Newstyle Then
                            ESTR.Font = New Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(fonttoload, False, False, Nothing, Nothing)
                        End If
                    Next
                    ''st.Item(str.Name)
                    'str.Font = New Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(fonttoload, False, False, Nothing, Nothing)
                    'tm.AddNewlyCreatedDBObject(str, True)
                    tm.Commit()
                    Return True
                End Try
            End Using
            Return True
        Catch
            tm.Abort()
            Return False
        End Try
        tm.Dispose()
    End Function
End Class
