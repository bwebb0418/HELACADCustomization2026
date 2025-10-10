Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Windows


Public Class ClsVLAX
    Public ReadOnly Property Thisdrawing() As AcadDocument
        Get
            Return DocumentExtension.GetAcadDocument(Application.DocumentManager.MdiActiveDocument)
        End Get
    End Property

    ' VLAX.CLS v1.4 (Last updated 8/27/2001)
    ' Copyright 1999-2001 by Frank Oquendo
    '
    ' Permission to use, copy, modify, and distribute this software
    ' for any purpose and without fee is hereby granted, provided
    ' that the above copyright notice appears in all copies and
    ' that both that copyright notice and the limited warranty and
    ' restricted rights notice below appear in all supporting
    ' documentation.
    '
    ' FRANK OQUENDO (THE AUTHOR) PROVIDES THIS PROGRAM "AS IS" AND WITH
    ' ALL FAULTS. THE AUTHOR SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY
    ' OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  THE AUTHOR
    ' DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
    ' UNINTERRUPTED OR ERROR FREE.
    '
    ' Use, duplication, or disclosure by the U.S. Government is subject to
    ' restrictions set forth in FAR 52.227-19 (Commercial Computer
    ' Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
    ' (Rights in Technical Data and Computer Software), as applicable.
    '
    ' VLAX.cls allows developers to evaluate AutoLISP expressions from
    ' Visual Basic or VBA
    '
    ' Notes:
    ' All code for this class module is publicly available througout various posts
    ' at news://discussion.autodesk.com/autodesk.autocad.customization.vba. I do not
    ' claim copyright or authorship on code presented in these posts, only on this
    ' compilation of that code. In addition, a great big "Thank you!" to Cyrille Fauvel
    ' demonstrating the use of the VisualLISP ActiveX Module.
    '
    ' Dependencies:
    ' Use of this class module requires the following application:
    ' 1. VisualLISP

    Private VL As Object
    Private VLF As Object

    Public Sub New()

        VL = Thisdrawing.Application.GetInterfaceObject("VL.Application.16")
        VLF = VL.ActiveDocument.Functions

    End Sub

    Private Sub Class_finalize()

        VLF = Nothing
        VL = Nothing

    End Sub

    Public Function EvalLispExpression(ByVal lispStatement As String)

        Dim sym As Object, ret As Object, retval

        sym = VLF.Item("read").funcall(lispStatement)
        On Error Resume Next
        retval = VLF.Item("eval").funcall(sym)
        If IsError(True) Then
            EvalLispExpression = ""
        Else
            EvalLispExpression = retval
        End If

    End Function

    Public Sub SetLispSymbol(ByVal symbolName As String, ByVal value As String)

        Dim sym As Object, ret, symValue

        symValue = value
        sym = VLF.Item("read").funcall(symbolName)
        ret = VLF.Item("set").funcall(sym, symValue)
        EvalLispExpression("(defun translate-variant (data) (cond ((= (type data) 'list) (mapcar 'translate-variant data)) ((= (type data) 'variant) (translate-variant (vlax-variant-value data))) ((= (type data) 'safearray) (mapcar 'translate-variant (vlax-safearray->list data))) (t data)))")
        EvalLispExpression("(setq " & symbolName & "(translate-variant " & symbolName & "))")
        EvalLispExpression("(setq translate-variant nil)")

    End Sub

    Public Function GetLispSymbol(ByVal symbolName As String) ', ByVal value As String)

        Dim sym As Object

        sym = VLF.Item("read").funcall(symbolName)
        GetLispSymbol = VLF.Item("eval").funcall(sym)

    End Function

    Public Function GetLispList(ByVal symbolName As String) As Object

        Dim sym As Object, list As Object
        Dim Count, elements(), I As Long

        sym = VLF.Item("read").funcall(symbolName)
        list = VLF.Item("eval").funcall(sym)

        Count = VLF.Item("length").funcall(list)

        ReDim elements(0 To Count - 1)

        For I = 0 To Count - 1
            elements(I) = VLF.Item("nth").funcall(I, list)
        Next

        GetLispList = elements

    End Function

    Public Sub NullifySymbol(ByVal ParamArray symbolName())

        Dim I As Integer

        For I = LBound(symbolName) To UBound(symbolName)
            EvalLispExpression("(setq " & CStr(symbolName(I)) & " nil)")
        Next

    End Sub

End Class
