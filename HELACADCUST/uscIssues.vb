Public Class uscIssues
    Dim clsiss As clsIssues

    Private Sub cmdIFRS_Click(sender As System.Object, e As System.EventArgs) Handles cmdIFRS.Click
        clsiss = New clsIssues
        clsiss.dte = dtpIssueDate.Text
        If rdoother.Checked = True Then
            clsiss.IFRS("")
            'ElseIf rdo30.Checked = True Then
            '    clsiss.IFRS("30%")
        ElseIf rdo50.Checked = True Then
            clsiss.IFRS("50%")
        ElseIf rdo75.Checked = True Then
            clsiss.IFRS("75%")
        ElseIf rdo100.Checked = True Then
            clsiss.IFRS("100%")
        ElseIf rdofinal.Checked = True Then
            clsiss.IFRS("FINAL")
        End If
        clsiss.dte = Nothing
    End Sub

    Private Sub cmdifr_Click(sender As System.Object, e As System.EventArgs) Handles cmdifr.Click
        clsiss = New clsIssues
        clsiss.dte = dtpIssueDate.Text
        If rdoother.Checked = True Then
            clsiss.IFR("")
            'ElseIf rdo30.Checked = True Then
            '   clsiss.IFR("30%")
        ElseIf rdo50.Checked = True Then
            clsiss.IFR("50%")
        ElseIf rdo75.Checked = True Then
            clsiss.IFR("75%")
        ElseIf rdo100.Checked = True Then
            clsiss.IFR("100%")
        ElseIf rdofinal.Checked = True Then
            clsiss.IFR("FINAL")
        End If
        clsiss.dte = Nothing
    End Sub

    Private Sub cmdift_Click(sender As System.Object, e As System.EventArgs) Handles cmdift.Click
        clsiss = New clsIssues
        clsiss.dte = dtpIssueDate.Text
        clsiss.ift()
        clsiss.dte = Nothing
    End Sub

    Private Sub cmdIFC_Click(sender As System.Object, e As System.EventArgs) Handles cmdIFC.Click
        clsiss = New clsIssues
        clsiss.dte = dtpIssueDate.Text
        clsiss.IFC()
        clsiss.dte = Nothing
    End Sub

    Private Sub cmdIFI_Click(sender As System.Object, e As System.EventArgs) Handles cmdIFI.Click
        clsiss = New clsIssues
        clsiss.dte = dtpIssueDate.Text
        clsiss.ifi()
        clsiss.dte = Nothing
    End Sub

    Private Sub cmdifa_Click(sender As Object, e As EventArgs) Handles cmdifa.Click
        clsiss = New clsIssues
        clsiss.dte = dtpIssueDate.Text
        clsiss.IFA()
        clsiss.dte = Nothing
    End Sub
End Class
