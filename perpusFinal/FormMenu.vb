﻿Public Class FormMenu
    Private Sub PetugasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PetugasToolStripMenuItem.Click
        EmployeeForm.ShowDialog()

    End Sub

    Private Sub AnggotaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnggotaToolStripMenuItem.Click
        MemberForm.ShowDialog()

    End Sub

    Private Sub BukuToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BukuToolStripMenuItem.Click
        BookForm.ShowDialog()

    End Sub

    Private Sub LoanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoanToolStripMenuItem.Click
        LoanForm.ShowDialog()

    End Sub
End Class