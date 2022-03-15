Imports System.Data.SqlClient
Public Class FormLogin
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Koneksi()
        Cmd = New SqlCommand("Select * From TB_PETUGAS where IdPetugas='" & TextBox1.Text & "'and PasswordPetugas='" & TextBox2.Text & "'", Conn)
        Dr = Cmd.ExecuteReader
        Dr.Read()
        If Dr.HasRows Then
            Me.Hide()
            FormMenu.Show()

        Else
            MsgBox("Id atau Password yang anda masukkan Salah!")
            TextBox1.Text = ""
            TextBox2.Text = ""

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub FormLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
