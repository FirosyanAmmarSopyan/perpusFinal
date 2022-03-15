Imports System.Data.SqlClient
Public Class ReturnForm
    Sub KondisiAwal()

        Call Koneksi()
        Da = New SqlDataAdapter("Select * From TB_PEMINJAMAN", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "TB_PEMINJAMAN")
        DataGridView1.DataSource = (Ds.Tables("TB_PEMINJAMAN"))
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        ListBox1.Text = ""
        ListBox2.Text = ""

        GenerateIdReturn()
        ' ComboBox1.Text = ""
    End Sub
    Sub GenerateIdReturn()
        Call Koneksi()
        Cmd = New SqlCommand("Select * from TB_PENGEMBALIAN where IdPengembalian in (select max(IdPengembalian) from TB_PENGEMBALIAN)", Conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        Dr = Cmd.ExecuteReader
        Dr.Read()
        If Not Dr.HasRows Then
            UrutanKode = "RTN-" + "000001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Dr.GetString(0), 6) + 1
            UrutanKode = "RTN-" + Microsoft.VisualBasic.Right("00000" & Hitung, 6)
        End If
        TextBox1.Text = UrutanKode
    End Sub

End Class