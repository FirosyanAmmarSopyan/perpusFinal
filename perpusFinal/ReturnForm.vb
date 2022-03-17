Imports System.Data.SqlClient
Public Class ReturnForm
    Sub KondisiAwal()

        Call Koneksi()
        Da = New SqlDataAdapter("Select * From TB_PENGEMBALIAN", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "TB_PENGEMBALIAN")
        DataGridView1.DataSource = (Ds.Tables("TB_PENGEMBALIAN"))
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""


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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call KondisiAwal()

    End Sub

    Private Sub ReturnForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Then
            MsgBox("Pastikan Di pilih terlebih dahulu")
        Else

            Call Koneksi()
            Cmd = New SqlCommand("delete TB_PENGEMBALIAN where IdPengembalian='" & TextBox1.Text & "' ", Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Delete")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or DateTimePicker1.Value.Date.ToString = "" Or DateTimePicker2.Value.Date.ToString = "" Or DateTimePicker3.Value.Date.ToString = "" Then
            MsgBox("Pastikan Di pilih terlebih dahulu")
        Else
            Call Koneksi()
            Cmd = New SqlCommand("update TB_PENGEMBALIAN set IdPinjam='" & TextBox2.Text & "', IdBuku='" & TextBox3.Text & "', JudulBuku='" & TextBox4.Text & "', IdAnggota='" & TextBox5.Text & "', NamaAnggota='" & TextBox6.Text & "', IdPetugas='" & TextBox7.Text & "', NamaPetugas='" & TextBox8.Text & "', MoneyFine = '" & TextBox10.Text & "' where IdPengembalian='" & TextBox1.Text & "' ", Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Edit")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Then
            MsgBox("Pastikan Di pilih terlebih dahulu")
        Else
            Call Koneksi()
            Cmd = New SqlCommand("select * from TB_PEMINJAMAN where IdPinjam = '" & TextBox2.Text & "' ", Conn)
            Dr = Cmd.ExecuteReader
            Dr.Read()
            If Dr.HasRows Then
                Cmd = New SqlCommand("insert into TB_PENGEMBALIAN values('" & TextBox1.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & TextBox5.Text & "', '" & TextBox6.Text & "' , '" & TextBox7.Text & "', '" & TextBox8.Text & "', '" & DateTimePicker1.Value.Date.ToString & "', '" & DateTimePicker2.Value.Date.ToString & "','" & Dr("QtyBuku").ToString & "', '" & TextBox10.Text & "' )", Conn)
                Dr.Close()
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil Di Buat")
                Call KondisiAwal()
            End If

        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call Koneksi()
        Cmd = New SqlCommand("Select * from TB_PEMINJAMAN where IdPinjam = '" & TextBox2.Text & "'", Conn)
        Dr = Cmd.ExecuteReader
        Dr.Read()
        If Dr.HasRows Then
            TextBox3.Text = Dr("IdBuku")
            TextBox4.Text = Dr("JudulBuku")
            TextBox5.Text = Dr("IdAnggota")
            TextBox6.Text = Dr("NamaAnggota")
            TextBox7.Text = Dr("IdPetugas")
            TextBox8.Text = Dr("NamaPetugas")
            DateTimePicker1.Value = Dr("TanggalPinjam")
            DateTimePicker2.Value = Dr("TanggalKembali")

        Else
            MsgBox("Id PINJAM Tidak Ditemukan !")
            TextBox3.Text = ""
        End If


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer
        On Error Resume Next ' cara ngatasi error e

        i = DataGridView1.CurrentRow.Index
        'variabel i kudu int, soale ngitung index datagrid e, kek index e array, dimulai dari 0
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
        TextBox4.Text = DataGridView1.Item(3, i).Value
        TextBox5.Text = DataGridView1.Item(4, i).Value
        TextBox6.Text = DataGridView1.Item(5, i).Value
        TextBox7.Text = DataGridView1.Item(6, i).Value
        TextBox8.Text = DataGridView1.Item(7, i).Value
        TextBox7.Text = DataGridView1.Item(8, i).Value
        DateTimePicker1.Text = DataGridView1.Item(9, i).Value
        DateTimePicker2.Text = DataGridView1.Item(10, i).Value

        TextBox10.Text = DataGridView1.Item(11, i).Value
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        Call Koneksi()
        Da = New SqlDataAdapter("select * from TB_PENGEMBALIAN where IdPengembalian like '%" & TextBox8.Text & "%' or IdPinjam like '%" & TextBox8.Text & "%' or IdBuku like '%" & TextBox8.Text & "%' ", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "TB_PENGEMBALIAN")
        DataGridView1.DataSource = (Ds.Tables("TB_PENGEMBALIAN"))
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class