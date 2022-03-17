Imports System.Data.SqlClient
Public Class LoanForm
    Sub KondisiAwal()
        Call IdAnggotaOtomatis()
        Call IdPetugasOtomatis()
        Call Koneksi()
        Da = New SqlDataAdapter("Select * From TB_PEMINJAMAN", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "TB_PEMINJAMAN")
        DataGridView1.DataSource = (Ds.Tables("TB_PEMINJAMAN"))
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        'TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        ListBox1.Text = ""
        ListBox2.Text = ""

        GenerateIdLoan()
        ' ComboBox1.Text = ""
    End Sub

    Sub IdAnggotaOtomatis()
        Call Koneksi()
        ListBox1.Items.Clear()
        Cmd = New SqlCommand("Select * from TB_ANGGOTA", Conn)
        Dr = Cmd.ExecuteReader
        Do While Dr.Read
            ListBox1.Items.Add(Dr.Item(0))
        Loop

    End Sub
    Sub IdPetugasOtomatis()
        Call Koneksi()
        ListBox2.Items.Clear()
        Cmd = New SqlCommand("Select * from TB_PETUGAS", Conn)
        Dr = Cmd.ExecuteReader
        Do While Dr.Read
            ListBox2.Items.Add(Dr.Item(0))
        Loop

    End Sub


    Sub GenerateIdLoan()
        Call Koneksi()
        Cmd = New SqlCommand("Select * from TB_PEMINJAMAN where IdPinjam in (select max(IdPinjam) from TB_PEMINJAMAN)", Conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        Dr = Cmd.ExecuteReader
        Dr.Read()
        If Not Dr.HasRows Then
            UrutanKode = "LOAN-" + "000001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Dr.GetString(0), 6) + 1
            UrutanKode = "LOAN-" + Microsoft.VisualBasic.Right("00000" & Hitung, 6)
        End If
        TextBox1.Text = UrutanKode
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Pastikan Di isi terlebih dahulu")
        Else
            Call Koneksi()
            'call koneksi langsung nak form load ben ga manggil satu satu iyo gampang wes nek iku
            Cmd = New SqlCommand("insert into TB_PEMINJAMAN values('" & TextBox1.Text & "', '" & DateTimePicker1.Value.Date.ToString & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & TextBox5.Text & "', '" & ListBox1.SelectedItem & "', '" & TextBox6.Text & "', '" & ListBox2.SelectedItem & "' , '" & TextBox7.Text & "', '" & DateTimePicker2.Value.Date.ToString & "')", Conn)
            Cmd.ExecuteNonQuery() 'untuk read data yang sudah dibikin dicommand atas
            'buat penjumlahan stok
            Cmd = New SqlCommand("update TB_BUKU set JumlahBuku = " & Val(TextBox4.Text) - Val(TextBox5.Text) & " where IdBuku = '" & TextBox2.Text & "'  ", Conn)
            Cmd.ExecuteNonQuery()

            MsgBox("Data Berhasil Di Buat")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Pastikan Di pilih terlebih dahulu")
        Else

            Call Koneksi()
            Cmd = New SqlCommand("delete TB_PEMINJAMAN where IdPinjam='" & TextBox1.Text & "' ", Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Delete")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Pastikan Di isi terlebih dahulu")
        Else
            Call Koneksi()
            Cmd = New SqlCommand("update TB_PEMINJAMAN set TanggalPinjam='" & DateTimePicker1.Value.Date.ToString & "', IdBuku='" & TextBox2.Text & "', JudulBuku='" & TextBox3.Text & "', JumlahBuku='" & TextBox4.Text & "', QtyBuku='" & TextBox5.Text & "', IdAnggota='" & ListBox1.SelectedItem & "', NamaAnggota='" & TextBox6.Text & "', IdPetugas='" & ListBox2.SelectedItem & "', NamaPetugas='" & TextBox7.Text & "', TanggalPengembalian='" & DateTimePicker2.Value.Date.ToString & "'  where IdPinjam='" & TextBox1.Text & "' ", Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Edit")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call KondisiAwal()

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub LoanForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        TextBox5.Text = "1"

    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer
        On Error Resume Next ' cara ngatasi error e

        i = DataGridView1.CurrentRow.Index
        'variabel i kudu int, soale ngitung index datagrid e, kek index e array, dimulai dari 0
        TextBox1.Text = DataGridView1.Item(0, i).Value
        DateTimePicker1.Text = DataGridView1.Item(1, i).Value
        TextBox2.Text = DataGridView1.Item(2, i).Value
        TextBox3.Text = DataGridView1.Item(3, i).Value
        TextBox4.Text = DataGridView1.Item(4, i).Value
        TextBox5.Text = DataGridView1.Item(5, i).Value
        ListBox1.Text = DataGridView1.Item(6, i).Value
        TextBox6.Text = DataGridView1.Item(7, i).Value
        ListBox2.Text = DataGridView1.Item(8, i).Value
        TextBox7.Text = DataGridView1.Item(9, i).Value
        DateTimePicker2.Text = DataGridView1.Item(10, i).Value

    End Sub



    Private Sub ListBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseClick
        Call Koneksi()

        Dim lb As ListBox = DirectCast(sender, ListBox)
        Dim selectedIdAnggota = lb.SelectedItem
        Cmd = New SqlCommand("Select * from TB_ANGGOTA where IdAnggota='" & selectedIdAnggota & "'", Conn)
        Dr = Cmd.ExecuteReader
        Dr.Read()
        If Dr.HasRows Then
            TextBox6.Text = Dr("NamaAnggota")
        End If

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged

    End Sub

    Private Sub ListBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles ListBox2.MouseClick
        Call Koneksi()
        Dim lb2 As ListBox = DirectCast(sender, ListBox)
        Dim selectedIdPetugas = lb2.SelectedItem
        Cmd = New SqlCommand("Select * from TB_PETUGAS where IdPetugas='" & selectedIdPetugas & "'", Conn)
        Dr = Cmd.ExecuteReader
        Dr.Read()
        If Dr.HasRows Then
            TextBox7.Text = Dr("NamaPetugas")
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call Koneksi()
        Cmd = New SqlCommand("Select * from TB_BUKU where IdBuku = '" & TextBox2.Text & "'", Conn)
        Dr = Cmd.ExecuteReader
        Dr.Read()
        If Dr.HasRows Then
            TextBox3.Text = Dr("JudulBuku")
            TextBox4.Text = Dr("JumlahBuku")

        Else
            MsgBox("Id Buku Tidak Ditemukan !")
            TextBox3.Text = ""
        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        Call Koneksi()
        Da = New SqlDataAdapter("select * from TB_PEMINJAMAN where IdPinjam like '%" & TextBox8.Text & "%' or IdBuku like '%" & TextBox8.Text & "%' or JudulBuku like '%" & TextBox8.Text & "%' ", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "TB_PEMINJAMAN")
        DataGridView1.DataSource = (Ds.Tables("TB_PEMINJAMAN"))
    End Sub
End Class

