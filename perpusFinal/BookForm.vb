Imports System.Data.SqlClient
Public Class BookForm
    Sub KondisiAwal()

        Call Koneksi()
        Da = New SqlDataAdapter("Select * From TB_BUKU", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "TB_BUKU")
        DataGridView1.DataSource = (Ds.Tables("TB_BUKU"))
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        GenerateIdBuku()
        ' ComboBox1.Text = ""
    End Sub

    Sub GenerateIdBuku()
        Call Koneksi()
        Cmd = New SqlCommand("Select * from TB_BUKU where IdBuku in (select max(IdBuku) from TB_BUKU)", Conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        Dr = Cmd.ExecuteReader
        Dr.Read()
        If Not Dr.HasRows Then
            UrutanKode = "BOOK-" + "000001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Dr.GetString(0), 6) + 1
            UrutanKode = "BOOK-" + Microsoft.VisualBasic.Right("00000" & Hitung, 6)
        End If
        TextBox1.Text = UrutanKode
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MsgBox("Pastikan Di isi terlebih dahulu")
        Else
            Call Koneksi()
            Cmd = New SqlCommand("insert into TB_BUKU values('" & TextBox1.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & TextBox5.Text & "', '" & TextBox6.Text & "')", Conn)
            Cmd.ExecuteNonQuery() 'untuk read data yang sudah dibikin dicommand atas
            MsgBox("Data Berhasil Di Buat")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub BookForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MsgBox("Pastikan Di isi terlebih dahulu")
        Else
            Call Koneksi()
            Cmd = New SqlCommand("update TB_BUKU set JudulBuku='" & TextBox2.Text & "', PenerbitBuku='" & TextBox3.Text & "', PengarangBuku='" & TextBox4.Text & "', JumlahBuku='" & TextBox5.Text & "', TahunBuku='" & TextBox6.Text & "'  where IdBuku='" & TextBox1.Text & "' ", Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Edit")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MsgBox("Pastikan Di Pilih terlebih dahulu terlebih dahulu")
        Else

            Call Koneksi()
            Cmd = New SqlCommand("delete TB_BUKU where IdBuku='" & TextBox1.Text & "' ", Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Delete")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call KondisiAwal()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()

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


    End Sub
End Class