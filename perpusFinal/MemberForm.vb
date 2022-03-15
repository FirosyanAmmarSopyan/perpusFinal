Imports System.Data.SqlClient
Public Class MemberForm

    Sub KondisiAwal()

        Call Koneksi()
        Da = New SqlDataAdapter("Select IdAnggota,NamaAnggota,AlamatAnggota,TeleponAnggota From TB_ANGGOTA", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "TB_ANGGOTA")
        DataGridView1.DataSource = (Ds.Tables("TB_ANGGOTA"))
        ' ComboBox1.Items.Clear()
        ' ComboBox1.Items.Add("ADMIN")
        ' ComboBox1.Items.Add("USER")
        ' ListBox1.Items.Clear()
        ' ListBox1.Items.Add("ADMIN")
        ' ListBox1.Items.Add("SUPERMASTER")
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ' ComboBox1.Text = ""
        GenerateIdMember()

    End Sub

    Sub GenerateIdMember()
        Call Koneksi()
        Cmd = New SqlCommand("Select * from TB_ANGGOTA where IdAnggota in (select max(IdAnggota) from TB_ANGGOTA)", Conn)
        Dim UrutanKode As String
        Dim Hitung As Long
        Dr = Cmd.ExecuteReader
        Dr.Read()
        If Not Dr.HasRows Then
            UrutanKode = "M-" + "000001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Dr.GetString(0), 6) + 1
            UrutanKode = "M-" + Microsoft.VisualBasic.Right("00000" & Hitung, 6)
        End If
        TextBox1.Text = UrutanKode
    End Sub


    Private Sub MemberForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Pastikan Di isi terlebih dahulu")
        Else
            Call Koneksi()
            Cmd = New SqlCommand("insert into TB_ANGGOTA values('" & TextBox1.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "')", Conn)
            Cmd.ExecuteNonQuery() 'untuk read data yang sudah dibikin dicommand atas
            MsgBox("Data Berhasil Di Buat")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Pastikan Di isi terlebih dahulu")
        Else
            Call Koneksi()
            Cmd = New SqlCommand("update TB_ANGGOTA set NamaAnggota='" & TextBox2.Text & "', AlamatAnggota='" & TextBox3.Text & "', TeleponAnggota='" & TextBox4.Text & "' where IdAnggota='" & TextBox1.Text & "' ", Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Edit")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Pastikan Di Pilih terlebih dahulu terlebih dahulu")
        Else

            Call Koneksi()
            Cmd = New SqlCommand("delete TB_ANGGOTA where IdAnggota='" & TextBox1.Text & "' ", Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Delete")
            Call KondisiAwal()
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

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call KondisiAwal()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()

    End Sub
End Class