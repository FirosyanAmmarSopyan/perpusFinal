Imports System.Data.SqlClient
Module Module1
    Public Conn As SqlConnection
    Public Da As SqlDataAdapter
    Public Ds As DataSet
    Public Dr As SqlDataReader
    Public Cmd As SqlCommand
    Public MyDb As String

    Public Sub Koneksi()
        MyDb = "data source=LAPTOP-8GJ9N7KR; initial catalog=DBperpus;integrated security=true"
        Conn = New SqlConnection(MyDb)
        If Conn.State = ConnectionState.Closed Then Conn.Open()

    End Sub

End Module
