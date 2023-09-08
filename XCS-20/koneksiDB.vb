Imports System.Data.SqlClient
Module koneksiDB
    Public Function Database() As String
        Dim db As String = "Data Source=192.168.1.188\SQLEXPRESS;
            initial catalog=TESE;
            Persist Security Info=True;
            User ID=tese;
            Password=Sanindo123;
            Connect Timeout=15000;
            Max Pool Size=15000;
            Pooling=True"
        Return db
    End Function

    Public Function bacaData(query As String) As DataSet
        Try
            Dim konek As New SqlConnection(Database)
            Dim sc As New SqlCommand(query, konek)
            Dim adapter As New SqlDataAdapter(sc)
            Dim ds As New DataSet

            If konek.State = ConnectionState.Closed Then konek.Open()

            adapter.Fill(ds)
            'koneksi.Close()
            Return ds
        Catch ex As Exception
            MsgBox("Database connection Error!")
        End Try
    End Function
End Module
