Imports System.Data.SqlClient

Public Class DatabaseHelper
    Public conn As New SqlConnection

    Public Function OpenConnection() As Boolean
        Dim strMySQLConnectionString As String
        strMySQLConnectionString = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=""D:\Projects\Online Bookshop MVC\Online Bookshop MVC\App_Data\DBBookShop.mdf"";Integrated Security=True;Connect Timeout=30"

        Try
            conn.ConnectionString = strMySQLConnectionString
            conn.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub CloseConnection()
        conn.Close()
    End Sub
End Class
