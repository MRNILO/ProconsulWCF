Imports System.Data.Odbc

Module ODBC
    Public Class CConexion

        Public ODBCconStr As String
        Public ODBCcon_A As IDbConnection = New OdbcConnection(ODBCconStr)
        Public ODBCcon_B As IDbConnection = New OdbcConnection(ODBCconStr)
        Public ODBC_CMD As IDbCommand = ODBCcon_A.CreateCommand()
        Public ODBC_DA As IDbDataAdapter = New OdbcDataAdapter
        Public ODBC_DR As OdbcDataReader
    End Class
    Public Function ODBCGetDataset(ByVal LocalSQL As String, Optional ByVal NumEmp As Integer = 11) As DataSet
        Dim ODBC_DS As New DataSet
        Dim Clase As CConexion
        Clase = New CConexion

        '
        Clase.ODBCconStr += "DSN=EK_ADM" & NumEmp & "_11;UID=user_read;PWD=lectura"

        Clase.ODBCcon_A.ConnectionString = Clase.ODBCconStr
        Clase.ODBCcon_B.ConnectionString = Clase.ODBCconStr

        Try
            'ConsODBCConStr(NumEmp)

            Clase.ODBCcon_A.ConnectionString = Clase.ODBCcon_A.ConnectionString

            Clase.ODBCcon_A.Open()
            Clase.ODBC_CMD.CommandText = LocalSQL
            Clase.ODBC_DA.SelectCommand = Clase.ODBC_CMD
            Clase.ODBC_DA.Fill(ODBC_DS)

            ' No es necesario para Datasets de ReadOnly
            'ODBC_DA.FillSchema(ODBC_DS, SchemaType.Source)
            'Dim rowsFilled As Long = ODBC_DA.Fill(ODBC_DS)
            Clase.ODBCcon_A.Close()

        Catch ex As Exception
            System.IO.File.WriteAllText("log.txt", ex.Message)
        Finally
            ODBCGetDataset = ODBC_DS.Copy
            With ODBC_DS : .Clear() : .Dispose() : End With
        End Try
    End Function
End Module

