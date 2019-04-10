Imports MySql.Data.MySqlClient

Module MySqlSAC
    Dim MYSQL_DS As New DataSet
    Dim MySVR As String = "192.168.1.17"
    Dim MyID As String = "root"
    Dim MyPass As String = "edifiroot"
    Dim MyDB As String = "sac_nuevo"

    Dim StrCon As String = "server=" & MySVR & ";port=3306;user id=" & MyID & ";password=" & MyPass & ";database=" & MyDB & ";charset=utf8;"

    Dim MySQLCon_A As New MySqlConnection(StrCon)
    Dim MySQLCon_B As New MySqlConnection(StrCon)

    ' Objetos de MYSQL
    Dim MYSQL_DA As New MySqlDataAdapter
    Dim MYSQL_CMD As New MySqlCommand
    Dim MYSQL_DR As MySqlDataReader
    Dim MYSQLTran As MySqlTransaction

    Dim _bolInTransaction As Boolean


    Enum TipoTransaccion
        OpenCon_BeginTrans = 0 ' OpenCon And BEGIN
        ContCon_Transaction = 1 ' Transaccion acumulada, Coninua
        CloseCon_CommitTrans = 2 ' CierraCon and COMMIT
        UniqueTransaction = 3 ' Transaccion unica y completa
    End Enum
    Public Function MYSQlGetDataset(ByVal localSQL As String) As DataSet
        MYSQL_DS = New DataSet
        Try
            MySQLCon_A.Open()
            MYSQL_DA = New MySqlDataAdapter(localSQL, MySQLCon_A)
            MYSQL_DA.Fill(MYSQL_DS)

        Catch ex As Exception


        Finally
            MYSQlGetDataset = MYSQL_DS.Copy
            MySQLCon_A.Close()
            MYSQL_DA.Dispose()
            With MYSQL_DS : .Clear() : .Dispose() : End With
        End Try
    End Function
    Public Sub MySQLBeginTransaction(Optional ByVal OpenConn As Boolean = True)
        If Not _bolInTransaction Then
            If MySQLCon_B.State = ConnectionState.Closed And OpenConn Then MySQLCon_B.Open()

            MYSQLTran = MySQLCon_B.BeginTransaction
            _bolInTransaction = True
        End If
    End Sub
    Public Sub MySQLCommitTransaction(Optional ByVal CloseConn As Boolean = False)
        With MYSQLTran
            If _bolInTransaction Then .Commit() : .Dispose() : _bolInTransaction = False

            If CloseConn Then MySQLCon_B.Close()
        End With
    End Sub

    Public Function MySQLExecSQL(ByVal LocalSQL As String, ByVal TransactionStep As TipoTransaccion) As Boolean
        ' TransactionStep 
        '  0 = OpenCon And BEGIN
        '  1 = Transaccion acumulada, Coninua
        '  2 = CierraCon and COMMIT
        '  3 = Transaccion unica y completa

        Try
            Select Case TransactionStep
                Case TipoTransaccion.OpenCon_BeginTrans, TipoTransaccion.UniqueTransaction '0, 3
                    If _bolInTransaction Then
                        If TransactionStep = TipoTransaccion.UniqueTransaction Then
                            Throw New Exception("No puede activarse una transaccion unica (3) dentro de una transaccion anidada previa (BeginTrans)")
                        End If
                    End If

                    MySQLBeginTransaction()
                    ' Si es una transaccion unica, entonces quitar bandera de transa global
                    If TransactionStep = TipoTransaccion.UniqueTransaction Then _bolInTransaction = False

                    'If MySQLCon_B.State = ConnectionState.Closed Then MySQLCon_B.Open()

                    '' Abrir la transaccion si es necesaria
                    'If Not _bolInTransaction Then MYSQLTran = MySQLCon_B.BeginTransaction
                Case Else ' 1 y 2 Continuacion de execute SQL's, y transaccion abierta

            End Select

            With MYSQL_CMD
                .Connection = MySQLCon_B
                .CommandText = LocalSQL
                .ExecuteNonQuery()
            End With

            ' Cerrar la transaccion si se pide
            Select Case TransactionStep
                Case TipoTransaccion.CloseCon_CommitTrans, TipoTransaccion.UniqueTransaction '2, 3 ' Transaccion unica terminacion obligada
                    ' Si es una transaccion Unica entonces activar bandera de transa global para terminar transaccion
                    If TransactionStep = TipoTransaccion.UniqueTransaction Then _bolInTransaction = True
                    If _bolInTransaction Then MySQLCommitTransaction()

                    MySQLExecSQL = True
                    If MySQLCon_B.State = ConnectionState.Open Then MySQLCon_B.Close()
                    MYSQL_CMD.Dispose()

                Case Else ' Case 1 (Transaccion incompleta y conexion abierta(mas querys))
                    MySQLExecSQL = True
            End Select


        Catch ex As Exception

            MySQLExecSQL = False
            ' Si esta en transaccion global, entonces no cerrar la conexion
            Select Case TransactionStep
                Case TipoTransaccion.OpenCon_BeginTrans, TipoTransaccion.ContCon_Transaction, TipoTransaccion.CloseCon_CommitTrans '0, 1, 2

                Case TipoTransaccion.UniqueTransaction '3
                    If Not _bolInTransaction Then
                        MYSQLTran.Rollback()
                        MySQLCon_B.Close()
                    End If
            End Select
            MYSQL_CMD.Dispose()
        End Try
    End Function
End Module
