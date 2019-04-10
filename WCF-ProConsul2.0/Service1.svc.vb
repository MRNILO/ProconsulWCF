Imports System.Data.Odbc
Imports System.IO
Imports System.ComponentModel
Imports MySql.Data.MySqlClient
Imports System.ServiceModel.Activation

' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de clase "Service1" en el código, en svc y en el archivo de configuración.
' NOTA: para iniciar el Cliente de prueba WCF para probar este servicio, seleccione Service1.svc o Service1.svc.vb en el Explorador de soluciones e inicie la depuración.
<AspNetCompatibilityRequirements(RequirementsMode:=AspNetCompatibilityRequirementsMode.Allowed)>
Public Class Service1
    Implements IService1

    Public cn As OdbcConnection = New OdbcConnection("dsn=EK_ADM11_11;uid=user_read;pwd=lectura;")

    Dim ConexionStr As String = "server=192.168.1.17;port=3306;user id=root;password=edifiroot;database=proconsul;charset=utf8"
    Dim ConexiongEDIFICASA As String = "server=192.168.1.17;port=3306;user id=root;password=edifiroot;database=gedificasa;charset=utf8"
    Dim ConexiongComisiones As String = "server=192.168.1.17;port=3306;user id=root;password=edifiroot;database=comisiones;charset=utf8"


    Dim Conexion As New MySqlConnection(ConexionStr)

    Dim ConexionEnkontrol As New OdbcConnection("dsn=EK_ADM11_11;uid=user_read;pwd=lectura;")
    Dim ConexionEnkontrol_18 As New OdbcConnection("dsn=EK_ADM11_18;uid=user_read;pwd=lectura;")

    Dim ConexionGedificasas As New MySqlConnection(ConexiongEDIFICASA)
    Dim ConexionComisiones As New MySqlConnection(ConexiongComisiones)

    Dim ConexionSAC As New MySqlConnection("server=192.168.1.17;port=3306;user id=root;password=edifiroot;database=sac_nuevo;charset=utf8;")

    Public Sub New()
    End Sub
    Public Sub Registro_Log(ByVal mensaje As String, ByVal Nombre_Funcion As String)
        Dim sw As StreamWriter = File.AppendText("C:\Logs\Log WCF-ProConsul2.0.txt")
        Try
            sw.WriteLine("Mensaje de Error: " + mensaje + " Fecha: " + Now.ToShortDateString + " Hora: " + Now.ToLongTimeString + " Nombre Función: " + Nombre_Funcion)
            sw.Flush()
            sw.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Function ObtenerDSSAC(ByVal SQL As String) As DataSet Implements IService1.ObtenerDSSAC
        Return MySqlSAC.MYSQlGetDataset(SQL)
    End Function
    Function ObtenerTerreno(ByVal CC As String) As Boolean Implements IService1.ObtenerTerreno
        Dim Resultado As Boolean = False
        Dim cmd As New MySqlCommand("SELECT * FROM CCterrenos WHERE CC='" + CC.ToString + "'", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader

        While reader.Read
            Resultado = True
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function ObtenerPCRU(ByVal smza As String) As String Implements IService1.ObtenerPCRU

        Dim DS As New DataSet
        DS = MySqlProConsul2.MYSQlGetDataset("SELECT smza25.id_smza FROM smza25 WHERE smza='" + smza.ToString + "'")

        Try
            If DS.Tables(0).Rows(0).Item(0) > 0 Then
                Return "PCRU25"
            End If
        Catch ex As Exception

        End Try
        DS = MySqlProConsul2.MYSQlGetDataset("SELECT smza27.id_smza FROM smza27 WHERE smza='" + smza.ToString + "'")

        Try
            If DS.Tables(0).Rows(0).Item(0) > 0 Then
                Return "PCRU27"
            End If
        Catch ex As Exception

        End Try

        Return "NO"


    End Function
    Function Obtener_limite_bonoContrato(ByVal CC As String) As Integer Implements IService1.Obtener_limite_bonoContrato
        Dim cmd As New MySqlCommand("SELECT pro_contratos_nuevo.Bono FROM pro_contratos_nuevo WHERE CC='" + CC.ToString + "' ORDER BY Bono ASC LIMIT 1", ConexionGedificasas)
        'cmd.CommandType = CommandType.StoredProcedure
        'cmd.Parameters.AddWithValue("?PNumcte", numcte)
        ConexionGedificasas.Close()
        ConexionGedificasas.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As New Integer
        While reader.Read

            Try
                Aux = reader.Item(0)

            Catch ex As Exception
                Aux = 0
            End Try

        End While
        ConexionGedificasas.Close()
        Return Aux
    End Function
    Function Inserta_reportes_check(ByVal id_categoria As Integer, ByVal id_subcategoria As Integer, ByVal id_subsubcategoria As Integer, ByVal id_subsubsubcategoria As Integer, ByVal id_subsubsubsubcategoria As Integer, ByVal NUMCTE As String, ByVal Observacioens As String, ByVal fotografia As String) As Boolean Implements IService1.Inserta_reportes_check

        Dim cmd As New MySqlCommand("Inserta_reportecheck", ConexionSAC)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_categoria", id_categoria)
        cmd.Parameters.AddWithValue("Pid_subcategoria", id_subcategoria)
        cmd.Parameters.AddWithValue("Pid_subsubcategoria", id_subsubcategoria)
        cmd.Parameters.AddWithValue("Pid_subsubsubcategoria", id_subsubsubcategoria)
        cmd.Parameters.AddWithValue("Pid_subsubsubsubcategoria", id_subsubsubsubcategoria)
        cmd.Parameters.AddWithValue("PNUMCTE", NUMCTE)
        cmd.Parameters.AddWithValue("PObservacioens", Observacioens)
        cmd.Parameters.AddWithValue("Pfotografia", fotografia)
        Conexion.Close()
        Try
            ConexionSAC.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                ConexionSAC.Close()
                Return True
            End If
        Catch ex As Exception
            ConexionSAC.Close()
            Return False
        End Try
        ConexionSAC.Close()
        Return False
    End Function
#Region "Asesores"
#Region "Notificaciones"
    Function Comprobar_Notificaciones(ByVal Empleado As Integer) As CNotificaciones Implements IService1.Comprobar_Notificaciones
        Dim DS As DataSet = MySqlProConsul2.MYSQlGetDataset("Select notificaciones.id_notificacion, notificaciones.Visto, notificaciones.Mensaje, notificaciones.Prioridad, notificaciones.empleado FROM notificaciones WHERE empleado=" + Empleado.ToString + " And Visto=0")
        Dim notificacion As New CNotificaciones
        If DS.Tables(0).Rows.Count > 0 Then
            notificacion.id_notificacion = DS.Tables(0).Rows(0).Item("id_notificacion")
            notificacion.Mensaje = DS.Tables(0).Rows(0).Item("Mensaje")
        Else
            notificacion.Mensaje = "Sin Mensajes"
        End If
        Return notificacion
    End Function
    Function Cambiar_a_Visto_notificacion(ByVal id_notificacion As Integer) As Boolean Implements IService1.Cambiar_a_Visto_notificacion
        If MySqlProConsul2.MySQLExecSQL("UPDATE notificaciones Set Visto=1 WHERE id_notificacion=" + id_notificacion.ToString, MySqlProConsul2.TipoTransaccion.UniqueTransaction) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function Obtener_ultimas_Notificaciones(ByVal Empleado As Integer) As CUltimasNotificaciones() Implements IService1.Obtener_ultimas_Notificaciones

        Dim dt As DataTable = MySqlProConsul2.MYSQlGetDataset("Select notificaciones.Mensaje, notificaciones.Prioridad, notificaciones.empleado, notificaciones.FechaUltima, notificaciones.id_notificacion, notificaciones.Visto FROM notificaciones WHERE empleado=" + Empleado.ToString + " ORDER BY id_notificacion DESC").Tables(0)

        Dim Res(dt.Rows.Count - 1) As CUltimasNotificaciones
        For I = 0 To dt.Rows.Count - 1
            Res(I) = New CUltimasNotificaciones
            Res(I).Mensaje = (dt.Rows(I).Item("Mensaje"))
            Res(I).Prioridad = (dt.Rows(I).Item("Prioridad"))
            Res(I).empleado = (dt.Rows(I).Item("empleado"))
            Res(I).FechaUltima = dt.Rows(I).Item("FechaUltima")
            Res(I).id_notificacion = (dt.Rows(I).Item("id_notificacion"))
            Res(I).Visto = (dt.Rows(I).Item("Visto"))
        Next

        Return Res
    End Function
#End Region
#Region "Account"
    Function LogIn(ByVal Usuario As String, ByVal Password As String) As CUsuario Implements IService1.LogIn
        Dim Respuesta As New CUsuario
        Try
            Respuesta = Obtener_Nombre(Usuario, Password)
            If Respuesta.Nombre_Usuario = "Usuario no encontrado." Then

                Respuesta.Nivel = 0
                Respuesta.Tipo = "-"
                Respuesta.Desc_Valor = "-"
                Return Respuesta
            Else
                Dim Tipo = Obtener_Tipo_Usuario(Usuario)
                Respuesta.Empleado = Usuario
                Respuesta.Desc_Valor = Tipo.Desc_Valor
                Respuesta.Nivel = Tipo.Nivel
                Respuesta.Tipo = Tipo.Tipo
            End If

        Catch ex As Exception
            Registro_Log(ex.Message, "LogIN")
        End Try
        Return Respuesta
    End Function
    Function Obtener_Nombre(ByVal empleado As Integer, ByVal Password As String) As CUsuario
        Dim Resultado As New CUsuario
        Dim mystring As String = "SELECT Nombre=(nom_empleado+' '+ap_paterno_empleado+' '+ap_materno_empleado)  FROM  dba.sm_agente  WHERE  empleado=" + Trim(empleado.ToString) + " and Password=" + Trim(Password.ToString) + " and status_agente='A';  "
        Dim cmd As OdbcCommand = New OdbcCommand(mystring, cn)
        Dim Parametro As New OdbcParameter
        Dim DS As New DataSet
        Dim RD As New OdbcDataAdapter(cmd)

        Try
            cn.Open()
            RD.Fill(DS)
            Resultado.Nombre_Usuario = DS.Tables(0).Rows(0).Item("Nombre").ToString
        Catch ex As Exception
            Registro_Log(ex.Message, "Obtener_nombre")
            Resultado.Nombre_Usuario = "Usuario no encontrado."
            Resultado.Nivel = 0
        End Try


        cn.Close()
        Return Resultado
    End Function
    Function Obtener_Tipo_Usuario(ByVal Empleado As Integer) As CUsuario
        Dim Resultado As New CUsuario
        Dim DS As DataSet

        Try
            DS = MySqlProConsul2.MYSQlGetDataset("SELECT usuarios.Empleado, tipo_usuarios.Descripcion, usuarios.id_tipo, usuarios.nivel, usuarios.Desc_Nombre FROM usuarios INNER JOIN tipo_usuarios ON usuarios.id_tipo = tipo_usuarios.id_tipo WHERE Empleado=" + Trim(Empleado.ToString))
            Resultado.Nivel = DS.Tables(0).Rows(0).Item("nivel")
            Resultado.Tipo = DS.Tables(0).Rows(0).Item("Descripcion")
            Resultado.Desc_Valor = DS.Tables(0).Rows(0).Item("Desc_Nombre")

        Catch ex As Exception
            Resultado.Nivel = 1
            Resultado.Tipo = "Asesor"
            Resultado.Desc_Valor = "-Sin Desc_Valor-"

        End Try
        Return Resultado
    End Function
#End Region
#Region "Comisiones"
    Function Obtener_Estado_de_Cuenta(ByVal Fecha_Inicial As Date, ByVal Fecha_Final As Date, ByVal Empleado As Integer) As List(Of CEstadoCuenta) Implements IService1.Obtener_Estado_de_Cuenta
        Dim Datos As New DataSet
        Dim Resultado As New List(Of CEstadoCuenta)
        Dim Aux As CEstadoCuenta
        Try
            Datos = MySqlComi.MYSQlGetDataset("SELECT comisiones.numcte, comisiones.Fecha_Pago, comisiones.Cantidad_Pagada_Total, tipopago.Descripcion AS `Tipo de Pago` FROM comisiones INNER JOIN tipopago ON comisiones.id_Tipo_Pago = tipopago.id_Tipo_Pago WHERE Empleado=" + Empleado.ToString + " and Fecha_Pago BETWEEN '" + Fecha_Inicial.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "' and Pagado=1 and id_tipo_comision=1 and Fecha_Pago > '" + Now.AddYears(-1).ToString("yyyy/MM/dd") + "';")
        Catch ex As Exception
            Registro_Log(ex.Message, "Obtener_Estado_de_Cuenta")
        End Try
        Try
            For I = 0 To Datos.Tables(0).Rows.Count - 1
                Aux = New CEstadoCuenta
                Aux.Numcte = Datos.Tables(0).Rows(I).Item("numcte")
                Aux.Cantidad_Pagada = Datos.Tables(0).Rows(I).Item("Cantidad_Pagada_Total")
                Aux.Fecha_pago = Datos.Tables(0).Rows(I).Item("Fecha_Pago")
                Aux.TipoPago = Datos.Tables(0).Rows(I).Item("Tipo de Pago")
                Try
                    Aux.NombreCliente = Obtener_Nombre_Cliente(Aux.Numcte)
                Catch exs As Exception
                    Aux.NombreCliente = "Error en nombre."
                    Registro_Log(exs.Message, "Obtener_Estado_de_Cuenta")
                End Try
                Resultado.Add(Aux)
            Next
        Catch ex As Exception
            Registro_Log(ex.Message, "Obtener_Estado_de_Cuenta")
        End Try
        Return Resultado
    End Function
    Function Obtener_reporte_semanal(ByVal empleado As Integer) As List(Of CEstadoCuenta) Implements IService1.Obtener_reporte_semanal
        Dim Resultado As New List(Of CEstadoCuenta)
        'Dim Complementos As Integer = MySqlComi.MYSQlGetDataset("SELECT periodos.Complementos FROM periodos WHERE Activo=1").Tables(0).Rows(0).Item(0)
        Dim id_periodo As Integer = MySqlComi.MYSQlGetDataset("SELECT periodos.id_periodo FROM periodos WHERE Activo=1").Tables(0).Rows(0).Item(0)
        Dim Datos As DataSet
        Dim Res As New CEstadoCuenta

        Datos = MySqlComi.MYSQlGetDataset("SELECT comisiones.Observaciones,comisiones.numcte, comisiones.Fecha_Pago, tipopago.Descripcion, comisiones.Cantidad_Pagada_Total FROM comisiones INNER JOIN tipopago ON comisiones.id_Tipo_Pago = tipopago.id_Tipo_Pago WHERE id_tipo_comision=1 and id_periodo IN (" + (id_periodo - 1).ToString + "," + id_periodo.ToString + ") and Empleado=" + empleado.ToString + " and Cantidad_Pagada_Total!=0")


        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res = New CEstadoCuenta
            Res.Numcte = Datos.Tables(0).Rows(I).Item("Numcte")

            Res.NombreCliente = Obtener_Nombre_Cliente(Res.Numcte)
            Res.Fecha_pago = Datos.Tables(0).Rows(I).Item("Fecha_Pago")
            Res.TipoPago = Datos.Tables(0).Rows(I).Item("Descripcion")
            Res.Observaciones = Datos.Tables(0).Rows(I).Item("Observaciones")
            'Dice Monica que siempre no 25/10/2016
            'If Res.TipoPago Like "PENALIZACION" Then
            '    Res.Cantidad_Pagada = (Datos.Tables(0).Rows(I).Item("Cantidad_Pagada_Total"))
            'Else
            Res.Cantidad_Pagada = (Datos.Tables(0).Rows(I).Item("Cantidad_Pagada_Total") * 0.92)
            'End If
            Resultado.Add(Res)
        Next

        Return Resultado

    End Function
    Function Inserta_comisionesgerencia(ByVal CC As String, ByVal Cantidad As Integer) As Boolean Implements IService1.Inserta_comisionesgerencia

        Dim cmd As New MySqlCommand("InsertaPagoGerencia", ConexionComisiones)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("PCC", CC)
        cmd.Parameters.AddWithValue("PCantidad", Cantidad)
        ConexionComisiones.Close()
        Try
            ConexionComisiones.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                ConexionComisiones.Close()
                Return True
            End If
        Catch ex As Exception
            ConexionComisiones.Close()
            Return False
        End Try
        ConexionComisiones.Close()
        Return False
    End Function
    Function Elimina_comisionesgerencia(ByVal CC As String) As Boolean Implements IService1.Elimina_comisionesgerencia

        Dim cmd As New MySqlCommand("EliminaPagoGerencia", ConexionComisiones)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("PCC", CC)
        ConexionComisiones.Close()
        Try
            ConexionComisiones.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                ConexionComisiones.Close()
                Return True
            End If
        Catch ex As Exception
            ConexionComisiones.Close()
            Return False
        End Try
        ConexionComisiones.Close()
        Return False
    End Function

#End Region
#Region "Contratos"
    Function Obtener_promociones(ByVal CC As String, ByVal SM As String) As List(Of CPromocionesContrato) Implements IService1.Obtener_promociones
        Dim Resultado As New List(Of CPromocionesContrato)
        Dim cmd As New MySqlCommand("SELECT * FROM promociones_contratos WHERE CC='" + CC.ToString + "' and (SM='TODAS' or SM='" + SM.ToString + "')", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CPromocionesContrato
        While reader.Read
            Aux = New CPromocionesContrato
            Aux.id_promocion = DirectCast(reader.Item("id_promocion"), Integer)
            Aux.CC = DirectCast(reader.Item("CC"), String)
            Aux.SM = DirectCast(reader.Item("SM"), String)
            Aux.Costo = reader.Item("Costo")
            Aux.textoCombo = DirectCast(reader.Item("TextoCombo"), String)
            Aux.textContrato = DirectCast(reader.Item("textContrato"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_promocione(ByVal id_promocion As Integer) As CPromocionesContrato Implements IService1.Obtener_promocione

        Dim cmd As New MySqlCommand("SELECT * FROM promociones_contratos WHERE id_promocion=" + id_promocion.ToString + "", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CPromocionesContrato
        While reader.Read
            Aux = New CPromocionesContrato
            Aux.id_promocion = DirectCast(reader.Item("id_promocion"), Integer)
            Aux.CC = DirectCast(reader.Item("CC"), String)
            Aux.SM = DirectCast(reader.Item("SM"), String)
            Aux.Costo = reader.Item("Costo")
            Aux.textoCombo = DirectCast(reader.Item("TextoCombo"), String)
            Aux.textContrato = DirectCast(reader.Item("textContrato"), String)

        End While
        Conexion.Close()
        Return Aux
    End Function
    Function Obtener_Equipamientos(ByVal CC As String, ByVal SM As String) As List(Of CEquipamiento) Implements IService1.Obtener_Equipamientos
        Dim Resultado As New List(Of CEquipamiento)
        Dim cmd As New MySqlCommand("SELECT * FROM promociones WHERE CC='" + CC.ToString + "' and (SM='TODAS' or SM='" + SM.ToString + "')", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CEquipamiento
        While reader.Read
            Aux = New CEquipamiento
            Aux.id_promocion = DirectCast(reader.Item("id_promocion"), Integer)
            Aux.CC = DirectCast(reader.Item("CC"), String)
            Aux.SM = DirectCast(reader.Item("SM"), String)
            Aux.Precio = reader.Item("Precio")
            Aux.TextoCombo = DirectCast(reader.Item("TextoCombo"), String)
            Aux.TextoContrato = DirectCast(reader.Item("TextoContrato"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_Equipamiento(ByVal id_promocion As Integer) As CEquipamiento Implements IService1.Obtener_Equipamiento
        Dim Resultado As New CEquipamiento
        Dim cmd As New MySqlCommand("SELECT * FROM promociones WHERE id_promocion=" + id_promocion.ToString + "", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CEquipamiento
        While reader.Read
            Aux = New CEquipamiento
            Aux.id_promocion = DirectCast(reader.Item("id_promocion"), Integer)
            Aux.CC = DirectCast(reader.Item("CC"), String)
            Aux.SM = DirectCast(reader.Item("SM"), String)
            Aux.Precio = reader.Item("Precio")
            Aux.TextoCombo = DirectCast(reader.Item("TextoCombo"), String)
            Aux.TextoContrato = DirectCast(reader.Item("TextoContrato"), String)
            Resultado = Aux
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_promociones(ByVal CC As String, ByVal SM As String, ByVal Precio As String, ByVal TextoCombo As String, ByVal TextoContrato As String) As Boolean Implements IService1.Inserta_promociones

        Dim cmd As New MySqlCommand("INSERT INTO promociones (CC, SM, Precio, TextoCombo, TextoContrato ) VALUES (@PCC, @PSM, @PPrecio, @PTextoCombo, @PTextoContrato )", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PCC", CC)
        cmd.Parameters.AddWithValue("@PSM", SM)
        cmd.Parameters.AddWithValue("@PPrecio", Precio)
        cmd.Parameters.AddWithValue("@PTextoCombo", TextoCombo)
        cmd.Parameters.AddWithValue("@PTextoContrato", TextoContrato)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_promociones(ByVal id_promocion As Integer, CC As String, ByVal SM As String, ByVal Precio As String, ByVal TextoCombo As String, ByVal TextoContrato As String) As Boolean Implements IService1.Actualiza_promociones

        Dim cmd As New MySqlCommand("Actualiza_Equipamiento", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("PCC", CC)
        cmd.Parameters.AddWithValue("PSM", SM)
        cmd.Parameters.AddWithValue("PPrecio", Precio)
        cmd.Parameters.AddWithValue("PTextoCombo", TextoCombo)
        cmd.Parameters.AddWithValue("PTextoContrato", TextoContrato)
        cmd.Parameters.AddWithValue("Pid_promocion", id_promocion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_promociones(ByVal id_promocion As Integer) As Boolean Implements IService1.Elimina_promociones

        Dim cmd As New MySqlCommand("Elimina_Equipamiento", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_promocion", id_promocion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Listar_Equipamientos() As List(Of CEquipamiento) Implements IService1.Listar_Equipamientos
        Dim Resultado As New List(Of CEquipamiento)
        Dim cmd As New MySqlCommand("Listar_equipamientos", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CEquipamiento
        While reader.Read
            Aux = New CEquipamiento
            Aux.id_promocion = DirectCast(reader.Item("id_promocion"), Integer)
            Aux.CC = DirectCast(reader.Item("CC"), String)
            Aux.SM = DirectCast(reader.Item("SM"), String)
            Aux.Precio = reader.Item("Precio")
            Aux.TextoCombo = DirectCast(reader.Item("TextoCombo"), String)
            Aux.TextoContrato = DirectCast(reader.Item("TextoContrato"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Listar_ContratoDatos() As List(Of CDatosContrato) Implements IService1.Listar_ContratoDatos
        Dim Resultado As New List(Of CDatosContrato)
        Dim cmd As New MySqlCommand("listacontratos", ConexionGedificasas)
        cmd.CommandType = CommandType.StoredProcedure
        ConexionGedificasas.Close()
        ConexionGedificasas.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CDatosContrato
        While reader.Read
            Aux = New CDatosContrato
            Aux.id_contrato = DirectCast(reader.Item("id_contrato"), Integer)
            Aux.CC = DirectCast(reader.Item("CC"), String)
            Aux.SM = DirectCast(reader.Item("SM"), String)
            Aux.TC = DirectCast(reader.Item("TC"), Integer)
            Aux.INFONAVIT = reader.Item("INFONAVIT")
            Aux.FOVISSSTE = reader.Item("FOVISSSTE")
            Aux.ISSEG = reader.Item("ISSEG")
            Aux.Fecha_DTU = DirectCast(reader.Item("Fecha_DTU"), Date)
            Aux.CPenalizaPrevio = DirectCast(reader.Item("CPenalizaPrevio"), Decimal)
            Aux.CEnganche = DirectCast(reader.Item("CEnganche"), Decimal)
            Aux.CPenalizaIngresado = DirectCast(reader.Item("CPenalizaIngresado"), Decimal)
            Aux.FormatoAdicional2 = DirectCast(reader.Item("FormatoAdicional2"), String)
            Aux.FormatoAdicional = DirectCast(reader.Item("FormatoAdicional"), String)
            Aux.PrecioCasa = DirectCast(reader.Item("PrecioCasa"), Integer)
            Aux.PrecioAdicional = DirectCast(reader.Item("PrecioAdicional"), Integer)
            Aux.Mtr_Casa = reader.Item("Mtr_Casa")
            Aux.Activo = reader.Item("Activo")
            Aux.Bono = DirectCast(reader.Item("Bono"), Integer)
            Resultado.Add(Aux)
        End While
        ConexionGedificasas.Close()
        Return Resultado
    End Function
    Function Inserta_pro_contratos_nuevo(ByVal CC As String, ByVal SM As String, ByVal TC As Integer, ByVal INFONAVIT As String, ByVal FOVISSSTE As String, ByVal ISSEG As String, ByVal Fecha_DTU As Date, ByVal CPenalizaPrevio As Decimal, ByVal CEnganche As Decimal, ByVal CPenalizaIngresado As Decimal, ByVal FormatoAdicional2 As String, ByVal FormatoAdicional As String, ByVal PrecioCasa As Integer, ByVal PrecioAdicional As Integer, ByVal Mtr_Casa As String, ByVal Activo As String, ByVal Bono As Integer) As Boolean Implements IService1.Inserta_pro_contratos_nuevo

        Dim cmd As New MySqlCommand("Inserta_Nuevo_DatoContrato", ConexionGedificasas)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("PCC", CC)
        cmd.Parameters.AddWithValue("PSM", SM)
        cmd.Parameters.AddWithValue("PTC", TC)
        cmd.Parameters.AddWithValue("PINFONAVIT", INFONAVIT)
        cmd.Parameters.AddWithValue("PFOVISSSTE", FOVISSSTE)
        cmd.Parameters.AddWithValue("PISSEG", ISSEG)
        cmd.Parameters.AddWithValue("PFecha_DTU", Fecha_DTU)
        cmd.Parameters.AddWithValue("PCPenalizaPrevio", CPenalizaPrevio)
        cmd.Parameters.AddWithValue("PCEnganche", CEnganche)
        cmd.Parameters.AddWithValue("PCPenalizaIngresado", CPenalizaIngresado)
        cmd.Parameters.AddWithValue("PFormatoAdicional2", FormatoAdicional2)
        cmd.Parameters.AddWithValue("PFormatoAdicional", FormatoAdicional)
        cmd.Parameters.AddWithValue("PPrecioCasa", PrecioCasa)
        cmd.Parameters.AddWithValue("PPrecioAdicional", PrecioAdicional)
        cmd.Parameters.AddWithValue("PMtr_Casa", Mtr_Casa)
        cmd.Parameters.AddWithValue("PActivo", Activo)
        cmd.Parameters.AddWithValue("PBono", Bono)
        ConexionGedificasas.Close()
        Try
            ConexionGedificasas.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                ConexionGedificasas.Close()
                Return True
            End If
        Catch ex As Exception
            ConexionGedificasas.Close()
            Return False
        End Try
        ConexionGedificasas.Close()
        Return False
    End Function
    Function Actualiza_pro_contratos_nuevo(ByVal id_contrato As Integer, ByVal CC As String, ByVal SM As String, ByVal TC As Integer, ByVal INFONAVIT As String, ByVal FOVISSSTE As String, ByVal ISSEG As String, ByVal Fecha_DTU As Date, ByVal CPenalizaPrevio As Decimal, ByVal CEnganche As Decimal, ByVal CPenalizaIngresado As Decimal, ByVal FormatoAdicional2 As String, ByVal FormatoAdicional As String, ByVal PrecioCasa As Integer, ByVal PrecioAdicional As Integer, ByVal Mtr_Casa As String, ByVal Activo As String, ByVal Bono As Integer) As Boolean Implements IService1.Actualiza_pro_contratos_nuevo

        Dim cmd As New MySqlCommand("Actualiza_Datoscontrato", ConexionGedificasas)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("PCC", CC)
        cmd.Parameters.AddWithValue("PSM", SM)
        cmd.Parameters.AddWithValue("PTC", TC)
        cmd.Parameters.AddWithValue("PINFONAVIT", INFONAVIT)
        cmd.Parameters.AddWithValue("PFOVISSSTE", FOVISSSTE)
        cmd.Parameters.AddWithValue("PISSEG", ISSEG)
        cmd.Parameters.AddWithValue("PFecha_DTU", Fecha_DTU)
        cmd.Parameters.AddWithValue("PCPenalizaPrevio", CPenalizaPrevio)
        cmd.Parameters.AddWithValue("PCEnganche", CEnganche)
        cmd.Parameters.AddWithValue("PCPenalizaIngresado", CPenalizaIngresado)
        cmd.Parameters.AddWithValue("PFormatoAdicional2", FormatoAdicional2)
        cmd.Parameters.AddWithValue("PFormatoAdicional", FormatoAdicional)
        cmd.Parameters.AddWithValue("PPrecioCasa", PrecioCasa)
        cmd.Parameters.AddWithValue("PPrecioAdicional", PrecioAdicional)
        cmd.Parameters.AddWithValue("PMtr_Casa", Mtr_Casa)
        cmd.Parameters.AddWithValue("PActivo", Activo)
        cmd.Parameters.AddWithValue("PBono", Bono)
        cmd.Parameters.AddWithValue("Pid_contrato", id_contrato)
        ConexionGedificasas.Close()
        Try
            ConexionGedificasas.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                ConexionGedificasas.Close()
                Return True
            End If
        Catch ex As Exception
            ConexionGedificasas.Close()
            Return False
        End Try
        ConexionGedificasas.Close()
        Return False
    End Function
    Function Elimina_pro_contratos_nuevo(ByVal id_contrato As Integer) As Boolean Implements IService1.Elimina_pro_contratos_nuevo

        Dim cmd As New MySqlCommand("Elimina_datoscontrato", ConexionGedificasas)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_contrato", id_contrato)
        ConexionGedificasas.Close()
        Try
            ConexionGedificasas.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                ConexionGedificasas.Close()
                Return True
            End If
        Catch ex As Exception
            ConexionGedificasas.Close()
            Return False
        End Try
        ConexionGedificasas.Close()
        Return False
    End Function
    Function Obtener_plazosTerrenos() As List(Of CPlazosTerreno) Implements IService1.Obtener_plazosTerrenos
        Dim Resultado As New List(Of CPlazosTerreno)
        Dim cmd As New MySqlCommand("SELECT * FROM plazosTerrenos", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CPlazosTerreno
        While reader.Read
            Aux = New CPlazosTerreno
            Aux.id_plazo = DirectCast(reader.Item("id_plazo"), Integer)
            Aux.plazo = DirectCast(reader.Item("plazo"), Integer)
            Aux.precioMetro = DirectCast(reader.Item("precioMetro"), Decimal)
            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function
    Function Obtener_plazoTerreno(ByVal id_plazo As Integer) As CPlazosTerreno Implements IService1.Obtener_plazoTerreno
        Dim Resultado As New CPlazosTerreno
        Dim cmd As New MySqlCommand("SELECT * FROM plazosTerrenos WHERE id_plazo = " + id_plazo.ToString + "", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As New CPlazosTerreno
        While reader.Read

            Aux.id_plazo = DirectCast(reader.Item("id_plazo"), Integer)
            Aux.plazo = DirectCast(reader.Item("plazo"), Integer)
            Aux.precioMetro = DirectCast(reader.Item("precioMetro"), Decimal)
            Resultado = Aux
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_Datos_Contrato_Nuevo(ByVal CC As String, ByVal SM As String, ByVal TC As Integer, ByVal INFONAVIT As Integer, ByVal FOVISSSTE As Integer, ByVal ISSEG As Integer, Optional ByVal Empresa As Integer = 11) As CDatosContratoNuevo Implements IService1.Obtener_Datos_Contrato_Nuevo
        Dim Res As New CDatosContratoNuevo

        'Dim DatosContrato = MYSQLGEDIFI.MYSQlGetDataset("SELECT * FROM pro_contratos_nuevo WHERE CC=" + CC.ToString + " and SM='" + SM + "' and INFONAVIT=" + INFONAVIT.ToString + " and FOVISSSTE=" + FOVISSSTE.ToString + " and ISSEG=""" + ISSEG.ToString + """ and Activo=1 ").Tables(0)
        Dim DatosContrato = MYSQLGEDIFI.MYSQlGetDataset("SELECT * FROM pro_contratos_nuevo WHERE CC=" + CC.ToString + " and SM='" + SM + "' and TC=" + TC.ToString + " and Activo=1 ").Tables(0)
        Dim InfoCredito = MYSQLGEDIFI.MYSQlGetDataset("SELECT * FROM pro_contratos_tc WHERE TC=" + TC.ToString + "").Tables(0)


        Res.TC_Abreviatura = InfoCredito.Rows(0).Item("Abreviatura")
        Res.TC_Nombre = InfoCredito.Rows(0).Item("Nombre_Completo")
        Res.Fecha_DTU = DatosContrato.Rows(0).Item("Fecha_DTU")
        Res.Precio_Adicional = DatosContrato.Rows(0).Item("PrecioAdicional")
        Res.Precio_Casa = DatosContrato.Rows(0).Item("PrecioCasa")
        Res.Pen_Previo = DatosContrato.Rows(0).Item("CPenalizaPrevio")
        Res.Pen_Final = DatosContrato.Rows(0).Item("CPenalizaIngresado")
        Res.Formato_adicional = DatosContrato.Rows(0).Item("FormatoAdicional")
        Res.Formato_adicional2 = DatosContrato.Rows(0).Item("FormatoAdicional2")
        Res.Mtrs_Casa = DatosContrato.Rows(0).Item("Mtr_Casa")
        Dim DatosPrototipo As DataTable
        If ObtenerTerreno(CC) Then
            DatosPrototipo = ODBCGetDataset("SELECT Nom_Tipocasa,M2_Constr FROM dba.sm_prototipo WHERE id_num_tipocasa=" + ODBCGetDataset("SELECT id_num_tipocasa,id_num_smza FROM dba.sm_fraccionamiento_lote WHERE id_cve_fracc='" + CC.ToString + "';", Empresa).Tables(0).Rows(0).Item("id_Num_Tipocasa").ToString + ";", Empresa).Tables(0)
            Res.Mtrs_Construccion = "0.0"
        Else
            Try
                DatosPrototipo = ODBCGetDataset("SELECT Nom_Tipocasa,M2_Constr FROM dba.sm_prototipo WHERE id_num_tipocasa=" + ODBCGetDataset("SELECT id_num_tipocasa,id_num_smza FROM dba.sm_fraccionamiento_lote WHERE id_cve_fracc='" + CC.ToString + "'  and id_num_smza LIKE '%" + SM.ToString + "';", Empresa).Tables(0).Rows(0).Item("id_Num_Tipocasa").ToString + ";", Empresa).Tables(0)
                Res.Mtrs_Construccion = DatosPrototipo.Rows(0).Item("M2_Constr")
            Catch ex As Exception
                Res.Mtrs_Construccion = 0
            End Try

        End If

        Res.Nombre_CC = ODBCGetDataset("SELECT Nom_Fracc FROM dba.sm_fraccionamiento WHERE id_cve_fracc='" + CC.ToString + "';", Empresa).Tables(0).Rows(0).Item(0)

        Res.Modelo_casa = DatosPrototipo.Rows(0).Item("Nom_Tipocasa")
        Res.Cantidad_Enganche = DatosContrato.Rows(0).Item("CEnganche")
        Res.Bono = DatosContrato.Rows(0).Item("Bono")
        Return Res

    End Function
    Function Obtener_Tipos_de_Credito() As CCreditos() Implements IService1.Obtener_Tipos_de_Credito
        Dim DT As New DataTable

        DT = MYSQLGEDIFI.MYSQlGetDataset("SELECT * FROM pro_contratos_tc").Tables(0)

        Dim Res(DT.Rows.Count - 1) As CCreditos
        For I = 0 To DT.Rows.Count - 1
            Res(I) = New CCreditos
            Res(I).TC = (DT.Rows(I).Item("TC"))
            Res(I).Abreviatura = (DT.Rows(I).Item("Abreviatura"))
            Res(I).NombreCompleto = (DT.Rows(I).Item("Nombre_Completo"))
        Next

        Return Res
    End Function

    Function Obtener_Tipos_de_CreditoCC(ByVal CC As String, ByVal SM As String) As CCreditos() Implements IService1.Obtener_Tipos_de_CreditoCC
        Dim DT As New DataTable

        'DT = MYSQLGEDIFI.MYSQlGetDataset("SELECT * FROM pro_contratos_tc").Tables(0)

        DT = MYSQLGEDIFI.MYSQlGetDataset("SELECT DISTINCT TC.TC, TC.Abreviatura, TC.Nombre_Completo
                                          FROM pro_contratos_tc TC
                                          INNER JOIN pro_contratos_nuevo CN ON CN.TC = TC.TC
                                          WHERE CN.CC = '" & CC & "' AND CN.SM = '" & SM & "' AND CN.Activo = 1").Tables(0).Copy

        Dim Res(DT.Rows.Count - 1) As CCreditos
        For I = 0 To DT.Rows.Count - 1
            Res(I) = New CCreditos
            Res(I).TC = (DT.Rows(I).Item("TC"))
            Res(I).Abreviatura = (DT.Rows(I).Item("Abreviatura"))
            Res(I).NombreCompleto = (DT.Rows(I).Item("Nombre_Completo"))
        Next

        Return Res
    End Function
#End Region
#Region "Generales"
    Function Obtener_Credito_Porcentaje(ByVal Numcte As Integer) As String Implements IService1.Obtener_Credito_Porcentaje
        Dim Res As String = ""
        Try
            Select Case ODBC.ODBCGetDataset("SELECT id_ruta From dba.sm_cliente WHERE numcte=" + Numcte.ToString + ";  ", 11).Tables(0).Rows(0).Item(0)
                Case 19, 1, 2, 5, 12
                    'Sin Subsidio 2.5%
                    Res = "2.5%"
                Case 3, 15, 22
                    'Subsidio 2.0%
                    Res = "2.0%"
                Case 11
                    'No infonavit
                    Res = "3.0%"
                Case 9
                    'Contado 2.5%
                    Res = "2.5%"
                Case Else
                    'No infonavit
                    Res = "3.0%"
            End Select
            Res = Res + " Credito:" + ODBC.ODBCGetDataset("SELECT dba.sm_ruta.Nombre FROM dba.sm_cliente,dba.sm_ruta WHERE dba.sm_cliente.id_ruta=dba.sm_ruta.id_ruta AND dba.sm_cliente.numcte=" + Numcte.ToString + ";", 11).Tables(0).Rows(0).Item(0)
        Catch ex As Exception
            Res = "Sin Credito"
        End Try

        Return Res
    End Function
    Function Obtener_Asesores_Activos() As CAsesoresActivos() Implements IService1.Obtener_Asesores_Activos
        Dim Datos As New DataSet
        Dim dt As New DataTable
        Datos = ODBCGetDataset("SELECT Empleado, Nom_Empleado, Ap_Paterno_Empleado, Ap_Materno_Empleado, Direccion_Archivo as Lider FROm dba.sm_agente WHERE  status_agente='A' and Direccion_Archivo!='ADMINISTRATIVO' and Empleado !=9999  ORDER BY Lider;", 11)
        dt = Datos.Tables(0)
        Dim Res(dt.Rows.Count - 1) As CAsesoresActivos
        For I = 0 To dt.Rows.Count - 1
            Res(I) = New CAsesoresActivos
            Try
                Res(I).Empleado = (dt.Rows(I).Item("Empleado"))
            Catch ex0 As Exception
                Res(I).Empleado = 0.0
            End Try
            Try
                Res(I).Nom_Empleado = (dt.Rows(I).Item("Nom_Empleado"))
            Catch ex1 As Exception
                Res(I).Nom_Empleado = ""
            End Try
            Try
                Res(I).Ap_Paterno_Empleado = (dt.Rows(I).Item("Ap_Paterno_Empleado"))
            Catch ex2 As Exception
                Res(I).Ap_Paterno_Empleado = ""
            End Try
            Try
                Res(I).Ap_Materno_Empleado = (dt.Rows(I).Item("Ap_Materno_Empleado"))
            Catch ex3 As Exception
                Res(I).Ap_Materno_Empleado = ""
            End Try
            Try
                Res(I).Lider = (dt.Rows(I).Item("Lider"))
            Catch ex4 As Exception
                Res(I).Lider = ""
            End Try
        Next
        Return Res
    End Function
    Function Obtener_Clientes_Activos(ByVal Empleado As Integer) As CClientesActivos() Implements IService1.Obtener_Clientes_Activos
        Dim Datos As New DataSet
        Dim dt As New DataTable
        Datos = ODBCGetDataset("SELECT 	dba.sm_cliente.numcte, 	NombreCliente = ( 		nom_cte + ' ' + ap_paterno_cte + ' ' + ap_materno_cte 	), 	dba.sm_fraccionamiento_lote.id_num_lote, 	dba.sm_fraccionamiento_lote.id_cve_fracc, 	dba.sm_fraccionamiento.nom_fracc, 	dba.sm_fraccionamiento_lote.id_num_mza, 	dba.sm_fraccionamiento_lote.Dir_casa, 	dba.sm_cliente.id_num_etapa, 	dba.sm_etapa.nom_etapa, 	Valor_credito, 	Valor_Total, dba.sm_fraccionamiento_lote.id_num_interior FROM 	dba.sm_cliente, 	dba.sm_fraccionamiento_lote, 	dba.sm_agente, 	dba.sm_etapa, 	dba.sm_fraccionamiento WHERE 	dba.sm_cliente.lote_id *= dba.sm_fraccionamiento_lote.lote_id AND dba.sm_cliente.id_cve_fracc*=dba.sm_fraccionamiento.id_cve_fracc AND dba.sm_cliente.empleado = dba.sm_agente.empleado AND dba.sm_cliente.id_num_etapa = dba.sm_etapa.id_num_etapa AND dba.sm_cliente.empleado = " + Empleado.ToString + " AND dba.sm_cliente.id_num_etapa < 20 AND dba.sm_cliente.status_cte != 'C';", 11)
        dt = Datos.Tables(0)
        Dim Res(dt.Rows.Count - 1) As CClientesActivos
        For I = 0 To dt.Rows.Count - 1
            Res(I) = New CClientesActivos
            Try
                Res(I).numcte = (dt.Rows(I).Item("numcte"))
            Catch ex0 As Exception
                Res(I).numcte = 0.0
            End Try
            Try
                Res(I).NombreCliente = (dt.Rows(I).Item("NombreCliente"))
            Catch ex1 As Exception
                Res(I).NombreCliente = ""
            End Try
            Try
                Res(I).lote_id = (dt.Rows(I).Item("id_num_lote"))
            Catch ex2 As Exception
                Res(I).lote_id = 0.0
            End Try
            Try
                Res(I).id_num_mza = (dt.Rows(I).Item("id_num_mza"))
            Catch ex3 As Exception
                Res(I).id_num_mza = ""
            End Try
            Try
                Res(I).CC = (dt.Rows(I).Item("id_cve_fracc"))
            Catch ex3 As Exception
                Res(I).id_num_mza = ""
            End Try
            Try
                Res(I).DirCasa = (dt.Rows(I).Item("Dir_Casa"))
            Catch ex3 As Exception
                Res(I).id_num_mza = ""
            End Try
            Try
                Res(I).Fracc = (dt.Rows(I).Item("Nom_Fracc"))
            Catch ex3 As Exception
                Res(I).id_num_mza = ""
            End Try
            Try
                Res(I).id_num_etapa = (dt.Rows(I).Item("id_num_etapa"))
            Catch ex4 As Exception
                Res(I).id_num_etapa = 0.0
            End Try
            Try
                Res(I).nom_etapa = (dt.Rows(I).Item("nom_etapa"))
            Catch ex5 As Exception
                Res(I).nom_etapa = ""
            End Try
            Try
                Res(I).Valor_credito = (dt.Rows(I).Item("Valor_credito"))
            Catch ex6 As Exception
                Res(I).Valor_credito = 0.0
            End Try
            Try
                Res(I).Valor_Total = (dt.Rows(I).Item("Valor_Total"))
            Catch ex7 As Exception
                Res(I).Valor_Total = 0.0
            End Try
            Try
                Res(I).NumeroOficial = (dt.Rows(I).Item("id_num_interior"))
            Catch ex5 As Exception
                Res(I).NumeroOficial = ""
            End Try
        Next
        Return Res
    End Function
    Function Obtener_cuentaDepodito_Cte(ByVal Numcte As Integer) As Integer Implements IService1.Obtener_cuentaDepodito_Cte
        Return ODBCGetDataset("SELECT dba.sm_cliente_adicionales.Cuenta_deposito FROM dba.sm_cliente_adicionales WHERE numcte=" + Numcte.ToString + ";", 11).Tables(0).Rows(0).Item(0)
    End Function
    Function Valida_Modificacion(ByVal Empleado As Integer) As Boolean
        Dim DS As New DataSet
        DS = MySqlProConsul2.MYSQlGetDataset("SELECT empleado_modificadatos.empleado, empleado_modificadatos.modifico FROM empleado_modificadatos WHERE empleado=" + Empleado.ToString + ";")
        If DS.Tables(0).Rows.Count > 0 Then
            If DS.Tables(0).Rows(0).Item("modifico") > 0 Then
                Return True
            End If

        End If
        Return False
    End Function
    Function MyDatos(empleado As Integer, FNacimiento As Date, SSexo As String, SNacionalidad As String, Tnacionalidad As String,
                    SCivil As String, tbCiudad As String, tbEstado As String, tbDir As String, tbCelular As String,
                    tbTel As String, tbEmail As String, tbRefNom1 As String, tbRefParentesco1 As String,
                    tbRefTel1 As String, tbRefNom2 As String, tbRefParentesco2 As String, tbRefTel2 As String,
                    tbRFC As String, tbCURP As String, tbIFE As String, tbCIFE As String, tbManejo As String,
                    FVenceManejo As Date, tbNSS As String, cbPrimaria As String, cbSecundaria As String,
                    cbPreparatoria As String, cbLicenciatura As String, tbLicName As String, cbMaestria As String,
                    tbMasName As String) As Boolean Implements IService1.MyDatos
        Dim Res As Boolean

        Dim Primaria = 0
        Dim Secundaria = 0
        Dim Preparatoria = 0
        Dim Licenciatura = 0
        Dim Maestria = 0

        If cbPrimaria = "True" Then
            Primaria = 1
        End If

        If cbSecundaria = "True" Then
            Secundaria = 1
        End If

        If cbPreparatoria = "True" Then
            Preparatoria = 1
        End If
        If cbLicenciatura = "True" Then
            Licenciatura = 1
        End If
        If cbMaestria = "True" Then
            Maestria = 1
        End If


        If SNacionalidad = "Mexicana" Then
        Else
            SNacionalidad = Tnacionalidad
        End If

        If Verfica_datos_empleado(empleado) Then
            'Update
            If Valida_Modificacion(empleado) Then

            Else
                Res = MySqlProConsul2.MySQLExecSQL("UPDATE empleado_datos SET empleado_datos.fecha_nacimiento='" + FNacimiento.ToString("yyyy/MM/dd") + "'," +
" empleado_datos.sexo='" + SSexo.ToString + "'," +
" empleado_datos.nacionalidad='" + SNacionalidad.ToString + "'," +
" empleado_datos.estado_civil='" + SCivil.ToString + "'," +
" empleado_datos.ciudad='" + tbCiudad.ToString + "'," +
" empleado_datos.estado='" + tbEstado.ToString + "'," +
" empleado_datos.domicilio='" + tbDir.ToString + "'," +
" empleado_datos.celular='" + tbCelular.ToString + "'," +
" empleado_datos.telfijo='" + tbTel.ToString + "'," +
" empleado_datos.email='" + tbEmail.ToString + "'," +
" empleado_datos.ref1nombre='" + tbRefNom1.ToString + "'," +
" empleado_datos.red1parentesco='" + tbRefParentesco1.ToString + "'," +
" empleado_datos.red1tel='" + tbRefTel1.ToString + "'," +
" empleado_datos.red2nombre='" + tbRefNom2.ToString + "'," +
" empleado_datos.red2parentesco='" + tbRefParentesco2.ToString + "'," +
" empleado_datos.red2tel='" + tbRefTel2.ToString + "'," +
" empleado_datos.rfc='" + tbRFC.ToString + "'," +
" empleado_datos.curp='" + tbCURP.ToString + "'," +
" empleado_datos.nife='" + tbIFE.ToString + "'," +
" empleado_datos.claveelector='" + tbCIFE.ToString + "'," +
" empleado_datos.licmanejo='" + tbManejo.ToString + "'," +
" empleado_datos.fechavence='" + FVenceManejo.ToString("yyyy/MM/dd") + "'," +
" empleado_datos.nss='" + tbNSS.ToString + "'," +
" empleado_datos.primaria='" + Primaria.ToString + "'," +
" empleado_datos.secundaria='" + Secundaria.ToString + "'," +
" empleado_datos.preparatoria='" + Preparatoria.ToString + "'," +
" empleado_datos.licenciatura='" + Licenciatura.ToString + "'," +
" empleado_datos.nomlicenciatura='" + tbLicName.ToString + "'," +
" empleado_datos.maestria='" + Maestria.ToString + "'," +
" empleado_datos.nommaestria='" + tbMasName.ToString + "'" +
         " WHERE" +
" empleado_datos.empleado=" + empleado.ToString + "", MySqlProConsul2.TipoTransaccion.UniqueTransaction)
                MySqlProConsul2.MySQLExecSQL("INSERT INTO empleado_modificaDatos (empleado,modifico) VALUES (" + empleado.ToString + ",1);", MySqlProConsul2.TipoTransaccion.UniqueTransaction)
            End If

        Else
            'Insert
            Res = MySqlProConsul2.MySQLExecSQL("INSERT INTO empleado_datos ( empleado_datos.empleado,  empleado_datos.fecha_nacimiento,  empleado_datos.sexo,  empleado_datos.nacionalidad,  empleado_datos.estado_civil,  empleado_datos.ciudad,  empleado_datos.estado,  empleado_datos.domicilio,  empleado_datos.celular,  empleado_datos.telfijo,  empleado_datos.email,  empleado_datos.ref1nombre,  empleado_datos.red1parentesco,  empleado_datos.red1tel,  empleado_datos.red2nombre,  empleado_datos.red2parentesco,  empleado_datos.red2tel,  empleado_datos.rfc,  empleado_datos.curp,  empleado_datos.nife,  empleado_datos.claveelector,  empleado_datos.licmanejo,  empleado_datos.fechavence,  empleado_datos.nss,  empleado_datos.primaria,  empleado_datos.secundaria,  empleado_datos.preparatoria, empleado_datos.licenciatura,  empleado_datos.nomlicenciatura,  empleado_datos.maestria, empleado_datos.nommaestria) VALUES ( " + empleado.ToString + ",'" + FNacimiento.ToString("yyyy/MM/dd") + "','" + SSexo + "','" + SNacionalidad + "','" + SCivil + "','" + tbCiudad + "','" + tbEstado + "','" + tbDir + "','" + tbCelular + "','" + tbTel + "','" + tbEmail + "','" + tbRefNom1 + "','" + tbRefParentesco1 + "','" + tbRefTel1 + "','" + tbRefNom2 + "','" + tbRefParentesco2 + "','" + tbRefTel2 + "','" + tbRFC + "','" + tbCURP + "','" + tbIFE.ToString + "','" + tbCIFE.ToString + "','" + tbManejo + "','" + FVenceManejo.ToString("yyyy/MM/dd") + "','" + tbNSS + "'," + Primaria.ToString + "," + Secundaria.ToString + "," + Preparatoria.ToString + "," + Licenciatura.ToString + ",'" + tbLicName + "'," + Maestria.ToString + ",'" + tbMasName + "')  ", MySqlProConsul2.TipoTransaccion.UniqueTransaction)
        End If

        Return Res
    End Function
    Function Obtener_DatosDetalle_Empleado(ByVal Empleado As Integer) As CDatosAsesorDetalle Implements IService1.Obtener_DatosDetalle_Empleado

        Dim dt As New DataTable

        dt = MySqlProConsul2.MYSQlGetDataset("SELECT empleado_datos.id_datos,empleado_datos.empleado, empleado_datos.fecha_nacimiento, empleado_datos.sexo, empleado_datos.nacionalidad, empleado_datos.estado_civil, empleado_datos.ciudad, empleado_datos.estado, empleado_datos.domicilio, empleado_datos.celular, empleado_datos.telfijo, empleado_datos.email, empleado_datos.ref1nombre, empleado_datos.red1parentesco, empleado_datos.red1tel, empleado_datos.red2nombre, empleado_datos.red2parentesco, empleado_datos.red2tel, empleado_datos.rfc, empleado_datos.curp, empleado_datos.nife, empleado_datos.claveelector, empleado_datos.licmanejo, empleado_datos.fechavence, empleado_datos.nss, empleado_datos.primaria, empleado_datos.secundaria, empleado_datos.preparatoria, empleado_datos.licenciatura, empleado_datos.nomlicenciatura, empleado_datos.maestria, empleado_datos.nommaestria FROM empleado_datos WHERE empleado=" + Empleado.ToString + "").Tables(0)
        Dim Res As New CDatosAsesorDetalle
        If dt.Rows.Count > 0 Then
            Res.id_datos = (dt.Rows(0).Item("id_datos"))
            Res.empleado = (dt.Rows(0).Item("empleado"))
            Res.fecha_nacimiento = dt.Rows(0).Item("fecha_nacimiento")
            Res.sexo = (dt.Rows(0).Item("sexo"))
            Res.nacionalidad = (dt.Rows(0).Item("nacionalidad"))
            Res.estado_civil = (dt.Rows(0).Item("estado_civil"))
            Res.ciudad = (dt.Rows(0).Item("ciudad"))
            Res.estado = (dt.Rows(0).Item("estado"))
            Res.domicilio = (dt.Rows(0).Item("domicilio"))
            Res.celular = (dt.Rows(0).Item("celular"))
            Res.telfijo = (dt.Rows(0).Item("telfijo"))
            Res.email = (dt.Rows(0).Item("email"))
            Res.ref1nombre = (dt.Rows(0).Item("ref1nombre"))
            Res.red1parentesco = (dt.Rows(0).Item("red1parentesco"))
            Res.red1tel = (dt.Rows(0).Item("red1tel"))
            Res.red2nombre = (dt.Rows(0).Item("red2nombre"))
            Res.red2parentesco = (dt.Rows(0).Item("red2parentesco"))
            Res.red2tel = (dt.Rows(0).Item("red2tel"))
            Res.rfc = (dt.Rows(0).Item("rfc"))
            Res.curp = (dt.Rows(0).Item("curp"))
            Res.nife = (dt.Rows(0).Item("nife"))
            Res.claveelector = (dt.Rows(0).Item("claveelector"))
            Res.licmanejo = (dt.Rows(0).Item("licmanejo"))
            Res.fechavence = dt.Rows(0).Item("fechavence")
            Res.nss = (dt.Rows(0).Item("nss"))
            Res.primaria = (dt.Rows(0).Item("primaria"))
            Res.secundaria = (dt.Rows(0).Item("secundaria"))
            Res.preparatoria = (dt.Rows(0).Item("preparatoria"))
            Res.licenciatura = (dt.Rows(0).Item("licenciatura"))
            Res.nomlicenciatura = (dt.Rows(0).Item("nomlicenciatura"))
            Res.maestria = (dt.Rows(0).Item("maestria"))
            Res.nommaestria = (dt.Rows(0).Item("nommaestria"))
        End If



        Return Res
    End Function
    Function Verfica_datos_empleado(ByVal Empleado As Integer) As Boolean
        Dim DT = MySqlProConsul2.MYSQlGetDataset("SELECT id_datos FROM empleado_datos WHERE empleado=" + Empleado.ToString).Tables(0)
        If DT.Rows.Count > 0 Then
            If DT.Rows(0).Item(0) > 0 Then
                Return True
            End If
        End If
        Return False
    End Function
    Function Verifica_Conectividad() As Boolean Implements IService1.Verifica_Conectividad
        Return True
    End Function
    Function Obtener_Nombre_Cliente(ByVal Numcte As Integer) As String Implements IService1.Obtener_Nombre_Cliente
        Dim Resultado As String = ""
        Try
            Resultado = ODBCGetDataset("SELECT nombre=(nom_cte+' '+ap_paterno_cte+' '+ap_materno_cte) FROM dba.sm_cliente WHERE numcte=" + Numcte.ToString + ";", 11).Tables(0).Rows(0).Item("Nombre").ToString
        Catch ex As Exception
            Registro_Log(ex.Message, "Obtener_nombre_del_cliente")
        End Try
        Return Verifica_String(Resultado)
    End Function
    Function Verifica_String(ByVal Texto As String) As String
        Texto = Texto.Replace("╤", "Ñ")
        Texto = Texto.Replace("Θ", "é")
        Return Texto
    End Function
    Function Obtener_Nombre_Asesor(ByVal Empleado As Integer) As String
        Dim Resultado As String = ""
        Try
            Resultado = ODBCGetDataset("SELECT NombreEmpleado=(ap_paterno_empleado+' '+ap_materno_empleado+' '+nom_empleado) FROM dba.sm_agente WHERE empleado=" + Empleado.ToString + ";", 11).Tables(0).Rows(0).Item("NombreEmpleado").ToString
        Catch ex As Exception
            Registro_Log(ex.Message, "Obtener_Nombre_Asesor")
        End Try
        Return Resultado
    End Function
    Function Obtener_datos_nom_fracc() As CDatosFracc() Implements IService1.Obtener_datos_nom_fracc
        Dim Datos As New DataSet
        Dim DTA As New DataTable
        Dim DTB As New DataTable
        Dim DTC As New DataTable
        Dim DTR As New DataTable

        Try
            DTA = ODBC.ODBCGetDataset("SELECT id_cve_fracc,(id_cve_fracc+' '+Nom_Fracc)as Fraccionamiento FROM dba.sm_fraccionamiento WHERE Status_Fracc='A' ORDER BY Fraccionamiento;", 11).Tables(0).Copy
            DTB = ODBC.ODBCGetDataset("SELECT id_cve_fracc,(id_cve_fracc+' '+Nom_Fracc)as Fraccionamiento FROM dba.sm_fraccionamiento WHERE Status_Fracc='A' ORDER BY Fraccionamiento;", 18).Tables(0).Copy
            DTC = MYSQLGEDIFI.MYSQlGetDataset("SELECT DISTINCT(CC) FROM pro_contratos_nuevo ORDER BY CC").Tables(0).Copy

            Dim rowA As DataRow
            For Each rowB As DataRow In DTB.Rows
                rowA = DTA.NewRow
                rowA("id_cve_fracc") = rowB("id_cve_fracc")
                rowA("Fraccionamiento") = rowB("Fraccionamiento")
                DTA.Rows.Add(rowA)
            Next

            Dim Index As Integer = 0
            Dim rowR As DataRow
            Dim RowResult() As DataRow

            DTR.Columns.AddRange({New DataColumn("id_cve_fracc", GetType(String)), New DataColumn("Fraccionamiento", GetType(String))})

            For Each rowC As DataRow In DTC.Rows
                RowResult = DTA.Select("id_cve_fracc = " & rowC("CC"))

                For Each row As DataRow In RowResult
                    Index = DTA.Rows.IndexOf(row)

                    rowR = DTR.NewRow()
                    rowR("id_cve_fracc") = DTA.Rows(Index).Item("id_cve_fracc")
                    rowR("Fraccionamiento") = DTA.Rows(Index).Item("Fraccionamiento")

                    DTR.Rows.Add(rowR)
                Next
            Next

            RowResult = Nothing

            Datos.Tables.Add(DTR)
        Catch ex As Exception
            Registro_Log(ex.Message, "Obtener_datos_nom_fracc")
        End Try

        Dim Res(Datos.Tables(0).Rows.Count - 1) As CDatosFracc
        For I = 0 To Res.Count - 1
            Res(I) = New CDatosFracc
            Res(I).id_cve_fracc = Datos.Tables(0).Rows(I).Item("id_cve_fracc")
            Res(I).Nom_Fracc = Datos.Tables(0).Rows(I).Item("Fraccionamiento")
        Next
        Return Res
    End Function
    Function Obtener_Smza(ByVal CC As String) As List(Of String) Implements IService1.Obtener_Smza
        Dim Datos As New DataSet
        Try
            'Datos = ODBC.ODBCGetDataset("SELECT DISTINCT(id_num_Smza) FROM dba.sm_fraccionamiento_lote WHERE id_cve_fracc='" + CC.ToString + "';  ", 11)
            Datos = MYSQLGEDIFI.MYSQlGetDataset("SELECT DISTINCT pro_contratos_nuevo.SM FROM pro_contratos_nuevo WHERE CC='" + CC.ToString + "' AND Activo=1")

        Catch ex As Exception
            Registro_Log(ex.Message, "Obtener_Smza")
        End Try

        Dim Res As New List(Of String)
        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res.Add(Datos.Tables(0).Rows(I).Item(0))
        Next
        Return Res
    End Function
#End Region
#End Region
#Region "Dashboard"


    Public Function Obtener_Ventas_Por_Semana_Entre_Fechas(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As CDatosVentaSemanal() Implements IService1.Obtener_Ventas_Por_Semana_Entre_Fechas
        Dim Datos As New DataSet
        Dim Cantidad As Integer = 0
        Try
            Datos = ODBCGetDataset("SELECT datepart(week,fec_registo) as NSemana,COUNT(numcte) as Cantidad FROm dba.sm_cliente WHERE lote_id>0 and Status_Cte!='C' and id_num_etapa>6 and fec_registo BETWEEN '" + Fecha_inicio.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "' GROUP BY Nsemana ORDER BY NSemana ASC;", 11)
        Catch ex As Exception
            Return Nothing
        End Try
        Dim Res(Datos.Tables(0).Rows.Count - 1) As CDatosVentaSemanal
        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res(I) = New CDatosVentaSemanal
            Res(I).NSemana = DirectCast(Datos.Tables(0).Rows(I).Item("NSemana"), Integer)


            Cantidad += DirectCast(Datos.Tables(0).Rows(I).Item("Cantidad"), Integer)
            Res(I).CantidadVentas = Cantidad
            'If I = 0 Then
            '    Res(I).CantidadVentas = DirectCast(Datos.Tables(0).Rows(I).Item("Cantidad"), Integer)
            'Else
            '    Res(I).CantidadVentas = (DirectCast(Datos.Tables(0).Rows(I).Item("Cantidad"), Integer)) + Res(I - 1).CantidadVentas
            'End If
        Next
        Return Res
    End Function
    Public Function Obtener_Ventas_Por_Semana_Entre_Fechas_Barras(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As CDatosVentaSemanal() Implements IService1.Obtener_Ventas_Por_Semana_Entre_Fechas_Barras
        Dim Datos As New DataSet
        Try
            Datos = ODBCGetDataset("SELECT datepart(week,fec_registo) as NSemana,COUNT(numcte) as Cantidad FROm dba.sm_cliente WHERE lote_id>0 and Status_Cte!='C' and id_num_etapa>6 and fec_registo BETWEEN '" + Fecha_inicio.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "' GROUP BY Nsemana;", 11)
        Catch ex As Exception
            Return Nothing
        End Try
        Dim Res(Datos.Tables(0).Rows.Count - 1) As CDatosVentaSemanal
        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res(I) = New CDatosVentaSemanal
            Res(I).NSemana = DirectCast(Datos.Tables(0).Rows(I).Item("NSemana"), Integer)
            Res(I).CantidadVentas = DirectCast(Datos.Tables(0).Rows(I).Item("Cantidad"), Integer)
        Next
        Return Res
    End Function
    Function Obtener_Total_Vendido_Ubicado(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As Integer Implements IService1.Obtener_Total_Vendido_Ubicado
        Dim Datos As New DataSet
        Try
            Datos = ODBCGetDataset("SELECT COUNT(Numcte) As Cantidad FROm dba.sm_cliente WHERE fec_registo BETWEEN '" + Fecha_inicio.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "' and id_num_etapa>6 and Status_cte!='C' and lote_id>0;", 11)
        Catch ex As Exception
            Return Nothing
        End Try
        If Datos.Tables(0).Rows.Count > 0 Then
            Return DirectCast(Datos.Tables(0).Rows(0).Item(0), Integer)
        End If
        Return Nothing
    End Function
    Function Obtener_Total_habitabilidad(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As Integer Implements IService1.Obtener_Total_habitabilidad
        Dim Datos As New DataSet
        Try
            Datos = ODBCGetDataset("SELECT COUNT(lote_id) As Cantidad FROM dba.sm_fraccionamiento_lote WHERE Fec_Habitabilidad BETWEEN '" + Fecha_inicio.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "';", 11)
        Catch ex As Exception
            Return Nothing
        End Try
        If Datos.Tables(0).Rows.Count > 0 Then
            Return DirectCast(Datos.Tables(0).Rows(0).Item(0), Integer)
        End If
        Return Nothing
    End Function
    Public Function Obtener_Habitabilidad_Por_Semana_Entre_Fechas(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As CDatosVentaSemanal() Implements IService1.Obtener_Habitabilidad_Por_Semana_Entre_Fechas
        Dim Datos As New DataSet
        Try
            Datos = ODBCGetDataset("SELECT datepart(week,fec_habitabilidad) as NSemana,COUNT(lote_id) As Cantidad FROM dba.sm_fraccionamiento_lote WHERE Fec_Habitabilidad BETWEEN '" + Fecha_inicio.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "' GROUP BY NSemana ORDER BY NSemana;", 11)
        Catch ex As Exception
            Return Nothing
        End Try
        Dim Res(Datos.Tables(0).Rows.Count - 1) As CDatosVentaSemanal
        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res(I) = New CDatosVentaSemanal
            Res(I).NSemana = DirectCast(Datos.Tables(0).Rows(I).Item("NSemana"), Integer)
            If I = 0 Then
                Res(I).CantidadVentas = DirectCast(Datos.Tables(0).Rows(I).Item("Cantidad"), Integer)
            Else
                Res(I).CantidadVentas = (DirectCast(Datos.Tables(0).Rows(I).Item("Cantidad"), Integer)) + Res(I - 1).CantidadVentas
            End If

        Next
        Return Res
    End Function
    Public Function Obtener_Firmas_X_Semana_Entre_Fechas(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As CDatosVentaSemanal() Implements IService1.Obtener_Firmas_X_Semana_Entre_Fechas
        Dim Datos As New DataSet
        Try
            Datos = ODBCGetDataset("SELECT datepart(week,fec_liberacion) as NSemana, COUNT(numcte) as Cantidad FROM dba.sm_cliente_etapa WHERE id_num_etapa=18 and fec_liberacion BETWEEN '" + Fecha_inicio.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "' GROUP BY NSemana ORDER BY NSemana;", 11)
        Catch ex As Exception
            Return Nothing
        End Try
        Dim Res(Datos.Tables(0).Rows.Count - 1) As CDatosVentaSemanal
        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res(I) = New CDatosVentaSemanal
            Res(I).NSemana = DirectCast(Datos.Tables(0).Rows(I).Item("NSemana"), Integer)
            'Para Acumulados
            'If I = 0 Then
            '    Res(I).CantidadVentas = DirectCast(Datos.Tables(0).Rows(I).Item("Cantidad"), Integer)
            'Else
            '    Res(I).CantidadVentas = (DirectCast(Datos.Tables(0).Rows(I).Item("Cantidad"), Integer)) + Res(I - 1).CantidadVentas
            'End If
            Res(I).CantidadVentas = DirectCast(Datos.Tables(0).Rows(I).Item("Cantidad"), Integer)
        Next
        Return Res
    End Function
    Function Obtener_Total_Cancelados_o_SinUbicacion(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As Integer Implements IService1.Obtener_Total_Cancelados_o_SinUbicacion
        Dim Datos As New DataSet
        Try
            Datos = ODBCGetDataset("SELECT COUNT(numcte) as Cantidad FROm dba.sm_cliente WHERE (lote_id is Null or Status_cte='C') and fec_registo BETWEEN '" + Fecha_inicio.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "';", 11)
        Catch ex As Exception
            Return Nothing
        End Try
        If Datos.Tables(0).Rows.Count > 0 Then
            Return DirectCast(Datos.Tables(0).Rows(0).Item(0), Integer)
        End If
        Return Nothing
    End Function
#End Region
#Region "Adminsitrativo"
    Function Obtener_comisiones_cliente(ByVal numcte As String) As List(Of CComisionesCliente) Implements IService1.Obtener_comisiones_cliente
        Dim Resultado As New List(Of CComisionesCliente)
        Dim cmd As New MySqlCommand("SELECT comisiones.id_comision, tipo_comsion.Descripcion AS Pagado_A, comisiones.id_periodo, comisiones.Fecha_Pago, comisiones.Cantidad_Pagada_Total, tipopago.Descripcion AS Tipo_pago, comisiones.Empleado, comisiones.Lider, comisiones.Gerente, admin.NombreCompleto, comisiones.Observaciones FROM comisiones INNER JOIN tipo_comsion ON comisiones.id_tipo_comision = tipo_comsion.id_tipo_comision INNER JOIN tipopago ON comisiones.id_Tipo_Pago = tipopago.id_Tipo_Pago INNER JOIN admin ON comisiones.Adm = admin.id_admin WHERE numcte=?PNumcte ORDER BY Pagado_A", ConexionComisiones)
        'cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("?PNumcte", numcte)
        ConexionComisiones.Close()
        ConexionComisiones.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CComisionesCliente
        While reader.Read
            Aux = New CComisionesCliente
            Aux.id_comision = DirectCast(reader.Item("id_comision"), Integer)
            Aux.Pagado_A = DirectCast(reader.Item("Pagado_A"), String)
            Aux.id_periodo = DirectCast(reader.Item("id_periodo"), Integer)
            Aux.Fecha_Pago = DirectCast(reader.Item("Fecha_Pago"), Date)
            Aux.Cantidad_Pagada_Total = DirectCast(reader.Item("Cantidad_Pagada_Total"), Integer)
            Aux.Tipo_pago = DirectCast(reader.Item("Tipo_pago"), String)
            Aux.Empleado = DirectCast(reader.Item("Empleado"), Integer)
            Aux.Lider = DirectCast(reader.Item("Lider"), String)
            Aux.Gerente = DirectCast(reader.Item("Gerente"), String)
            Aux.NombreCompleto = DirectCast(reader.Item("NombreCompleto"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Resultado.Add(Aux)
        End While
        ConexionComisiones.Close()
        Return Resultado
    End Function
    Function Obtener_pro_contratos_nuevo(ByVal id_contrato As Integer) As CContratos Implements IService1.Obtener_pro_contratos_nuevo

        Dim cmd As New MySqlCommand("SELECT * FROM pro_contratos_nuevo WHERE id_contrato=" + id_contrato.ToString, ConexionGedificasas)
        'cmd.CommandType = CommandType.StoredProcedure
        ConexionGedificasas.Close()
        ConexionGedificasas.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As New CContratos
        While reader.Read

            Aux.id_contrato = DirectCast(reader.Item("id_contrato"), Integer)
            Aux.CC = DirectCast(reader.Item("CC"), String)
            Aux.SM = DirectCast(reader.Item("SM"), String)
            Aux.TC = DirectCast(reader.Item("TC"), Integer)
            Aux.INFONAVIT = reader.Item("INFONAVIT")
            Aux.FOVISSSTE = reader.Item("FOVISSSTE")
            Aux.ISSEG = reader.Item("ISSEG")
            Aux.Fecha_DTU = DirectCast(reader.Item("Fecha_DTU"), Date)
            Aux.CPenalizaPrevio = DirectCast(reader.Item("CPenalizaPrevio"), Decimal)
            Aux.CEnganche = DirectCast(reader.Item("CEnganche"), Decimal)
            Aux.CPenalizaIngresado = DirectCast(reader.Item("CPenalizaIngresado"), Decimal)
            Aux.FormatoAdicional2 = DirectCast(reader.Item("FormatoAdicional2"), String)
            Aux.FormatoAdicional = DirectCast(reader.Item("FormatoAdicional"), String)
            Aux.PrecioCasa = DirectCast(reader.Item("PrecioCasa"), Integer)
            Aux.PrecioAdicional = DirectCast(reader.Item("PrecioAdicional"), Integer)
            Aux.Mtr_Casa = reader.Item("Mtr_Casa")
            Aux.Activo = reader.Item("Activo")
            Aux.Bono = DirectCast(reader.Item("Bono"), Integer)

        End While
        ConexionGedificasas.Close()
        Return Aux
    End Function
    Function CancelacionCliente(ByVal Numcte As Integer) As CDatosCancelacion Implements IService1.CancelacionCliente
        Dim Resultado As New CDatosCancelacion
        Dim DatosEnk As New DataSet
        Dim AnticipoCliente As Integer = 0

        Dim IDULTIMO As Integer = 0



        Try
            DatosEnk = ODBCGetDataset("SELECT  dba.sm_cliente.NUmcte, NombreCliente=(nom_cte+' '+ap_paterno_cte+' '+ap_materno_cte), dba.sm_cliente.EmPleado, NombreEmpleado=(nom_empleado+' '+ap_paterno_empleado+' '+ap_materno_empleado), dba.sm_fraccionamiento.nom_fracc, dba.sm_cliente.id_cve_fracc FROM dba.sm_cliente, dba.sm_fraccionamiento,dba.sm_agente WHERE dba.sm_cliente.empleado=dba.sm_agente.empleado AND dba.sm_cliente.id_cve_fracc=dba.sm_fraccionamiento.id_cve_fracc AND dba.sm_cliente.numcte=" + Numcte.ToString + ";")
        Catch ex As Exception

        End Try
        Resultado.Numcte = DatosEnk.Tables(0).Rows(0).Item("NUmcte")
        Resultado.NombreCliente = DatosEnk.Tables(0).Rows(0).Item("NombreCliente")
        Resultado.Empleado = DatosEnk.Tables(0).Rows(0).Item("EmPleado")
        Resultado.NombreEmpleado = DatosEnk.Tables(0).Rows(0).Item("NombreEmpleado")
        Resultado.Frente = DatosEnk.Tables(0).Rows(0).Item("nom_fracc")
        Resultado.CC = DatosEnk.Tables(0).Rows(0).Item("id_cve_fracc")
        AnticipoCliente = Obtener_anticipo_de_venta_cliente(Resultado.Numcte)
        Dim PagadoAsesor As Integer = 0
        Dim PagadoLider As Integer = 0
        Dim PagadoGerente As Integer = 0
        Try
            PagadoAsesor = MySqlComi.MYSQlGetDataset("SELECT SUM(comisiones.Cantidad_Pagada_Total) as Cantidad FROM comisiones WHERE numcte=" + Resultado.Numcte.ToString + " and id_tipo_comision=1 and Pagado=1 and Fecha_Pago>'2008/01/01'").Tables(0).Rows(0).Item(0)
        Catch ex As Exception
            PagadoAsesor = 0
        End Try
        Try
            PagadoLider = MySqlComi.MYSQlGetDataset("SELECT SUM(comisiones.Cantidad_Pagada_Total) as Cantidad FROM comisiones WHERE numcte=" + Resultado.Numcte.ToString + " and id_tipo_comision=2 and Pagado=1 and Fecha_Pago>'2008/01/01'").Tables(0).Rows(0).Item(0)
        Catch ex As Exception
            PagadoLider = 0
        End Try
        Try
            PagadoGerente = MySqlComi.MYSQlGetDataset("SELECT SUM(comisiones.Cantidad_Pagada_Total) as Cantidad FROM comisiones WHERE numcte=" + Resultado.Numcte.ToString + " and id_tipo_comision=3 and Pagado=1 and Fecha_Pago>'2008/01/01'").Tables(0).Rows(0).Item(0)
        Catch ex As Exception
            PagadoGerente = 0
        End Try

        Resultado.CantidadEmpleado = PagadoAsesor
        Resultado.CantidadLider = PagadoLider
        Resultado.CantidadGerente1 = (PagadoGerente / 2)
        Resultado.CantidadGerente2 = (PagadoGerente / 2)



        Resultado.CantidadDeDevolucion = (((AnticipoCliente - PagadoAsesor) - PagadoGerente) - PagadoLider)

        If Resultado.CantidadDeDevolucion < 0 Then
            'Faltan dineros

            Resultado.CantidadDeDevolucion = Resultado.CantidadDeDevolucion + PagadoGerente
            Resultado.PenalizacionGerente1 = (PagadoGerente / 2)
            Resultado.PenalizacionGerente2 = (PagadoGerente / 2)

            If Resultado.CantidadDeDevolucion < 0 Then
                'Falta por recuperar
                Resultado.CantidadDeDevolucion = Resultado.CantidadDeDevolucion + PagadoLider
                Resultado.PenalizacionLider = PagadoLider

                If Resultado.CantidadDeDevolucion < 0 Then
                    'Faltan Dineros
                    Resultado.CantidadDeDevolucion = Resultado.CantidadDeDevolucion + PagadoAsesor
                    Resultado.PenalizacionAsesor = PagadoAsesor - Resultado.CantidadDeDevolucion

                End If

            End If



            Resultado.CantidadDeDevolucion = 0
        End If


        Try
            Resultado.FechaPago = MySqlComi.MYSQlGetDataset("SELECT MAX(comisiones.Fecha_Pago) FROM comisiones WHERE numcte=" + Resultado.Numcte.ToString + "").Tables(0).Rows(0).Item(0)
        Catch ex As Exception
            Resultado.FechaPago = New Date
        End Try


        Resultado.NombreGerente1 = "BELEM OLVERA"
        Resultado.NombreGerente2 = "JAVIER FRANCO"


        Resultado.PenalizacionCliente = If(AnticipoCliente > (PagadoAsesor + PagadoGerente + PagadoLider), (PagadoAsesor + PagadoGerente + PagadoLider), AnticipoCliente)

        If Resultado.PenalizacionCliente > 0 Then
            Resultado.Penaliza = True
        Else
            Resultado.Penaliza = False
        End If

        Try
            MySqlProConsul2.MySQLExecSQL("INSERT INTO com_cancelaciones (numcte, P_Cliente, P_Asesor, P_Lider, P_Gerente, P_Administrativo, FechaCancelacion ) VALUES(" + Numcte.ToString + "," + Resultado.PenalizacionCliente.ToString + "," + Resultado.PenalizacionAsesor.ToString + "," + Resultado.PenalizacionLider.ToString + "," + (Resultado.CantidadGerente1 + Resultado.CantidadGerente2).ToString + ",0,'" + Now.ToString("yyyy/MM/dd") + "');", MySqlProConsul2.TipoTransaccion.UniqueTransaction)
        Catch ex As Exception

        End Try
        Try

            IDULTIMO = MySqlProConsul2.MYSQlGetDataset("SELECT MAX(com_cancelaciones.id_cancelacion) as ID FROM com_cancelaciones").Tables(0).Rows(0).Item(0)
        Catch ex As Exception

        End Try

        Resultado.Folio = IDULTIMO


        Return Resultado
    End Function

    Function Obtener_anticipo_de_venta_cliente(ByVal numcte As Integer) As Integer
        Dim Datos = ODBCGetDataset("SELECT SUM(MONTO*-1)as Cantidad FROM sx_movcltes WHERE TM IN (55,56,57) and CONCEPTO LIKE 'an%' and numcte=" + numcte.ToString + "", 11)
        Dim Cantidad As Integer
        Try
            Cantidad = Datos.Tables(0).Rows(0).Item(0)
        Catch ex As Exception
            Cantidad = 0
        End Try
        Return Cantidad
    End Function
    Function Obtener_Desgloce_Etapas_Cliente(ByVal numcte As Integer) As CEtapasCliente() Implements IService1.Obtener_Desgloce_Etapas_Cliente
        Dim DS As New DataSet

        Try
            DS = ODBCGetDataset("SELECT Nom_etapa, dba.sm_cliente_etapa.id_num_etapa, dba.sm_cliente_etapa.Fec_inicio, dba.sm_cliente_etapa.Fec_Liberacion, dba.sm_cliente_etapa.Observaciones FROM dba.sm_cliente_etapa,dba.sm_etapa WHERE  dba.sm_cliente_etapa.id_num_etapa=dba.sm_etapa.id_num_etapa and dba.sm_cliente_etapa.numcte=" + numcte.ToString + " ORDER BY 2;")

        Catch ex As Exception

        End Try
        Dim Resultado(DS.Tables(0).Rows.Count - 1) As CEtapasCliente
        If DS.Tables(0).Rows.Count > 0 Then
            For I = 0 To DS.Tables(0).Rows.Count - 1
                Resultado(I) = New CEtapasCliente
                Resultado(I).id_num_etapa = DS.Tables(0).Rows(I).Item("id_num_etapa")
                Resultado(I).Nom_etapa = DS.Tables(0).Rows(I).Item("Nom_etapa")
                Resultado(I).Fec_inicio = DS.Tables(0).Rows(I).Item("Fec_inicio")
                Try
                    Resultado(I).Fec_Liberacion = DS.Tables(0).Rows(I).Item("Fec_Liberacion")
                Catch ex As Exception

                End Try
                Try
                    Resultado(I).Observaciones = DS.Tables(0).Rows(I).Item("Observaciones")
                Catch ex As Exception
                    Resultado(I).Observaciones = ""
                End Try


            Next
        End If

        Return Resultado
    End Function
    Function Obtener_Datos_Generales_Cliente(ByVal numcte As Integer) As CGeneralesCliente Implements IService1.Obtener_Datos_Generales_Cliente
        Dim DS As New DataSet
        Dim Resultado As New CGeneralesCliente
        Try
            DS = ODBCGetDataset("SELECT dba.sm_cliente.numcte,dba.sm_fraccionamiento_lote.id_cve_fracc, NombreCliente=(nom_cte+' '+ap_paterno_cte+' '+ap_materno_cte), dba.sm_agente.empleado, NombreEmpleado=(nom_empleado+' '+ap_paterno_empleado+' '+ap_materno_empleado), Direccion_archivo As LIDER, dba.sm_cliente.id_num_etapa As EtapaActual, dba.sm_cliente.status_cte, dba.sm_cliente.lote_id, dba.sm_fraccionamiento_lote.id_num_mza,dba.sm_fraccionamiento_lote.id_num_lote,id_num_interior,dba.sm_fraccionamiento_lote.dir_casa, Valor_credito, Valor_total FROM dba.sm_cliente, dba.sm_fraccionamiento_lote, dba.sm_agente, WHERE  dba.sm_cliente.lote_id*=dba.sm_fraccionamiento_lote.lote_id and dba.sm_cliente.empleado=dba.sm_agente.empleado and dba.sm_cliente.numcte=" + numcte.ToString + ";")
        Catch ex As Exception

        End Try
        Try
            If DS.Tables(0).Rows.Count > 0 Then
                Resultado.numcte = DS.Tables(0).Rows(0).Item("numcte")
                Resultado.NombreCliente = DS.Tables(0).Rows(0).Item("NombreCliente")
                Resultado.empleado = DS.Tables(0).Rows(0).Item("empleado")
                Resultado.NombreEmpleado = DS.Tables(0).Rows(0).Item("NombreEmpleado")
                Resultado.LIDER = DS.Tables(0).Rows(0).Item("LIDER")
                Resultado.EtapaActual = DS.Tables(0).Rows(0).Item("EtapaActual")
                Resultado.status_cte = DS.Tables(0).Rows(0).Item("status_cte")
                Try
                    Resultado.lote_id = DS.Tables(0).Rows(0).Item("lote_id")
                Catch ex As Exception
                    Resultado.lote_id = 0
                End Try
                Try
                    Resultado.id_num_lote = DS.Tables(0).Rows(0).Item("id_num_lote")
                Catch ex As Exception
                    Resultado.id_num_lote = 0
                End Try
                Try
                    Resultado.id_num_interior = DS.Tables(0).Rows(0).Item("id_num_interior")
                Catch ex As Exception
                    Resultado.id_num_interior = 0
                End Try
                Try
                    Resultado.dir_casa = DS.Tables(0).Rows(0).Item("dir_casa")
                Catch ex As Exception
                    Resultado.dir_casa = "-Sin Ubicación-"
                End Try
                Try
                    Resultado.id_cve_fracc = DS.Tables(0).Rows(0).Item("id_cve_fracc")
                Catch ex As Exception
                    Resultado.id_cve_fracc = "-"
                End Try
                Try
                    Resultado.id_num_mza = DS.Tables(0).Rows(0).Item("id_num_mza")
                Catch ex As Exception
                    Resultado.id_num_mza = 0
                End Try
                Try
                    Resultado.Valor_credito = DS.Tables(0).Rows(0).Item("Valor_credito")
                Catch ex As Exception
                    Resultado.Valor_credito = 0
                End Try
                Try
                    Resultado.Valor_total = DS.Tables(0).Rows(0).Item("Valor_total")
                Catch ex As Exception
                    Resultado.Valor_total = 0
                End Try
            End If
        Catch ex As Exception

        End Try
        Return Resultado
    End Function
    Function buscarClienteAutoCompletar(ByVal Busqueda As String) As CClientesBusqueda() Implements IService1.buscarClienteAutoCompletar
        Dim DS As New DataSet
        Try
            DS = ODBCGetDataset("SELECT TOP 5 NombreCliente=(nom_cte+' '+ap_paterno_cte+' '+ap_materno_cte),Numcte FROM dba.sm_cliente WHERE NombreCliente LIKE '%" + Busqueda + "%' ORDER BY Numcte DESC;")
        Catch ex As Exception

        End Try
        Try
            If DS.Tables(0).Rows.Count > 0 Then
                Dim Resultado(DS.Tables(0).Rows.Count - 1) As CClientesBusqueda
                For I = 0 To Resultado.Count - 1
                    Resultado(I) = New CClientesBusqueda
                    Resultado(I).NombreCliente = DS.Tables(0).Rows(I).Item("NombreCliente")
                    Resultado(I).Numcte = DS.Tables(0).Rows(I).Item("Numcte")
                Next
                Return Resultado
            End If
        Catch ex As Exception

        End Try
        Return Nothing
    End Function
    Function Obtener_Reporte_concepto(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date, ByVal Etapa As Integer) As CreporteConepto() Implements IService1.Obtener_Reporte_concepto
        Dim Datos As New DataSet
        Dim dt As New DataTable
        Datos = ODBCGetDataset("SELECT dba.sm_cliente.numcte, NombreCliente=(nom_cte+' '+ap_paterno_cte+' '+ap_materno_cte), dba.sm_cliente.id_cve_fracc, dba.sm_fraccionamiento_lote.id_num_smza, dba.sm_fraccionamiento_lote.id_num_mza, dba.sm_fraccionamiento_lote.id_num_lote, dba.sm_fraccionamiento_lote.Cant_mts_excedente,Concepto, Monto FROM dba.sm_cliente, dba.sm_fraccionamiento_lote, dba.sm_cliente_etapa,sx_movcltes WHERE dba.sm_cliente.numcte=sx_movcltes.numcte AND dba.sm_cliente.lote_id=dba.sm_fraccionamiento_lote.lote_id AND dba.sm_cliente.numcte=dba.sm_cliente_etapa.numcte AND dba.sm_cliente_etapa.id_num_etapa=" + Etapa.ToString + " AND dba.sm_cliente_etapa.fec_liberacion BETWEEN '" + Fecha_inicio.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "';", 11)
        dt = Datos.Tables(0)
        Dim Res(dt.Rows.Count - 1) As CreporteConepto
        For I = 0 To dt.Rows.Count - 1
            Res(I) = New CreporteConepto
            Try
                Res(I).numcte = (dt.Rows(I).Item("numcte"))
            Catch ex0 As Exception
                Res(I).numcte = 0.0
            End Try
            Try
                Res(I).NombreCliente = (dt.Rows(I).Item("NombreCliente"))
            Catch ex1 As Exception
                Res(I).NombreCliente = ""
            End Try
            Try
                Res(I).id_cve_fracc = (dt.Rows(I).Item("id_cve_fracc"))
            Catch ex2 As Exception
                Res(I).id_cve_fracc = ""
            End Try
            Try
                Res(I).id_num_smza = (dt.Rows(I).Item("id_num_smza"))
            Catch ex3 As Exception
                Res(I).id_num_smza = ""
            End Try
            Try
                Res(I).id_num_mza = (dt.Rows(I).Item("id_num_mza"))
            Catch ex4 As Exception
                Res(I).id_num_mza = ""
            End Try
            Try
                Res(I).id_num_lote = (dt.Rows(I).Item("id_num_lote"))
            Catch ex5 As Exception
                Res(I).id_num_lote = ""
            End Try
            Try
                Res(I).Cant_mts_excedente = (dt.Rows(I).Item("Cant_mts_excedente"))
            Catch ex6 As Exception
                Res(I).Cant_mts_excedente = 0.0
            End Try
            Try
                Res(I).Concepto = (dt.Rows(I).Item("Concepto"))
            Catch ex7 As Exception
                Res(I).Concepto = ""
            End Try
            Try
                Res(I).Monto = (dt.Rows(I).Item("Monto"))
            Catch ex8 As Exception
                Res(I).Monto = 0.0
            End Try
        Next
        Return Res
    End Function
    Public Function Reporte_Concepto_Total(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date, ByVal Etapa As Integer) As List(Of CReporteMontos) Implements IService1.Reporte_Concepto_Total
        Dim DSClientes = ODBCGetDataset("SELECT dba.sm_cliente.numcte, NombreCliente=(nom_cte+' '+ap_paterno_cte+' '+ap_materno_cte), dba.sm_cliente.id_cve_fracc, dba.sm_fraccionamiento_lote.id_num_smza, dba.sm_fraccionamiento_lote.id_num_mza, dba.sm_fraccionamiento_lote.id_num_lote, dba.sm_fraccionamiento_lote.Cant_mts_excedente FROM dba.sm_cliente, dba.sm_fraccionamiento_lote, dba.sm_cliente_etapa WHERE dba.sm_cliente.lote_id=dba.sm_fraccionamiento_lote.lote_id AND dba.sm_cliente.numcte=dba.sm_cliente_etapa.numcte AND dba.sm_cliente_etapa.id_num_etapa=" + Etapa.ToString + " AND dba.sm_cliente_etapa.fec_liberacion BETWEEN '" + Fecha_inicio.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "';")
        Dim Resultado As New List(Of CReporteMontos)
        Dim aux As CReporteMontos
        For I = 0 To DSClientes.Tables(0).Rows.Count - 1
            aux = New CReporteMontos
            aux.Numcte = DSClientes.Tables(0).Rows(I).Item("numcte")
            aux.NombreCliente = DSClientes.Tables(0).Rows(I).Item("NombreCliente")
            aux.CC = DSClientes.Tables(0).Rows(I).Item("id_cve_fracc")
            aux.SM = DSClientes.Tables(0).Rows(I).Item("id_num_smza")
            aux.Mza = DSClientes.Tables(0).Rows(I).Item("id_num_mza")
            aux.Lote = DSClientes.Tables(0).Rows(I).Item("id_num_lote")
            Try
                aux.TerrenoExcedente = DSClientes.Tables(0).Rows(I).Item("Cant_mts_excedente")
            Catch ex As Exception
                aux.TerrenoExcedente = 0
            End Try

            Dim DSMov = ODBCGetDataset("SELECT Concepto, Monto FROM sx_movcltes WHERE  numcte=" + aux.Numcte.ToString + ";")
            Dim listamov As New List(Of MovimientosCliente)
            Dim mov As MovimientosCliente

            For J = 0 To DSMov.Tables(0).Rows.Count - 1
                mov = New MovimientosCliente
                mov.Concepto = DSMov.Tables(0).Rows(J).Item("Concepto")
                mov.Monto = DSMov.Tables(0).Rows(J).Item("Monto")
                listamov.Add(mov)
            Next

            aux.Movimientos = listamov

            Resultado.Add(aux)
        Next

        Return Resultado
    End Function
    Public Function ConvertToDataTable(Of T)(data As IList(Of T)) As DataTable
        Dim properties As PropertyDescriptorCollection = TypeDescriptor.GetProperties(GetType(T))
        Dim table As New DataTable()
        For Each prop As PropertyDescriptor In properties
            If prop.Name = "ExtensionData" Then
            Else
                table.Columns.Add(prop.Name, If(Nullable.GetUnderlyingType(prop.PropertyType), prop.PropertyType))
            End If
        Next
        For Each item As T In data
            Dim row As DataRow = table.NewRow()
            For Each prop As PropertyDescriptor In properties
                If prop.Name = "ExtensionData" Then
                Else
                    row(prop.Name) = If(prop.GetValue(item), DBNull.Value)
                End If
            Next
            table.Rows.Add(row)
        Next
        Return table

    End Function
#End Region
#Region "BI"
    Function Obtener_Ingresados_por_semana(ByVal Fecha_Inicio As Date, ByVal Fecha_Final As Date) As List(Of CGraficoBI) Implements IService1.Obtener_Ingresados_por_semana
        'SELECT datepart(week,fec_liberacion) as NSemana,COUNT(DISTINCT(dba.sm_cliente_etapa.numcte)) as Cantidad FROM dba.sm_cliente_etapa WHERE id_num_etapa IN (11,13) and fec_liberacion BETWEEN '2014/01/01' and '2014/06/12' GROUP BY NSemana;
        Dim Resultado As New List(Of CGraficoBI)
        Dim Datos As DataSet
        Dim Res As New CGraficoBI
        Dim ingresados As Integer = 0
        Datos = ODBCGetDataset("SELECT datepart(week,fec_liberacion) as NSemana,COUNT(DISTINCT(dba.sm_cliente_etapa.numcte)) as Cantidad FROM dba.sm_cliente_etapa WHERE id_num_etapa IN (11,13) and fec_liberacion BETWEEN '" + Fecha_Inicio.ToString("yyyy/MM/dd") + "' and '" + Fecha_Final.ToString("yyyy/MM/dd") + "' GROUP BY NSemana ORDER BY NSemana;")

        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res = New CGraficoBI
            Res.NSemana = Datos.Tables(0).Rows(I).Item("NSemana")
            ingresados += Datos.Tables(0).Rows(I).Item("Cantidad")
            Res.CantidadVentas = ingresados
            Resultado.Add(Res)
        Next

        Return Resultado
    End Function
    Public Function Obtener_VentasLM_BI(ByVal Fecha_inicio As Date) As List(Of CGraficoBI) Implements IService1.Obtener_VentasLM_BI
        Dim Resultado As New List(Of CGraficoBI)
        Dim Datos As DataSet
        Dim Res As New CGraficoBI

        Datos = MySqlProConsul2.MYSQlGetDataset("SELECT historicolm.NSemana, historicolm.f_inicio, historicolm.f_final, historicolm.Cantidad_Ventas, historicolm.Cantidad_Escriturados, historicolm.Cantidad_Cancelados, historicolm.Cantidad_Habitabilidad FROM historicolm WHERE  historicolm.f_inicio>= '" + Fecha_inicio.ToString("yyyy/MM/dd") + "'")

        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res = New CGraficoBI
            Res.NSemana = Datos.Tables(0).Rows(I).Item("NSemana")
            Res.CantidadVentas = Datos.Tables(0).Rows(I).Item("Cantidad_Ventas")
            Resultado.Add(Res)
        Next

        Return Resultado
    End Function
    Public Function Obtener_EscrituradosLM_BI(ByVal Fecha_inicio As Date) As List(Of CGraficoBI) Implements IService1.Obtener_EscrituradosLM_BI
        Dim Resultado As New List(Of CGraficoBI)
        Dim Datos As DataSet
        Dim Res As New CGraficoBI

        Datos = MySqlProConsul2.MYSQlGetDataset("SELECT historicolm.NSemana, historicolm.f_inicio, historicolm.f_final, historicolm.Cantidad_Ventas, historicolm.Cantidad_Escriturados, historicolm.Cantidad_Cancelados, historicolm.Cantidad_Habitabilidad FROM historicolm WHERE  historicolm.f_inicio>= '" + Fecha_inicio.ToString("yyyy/MM/dd") + "'")

        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res = New CGraficoBI
            Res.NSemana = Datos.Tables(0).Rows(I).Item("NSemana")
            Res.CantidadVentas = Datos.Tables(0).Rows(I).Item("Cantidad_Escriturados")
            Resultado.Add(Res)
        Next

        Return Resultado
    End Function
    Public Function Obtener_EscrituradosAcumuladoLM_BI(ByVal Fecha_inicio As Date) As List(Of CGraficoBI) Implements IService1.Obtener_EscrituradosAcumuladoLM_BI
        Dim Resultado As New List(Of CGraficoBI)
        Dim Datos As DataSet
        Dim Res As New CGraficoBI
        Dim Escriturados As Integer = 0

        Datos = MySqlProConsul2.MYSQlGetDataset("SELECT historicolm.NSemana, historicolm.f_inicio, historicolm.f_final, historicolm.Cantidad_Ventas, historicolm.Cantidad_Escriturados, historicolm.Cantidad_Cancelados, historicolm.Cantidad_Habitabilidad FROM historicolm WHERE  historicolm.f_inicio>= '" + Fecha_inicio.ToString("yyyy/MM/dd") + "'")

        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res = New CGraficoBI
            Res.NSemana = Datos.Tables(0).Rows(I).Item("NSemana")
            Escriturados += Datos.Tables(0).Rows(I).Item("Cantidad_Escriturados")
            Res.CantidadVentas = Escriturados
            Resultado.Add(Res)
        Next

        Return Resultado
    End Function
    Public Function Obtener_CanceladosLM_BI(ByVal Fecha_inicio As Date) As List(Of CGraficoBI) Implements IService1.Obtener_CanceladosLM_BI
        Dim Resultado As New List(Of CGraficoBI)
        Dim Datos As DataSet
        Dim Res As New CGraficoBI

        Datos = MySqlProConsul2.MYSQlGetDataset("SELECT historicolm.NSemana, historicolm.f_inicio, historicolm.f_final, historicolm.Cantidad_Ventas, historicolm.Cantidad_Escriturados, historicolm.Cantidad_Cancelados, historicolm.Cantidad_Habitabilidad FROM historicolm WHERE  historicolm.f_inicio>= '" + Fecha_inicio.ToString("yyyy/MM/dd") + "'")

        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res = New CGraficoBI
            Res.NSemana = Datos.Tables(0).Rows(I).Item("NSemana")
            Res.CantidadVentas = Datos.Tables(0).Rows(I).Item("Cantidad_Cancelados")
            Resultado.Add(Res)
        Next

        Return Resultado
    End Function
    Public Function Obtener_HabitabilidadLM_BI(ByVal Fecha_inicio As Date) As List(Of CGraficoBI) Implements IService1.Obtener_HabitabilidadLM_BI
        Dim Resultado As New List(Of CGraficoBI)
        Dim Datos As DataSet
        Dim Res As New CGraficoBI

        Datos = MySqlProConsul2.MYSQlGetDataset("SELECT historicolm.NSemana, historicolm.f_inicio, historicolm.f_final, historicolm.Cantidad_Ventas, historicolm.Cantidad_Escriturados, historicolm.Cantidad_Cancelados, historicolm.Cantidad_Habitabilidad FROM historicolm WHERE  historicolm.f_inicio>= '" + Fecha_inicio.ToString("yyyy/MM/dd") + "'")

        For I = 0 To Datos.Tables(0).Rows.Count - 1
            Res = New CGraficoBI
            Res.NSemana = Datos.Tables(0).Rows(I).Item("NSemana")
            Res.CantidadVentas = Datos.Tables(0).Rows(I).Item("Cantidad_Habitabilidad")
            Resultado.Add(Res)
        Next

        Return Resultado
    End Function
#End Region
#Region "PruebaABC"
    Function Inserta_usuarios(ByVal id_tipo As Integer, ByVal nivel As Integer, ByVal Desc_Nombre As String) As Boolean Implements IService1.Inserta_usuarios

        Dim cmd As New MySqlCommand("Inserta_Usuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_tipo", id_tipo)
        cmd.Parameters.AddWithValue("Pnivel", nivel)
        cmd.Parameters.AddWithValue("PDesc_Nombre", Desc_Nombre)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function


#End Region
#Region "ReportesComisiones"
    Function Obtener_Datos_Grid_Gerardo() As List(Of CReporteGerardoGrid) Implements IService1.Obtener_Datos_Grid_Gerardo
        Dim Resultado As New List(Of CReporteGerardoGrid)
        Dim aux As New CReporteGerardoGrid
        Dim ComisionesPeriodos As DataSet = MySqlComi.MYSQlGetDataset("CALL ComisionesVigentes")
        Dim DatosFaltantes As New DataSet


        For I = 0 To ComisionesPeriodos.Tables(0).Rows.Count - 1
            aux = New CReporteGerardoGrid
            aux.id_comision = ComisionesPeriodos.Tables(0).Rows(I).Item("id_comision")
            aux.Numcte = ComisionesPeriodos.Tables(0).Rows(I).Item("numcte")
            aux.Importe = ComisionesPeriodos.Tables(0).Rows(I).Item("Cantidad_Pagada_Total")
            aux.TipoPago = ComisionesPeriodos.Tables(0).Rows(I).Item("TipoPago")
            aux.Observacion = ComisionesPeriodos.Tables(0).Rows(I).Item("observaciones")


            If I = 0 Then
                DatosFaltantes = ODBCGetDataset("SELECT 	NombreCliente = ( ap_paterno_cte + ' ' + ap_materno_cte+' '+nom_cte ), 	NombreEmpleado = ( 		nom_empleado + ' ' + ap_paterno_empleado + ' ' + ap_materno_empleado 	),dba.sm_fraccionamiento_lote.id_num_smza, 	Nom_Fracc, 	Direccion = ( 		dba.sm_fraccionamiento_lote.Dir_Casa + ' ' + id_num_interior 	), 	dba.sm_fraccionamiento_lote.id_num_mza, 	dba.sm_fraccionamiento_lote.id_num_lote FROM 	dba.sm_cliente, 	dba.sm_fraccionamiento_lote, 	dba.sm_fraccionamiento, 	dba.sm_agente WHERE 	dba.sm_cliente.id_cve_fracc = dba.sm_fraccionamiento.id_cve_fracc AND dba.sm_cliente.lote_id *= dba.sm_fraccionamiento_lote.lote_id AND dba.sm_cliente.empleado = dba.sm_agente.empleado AND dba.sm_cliente.numcte = " + aux.Numcte.ToString + ";  ")
            Else
                If ComisionesPeriodos.Tables(0).Rows(I).Item("numcte") = ComisionesPeriodos.Tables(0).Rows(I - 1).Item("numcte") Then
                Else
                    DatosFaltantes = ODBCGetDataset("SELECT 	NombreCliente = ( ap_paterno_cte + ' ' + ap_materno_cte+' '+nom_cte ), 	NombreEmpleado = ( 		nom_empleado + ' ' + ap_paterno_empleado + ' ' + ap_materno_empleado 	),dba.sm_fraccionamiento_lote.id_num_smza, 	Nom_Fracc, 	Direccion = ( 		dba.sm_fraccionamiento_lote.Dir_Casa + ' ' + id_num_interior 	), 	dba.sm_fraccionamiento_lote.id_num_mza, 	dba.sm_fraccionamiento_lote.id_num_lote FROM 	dba.sm_cliente, 	dba.sm_fraccionamiento_lote, 	dba.sm_fraccionamiento, 	dba.sm_agente WHERE 	dba.sm_cliente.id_cve_fracc = dba.sm_fraccionamiento.id_cve_fracc AND dba.sm_cliente.lote_id *= dba.sm_fraccionamiento_lote.lote_id AND dba.sm_cliente.empleado = dba.sm_agente.empleado AND dba.sm_cliente.numcte = " + aux.Numcte.ToString + ";  ")
                End If
            End If


            aux.NombreCliente = DatosFaltantes.Tables(0).Rows(0).Item("NombreCliente")

            Select Case ComisionesPeriodos.Tables(0).Rows(I).Item("TipoComision")
                Case 1
                    aux.NombreAgente = DatosFaltantes.Tables(0).Rows(0).Item("NombreEmpleado")
                Case 2
                    Try
                        'aux.NombreAgente = Obtener_Nombrecompleto_lider(ComisionesPeriodos.Tables(0).Rows(I).Item("Lider")).NombreCompleto
                        aux.NombreAgente = Obtener_Nombre_Asesor(ComisionesPeriodos.Tables(0).Rows(I).Item("Lider"))
                    Catch ex As Exception
                        aux.NombreAgente = ComisionesPeriodos.Tables(0).Rows(I).Item("Lider")
                    End Try

                Case 3
                    aux.NombreAgente = ComisionesPeriodos.Tables(0).Rows(I).Item("Gerente")
                Case 4
                    aux.NombreAgente = ComisionesPeriodos.Tables(0).Rows(I).Item("NombreCompleto")
            End Select


            Try
                aux.SMZA = DatosFaltantes.Tables(0).Rows(0).Item("id_num_smza")
            Catch ex As Exception
                aux.SMZA = "-"
            End Try

            aux.Fracc = DatosFaltantes.Tables(0).Rows(0).Item("Nom_Fracc")

            Try
                aux.Domicilio = DatosFaltantes.Tables(0).Rows(0).Item("Direccion")
            Catch ex As Exception
                aux.Domicilio = "- Sin Lote-"
            End Try

            Try
                aux.Manzana = DatosFaltantes.Tables(0).Rows(0).Item("id_num_mza")
            Catch ex As Exception
                aux.Manzana = "- Sin Lote -"
            End Try
            Try
                aux.Lote = DatosFaltantes.Tables(0).Rows(0).Item("id_num_lote")
            Catch ex As Exception
                aux.Lote = "-Sin Lote-"
            End Try

            Resultado.Add(aux)
        Next

        Return Resultado
    End Function
    Function Obtener_Nombrecompleto_lider(ByVal desc_valor As String) As CNombreLiderGerente Implements IService1.Obtener_Nombrecompleto_lider

        Dim cmd As New MySqlCommand("SELECT nombrelidergerente.id_registro, nombrelidergerente.empleado, nombrelidergerente.desc_valor, nombrelidergerente.NombreCompleto FROM nombrelidergerente WHERE desc_valor LIKE '%" + desc_valor + "%'", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As New CNombreLiderGerente
        While reader.Read
            Aux.id_registro = DirectCast(reader.Item("id_registro"), Integer)
            Aux.empleado = DirectCast(reader.Item("empleado"), Integer)
            Aux.desc_valor = DirectCast(reader.Item("desc_valor"), String)
            Aux.NombreCompleto = DirectCast(reader.Item("NombreCompleto"), String)

        End While
        If String.IsNullOrEmpty(Aux.NombreCompleto) Then
            Aux.NombreCompleto = "-"
        End If
        Conexion.Close()
        Return Aux
    End Function
    Function Obtener_Detalles_comisionID(ByVal id_comision As Integer) As CComision Implements IService1.Obtener_Detalles_comisionID
        Dim Resultado As New CComision
        Dim NombreCliente As String = ""
        Dim NombreAsesor As String = ""
        Dim NombreLider As String = ""

        Dim DatosComision As DataSet = MySqlComi.MYSQlGetDataset("SELECT comisiones.id_comision, comisiones.numcte, comisiones.id_tipo_comision, comisiones.id_Promocion, comisiones.Premiado, comisiones.Pagado, comisiones.id_periodo, comisiones.Fecha_Pago, comisiones.Porcentaje_Com, comisiones.Cantidad_Pagada_Total, comisiones.id_reglaCom, comisiones.id_Tipo_Pago, comisiones.Empleado, comisiones.Lider, comisiones.Gerente, comisiones.Adm, comisiones.Observaciones, tipo_comsion.Descripcion, admin.NombreCompleto FROM comisiones INNER JOIN tipo_comsion ON comisiones.id_tipo_comision = tipo_comsion.id_tipo_comision INNER JOIN admin ON comisiones.Adm = admin.id_admin WHERE id_comision=" + id_comision.ToString + "")

        Resultado.id_comision = DatosComision.Tables(0).Rows(0).Item("id_comision")
        Resultado.numcte = DatosComision.Tables(0).Rows(0).Item("numcte")
        Resultado.NombreCliente = Obtener_Nombre_Cliente(Resultado.numcte)
        Resultado.Empleado = DatosComision.Tables(0).Rows(0).Item("Empleado")
        Resultado.Lider = Obtener_Nombre_Asesor(DatosComision.Tables(0).Rows(0).Item("Lider"))
        'Resultado.Gerente = Obtener_Nombrecompleto_lider(DatosComision.Tables(0).Rows(0).Item("Gerente")).NombreCompleto
        Resultado.Gerente = "MA DE BELEM OLVERA CANCHOLA"
        Resultado.Administrativo = DatosComision.Tables(0).Rows(0).Item("NombreCompleto")
        Resultado.Cantidad_Pagada_Total = DatosComision.Tables(0).Rows(0).Item("Cantidad_Pagada_Total")
        Resultado.Observacion = DatosComision.Tables(0).Rows(0).Item("Observaciones")



        Return Resultado
    End Function
    Function Modifica_Comision(ByVal id_comision As Integer, ByVal Cantidad_Pagada_Total As Integer, ByVal Observaciones As String) As Boolean Implements IService1.Modifica_Comision
        Try
            If MySqlComi.MySQLExecSQL("UPDATE comisiones SET Cantidad_Pagada_Total=" + Cantidad_Pagada_Total.ToString + ",Observaciones='" + Observaciones + "' WHERE id_comision=" + id_comision.ToString + "", MySqlComi.TipoTransaccion.UniqueTransaction) Then
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
        Return False
    End Function
    Function Resumen_fraccionamiento() As List(Of CResumenFracc) Implements IService1.Resumen_fraccionamiento
        Dim Resultado As New List(Of CResumenFracc)
        Dim aux As CResumenFracc
        Dim ComisionesPeriodos As DataSet = MySqlComi.MYSQlGetDataset("CALL ComisionesVigentes")

        For I = 0 To ComisionesPeriodos.Tables(0).Rows.Count - 1
            aux = New CResumenFracc
            aux.NombreFraccionamiento = obtener_fraccinamiento_cliente(ComisionesPeriodos.Tables(0).Rows(I).Item("numcte"))
            aux.Importe = ComisionesPeriodos.Tables(0).Rows(I).Item("Cantidad_Pagada_Total")
            aux.TipoPago = ComisionesPeriodos.Tables(0).Rows(I).Item("TipoPago")
            Resultado.Add(aux)



        Next
        Return Resultado
    End Function
    Function Resumen_asesor_comisiones() As List(Of CResumenAsesor) Implements IService1.Resumen_asesor_comisiones
        Dim Resultado As New List(Of CResumenAsesor)
        Dim aux As CResumenAsesor
        Dim ComisionesPeriodos As DataSet = MySqlComi.MYSQlGetDataset("CALL ComisionesVigentes")

        For I = 0 To ComisionesPeriodos.Tables(0).Rows.Count - 1
            aux = New CResumenAsesor
            'aux.NombreFraccionamiento = obtener_fraccinamiento_cliente(ComisionesPeriodos.Tables(0).Rows(I).Item("numcte"))
            aux.Importe = ComisionesPeriodos.Tables(0).Rows(I).Item("Cantidad_Pagada_Total")
            aux.TipoPago = ComisionesPeriodos.Tables(0).Rows(I).Item("TipoPago")


            Select Case ComisionesPeriodos.Tables(0).Rows(I).Item("TipoComision")
                Case 1
                    aux.NombreAsesor = Obtener_Nombre_Asesor(ComisionesPeriodos.Tables(0).Rows(I).Item("empleado"))
                Case 2
                    aux.NombreAsesor = Obtener_Nombre_Asesor(ComisionesPeriodos.Tables(0).Rows(I).Item("Lider"))
                Case 3
                    aux.NombreAsesor = ComisionesPeriodos.Tables(0).Rows(I).Item("Gerente")
                Case 4
                    aux.NombreAsesor = ComisionesPeriodos.Tables(0).Rows(I).Item("NombreCompleto")
            End Select





            Resultado.Add(aux)
        Next
        Return Resultado
    End Function
    Function Obtener_nombreliderOGerente(ByVal desc_valor As String) As String
        Dim DS = MySqlProConsul2.MYSQlGetDataset("SELECT nombrelidergerente.NombreCompleto FROM nombrelidergerente WHERE desc_valor='" + desc_valor + "'")
        Dim Resultado As String = ""
        Try
            If DS.Tables(0).Rows.Count > 0 Then
                Resultado = DS.Tables(0).Rows(0).Item("NombreCompleto")
            End If

        Catch ex As Exception
            Resultado = "-"
        End Try

        Return Resultado
    End Function
    Function obtener_fraccinamiento_cliente(ByVal numcte As Integer) As String
        Dim Resupuesta As String = ""
        Try
            Resupuesta = ODBCGetDataset("SELECT Nom_fracc  FROM   dba.sm_fraccionamiento, dba.sm_cliente WHERE  dba.sm_cliente.id_cve_fracc=dba.sm_fraccionamiento.id_cve_fracc and dba.sm_cliente.numcte=" + numcte.ToString + ";").Tables(0).Rows(0).Item(0)
        Catch ex As Exception
            Resupuesta = "- Sin Lote -"
        End Try
        Return Resupuesta
    End Function
#End Region
#Region "Comisiones"
    Function Actualiza_admin(ByVal id_admin As Integer, ByVal Empleado As Integer, ByVal NombreCompleto As String, ByVal Cantidad As Integer) As Boolean Implements IService1.Actualiza_admin

        Dim cmd As New MySqlCommand("UPDATE  admin SET Empleado=?PEmpleado, NombreCompleto=?PNombreCompleto, Cantidad=?PCantidad WHERE id_admin= ?Pid_admin;", ConexionComisiones)
        'cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("?PEmpleado", Empleado)
        cmd.Parameters.AddWithValue("?PNombreCompleto", NombreCompleto)
        cmd.Parameters.AddWithValue("?PCantidad", Cantidad)
        cmd.Parameters.AddWithValue("?Pid_admin", id_admin)
        ConexionComisiones.Close()
        Try
            ConexionComisiones.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                ConexionComisiones.Close()
                Return True
            End If
        Catch ex As Exception
            ConexionComisiones.Close()
            Return False
        End Try
        ConexionComisiones.Close()
        Return False
    End Function
    Function Obtener_admin() As List(Of CAdministrativos) Implements IService1.Obtener_admin
        Dim Resultado As New List(Of CAdministrativos)
        Dim cmd As New MySqlCommand("SELECT * FROM admin", ConexionComisiones)
        'cmd.CommandType = CommandType.StoredProcedure
        ConexionComisiones.Close()
        ConexionComisiones.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CAdministrativos
        While reader.Read
            Aux = New CAdministrativos
            Aux.id_admin = DirectCast(reader.Item("id_admin"), Integer)
            Aux.Empleado = DirectCast(reader.Item("Empleado"), Integer)
            Aux.NombreCompleto = DirectCast(reader.Item("NombreCompleto"), String)
            Aux.Cantidad = DirectCast(reader.Item("Cantidad"), Integer)
            Resultado.Add(Aux)
        End While
        ConexionComisiones.Close()
        Return Resultado
    End Function

    Function Inserta_comisiones_clasificaciones(ByVal DescripcionClasificacion As String) As Boolean Implements IService1.Inserta_comisiones_clasificaciones

        Dim cmd As New MySqlCommand("Inserta_Clasificacion", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("PDescripcionClasificacion", DescripcionClasificacion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_comisiones_clasificaciones(ByVal id_clasificacion As Integer) As Boolean Implements IService1.Elimina_comisiones_clasificaciones

        Dim cmd As New MySqlCommand("Elimina_clasificacion", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_clasificacion", id_clasificacion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_comisionesgerencia() As List(Of CPagosGerencia) Implements IService1.Obtener_comisionesgerencia
        Dim Resultado As New List(Of CPagosGerencia)
        Dim cmd As New MySqlCommand("SELECT * FROM comisionesgerencia", ConexionComisiones)
        'cmd.CommandType = CommandType.StoredProcedure
        ConexionComisiones.Close()
        ConexionComisiones.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CPagosGerencia
        While reader.Read
            Aux = New CPagosGerencia
            Aux.CC = DirectCast(reader.Item("CC"), String)
            Aux.Cantidad = DirectCast(reader.Item("Cantidad"), Integer)
            Resultado.Add(Aux)
        End While
        ConexionComisiones.Close()
        Return Resultado
    End Function
    Function Actualiza_comisiones_clasificaciones(ByVal id_clasificacion As Integer, ByVal DescripcionClasificacion As String) As Boolean Implements IService1.Actualiza_comisiones_clasificaciones

        Dim cmd As New MySqlCommand("Actualiza_clasificacion", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_clasificacion", id_clasificacion)
        cmd.Parameters.AddWithValue("PDescripcionClasificacion", DescripcionClasificacion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_comisiones_clasificaciones() As List(Of CClasificaciones) Implements IService1.Obtener_comisiones_clasificaciones
        Dim Resultado As New List(Of CClasificaciones)
        Dim cmd As New MySqlCommand("SELECT * FROM comisiones_clasificaciones", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CClasificaciones
        While reader.Read
            Aux = New CClasificaciones
            Aux.id_clasificacion = DirectCast(reader.Item("id_clasificacion"), Integer)
            Aux.DescripcionClasificacion = DirectCast(reader.Item("DescripcionClasificacion"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_comisiones_clasificacion(ByVal id_clasificacion As Integer) As CClasificaciones Implements IService1.Obtener_comisiones_clasificacion

        Dim cmd As New MySqlCommand("SELECT * FROM comisiones_clasificaciones WHERE id_clasificacion=" + id_clasificacion.ToString, Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As New CClasificaciones
        While reader.Read
            Aux = New CClasificaciones
            Aux.id_clasificacion = DirectCast(reader.Item("id_clasificacion"), Integer)
            Aux.DescripcionClasificacion = DirectCast(reader.Item("DescripcionClasificacion"), String)

        End While
        Conexion.Close()
        Return Aux
    End Function


    Function Obtener_esquemas_enkontrol() As List(Of CEsquemas) Implements IService1.Obtener_esquemas_enkontrol
        Dim Resultado As New List(Of CEsquemas)
        Dim DS As New DataSet
        Dim Aux As New CEsquemas
        Try
            DS = ODBCGetDataset("select id_num_relacion,Nom_Financinst FROM dba.sm_financiamiento_instcred WHERE id_num_relacion>=200;")
            For I = 0 To DS.Tables(0).Rows.Count - 1
                Aux = New CEsquemas
                Aux.id_num_relacion = DS.Tables(0).Rows(I).Item("id_num_relacion")
                Aux.Nom_Financinst = DS.Tables(0).Rows(I).Item("Nom_Financinst")
                Resultado.Add(Aux)
            Next
        Catch ex As Exception

        End Try
        Return Resultado
    End Function
    Function Obtener_Reglas_Comisiones() As List(Of CReglasComisiones) Implements IService1.Obtener_Reglas_Comisiones
        Dim Resultado As New List(Of CReglasComisiones)
        Dim cmd As New MySqlCommand("SELECT comisiones_reglas.id_regla, comisiones_reglas.empleado, comisiones_reglas.id_clasificacion, comisiones_reglas.etapaFinal, comisiones_reglas.Cantidad_PagoUnitario FROM comisiones_reglas ", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CReglasComisiones
        While reader.Read
            Aux = New CReglasComisiones
            Aux.id_regla = DirectCast(reader.Item("id_regla"), Integer)
            Aux.empleado = DirectCast(reader.Item("empleado"), Integer)
            Aux.id_clasificacion = DirectCast(reader.Item("id_clasificacion"), Integer)
            Aux.etapaFinal = DirectCast(reader.Item("etapaFinal"), Integer)
            Aux.Cantidad_PagoUnitario = DirectCast(reader.Item("Cantidad_PagoUnitario"), Integer)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_comisiones_reglas(ByVal empleado As Integer, ByVal id_clasificacion As Integer, ByVal etapaFinal As Integer, ByVal Cantidad_PagoUnitario As Integer) As Boolean Implements IService1.Inserta_comisiones_reglas

        Dim cmd As New MySqlCommand("Inserta_Regla", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pempleado", empleado)
        cmd.Parameters.AddWithValue("Pid_clasificacion", id_clasificacion)
        cmd.Parameters.AddWithValue("PetapaFinal", etapaFinal)
        cmd.Parameters.AddWithValue("PCantidad_PagoUnitario", Cantidad_PagoUnitario)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_comisiones_reglas(ByVal id_regla As Integer) As Boolean Implements IService1.Elimina_comisiones_reglas

        Dim cmd As New MySqlCommand("Elimina_regla", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_regla", id_regla)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function

    Function Comprobar_duplicador_esquemas(ByVal id_esquema As Integer) As Boolean
        Dim DS As New DataSet
        Try
            DS = MySqlProConsul2.MYSQlGetDataset("SELECT id_registro FROM comisiones_esquemas WHERE id_esquema=" + id_esquema.ToString + ";")
        Catch ex As Exception

        End Try

        If DS.Tables(0).Rows.Count > 0 Then
            Return True
        End If
        Return False
    End Function
    Function Inserta_comisiones_esquemas(ByVal id_esquema As Integer, ByVal id_clasificacion As Integer) As Boolean Implements IService1.Inserta_comisiones_esquemas

        If Comprobar_duplicador_esquemas(id_esquema) Then
        Else

            Dim cmd As New MySqlCommand("Inserta_comisiones_esquemas", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("Pid_esquema", id_esquema)
            cmd.Parameters.AddWithValue("Pid_clasificacion", id_clasificacion)
            Conexion.Close()
            Try
                Conexion.Open()
                If cmd.ExecuteNonQuery() > 0 Then
                    Conexion.Close()
                    Return True
                End If
            Catch ex As Exception
                Conexion.Close()
                Return False
            End Try
            Conexion.Close()
            Return False
        End If
        Return False
    End Function
    Function Elimina_comisiones_esquemas(ByVal id_registro As Integer) As Boolean Implements IService1.Elimina_comisiones_esquemas

        Dim cmd As New MySqlCommand("Elimina_comisiones_esquemas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_registro", id_registro)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_Comision(ByVal id_comisione As Integer) As Boolean Implements IService1.Elimina_Comision

        Dim cmd As New MySqlCommand("Elimina_comision", ConexionComisiones)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("PidComision", id_comisione)
        ConexionComisiones.Close()
        Try
            ConexionComisiones.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                ConexionComisiones.Close()
                Return True
            End If
        Catch ex As Exception
            ConexionComisiones.Close()
            Return False
        End Try
        ConexionComisiones.Close()
        Return False
    End Function
    Function Obtener_comisiones_esquemas() As List(Of CEsquemaRegla) Implements IService1.Obtener_comisiones_esquemas
        Dim Resultado As New List(Of CEsquemaRegla)
        Dim cmd As New MySqlCommand("SELECT * FROM comisiones_esquemas", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CEsquemaRegla
        While reader.Read
            Aux = New CEsquemaRegla
            Aux.id_registro = DirectCast(reader.Item("id_registro"), Integer)
            Aux.id_esquema = DirectCast(reader.Item("id_esquema"), Integer)
            Aux.id_clasificacion = DirectCast(reader.Item("id_clasificacion"), Integer)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_comisiones_esquema(ByVal id_clasificacion As Integer) As List(Of CEsquemaRegla) Implements IService1.Obtener_comisiones_esquema
        Dim Resultado As New List(Of CEsquemaRegla)
        Dim cmd As New MySqlCommand("SELECT * FROM comisiones_esquemas WHERE id_clasificacion=" + id_clasificacion.ToString, Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CEsquemaRegla
        While reader.Read
            Aux = New CEsquemaRegla
            Aux.id_registro = DirectCast(reader.Item("id_registro"), Integer)
            Aux.id_esquema = DirectCast(reader.Item("id_esquema"), Integer)
            Aux.id_clasificacion = DirectCast(reader.Item("id_clasificacion"), Integer)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function




    Function Obtener_Valores_Comisiones() As List(Of CValoresComisiones) Implements IService1.Obtener_Valores_Comisiones
        Dim Resultado As New List(Of CValoresComisiones)
        Dim cmd As New MySqlCommand("SELECT comisiones_porcentajes.id_porcentaje,comisiones_porcentajes.Descripcion_porcentaje, comisiones_porcentajes.Porcentaje FROM comisiones_porcentajes ", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CValoresComisiones
        While reader.Read
            Aux = New CValoresComisiones
            Aux.id_porcentaje = DirectCast(reader.Item("id_porcentaje"), Integer)
            Aux.Descripcion_porcentaje = DirectCast(reader.Item("Descripcion_porcentaje"), String)
            Aux.Porcentaje = reader.Item("Porcentaje")
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Actualiza_comisiones_porcentajes(ByVal id_valor As Integer, ByVal Porcentaje As String) As Boolean Implements IService1.Actualiza_comisiones_porcentajes

        Dim cmd As New MySqlCommand("Actualiza_valor", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_porcentaje", id_valor)
        cmd.Parameters.AddWithValue("PPorcentaje", Porcentaje)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_comisiones_porcentajes(ByVal id_porcentaje As Integer) As List(Of CValoresComisiones) Implements IService1.Obtener_comisiones_porcentajes
        Dim Resultado As New List(Of CValoresComisiones)
        Dim cmd As New MySqlCommand("SELECT * FROM comisiones_porcentajes WHERE id_porcentaje=" + id_porcentaje.ToString, Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CValoresComisiones
        While reader.Read
            Aux = New CValoresComisiones
            Aux.id_porcentaje = DirectCast(reader.Item("id_porcentaje"), Integer)
            Aux.Descripcion_porcentaje = DirectCast(reader.Item("Descripcion_porcentaje"), String)
            Aux.Porcentaje = reader.Item("Porcentaje")
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Programacion Entrega De vividenda"
    Function Inserta_programacionentrega(ByVal numcte As Integer, ByVal Hora As TimeSpan, ByVal FECHA As Date) As Boolean Implements IService1.Inserta_programacionentrega

        Dim cmd As New MySqlCommand("Inserta_Programacion", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pnumcte", numcte)
        cmd.Parameters.AddWithValue("PNombreCliente", Obtener_Nombre_Cliente(numcte))
        cmd.Parameters.AddWithValue("PHora", Hora)
        cmd.Parameters.AddWithValue("PFECHA", FECHA)
        Try
            cmd.Parameters.AddWithValue("PCC", ODBCGetDataset("SELECT id_cve_Fracc FROM dba.sm_cliente WHERE numcte=" + numcte.ToString + ";", 11).Tables(0).Rows(0).Item(0).ToString)
        Catch ex As Exception
            cmd.Parameters.AddWithValue("PCC", "-")
        End Try



        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Cambia_contratos(ByVal id_contrato As Integer, ByVal Activo As Integer) As Boolean Implements IService1.Cambia_contratos

        Dim cmd As New MySqlCommand("UPDATE pro_contratos_nuevo SET Activo=?PActivo WHERE id_contrato=?PidContrato", ConexionGedificasas)
        'cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("?PidContrato", id_contrato)
        cmd.Parameters.AddWithValue("?PActivo", Activo)
        ConexionGedificasas.Close()
        Try
            ConexionGedificasas.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                ConexionGedificasas.Close()
                Return True
            End If
        Catch ex As Exception
            ConexionGedificasas.Close()
            Return False
        End Try
        ConexionGedificasas.Close()
        Return False
    End Function
    Function Elimina_programacionentrega(ByVal id_programacion As Integer) As Boolean Implements IService1.Elimina_programacionentrega

        Dim cmd As New MySqlCommand("Elimina_Programacion", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_programacion", id_programacion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_programacionentrega(ByVal id_programacion As Integer, ByVal numcte As Integer, ByVal NombreCliente As String, ByVal Hora As TimeSpan, ByVal FECHA As Date) As Boolean Implements IService1.Actualiza_programacionentrega

        Dim cmd As New MySqlCommand("Actualiza_Programacion", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("Pid_programacion", id_programacion)
        cmd.Parameters.AddWithValue("Pnumcte", numcte)
        cmd.Parameters.AddWithValue("PNombreCliente", NombreCliente)
        cmd.Parameters.AddWithValue("PHora", Hora)
        cmd.Parameters.AddWithValue("PFECHA", FECHA)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_programacionentrega() As List(Of CProgramacionEntrega) Implements IService1.Obtener_programacionentrega
        Dim Resultado As New List(Of CProgramacionEntrega)
        Dim cmd As New MySqlCommand("SELECT * FROM programacionEntrega", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As CProgramacionEntrega
        While reader.Read
            Aux = New CProgramacionEntrega
            Aux.id_programacion = DirectCast(reader.Item("id_programacion"), Integer)
            Aux.numcte = DirectCast(reader.Item("numcte"), Integer)
            Aux.NombreCliente = DirectCast(reader.Item("NombreCliente"), String)
            If Aux.numcte = 25164 Then
                Aux.Hora = TimeSpan.Parse("16:00:00")
            Else
                Aux.Hora = reader.Item("Hora")
            End If

            Aux.FECHA = DirectCast(reader.Item("FECHA"), Date)
            Aux.CC = DirectCast(reader.Item("CC"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region

#Region "Reportes"
    Function Obtener_programacionEntregaAdm(ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CProgramacionEntregaAdm) Implements IService1.Obtener_programacionEntregaAdm
        Dim Resultado As New List(Of CProgramacionEntregaAdm)
        Dim FechaReprogramacion As New Date
        Dim Regular As New Regex("\d{4}(?:/\d{1,2}){2}")
        Dim LlamadaCalidad As String = ""
        Dim RevisionDeCalidad As String = ""
        Dim cmd As New OdbcCommand("SELECT dba.sm_fraccionamiento_lote.id_num_lote,	dba.sm_fraccionamiento.Nom_fracc, dba.sm_cliente.numcte AS NumeroCliente, 
                                           (ISNULL((SELECT Fec_registo FROM dba.sm_cliente_parentesco WHERE	numcte = NumeroCliente AND id_num_tiporelacion = 8),
                                           (SELECT Fec_liberacion AS EntregaCliente	FROM dba.sm_cliente_etapa WHERE	numcte = NumeroCliente AND id_num_etapa = 19))) AS EntregaCliente, 	
                                           (SELECT Fec_inicio AS EntregaCliente	FROM dba.sm_cliente_etapa WHERE	numcte = NumeroCliente AND id_num_etapa = 18) AS ProgramacionFirma, 	
                                           ISNULL((SELECT Observaciones	FROM dba.sm_cliente_etapa WHERE	id_num_etapa = 20 AND numcte = NumeroCliente), '-') AS Reprogramacion,
                                           NombreCliente = (Ap_paterno_cte + ' ' + ap_materno_cte + ' ' + Nom_cte),	dba.sm_fraccionamiento_lote.id_num_smza, dba.sm_cliente.lote_id, 	
                                           dba.sm_fraccionamiento_lote.id_num_mza, dba.sm_fraccionamiento_lote.id_num_interior,	dba.sm_fraccionamiento_lote.dir_casa,
                                           ISNULL((SELECT Observaciones AS EntregaCliente FROM dba.sm_cliente_etapa	WHERE numcte = NumeroCliente AND id_num_etapa = 19), '-') AS ObservacionProgramacion, 	
                                           ISNULL(CONVERT(CHAR (20),(SELECT	Fec_registo	FROM dba.sm_cliente_parentesco WHERE numcte = NumeroCliente	AND id_num_tiporelacion = 8), 104), '-', 'X') AS SS 
                                    FROM dba.sm_cliente, dba.sm_fraccionamiento, dba.sm_agente,	dba.sm_fraccionamiento_lote 
                                    WHERE dba.sm_cliente.id_cve_fracc = dba.sm_fraccionamiento.id_cve_fracc AND dba.sm_cliente.empleado = dba.sm_agente.Empleado AND dba.sm_cliente.lote_id = dba.sm_fraccionamiento_lote.lote_id AND EntregaCliente BETWEEN '" + FechaInicio.ToString("yyyy/MM/dd") + "' and '" + FechaFinal.ToString("yyyy/MM/dd") + "';", ConexionEnkontrol)
        'cmd.CommandType = CommandType.StoredProcedure       
        ConexionEnkontrol.Close()
        ConexionEnkontrol.Open()
        Dim reader As OdbcDataReader = cmd.ExecuteReader
        Dim Aux As CProgramacionEntregaAdm


        While reader.Read
            Aux = New CProgramacionEntregaAdm
            Aux.Nom_fracc = DirectCast(reader.Item("Nom_fracc"), String)
            Aux.NumeroCliente = DirectCast(reader.Item("NumeroCliente"), Decimal)

            'Try
            '    FechaReprogramacion = ODBCGetDataset("SELECT Fec_registo FROM dba.sm_cliente_parentesco WHERE numcte=" + Aux.NumeroCliente + " and id_num_tiporelacion=8;", 11).Tables(0).Rows(0).Item(0)
            'Catch ex As Exception
            '    FechaReprogramacion = New Date
            'End Try

            Try
                Aux.programacionFirma = DirectCast(reader.Item("ProgramacionFirma"), Date)
            Catch ex As Exception
                Aux.programacionFirma = New Date
            End Try

            Aux.EntregaCliente = DirectCast(reader.Item("EntregaCliente"), Date)
            Aux.PreEntrega = Resta5diasHabiles(Aux.EntregaCliente, Aux.Nom_fracc)
            Aux.NombreCliente = DirectCast(reader.Item("NombreCliente"), String)
            Aux.id_num_smza = DirectCast(reader.Item("id_num_smza"), String)
            Aux.id_num_mza = DirectCast(reader.Item("id_num_mza"), String)
            Aux.id_num_interior = DirectCast(reader.Item("id_num_interior"), String)
            Aux.dir_casa = DirectCast(reader.Item("dir_casa"), String)
            Aux.id_num_lote = reader.Item("id_num_lote")
            Aux.ObservacionProgramacion = DirectCast(reader.Item("ObservacionProgramacion"), String)

            'If reader.Item("Reprogramacion") Like "*REPROGRAMACION:*" Then
            '    Try
            '        Aux.EntregaCliente = Regular.Match(reader.Item("Reprogramacion")).Value
            '        Aux.Reprogramado = True
            '    Catch ex As Exception
            '        Aux.EntregaCliente = DirectCast(reader.Item("EntregaCliente"), Date)
            '    End Try

            '    Aux.EntregaCliente = FechaReprogramacion
            'End If

            Try
                If reader.Item("SS") <> "-" Then
                    Aux.Reprogramado = True
                Else
                    Aux.Reprogramado = False
                End If
            Catch ex As Exception
                Aux.Reprogramado = False
            End Try

            Try
                LlamadaCalidad = ODBCGetDataset("SELECT ISNULL(desc_valor,'-') AS LlamadaCalidad  FROm dba.sm_cliente_adicional WHERE numcte=" + Aux.NumeroCliente.ToString + " AND Desc_Referencia='LLAMADA DE CALIDAD';", 11).Tables(0).Rows(0).Item(0)
            Catch ex As Exception
                LlamadaCalidad = "-"
            End Try
            Try
                RevisionDeCalidad = ODBCGetDataset("SELECT ISNULL(desc_valor,'-') AS LlamadaCalidad  FROm dba.sm_cliente_adicional WHERE numcte=" + Aux.NumeroCliente.ToString + " AND Desc_Referencia='REVISION DE VIVIENDA';", 11).Tables(0).Rows(0).Item(0)
            Catch ex As Exception
                RevisionDeCalidad = "-"
            End Try

            Aux.LlamadaDeCalidad = LlamadaCalidad
            Aux.RevisioonDeDivivenda = RevisionDeCalidad

            If Aux.NumeroCliente = 25164 Then
                Aux.ObservacionProgramacion = "A LAS 4"
            End If

            'If Aux.NumeroCliente = 23666 Then
            '    Aux.ObservacionProgramacion = "A LAS 11:30 am"
            'End If
            Resultado.Add(Aux)
        End While
        ConexionEnkontrol.Close()
        Return Resultado
    End Function
    Function Resta5diasHabiles(ByRef FechaOriginal As Date, ByVal NomFracc As String) As Date
        Dim Resultado As Date = FechaOriginal
        Dim DiaSemana As Integer = 0

        For I = 0 To 4
            Resultado = Resultado.AddDays(-1)
            If Resultado.DayOfWeek = DayOfWeek.Sunday Then
                I = I - 1
            End If
        Next

        If Resultado.DayOfWeek = DayOfWeek.Sunday Then
            Resultado = Resultado.AddDays(-1)
        End If

        Return Resultado
    End Function
    Function Reporte_Penalizacion(ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CPagoyPenalizacion) Implements IService1.Reporte_Penalizacion
        Dim Resultado As New List(Of CPagoyPenalizacion)
        Dim DatosEnkontrol = ODBCGetDataset("SELECT 	dba.sm_cliente.numcte, 	NombreCliente = ( 		CONVERT (CHAR(20), dba.sm_cliente.numcte) + ' ' + nom_cte + ' ' + ap_paterno_cte + ' ' + ap_materno_cte 	), 	NombreFracc = ( 		dba.sm_cliente.id_cve_fracc + ' ' + dba.sm_fraccionamiento.nom_fracc 	), 	dba.sm_cliente.empleado, 	NombreEmpleado = ( 		CONVERT ( 			CHAR (20), 			dba.sm_cliente.empleado 		) + ' ' + nom_empleado + ' ' + ap_paterno_empleado + ' ' + ap_materno_empleado 	) FROM 	dba.sm_cliente, 	dba.sm_fraccionamiento, 	dba.sm_agente WHERE 	dba.sm_cliente.empleado = dba.sm_agente.empleado AND dba.sm_cliente.id_cve_fracc = dba.sm_fraccionamiento.id_cve_fracc AND dba.sm_cliente.fec_registo BETWEEN '" + FechaInicio.ToString("yyyy/MM/dd") + "' AND '" + FechaFinal.ToString("yyyy/MM/dd") + "' AND STATUS_cte='C';")
        Dim Aux As New CPagoyPenalizacion



        For I = 0 To DatosEnkontrol.Tables(0).Rows.Count - 1
            Dim Fila = DatosEnkontrol.Tables(0).Rows(I)
            Dim Numcte As Integer = 0
            Numcte = Fila.Item("numcte")

            Dim CantidadPagado = Obtener_CantidadPagada(Numcte)
            Dim CantidadPenalizado = Obtener_CantidadPenalizada(Numcte)


            Aux = New CPagoyPenalizacion
            Aux.NombreCliente = Fila.Item("NombreCliente")
            Aux.CC = Fila.Item("NombreFracc")
            Aux.Empleado = Fila.Item("Empleado")
            Aux.NombreEmpleado = Fila.Item("NombreEmpleado")
            Aux.PagadoTotal = CantidadPagado.CantidadPagada
            Aux.PenalizadoTotal = CantidadPenalizado.CantidadPagada
            Aux.FechaPago = CantidadPagado.FechaPago
            Aux.FechaPenalizado = CantidadPenalizado.FechaPago
            Resultado.Add(Aux)
        Next



        Return Resultado
    End Function
    Function Obtener_CantidadPagada(ByVal numcte As Integer) As CPagosPenaliza

        Dim cmd As New MySqlCommand("SELECT SUM(Cantidad_pagada_total) as Cantidad, MAX(Fecha_Pago) as FechaPago FROM comisiones WHERE numcte=?PNumcte and id_tipo_comision=1 and Cantidad_Pagada_Total>0;", ConexionComisiones)
        'cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("?PNumcte", numcte)
        ConexionComisiones.Close()
        ConexionComisiones.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As New CPagosPenaliza
        While reader.Read

            Try
                Aux.CantidadPagada = reader.Item("Cantidad")
                Aux.FechaPago = DirectCast(reader.Item("FechaPago"), Date)
            Catch ex As Exception
                Aux.CantidadPagada = 0
            End Try

        End While
        ConexionComisiones.Close()
        Return Aux
    End Function
    Function Obtener_CantidadPenalizada(ByVal numcte As Integer) As CPagosPenaliza

        Dim cmd As New MySqlCommand("SELECT SUM(Cantidad_pagada_total) as Cantidad, MAX(Fecha_Pago) as FechaPago FROM comisiones WHERE numcte=?PNumcte and id_tipo_comision=1 and Cantidad_Pagada_Total<0;", ConexionComisiones)
        'cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("?PNumcte", numcte)
        ConexionComisiones.Close()
        ConexionComisiones.Open()
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        Dim Aux As New CPagosPenaliza
        While reader.Read
            Try
                Aux.CantidadPagada = reader.Item("Cantidad")
                Aux.FechaPago = DirectCast(reader.Item("FechaPago"), Date)
            Catch ex As Exception
                Aux.CantidadPagada = 0
            End Try


        End While
        ConexionComisiones.Close()
        Return Aux
    End Function
    Function Obtener_loteids() As String Implements IService1.Obtener_loteids
        Dim Resultado As String = ""


        Dim DS = ODBCGetDataset("SELECT lote_id,id_num_mza,id_num_lote,id_num_interior,id_cve_fracc FROM dba.sm_fraccionamiento_lote;")

        Resultado += "["
        For I = 0 To DS.Tables(0).Rows.Count - 1
            Resultado += "{"
            Resultado += "lote_id:" + DS.Tables(0).Rows(I).Item("lote_id").ToString + ","
            Resultado += "id_num_mza:" + DS.Tables(0).Rows(I).Item("id_num_mza").ToString + ","
            Resultado += "id_num_lote:" + DS.Tables(0).Rows(I).Item("id_num_lote").ToString + ","
            Resultado += "id_num_interior:" + DS.Tables(0).Rows(I).Item("id_num_interior").ToString + ","
            Resultado += "id_cve_fracc:" + DS.Tables(0).Rows(I).Item("id_cve_fracc").ToString + ""
            Resultado += "}"
        Next
        Resultado += "]"


        Return Resultado
    End Function
    Class CPagosPenaliza
        Property CantidadPagada As Integer
        Property FechaPago As Date
    End Class

#End Region
#Region "cancelaciones"
    Public Function Obtener_datos_cancelaciones() As List(Of CCancelacionesDetalles) Implements IService1.Obtener_datos_cancelaciones
        Dim Resultado As New List(Of CCancelacionesDetalles)
        Dim Aux As New CCancelacionesDetalles
        Dim DatosEnkontrol As New DataSet
        Dim DatosComisiones As New DataSet
        Dim Cancelados As DataSet = MySqlProConsul2.MYSQlGetDataset("SELECT * FROM com_cancelaciones")

        For I = 0 To Cancelados.Tables(0).Rows.Count - 1


            'Datos Cancelados
            Dim Cancelado = Cancelados.Tables(0).Rows(I)

            Aux = New CCancelacionesDetalles
            Aux.Numcte = Cancelado.Item("numcte")

            Aux.P_Cliente = Cancelado.Item("P_Cliente")
            Aux.P_Asesor = Cancelado.Item("P_Asesor")
            Aux.P_Lider = Cancelado.Item("P_Lider")
            Aux.P_Gerente = Cancelado.Item("P_Gerente")
            Aux.P_Administrativo = Cancelado.Item("P_Administrativo")
            Aux.FechaCancelacion = Cancelado.Item("FechaCancelacion")

            'Datos de comisiones
            DatosComisiones = MySqlComi.MYSQlGetDataset("CALL Obtener_pagos_cliente(" + Aux.Numcte.ToString + ")")
            Dim Comision = DatosComisiones.Tables(0).Rows(0)

            Aux.PagoAsesor = Comision.Item("Pago_Asesor")
            Aux.PagoGerente = Comision.Item("Pago_Gerente")
            Aux.PagoAsesor = Comision.Item("Pago_Lider")
            Aux.P_Administrativo = Comision.Item("Pago_Adminsitrativo")
            Try
                Aux.FechaPago = Comision.Item("FechaPago")
            Catch ex As Exception
                Aux.FechaPago = New Date
            End Try


            'Datos de enkontrol
            DatosEnkontrol = ODBC.ODBCGetDataset("SELECT NombreCliente=(Nom_cte+' '+Ap_paterno_cte+' '+ap_materno_cte+' '), dba.sm_cliente.id_cve_fracc as CC, dba.sm_fraccionamiento.Nom_fracc, dba.sm_cliente.Empleado, NombreAsesor=(dba.sm_agente.Nom_empleado+' '+ dba.sm_agente.Ap_paterno_empleado+' '+ dba.sm_agente.Ap_materno_empleado), dba.sm_agente.Direccion_Archivo as LIDER FROM  dba.sm_cliente, dba.sm_fraccionamiento, dba.sm_agente WHERE dba.sm_cliente.id_cve_fracc=dba.sm_fraccionamiento.id_cve_fracc AND dba.sm_cliente.empleado=dba.sm_agente.Empleado AND dba.sm_cliente.numcte=" + Aux.Numcte.ToString + ";", 11)
            Dim Dato = DatosEnkontrol.Tables(0).Rows(0)
            Aux.NombreAsesor = Dato.Item("NombreAsesor")
            Aux.NombreCliente = Dato.Item("NombreCliente")
            Aux.NombreLider = Obtener_Nombre_Asesor(Dato.Item("LIDER"))
            Aux.Empleado = Dato.Item("Empleado")
            'Aquí siempre es BELEM
            Aux.NombreGerente = "BELEM OLVERA CANCHOLA"
            Aux.CC = Dato.Item("CC")
            Aux.Fraccionamiento = Dato.Item("Nom_fracc")


            Aux.id_cancelacion = Cancelado.Item("id_cancelacion")

            Resultado.Add(Aux)


        Next

        Return Resultado
    End Function

#End Region

#Region "SubirFotosSAC"
    Function ObtenerDSEnkontrol(ByVal SQL As String) As DataTable Implements IService1.ObtenerDSEnkontrol
        Return ODBCGetDataset(SQL).Tables(0)
    End Function
#End Region
    Function Obtener_datos_cliente_contrato_fovissste(ByVal numcte As Integer) As CDatosClienteContratoFovissste Implements IService1.Obtener_datos_cliente_contrato_fovissste

        Dim cmd As New OdbcCommand("SELECT NombreCliente=(Nom_cte+' '+Ap_paterno_cte+' '+ap_materno_cte+' '), dba.sm_cliente.Ciudad, dba.sm_cliente.Fec_Nac, Edo_civil, dba.sm_cliente_empresa.puesto, dba.sm_cliente.Dir_Casa, dba.sm_cliente.RFC_cte, id_num_relacion, dba.sm_cliente.valor_total, dba.sm_fraccionamiento_lote.id_num_lote, dba.sm_fraccionamiento_lote.id_num_interior, dba.sm_fraccionamiento_lote.id_num_mza, dba.sm_fraccionamiento.nom_fracc, ISNULL(dba.sm_fraccionamiento_lote.Desc_Colindancia,'-') as Desc_Colindancia, dba.sm_fraccionamiento_lote.Cant_superficie  FROM  dba.sm_cliente, dba.sm_cliente_empresa, dba.sm_fraccionamiento_lote, dba.sm_fraccionamiento WHERE dba.sm_cliente.numcte=dba.sm_cliente_empresa.numcte AND dba.sm_cliente.lote_id=dba.sm_fraccionamiento_lote.lote_id AND dba.sm_cliente.id_cve_fracc=dba.sm_fraccionamiento.id_cve_fracc AND dba.sm_cliente.numcte=" + numcte.ToString + ";", cn)
        'cmd.CommandType = CommandType.StoredProcedure
        cn.Close()
        cn.Open()
        Dim reader As OdbcDataReader = cmd.ExecuteReader
        Dim Aux As New CDatosClienteContratoFovissste
        While reader.Read
            Aux = New CDatosClienteContratoFovissste
            Aux.NombreCliente = DirectCast(reader.Item("NombreCliente"), String)
            Aux.Ciudad = DirectCast(reader.Item("Ciudad"), String)
            Aux.Fec_Nac = DirectCast(reader.Item("Fec_Nac"), Date)
            Aux.Edo_civil = DirectCast(reader.Item("Edo_civil"), String)
            Aux.puesto = DirectCast(reader.Item("puesto"), String)
            Aux.Dir_Casa = DirectCast(reader.Item("Dir_Casa"), String)
            Aux.RFC_cte = DirectCast(reader.Item("RFC_cte"), String)
            Aux.id_num_relacion = DirectCast(reader.Item("id_num_relacion"), Decimal)
            Aux.valor_total = DirectCast(reader.Item("valor_total"), Decimal)
            Aux.id_num_lote = DirectCast(reader.Item("id_num_lote"), String)
            Aux.id_num_interior = DirectCast(reader.Item("id_num_interior"), String)
            Aux.id_num_mza = DirectCast(reader.Item("id_num_mza"), String)
            Aux.nom_fracc = DirectCast(reader.Item("nom_fracc"), String)
            Aux.Desc_Colindancia = reader.Item("Desc_Colindancia")
            Aux.Cant_superficie = DirectCast(reader.Item("Cant_superficie"), Decimal)

        End While
        cn.Close()
        Return Aux
    End Function
End Class

Public Class CConexion
    Public ODBCconStr As String
    Public ODBCcon_A As IDbConnection = New OdbcConnection(ODBCconStr)
    Public ODBCcon_B As IDbConnection = New OdbcConnection(ODBCconStr)
    Public ODBC_CMD As IDbCommand = ODBCcon_A.CreateCommand()
    Public ODBC_DA As IDbDataAdapter = New OdbcDataAdapter
    Public ODBC_DR As OdbcDataReader
End Class
