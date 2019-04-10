' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de interfaz "IService1" en el código y en el archivo de configuración a la vez.
<ServiceContract()>
Public Interface IService1


    <OperationContract()>
    Function LogIn(ByVal Usuario As String, ByVal Password As String) As CUsuario
    <OperationContract()>
    Function Obtener_Estado_de_Cuenta(ByVal Fecha_Inicial As Date, ByVal Fecha_Final As Date, ByVal Empleado As Integer) As List(Of CEstadoCuenta)
    <OperationContract()>
    Function Obtener_reporte_semanal(ByVal empleado As Integer) As List(Of CEstadoCuenta)
    <OperationContract()>
    Function Obtener_datos_nom_fracc() As CDatosFracc()
    <OperationContract()>
    Function Obtener_Smza(ByVal id_cve_fracc As String) As List(Of String)
    <OperationContract()>
    Function Obtener_Tipos_de_Credito() As CCreditos()
    <OperationContract()>
    Function Obtener_Tipos_de_CreditoCC(ByVal CC As String, ByVal SM As String) As CCreditos()
    <OperationContract()>
    Function Obtener_Datos_Contrato_Nuevo(ByVal CC As String, ByVal SM As String, ByVal TC As Integer, ByVal INFONAVIT As Integer, ByVal FOVISSSTE As Integer, ByVal ISSEG As Integer, Optional ByVal Empresa As Integer = 11) As CDatosContratoNuevo
    <OperationContract()>
    Function Verifica_Conectividad() As Boolean
    <OperationContract()>
    Function Obtener_Ventas_Por_Semana_Entre_Fechas(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As CDatosVentaSemanal()
    <OperationContract()>
    Function Obtener_Total_Vendido_Ubicado(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As Integer
    <OperationContract()>
    Function Obtener_Total_habitabilidad(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As Integer
    <OperationContract()>
    Function Obtener_Habitabilidad_Por_Semana_Entre_Fechas(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As CDatosVentaSemanal()
    <OperationContract()>
    Function Obtener_Ventas_Por_Semana_Entre_Fechas_Barras(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As CDatosVentaSemanal()
    <OperationContract()>
    Function Obtener_Firmas_X_Semana_Entre_Fechas(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As CDatosVentaSemanal()
    <OperationContract()>
    Function Obtener_Total_Cancelados_o_SinUbicacion(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date) As Integer
    <OperationContract()>
    Function MyDatos(empleado As Integer, FNacimiento As Date, SSexo As String, SNacionalidad As String, Tnacionalidad As String,
                     SCivil As String, tbCiudad As String, tbEstado As String, tbDir As String, tbCelular As String,
                     tbTel As String, tbEmail As String, tbRefNom1 As String, tbRefParentesco1 As String,
                     tbRefTel1 As String, tbRefNom2 As String, tbRefParentesco2 As String, tbRefTel2 As String,
                     tbRFC As String, tbCURP As String, tbIFE As String, tbCIFE As String, tbManejo As String,
                     FVenceManejo As Date, tbNSS As String, cbPrimaria As String, cbSecundaria As String,
                     cbPreparatoria As String, cbLicenciatura As String, tbLicName As String, cbMaestria As String,
                     tbMasName As String) As Boolean
    <OperationContract()>
    Function Obtener_DatosDetalle_Empleado(ByVal Empleado As Integer) As CDatosAsesorDetalle
    <OperationContract()>
    Function Comprobar_Notificaciones(ByVal Empleado As Integer) As CNotificaciones
    <OperationContract()>
    Function Cambiar_a_Visto_notificacion(ByVal id_notificacion As Integer) As Boolean
    <OperationContract()>
    Function Obtener_Clientes_Activos(ByVal Empleado As Integer) As CClientesActivos()
    <OperationContract()>
    Function Obtener_ultimas_Notificaciones(ByVal Empleado As Integer) As CUltimasNotificaciones()
    <OperationContract()>
    Function Obtener_Asesores_Activos() As CAsesoresActivos()
    <OperationContract()>
    Function Reporte_Concepto_Total(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date, ByVal Etapa As Integer) As List(Of CReporteMontos)
    <OperationContract()>
    Function Obtener_Reporte_concepto(ByVal Fecha_inicio As Date, ByVal Fecha_Final As Date, ByVal Etapa As Integer) As CreporteConepto()
    <OperationContract()>
    Function buscarClienteAutoCompletar(ByVal Busqueda As String) As CClientesBusqueda()
    <OperationContract()>
    Function Obtener_Datos_Generales_Cliente(ByVal numcte As Integer) As CGeneralesCliente
    <OperationContract()>
    Function Obtener_Desgloce_Etapas_Cliente(ByVal numcte As Integer) As CEtapasCliente()
    <OperationContract()>
    Function CancelacionCliente(ByVal Numcte As Integer) As CDatosCancelacion
    <OperationContract()>
    Function Inserta_usuarios(ByVal id_tipo As Integer, ByVal nivel As Integer, ByVal Desc_Nombre As String) As Boolean
    <OperationContract()>
    Function Inserta_pro_contratos_nuevo(ByVal CC As String, ByVal SM As String, ByVal TC As Integer, ByVal INFONAVIT As String, ByVal FOVISSSTE As String, ByVal ISSEG As String, ByVal Fecha_DTU As Date, ByVal CPenalizaPrevio As Decimal, ByVal CEnganche As Decimal, ByVal CPenalizaIngresado As Decimal, ByVal FormatoAdicional2 As String, ByVal FormatoAdicional As String, ByVal PrecioCasa As Integer, ByVal PrecioAdicional As Integer, ByVal Mtr_Casa As String, ByVal Activo As String, ByVal Bono As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_pro_contratos_nuevo(ByVal id_contrato As Integer, ByVal CC As String, ByVal SM As String, ByVal TC As Integer, ByVal INFONAVIT As String, ByVal FOVISSSTE As String, ByVal ISSEG As String, ByVal Fecha_DTU As Date, ByVal CPenalizaPrevio As Decimal, ByVal CEnganche As Decimal, ByVal CPenalizaIngresado As Decimal, ByVal FormatoAdicional2 As String, ByVal FormatoAdicional As String, ByVal PrecioCasa As Integer, ByVal PrecioAdicional As Integer, ByVal Mtr_Casa As String, ByVal Activo As String, ByVal Bono As Integer) As Boolean
    <OperationContract()>
    Function Listar_ContratoDatos() As List(Of CDatosContrato)
    <OperationContract()>
    Function Elimina_pro_contratos_nuevo(ByVal id_contrato As Integer) As Boolean
    <OperationContract()>
    Function Inserta_promociones(ByVal CC As String, ByVal SM As String, ByVal Precio As String, ByVal TextoCombo As String, ByVal TextoContrato As String) As Boolean
    <OperationContract()>
    Function Actualiza_promociones(ByVal id_promocion As Integer, CC As String, ByVal SM As String, ByVal Precio As String, ByVal TextoCombo As String, ByVal TextoContrato As String) As Boolean
    <OperationContract()>
    Function Elimina_promociones(ByVal id_promocion As Integer) As Boolean
    <OperationContract()>
    Function Listar_Equipamientos() As List(Of CEquipamiento)
    <OperationContract()>
    Function Obtener_Equipamientos(ByVal CC As String, ByVal SM As String) As List(Of CEquipamiento)
    <OperationContract()>
    Function Obtener_Equipamiento(ByVal id_promocion As Integer) As CEquipamiento
    <OperationContract()>
    Function Obtener_Datos_Grid_Gerardo() As List(Of CReporteGerardoGrid)
    <OperationContract()>
    Function Obtener_Detalles_comisionID(ByVal id_comision As Integer) As CComision
    <OperationContract()>
    Function Modifica_Comision(ByVal id_comision As Integer, ByVal Cantidad_Pagada_Total As Integer, ByVal Observaciones As String) As Boolean
    <OperationContract()>
    Function Obtener_Nombrecompleto_lider(ByVal desc_valor As String) As CNombreLiderGerente
    <OperationContract()>
    Function Resumen_fraccionamiento() As List(Of CResumenFracc)
    <OperationContract()>
    Function Resumen_asesor_comisiones() As List(Of CResumenAsesor)
    <OperationContract()>
    Function Obtener_pro_contratos_nuevo(ByVal id_contrato As Integer) As CContratos
    <OperationContract()>
    Function Obtener_Valores_Comisiones() As List(Of CValoresComisiones)
    <OperationContract()>
    Function Actualiza_comisiones_porcentajes(ByVal id_valor As Integer, ByVal Porcentaje As String) As Boolean
    <OperationContract()>
    Function Obtener_comisiones_porcentajes(ByVal id_porcentaje As Integer) As List(Of CValoresComisiones)
    <OperationContract()>
    Function Obtener_esquemas_enkontrol() As List(Of CEsquemas)
    <OperationContract()>
    Function Inserta_comisiones_clasificaciones(ByVal DescripcionClasificacion As String) As Boolean
    <OperationContract()>
    Function Elimina_comisiones_clasificaciones(ByVal id_clasificacion As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_comisiones_clasificaciones(ByVal id_clasificacion As Integer, ByVal DescripcionClasificacion As String) As Boolean
    <OperationContract()>
    Function Obtener_comisiones_clasificaciones() As List(Of CClasificaciones)
    <OperationContract()>
    Function Obtener_comisiones_clasificacion(ByVal id_clasificacion As Integer) As CClasificaciones
    <OperationContract()>
    Function Inserta_comisiones_esquemas(ByVal id_esquema As Integer, ByVal id_clasificacion As Integer) As Boolean
    <OperationContract()>
    Function Elimina_comisiones_esquemas(ByVal id_registro As Integer) As Boolean
    <OperationContract()>
    Function Obtener_comisiones_esquemas() As List(Of CEsquemaRegla)
    <OperationContract()>
    Function Obtener_comisiones_esquema(ByVal id_clasificacion As Integer) As List(Of CEsquemaRegla)
    <OperationContract()>
    Function Obtener_Reglas_Comisiones() As List(Of CReglasComisiones)
    <OperationContract()>
    Function Inserta_comisiones_reglas(ByVal empleado As Integer, ByVal id_clasificacion As Integer, ByVal etapaFinal As Integer, ByVal Cantidad_PagoUnitario As Integer) As Boolean
    <OperationContract()>
    Function Elimina_comisiones_reglas(ByVal id_regla As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_admin(ByVal id_admin As Integer, ByVal Empleado As Integer, ByVal NombreCompleto As String, ByVal Cantidad As Integer) As Boolean
    <OperationContract()>
    Function Obtener_admin() As List(Of CAdministrativos)
    <OperationContract()>
    Function Inserta_programacionentrega(ByVal numcte As Integer, ByVal Hora As TimeSpan, ByVal FECHA As Date) As Boolean
    <OperationContract()>
    Function Elimina_programacionentrega(ByVal id_programacion As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_programacionentrega(ByVal id_programacion As Integer, ByVal numcte As Integer, ByVal NombreCliente As String, ByVal Hora As TimeSpan, ByVal FECHA As Date) As Boolean
    <OperationContract()>
    Function Obtener_programacionentrega() As List(Of CProgramacionEntrega)
    <OperationContract()>
    Function Obtener_Nombre_Cliente(ByVal Numcte As Integer) As String
    <OperationContract()>
    Function Cambia_contratos(ByVal id_contrato As Integer, ByVal Activo As Integer) As Boolean
    <OperationContract()>
    Function Inserta_comisionesgerencia(ByVal CC As String, ByVal Cantidad As Integer) As Boolean
    <OperationContract()>
    Function Elimina_comisionesgerencia(ByVal CC As String) As Boolean
    <OperationContract()>
    Function Elimina_Comision(ByVal id_comisione As Integer) As Boolean
    <OperationContract()>
    Function Obtener_comisionesgerencia() As List(Of CPagosGerencia)
    <OperationContract()>
    Function Obtener_promocione(ByVal id_promocion As Integer) As CPromocionesContrato
    <OperationContract()>
    Function Reporte_Penalizacion(ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CPagoyPenalizacion)
    '<OperationContract()>
    'Function Obtener_Reporte_Penaliza(ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CReportePenalizacion)
    <OperationContract()>
    Function Obtener_comisiones_cliente(ByVal numcte As String) As List(Of CComisionesCliente)
    <OperationContract()>
    Function Obtener_datos_cancelaciones() As List(Of CCancelacionesDetalles)
    <OperationContract()>
    Function Obtener_programacionEntregaAdm(ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CProgramacionEntregaAdm)
    <OperationContract()>
    Function Obtener_datos_cliente_contrato_fovissste(ByVal numcte As Integer) As CDatosClienteContratoFovissste
    <OperationContract()>
    Function Obtener_cuentaDepodito_Cte(ByVal Numcte As Integer) As Integer
    <OperationContract()>
    Function Obtener_Credito_Porcentaje(ByVal Numcte As Integer) As String
    <OperationContract()>
    Function Obtener_promociones(ByVal CC As String, ByVal SM As String) As List(Of CPromocionesContrato)
    <OperationContract()>
    Function ObtenerDSEnkontrol(ByVal SQL As String) As DataTable
    <OperationContract()>
    Function Obtener_loteids() As String
    <OperationContract()>
    Function ObtenerDSSAC(ByVal SQL As String) As DataSet
    <OperationContract()>
    Function Inserta_reportes_check(ByVal id_categoria As Integer, ByVal id_subcategoria As Integer, ByVal id_subsubcategoria As Integer, ByVal id_subsubsubcategoria As Integer, ByVal id_subsubsubsubcategoria As Integer, ByVal NUMCTE As String, ByVal Observacioens As String, ByVal fotografia As String) As Boolean
    <OperationContract()>
    Function ObtenerTerreno(ByVal CC As String) As Boolean
    <OperationContract()>
    Function Obtener_plazosTerrenos() As List(Of CPlazosTerreno)
    <OperationContract()>
    Function Obtener_plazoTerreno(ByVal id_plazo As Integer) As CPlazosTerreno
    <OperationContract()>
    Function ObtenerPCRU(ByVal smza As String) As String
    <OperationContract()>
    Function Obtener_limite_bonoContrato(ByVal CC As String) As Integer

#Region "BI"
    <OperationContract()>
    Function Obtener_HabitabilidadLM_BI(ByVal Fecha_inicio As Date) As List(Of CGraficoBI)
    <OperationContract()>
    Function Obtener_CanceladosLM_BI(ByVal Fecha_inicio As Date) As List(Of CGraficoBI)
    <OperationContract()>
    Function Obtener_EscrituradosLM_BI(ByVal Fecha_inicio As Date) As List(Of CGraficoBI)
    <OperationContract()>
    Function Obtener_VentasLM_BI(ByVal Fecha_inicio As Date) As List(Of CGraficoBI)
    <OperationContract()>
    Function Obtener_EscrituradosAcumuladoLM_BI(ByVal Fecha_inicio As Date) As List(Of CGraficoBI)
    <OperationContract()>
    Function Obtener_Ingresados_por_semana(ByVal Fecha_Inicio As Date, ByVal Fecha_Final As Date) As List(Of CGraficoBI)
#End Region
    ' TODO: agregue aquí sus operaciones de servicio

End Interface

' Utilice un contrato de datos, como se ilustra en el ejemplo siguiente, para agregar tipos compuestos a las operaciones de servicio.
<DataContract()>
Public Class CPlazosTerreno
    <DataMember()>
    Public id_plazo As Integer
    <DataMember()>
    Public plazo As Integer
    <DataMember()>
    Public precioMetro As Decimal
End Class

<DataContract()>
Public Class CDatosClienteContratoFovissste
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public Ciudad As String
    <DataMember()>
    Public Fec_Nac As Date
    <DataMember()>
    Public Edo_civil As String
    <DataMember()>
    Public puesto As String
    <DataMember()>
    Public Dir_Casa As String
    <DataMember()>
    Public RFC_cte As String
    <DataMember()>
    Public id_num_relacion As Decimal
    <DataMember()>
    Public valor_total As Decimal
    <DataMember()>
    Public id_num_lote As String
    <DataMember()>
    Public id_num_interior As String
    <DataMember()>
    Public id_num_mza As String
    <DataMember()>
    Public nom_fracc As String
    <DataMember()>
    Public Desc_Colindancia As String
    <DataMember()>
    Public Cant_superficie As Decimal
End Class

<DataContract()>
Public Class CProgramacionEntregaAdm
    <DataMember()>
    Public Nom_fracc As String
    <DataMember()>
    Public NumeroCliente As Decimal
    <DataMember()>
    Public programacionFirma As Date
    <DataMember()>
    Public PreEntrega As Date
    <DataMember()>
    Public EntregaCliente As Date
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public id_num_smza As String
    <DataMember()>
    Public id_num_mza As String
    <DataMember()>
    Public id_num_interior As String
    <DataMember()>
    Public dir_casa As String
    <DataMember()>
    Public ObservacionProgramacion As String
    <DataMember()>
    Public Reprogramado As Boolean = False
    <DataMember()>
    Public LlamadaDeCalidad As String
    <DataMember()>
    Public RevisioonDeDivivenda As String
    <DataMember()>
    Public id_num_lote As String
End Class

<DataContract()>
Public Class CCancelacionesDetalles
    <DataMember()>
    Public id_cancelacion As Integer
    <DataMember()>
    Public CC As String
    <DataMember()>
    Public Fraccionamiento As String
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public AnticipoDeVenta As Integer
    <DataMember()>
    Public Empleado As Integer
    <DataMember()>
    Public NombreAsesor As String
    <DataMember()>
    Public NombreLider As String
    <DataMember()>
    Public NombreGerente As String
    <DataMember()>
    Public PagoAsesor As Integer
    <DataMember()>
    Public PagoLider As Integer
    <DataMember()>
    Public PagoGerente As Integer
    <DataMember()>
    Public PagoAdministrativo As Integer
    <DataMember()>
    Public FechaPago As Date
    <DataMember()>
    Public Numcte As Integer
    <DataMember()>
    Public P_Cliente As Integer
    <DataMember()>
    Public P_Asesor As Integer
    <DataMember()>
    Public P_Lider As Integer
    <DataMember()>
    Public P_Gerente As Integer
    <DataMember()>
    Public P_Administrativo As Integer
    <DataMember()>
    Public FechaCancelacion As Date
End Class

<DataContract()>
Public Class CComisionesCliente
    <DataMember()>
    Public id_comision As Integer
    <DataMember()>
    Public Pagado_A As String
    <DataMember()>
    Public id_periodo As Integer
    <DataMember()>
    Public Fecha_Pago As Date
    <DataMember()>
    Public Cantidad_Pagada_Total As Integer
    <DataMember()>
    Public Tipo_pago As String
    <DataMember()>
    Public Empleado As Integer
    <DataMember()>
    Public Lider As String
    <DataMember()>
    Public Gerente As String
    <DataMember()>
    Public NombreCompleto As String
    <DataMember()>
    Public Observaciones As String
End Class

<DataContract()>
Public Class CPagoyPenalizacion
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public CC As String
    <DataMember()>
    Public PagadoTotal As Integer
    <DataMember()>
    Public PenalizadoTotal As Integer
    <DataMember()>
    Public Empleado As Integer
    <DataMember()>
    Public NombreEmpleado As String
    <DataMember()>
    Public FechaPago As Date
    <DataMember()>
    Public FechaPenalizado As Date
End Class

<DataContract()>
Public Class CPagosGerencia
    <DataMember()>
    Public CC As String
    <DataMember()>
    Public Cantidad As Integer
End Class

<DataContract()>
Public Class CProgramacionEntrega
    <DataMember()>
    Public id_programacion As Integer
    <DataMember()>
    Public numcte As Integer
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public Hora As TimeSpan
    <DataMember()>
    Public CC As String
    <DataMember()>
    Public FECHA As Date
End Class

<DataContract()>
Public Class CReportePenalizacion
    <DataMember()>
    Public NombreFraccionamiento As String
    <DataMember()>
    Public Importe As Integer

End Class

<DataContract()>
Public Class CAdministrativos
    <DataMember()>
    Public id_admin As Integer
    <DataMember()>
    Public Empleado As Integer
    <DataMember()>
    Public NombreCompleto As String
    <DataMember()>
    Public Cantidad As Integer
End Class

<DataContract()>
Public Class CReglasComisiones
    <DataMember()>
    Public id_regla As Integer
    <DataMember()>
    Public empleado As Integer
    <DataMember()>
    Public id_clasificacion As Integer
    <DataMember()>
    Public etapaFinal As Integer
    <DataMember()>
    Public Cantidad_PagoUnitario As Integer
End Class

<DataContract()>
Public Class CEsquemaRegla
    <DataMember()>
    Public id_registro As Integer
    <DataMember()>
    Public id_esquema As Integer
    <DataMember()>
    Public id_clasificacion As Integer
End Class

<DataContract()>
Public Class CClasificaciones
    <DataMember()>
    Public id_clasificacion As Integer
    <DataMember()>
    Public DescripcionClasificacion As String
End Class

<DataContract()>
Public Class CEsquemas
    <DataMember()>
    Public id_num_relacion As Decimal
    <DataMember()>
    Public Nom_Financinst As String
End Class

<DataContract()>
Public Class CValoresComisiones
    <DataMember()>
    Public id_porcentaje As Integer
    <DataMember()>
    Public Descripcion_porcentaje As String
    <DataMember()>
    Public Porcentaje As String
End Class

<DataContract()>
Public Class CContratos
    <DataMember()>
    Public id_contrato As Integer
    <DataMember()>
    Public CC As String
    <DataMember()>
    Public SM As String
    <DataMember()>
    Public TC As Integer
    <DataMember()>
    Public INFONAVIT As String
    <DataMember()>
    Public FOVISSSTE As String
    <DataMember()>
    Public ISSEG As String
    <DataMember()>
    Public Fecha_DTU As Date
    <DataMember()>
    Public CPenalizaPrevio As Decimal
    <DataMember()>
    Public CEnganche As Decimal
    <DataMember()>
    Public CPenalizaIngresado As Decimal
    <DataMember()>
    Public FormatoAdicional2 As String
    <DataMember()>
    Public FormatoAdicional As String
    <DataMember()>
    Public PrecioCasa As Integer
    <DataMember()>
    Public PrecioAdicional As Integer
    <DataMember()>
    Public Mtr_Casa As String
    <DataMember()>
    Public Activo As String
    <DataMember()>
    Public Bono As Integer
End Class

<DataContract()>
Public Class CResumenFracc
    <DataMember()>
    Public NombreFraccionamiento As String
    <DataMember()>
    Public Importe As Integer
    <DataMember()>
    Public TipoPago As String
End Class

<DataContract()>
Public Class CResumenAsesor
    <DataMember()>
    Public NombreAsesor As String
    <DataMember()>
    Public Importe As Integer
    <DataMember()>
    Public TipoPago As String
End Class

<DataContract()>
Public Class CNombreLiderGerente
    <DataMember()>
    Public id_registro As Integer
    <DataMember()>
    Public empleado As Integer
    <DataMember()>
    Public desc_valor As String
    <DataMember()>
    Public NombreCompleto As String
End Class

<DataContract()>
Public Class CComision
    <DataMember()>
    Public id_comision As Integer
    <DataMember()>
    Public numcte As Integer
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public Empleado As Integer
    <DataMember()>
    Public Lider As String
    <DataMember()>
    Public Gerente As String
    <DataMember()>
    Public Administrativo As String
    <DataMember()>
    Public Cantidad_Pagada_Total As Integer
    <DataMember()>
    Public Observacion As String
End Class

<DataContract()>
Public Class CReporteGerardoGrid
    <DataMember()>
    Public id_comision As Integer
    <DataMember()>
    Public Numcte As Integer
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public NombreAgente As String
    <DataMember()>
    Public TipoPago As String
    <DataMember()>
    Public Importe As Integer
    <DataMember()>
    Public Observacion As String
    <DataMember()>
    Public Fracc As String
    <DataMember()>
    Public SMZA As String
    <DataMember()>
    Public Domicilio As String
    <DataMember()>
    Public Manzana As String
    <DataMember()>
    Public Lote As String
End Class

<DataContract()>
Public Class CEquipamiento
    <DataMember()>
    Public id_promocion As Integer
    <DataMember()>
    Public CC As String
    <DataMember()>
    Public SM As String
    <DataMember()>
    Public Precio As Double
    <DataMember()>
    Public TextoCombo As String
    <DataMember()>
    Public TextoContrato As String
End Class

<DataContract()>
Public Class CPromocionesContrato
    <DataMember()>
    Public id_promocion As Integer
    <DataMember()>
    Public CC As String
    <DataMember()>
    Public SM As String
    <DataMember()>
    Public textoCombo As String
    <DataMember()>
    Public textContrato As String
    <DataMember()>
    Public Costo As Decimal
    <DataMember()>
    Public Activo As String
End Class

<DataContract()>
Public Class CDatosContrato
    <DataMember()>
    Public id_contrato As Integer
    <DataMember()>
    Public CC As String
    <DataMember()>
    Public SM As String
    <DataMember()>
    Public TC As Integer
    <DataMember()>
    Public INFONAVIT As String
    <DataMember()>
    Public FOVISSSTE As String
    <DataMember()>
    Public ISSEG As String
    <DataMember()>
    Public Fecha_DTU As Date
    <DataMember()>
    Public CPenalizaPrevio As Decimal
    <DataMember()>
    Public CEnganche As Decimal
    <DataMember()>
    Public CPenalizaIngresado As Decimal
    <DataMember()>
    Public FormatoAdicional2 As String
    <DataMember()>
    Public FormatoAdicional As String
    <DataMember()>
    Public PrecioCasa As Integer
    <DataMember()>
    Public PrecioAdicional As Integer
    <DataMember()>
    Public Mtr_Casa As String
    <DataMember()>
    Public Activo As String
    <DataMember()>
    Public Bono As Integer
End Class

<DataContract()>
Public Class CUsuariosProConsul
    <DataMember()>
    Public Empleado As Integer
    <DataMember()>
    Public id_tipo As Integer
    <DataMember()>
    Public nivel As Integer
    <DataMember()>
    Public Desc_Nombre As String
End Class

<DataContract()>
Public Class CDatosCancelacion
    <DataMember()>
    Public Folio As Integer = 0
    <DataMember()>
    Public Numcte As Integer = 0
    <DataMember()>
    Public NombreCliente As String = "-"
    <DataMember()>
    Public Frente As String = "-"
    <DataMember()>
    Public CC As String = "-"
    <DataMember()>
    Public CantidadDeDevolucion As Integer = 0
    <DataMember()>
    Public CantidadDevolucionPesos As String = "-"
    <DataMember()>
    Public Motivo As String = "-"
    <DataMember()>
    Public Penaliza As Boolean
    <DataMember()>
    Public RazonNoPenaliza As String = "-"
    <DataMember()>
    Public PenalizacionCliente As Integer = 0
    <DataMember()>
    Public Empleado As Integer = 0
    <DataMember()>
    Public NombreEmpleado As String = "-"
    <DataMember()>
    Public FechaPago As Date = New Date
    <DataMember()>
    Public CantidadEmpleado As Integer = 0
    <DataMember()>
    Public CantidadLider As Integer = 0
    <DataMember()>
    Public PenalizacionAsesor As Integer = 0
    <DataMember()>
    Public NumLider As Integer = 0
    <DataMember()>
    Public NombreLider As String = "-"
    <DataMember()>
    Public PenalizacionLider As Integer = 0
    <DataMember()>
    Public NumGer1 As Integer = 0
    <DataMember()>
    Public NombreGerente1 As String = "-"
    <DataMember()>
    Public CantidadGerente1 As Integer = 0
    <DataMember()>
    Public PenalizacionGerente1 As Integer = 0
    <DataMember()>
    Public PenalizaVentas As Integer = 0
    <DataMember()>
    Public NumGer2 As Integer = 0
    <DataMember()>
    Public NombreGerente2 As String = "-"
    <DataMember()>
    Public CantidadGerente2 As Integer = 0
    <DataMember()>
    Public PenalizacionGerente2 As Integer = 0

End Class

<DataContract()>
Public Class CEtapasCliente
    <DataMember()>
    Public Nom_etapa As String
    <DataMember()>
    Public id_num_etapa As Decimal
    <DataMember()>
    Public Fec_inicio As Date
    <DataMember()>
    Public Fec_Liberacion As Date
    <DataMember()>
    Public Observaciones As String
End Class

<DataContract()>
Public Class CGeneralesCliente
    <DataMember()>
    Public numcte As Decimal
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public empleado As Decimal
    <DataMember()>
    Public NombreEmpleado As String
    <DataMember()>
    Public LIDER As String
    <DataMember()>
    Public EtapaActual As Decimal
    <DataMember()>
    Public status_cte As String
    <DataMember()>
    Public id_cve_fracc As String
    <DataMember()>
    Public lote_id As Decimal
    <DataMember()>
    Public id_num_mza As String
    <DataMember()>
    Public id_num_lote As String
    <DataMember()>
    Public id_num_interior As String
    <DataMember()>
    Public dir_casa As String
    <DataMember()>
    Public Valor_credito As Decimal
    <DataMember()>
    Public Valor_total As Decimal
End Class

<DataContract()>
Public Class CClientesBusqueda
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public Numcte As Integer
End Class

<DataContract()>
Public Class CreporteConepto
    <DataMember()>
    Public numcte As Decimal
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public id_cve_fracc As String
    <DataMember()>
    Public id_num_smza As String
    <DataMember()>
    Public id_num_mza As String
    <DataMember()>
    Public id_num_lote As String
    <DataMember()>
    Public Cant_mts_excedente As Decimal
    <DataMember()>
    Public Concepto As String
    <DataMember()>
    Public Monto As Decimal
End Class

<DataContract()>
Public Class CReporteMontos
    <DataMember()>
    Public Numcte As Integer
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public CC As String
    <DataMember()>
    Public SM As String
    <DataMember()>
    Public Mza As String
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public Movimientos As List(Of MovimientosCliente)
    <DataMember()>
    Public TerrenoExcedente As Double
End Class

Public Class MovimientosCliente
    Public Concepto As String
    Public Monto As Double
End Class

<DataContract()>
Public Class CAsesoresActivos
    <DataMember()>
    Public Empleado As Decimal
    <DataMember()>
    Public Nom_Empleado As String
    <DataMember()>
    Public Ap_Paterno_Empleado As String
    <DataMember()>
    Public Ap_Materno_Empleado As String
    <DataMember()>
    Public Lider As String
End Class

<DataContract()>
Public Class CUltimasNotificaciones
    <DataMember()>
    Public Mensaje As String
    <DataMember()>
    Public Prioridad As Integer
    <DataMember()>
    Public empleado As Integer
    <DataMember()>
    Public FechaUltima As Date
    <DataMember()>
    Public id_notificacion As Integer
    <DataMember()>
    Public Visto As String
End Class

<DataContract()>
Public Class CClientesActivos
    <DataMember()>
    Public numcte As Decimal
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public lote_id As Integer
    <DataMember()>
    Public id_num_mza As String
    <DataMember()>
    Public CC As String
    <DataMember()>
    Public Fracc As String
    <DataMember()>
    Public DirCasa As String
    <DataMember()>
    Public id_num_etapa As Integer
    <DataMember()>
    Public nom_etapa As String
    <DataMember()>
    Public Valor_credito As Integer
    <DataMember()>
    Public Valor_Total As Integer
    <DataMember()>
    Public NumeroOficial As String
End Class

<DataContract()>
Public Class CNotificaciones
    <DataMember()>
    Public id_notificacion As Integer
    <DataMember()>
    Public Mensaje As String
End Class

#Region "BI"
<DataContract()>
Public Class CGraficoBI
    <DataMember()>
    Public NSemana As Integer
    <DataMember()>
    Public CantidadVentas As Integer
End Class
#End Region

<DataContract()>
Public Class CDatosAsesorDetalle
    <DataMember()>
    Public id_datos As Integer = 0
    <DataMember()>
    Public empleado As Integer = 0
    <DataMember()>
    Public fecha_nacimiento As Date = New Date
    <DataMember()>
    Public sexo As String = ""
    <DataMember()>
    Public nacionalidad As String = ""
    <DataMember()>
    Public estado_civil As String = ""
    <DataMember()>
    Public ciudad As String = ""
    <DataMember()>
    Public estado As String = ""
    <DataMember()>
    Public domicilio As String = ""
    <DataMember()>
    Public celular As String = ""
    <DataMember()>
    Public telfijo As String = ""
    <DataMember()>
    Public email As String = ""
    <DataMember()>
    Public ref1nombre As String = ""
    <DataMember()>
    Public red1parentesco As String = ""
    <DataMember()>
    Public red1tel As String = ""
    <DataMember()>
    Public red2nombre As String = ""
    <DataMember()>
    Public red2parentesco As String = ""
    <DataMember()>
    Public red2tel As String = ""
    <DataMember()>
    Public rfc As String = ""
    <DataMember()>
    Public curp As String = ""
    <DataMember()>
    Public nife As String = ""
    <DataMember()>
    Public claveelector As String = ""
    <DataMember()>
    Public licmanejo As String = ""
    <DataMember()>
    Public fechavence As Date = New Date
    <DataMember()>
    Public nss As String = ""
    <DataMember()>
    Public primaria As String = "False"
    <DataMember()>
    Public secundaria As String = "False"
    <DataMember()>
    Public preparatoria As String = "False"
    <DataMember()>
    Public licenciatura As String = "False"
    <DataMember()>
    Public nomlicenciatura As String = ""
    <DataMember()>
    Public maestria As String = "False"
    <DataMember()>
    Public nommaestria As String = ""
End Class

<DataContract()>
Public Class CDatosVentaSemanal
    <DataMember()>
    Public NSemana As Integer
    <DataMember()>
    Public CantidadVentas As Integer
End Class

<DataContract()>
Public Class CDatosFracc
    <DataMember()>
    Property id_cve_fracc As String
    <DataMember()>
    Property Nom_Fracc As String
End Class

<DataContract()>
Public Class CUsuario
    <DataMember()>
    Public Property Empleado As Integer
    <DataMember()>
    Public Property Nombre_Usuario() As String
    <DataMember()>
    Public Property Tipo() As String
    <DataMember()>
    Public Property Desc_Valor As String
    <DataMember()>
    Public Property Nivel() As Integer

End Class

<DataContract()>
Public Class CEstadoCuenta
    <DataMember()>
    Public Property Numcte As Integer
    <DataMember()>
    Public Property NombreCliente As String
    <DataMember()>
    Public Property Fecha_pago As Date
    <DataMember()>
    Public Property TipoPago As String
    <DataMember()>
    Public Property Observaciones As String
    <DataMember()>
    Public Property Cantidad_Pagada As Integer
End Class

<DataContract()>
Public Class CCreditos
    <DataMember()>
    Public TC As Integer
    <DataMember()>
    Public Abreviatura As String
    <DataMember()>
    Public NombreCompleto As String
End Class

<DataContract()>
Public Class CDatosContratoNuevo
    <DataMember()>
    Public TC_Abreviatura As String
    <DataMember()>
    Public Nombre_CC As String
    <DataMember()>
    Public TC_Nombre As String
    <DataMember()>
    Public Modelo_casa As String
    <DataMember()>
    Public Formato_adicional As String
    <DataMember()>
    Public Formato_adicional2 As String
    <DataMember()>
    Public Superficie As Double
    <DataMember()>
    Public Mtrs_Construccion As Double
    <DataMember()>
    Public Fecha_DTU As Date
    <DataMember()>
    Public Mtrs_Casa As Double
    <DataMember()>
    Public Precio_Casa As Integer
    <DataMember()>
    Public Precio_Adicional As Integer
    <DataMember()>
    Public Pen_Previo As Integer
    <DataMember()>
    Public Pen_Final As Integer
    <DataMember()>
    Public Precio_Total As String
    <DataMember()>
    Public Cantidad_Enganche As Double
    <DataMember()>
    Public Bono As Integer
End Class