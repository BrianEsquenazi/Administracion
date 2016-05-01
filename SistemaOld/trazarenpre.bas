Attribute VB_Name = "Module1"
' Ejemplo de Uso de Interface COM para servicio web PyAfipWs
' Trazabilidad de Precursores Quimicos RENPRE SEDRONAR INSSJP PAMI
' 2013 (C) Mariano Reingart <reingart@gmail.com>
' Licencia: GPLv3


Sub Main()

    Dim TrazaRenpre As Object, ok As Variant
    
    ' Crear la interfaz COM con el servicio web
    Set TrazaRenpre = CreateObject("TrazaRenpre")
    
    Debug.Print TrazaRenpre.Version, TrazaRenpre.InstallDir
    
    ' Establecer credenciales de seguridad
    TrazaRenpre.Username = "testwservice"
    TrazaRenpre.password = "testwservicepsw"
    
    
    
    
    ' Conectar al servidor (pruebas)
    ok = TrazaRenpre.Conectar()
    Debug.Print Err.Description
    Debug.Print TrazaRenpre.Excepcion
    Debug.Print TrazaRenpre.Traceback
    
    
    Debug.Print "Resultado:", TrazaRenpre.Resultado
    Debug.Print "CodigoTransaccion:", TrazaRenpre.CodigoTransaccion
    Debug.Print er
    
    
    ' datos de prueba
    Rem usuario = "9980342100037"
    Rem password = "Surfa0000"
    Rem gln_origen = 9992440700002#
    Rem gln_destino = 9980334210003#
    Rem f_operacion = "01/01/2012"
    Rem id_evento = 40                   ' 43: COMERCIALIZACION COMPRA, 44: COMERCIALIZACION VENTA
    Rem cod_producto = "00000000000024"  ' Acido Clorhidrico
    Rem n_cantidad = 200
    Rem n_documento_operacion = 1
    Rem m_entrega_parcial = ""
    Rem n_remito = 123
    Rem n_serie = 112
    
    Stop
    
    
GoTo da:
    
    
    ' datos de prueba de una importacion
    
    usuario = "9980342100037"
    password = "Surfa0000"
    gln_origen = "0300000000018"
    gln_destino = "9980342100037"
    f_operacion = "01/04/2014"
    id_evento = 45
    cod_producto = "00000000000024"  ' Acido Clorhidrico
    n_cantidad = 1500
    n_documento_operacion = 0
    m_entrega_parcial = ""
    
    n_remito = 0
    n_serie = 0
    n_frontera = 219
    n_tipodoc = 0
    n_tractor = ""
    n_Semi = ""
    n_serie = ""
    n_lote = ""
    n_despacho = "14001IC04017795V"
    n_permiso = ""
    
    n_djai = "13800DJAI080818R"
    n_certificado = ""
    n_tipodocii = ""
    n_nrodoc = 0
    n_calidad = 0
    n_Cufetransporte = ""
    
    
    
    ok = TrazaRenpre.SaveTransacciones( _
                         usuario, password, gln_origen, gln_destino, _
                         f_operacion = "01/01/2012", id_evento, cod_producto, n_cantidad, _
                         n_documento_operacion, m_entrega_parcial, n_remito, n_serie, _
                         n_frontera, n_tipodoc, n_tractor, n_Semi, _
                         n_serie, n_lote, n_despacho, n_permiso, _
                         n_djai, n_certificado, n_tipodocii, n_nrodoc _
                         )
    
    
    
    
    
    
    
da:
    
    
    ' datos de prueba de una fabricacion
    
    usuario = "9980342100037"
    password = "Surfa0000"
    gln_origen = "9980342100037"
    gln_destino = ""
    f_operacion = "01/04/2014"
    id_evento = 40
    cod_producto = "00000000000024"  ' Acido Clorhidrico
    n_cantidad = 300
    n_documento_operacion = ""
    m_entrega_parcial = ""
    
    n_remito = ""
    n_serie = 0
    
    ok = TrazaRenpre.SaveTransacciones( _
                         usuario, password, gln_origen, gln_destino, _
                         f_operacion, id_evento, cod_producto, n_cantidad, _
                         n_documento_operacion, m_entrega_parcial, n_remito, n_serie _
                         )
    
    
    
    Debug.Print Err.Description
    Debug.Print TrazaRenpre.Excepcion
    Debug.Print TrazaRenpre.Traceback
    
    
    
    ' Hubo error interno?
    If TrazaRenpre.Excepcion <> "" Then
        Debug.Print TrazaRenpre.Excepcion, TrazaRenpre.Traceback
        MsgBox TrazaRenpre.Traceback, vbCritical, "Excepcion:" & TrazaRenpre.Excepcion
    Else
        Debug.Print "Resultado:", TrazaRenpre.Resultado
        Debug.Print "CodigoTransaccion:", TrazaRenpre.CodigoTransaccion
        
        For Each er In TrazaRenpre.Errores
            Debug.Print er
            MsgBox er, vbExclamation, "Error en SendMedicamentos"
        Next
        
        MsgBox "Resultado: " & TrazaRenpre.Resultado & vbCrLf & _
                "CodigoTransaccion: " & TrazaRenpre.CodigoTransaccion, _
                vbInformation, "SaveTransacciones"
        
    End If
    
    ' Cancelo la transacción (anulación):
    codigo_transaccion = TrazaRenpre.CodigoTransaccion
    ok = TrazaRenpre.SendCancelacTransacc(usuario, password, codigo_transaccion)
    If ok Then
        Debug.Print "Resultado", TrazaRenpre.Resultado
        Debug.Print "CodigoTransaccion", TrazaRenpre.CodigoTransaccion
        MsgBox "Resultado: " & TrazaRenpre.Resultado & vbCrLf & _
                "CodigoTransaccion: " & TrazaRenpre.CodigoTransaccion, _
                vbInformation, "SendCancelacTransacc"
        For Each er In TrazaRenpre.Errores
            Debug.Print er
            MsgBox er, vbExclamation, "Error en SendCancelacTransacc"
        Next
    Else
        Debug.Print TrazaRenpre.XmlResponse
        MsgBox TrazaRenpre.Traceback, vbExclamation + vbCritical, "Excepcion en SendCancelacTransacc"
    End If
End Sub
