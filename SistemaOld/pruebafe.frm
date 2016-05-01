VERSION 5.00
Begin VB.Form Pruebafe 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Vencimiento 
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Text            =   "15/04/2011"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Fecha 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Text            =   "15/03/2011"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Numero 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "1"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Letra 
      Height          =   405
      Left            =   840
      TabIndex        =   2
      Text            =   "A"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Cae 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Pruebafe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WCuit As String

Private Sub Command1_Click()

    Dim WSAA As Object, WSFE As Object
    
    On Error GoTo ManejoError
    
    
    Stop
    
    If Trim(Cae.Text) <> "" Then
        Exit Sub
    End If
    
    
    ' Crear objeto interface Web Service Autenticación y Autorización
    Set WSAA = CreateObject("WSAA")
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA)
    tra = WSAA.CreateTRA()
    Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    Rem Path = CurDir() + "\"
    
    ZPath = "c:\salva\"
    ZNombre = "celugama"
    ZCuit = "30637671622"

        
    Rem ZPath = "c:\salva\"
    Rem ZNombre = "surfa"
    Rem ZCuit = "30549165083"
        
        
    ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
    Certificado = ZPath + ZNombre + ".crt" ' certificado de prueba
    ClavePrivada = ZPath + ZNombre + ".key" ' clave privada de prueba
    
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)
    Debug.Print cms
    
    ' Llamar al web service para autenticar:
    Rem ta = WSAA.CallWSAA(cms, "https://wsaahomo.afip.gov.ar/ws/services/LoginCms") ' Hologación
    ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms") ' Producción

    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
    Debug.Print ta
    Debug.Print "Token:", WSAA.Token
    Debug.Print "Sign:", WSAA.Sign
    
    ' Una vez obtenido, se puede usar el mismo token y sign por 6 horas
    ' (este período se puede cambiar)
    
    ' Crear objeto interface Web Service de Factura Electrónica
    Set WSFE = CreateObject("WSFE")
    ' Setear tocken y sing de autorización (pasos previos)
    WSFE.Token = WSAA.Token
    WSFE.Sign = WSAA.Sign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFE.Cuit = ZCuit
    
    ' Conectar al Servicio Web de Facturación
    Rem ok = WSFE.Conectar("https://wswhomo.afip.gov.ar/wsfe/service.asmx") ' homologación
    Rem ok = WSFE.Conectar("https://wsw.afip.gov.ar/wsfe/service.asmx") ' producción
    ok = WSFE.Conectar("https://servicios1.afip.gov.ar/wsfe/service.asmx") ' producción
    


    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFE.Dummy
    Debug.Print "appserver status", WSFE.AppServerStatus
    Debug.Print "dbserver status", WSFE.DbServerStatus
    Debug.Print "authserver status", WSFE.AuthServerStatus
    
    ' Recupera cantidad máxima de registros (opcional)
    qty = WSFE.RecuperarQty()
    
    ' Recupera último número de secuencia ID
    LastId = WSFE.UltNro()
    
    ' Recupero último número de comprobante para un punto de venta y tipo (opcional)
    
    ZZTipo = "03"
    
    If Letra.Text = "A" Then
        Select Case Val(ZZTipo)
            Case 3
                tipo_cbte = 1
                WTipo = "01"
            Case 4
                tipo_cbte = 2
                WTipo = "02"
            Case 5
                tipo_cbte = 3
                WTipo = "03"
            Case Else
        End Select
            Else
        Select Case Val(ZZTipo)
            Case 3
                tipo_cbte = 6
                WTipo = "06"
            Case 4
                tipo_cbte = 7
                WTipo = "07"
            Case 5
                tipo_cbte = 8
                WTipo = "08"
            Case Else
        End Select
    End If
    
    punto_vta = 3
    LastCBTE = WSFE.RecuperaLastCMP(punto_vta, tipo_cbte)
    
    ' Establezco los valores de la factura o lote a autorizar:
    
    WRazon = "GHISETTI CARMEN "
    WCuit = "27-03983626-8"
    Call Eval
    
    Rem WRazon = "Pellital S.A."
    Rem WCuit = "30-61052459-8"
    Rem Call Eval
        
        
    Rem Fecha = Format(Date, "yyyymmdd")
    
    id = LastId + 1
    presta_serv = 0
    
    If Len(WCuit) = 11 Then
        tipo_doc = 80
            Else
        tipo_doc = 96
    End If
    nro_doc = Left$(WCuit + Space$(11), 11)
    
    If Val(Numero.Text) <> LastCBTE + 1 Then
        m$ = "El numero de comprobante no es igua al correlativo indicado por la afip " + Str$(LastCBTE + 1)
        A% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
        Exit Sub
    End If
    
    cbt_desde = LastCBTE + 1
    cbt_hasta = LastCBTE + 1
    
    imp_total = 36.18
    imp_tot_conc = 0
    imp_neto = 29.9
    impto_liq = 6.28
    impto_liq_rni = 0
    imp_op_ex = 0
    
    ZZFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    fecha_cbte = ZZFecha
    
    ZZFecha = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
    fecha_venc_pago = ZZFecha
    
    
    ZZDesdeFecha = "0000000"
    ZZHastaFecha = "0000000"
    
    ' Fechas del período del servicio facturado (solo si presta_serv = 1)
    ZZFecha = Right$(ZZDesdeFecha, 4) + Mid$(ZZDesdeFecha, 4, 2) + Left$(ZZDesdeFecha, 2)
    Rem XXFecha = Format(ZZFecha, "yyyymmdd")
    fecha_serv_desde = ZZFecha
    
    
    
    ZZFecha = Right$(ZZHastaFecha, 4) + Mid$(ZZHastaFecha, 4, 2) + Left$(ZZHastaFecha, 2)
    Rem Fecha = Format(ZZFecha, "yyyymmdd")
    fecha_serv_hasta = ZZFecha
    
    
    fecha_serv_desde = ""
    fecha_serv_hasta = ""
    
    
    ' Llamo al WebService de Autorización para obtener el CAE
    Cae = WSFE.Aut(id, presta_serv, _
        tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
        impto_liq, impto_liq_rni, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta) ' si presta_serv = 0 no pasar estas fechas
    
    Debug.Print "Vencimiento ", WSFE.Vencimiento ' Fecha de vencimiento o vencimiento de la autorización
    Debug.Print "Resultado: ", WSFE.resultado ' A=Aceptado, R=Rechazado
    Debug.Print "Motivo de rechazo o advertencia", WSFE.Motivo ' 00= No hay error
    Debug.Print "Reprocesado?", WSFE.Reproceso ' S=Si, N=No
    
    ' Verifico que no haya rechazo o advertencia al generar el CAE
    If Cae = "" Then
        MsgBox "La página esta caida o la respuesta es inválida"
    ElseIf Cae = "NULL" Or WSFE.resultado <> "A" Then
        MsgBox "No se asignó CAE (Rechazado). Motivos: " & WSFE.Motivo, vbInformation + vbOKOnly
    ElseIf WSFE.Motivo <> "NULL" And WSFE.Motivo <> "00" Then
        MsgBox "Se asignó CAE pero con advertencias. Motivos: " & WSFE.Motivo, vbInformation + vbOKOnly
    End If
    
    ' Imprimo respuesta XML para depuración (errores de formato)
    Debug.Print WSFE.XmlResponse
    
    MsgBox "QTY: " & qty & vbCrLf & "LastId: " & LastId & vbCrLf & "LastCBTE:" & LastCBTE & vbCrLf & "CAE: " & Cae, vbInformation + vbOKOnly
    
    If WSFE.resultado = "A" Then
        Cae.Text = Cae
        Vencimiento.Text = WSFE.Vencimiento
    End If
    
    Exit Sub

ManejoError:
    ' Si hubo error:
    Debug.Print Err.Description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select

End Sub

Private Sub Eval()

    Es = WCuit

    x = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               x = x + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               x = x + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               x = x + Y: MinusOk = 0

        End If

    Next

    WCuit = x

End Sub



