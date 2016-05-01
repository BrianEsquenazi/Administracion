VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Proveedores"
   ClientHeight    =   7875
   ClientLeft      =   1635
   ClientTop       =   750
   ClientWidth     =   8640
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   8640
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "excel"
      Height          =   555
      Left            =   3720
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "recibo"
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "provi"
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Cambio 
      Caption         =   "Cambio de Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Menu Maestros 
      Caption         =   "Maestros"
      Begin VB.Menu Cuenta 
         Caption         =   "Ingreso de Cuentas Contables"
      End
      Begin VB.Menu Porveedores 
         Caption         =   "Ingreso de Proveedores"
      End
      Begin VB.Menu Bancos 
         Caption         =   "Ingresos de Bancos"
      End
      Begin VB.Menu CambioAdm 
         Caption         =   "Ingreso de Cambios"
      End
      Begin VB.Menu Tipoprv 
         Caption         =   "Ingreso de Rubros de Proveedores"
      End
      Begin VB.Menu EnvioEmailProve 
         Caption         =   "Envio de Email a Proveedores"
      End
   End
   Begin VB.Menu Nov 
      Caption         =   "Novedades"
      Begin VB.Menu movi 
         Caption         =   "A/B/M de Movimientos"
      End
      Begin VB.Menu ConsultaRemito 
         Caption         =   "Consulta de Remitos de Proveedores"
      End
      Begin VB.Menu Aplica 
         Caption         =   "Ingreso de Aplicacion de Comprobantes"
      End
      Begin VB.Menu Pagoprv 
         Caption         =   "Ingreso de Ordenes de Pagos a Proveedores"
      End
      Begin VB.Menu Depo 
         Caption         =   "Ingreso de Depositos"
      End
      Begin VB.Menu Reci 
         Caption         =   "Ingreso de Recibos"
      End
      Begin VB.Menu RecibosProvi 
         Caption         =   "Ingreso de Recibos Provisorios"
      End
      Begin VB.Menu ConChe 
         Caption         =   "Consulta de Cheques"
      End
      Begin VB.Menu cargaInteres 
         Caption         =   "Carga de Intereses"
      End
      Begin VB.Menu Modificainteres 
         Caption         =   "Modificacion de Interes ya cargados"
      End
      Begin VB.Menu SeleccionaREcibo 
         Caption         =   "Seleccion de Recibos a Aplicar diferencia de Cambio"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu CtaCtePrv1 
         Caption         =   "Consulta de Cuenta Corriente por Pantalla"
      End
      Begin VB.Menu CtaCtePrv 
         Caption         =   "Cuenta Corriente de Proveedores"
      End
      Begin VB.Menu SalCtaCtePrv 
         Caption         =   "Saldos de Cuenta Corriente de Proveedores"
      End
      Begin VB.Menu AsiRes 
         Caption         =   "Asiento Resumen"
      End
      Begin VB.Menu IvaComp 
         Caption         =   "Subdiario de Iva compras"
      End
      Begin VB.Menu Proyecc 
         Caption         =   "Proyeccion de Cobros"
      End
      Begin VB.Menu CtaCtePrvSel 
         Caption         =   "Cuenta Corriente de Proveedores Selectivo"
      End
      Begin VB.Menu Valcar 
         Caption         =   "Listado de Valores en Cartera"
      End
      Begin VB.Menu Valcarcuit 
         Caption         =   "Listado de Valores en Cartera por Cuit"
      End
      Begin VB.Menu Posdat 
         Caption         =   "Listado de Pagos Posdatados"
      End
      Begin VB.Menu Listdepo 
         Caption         =   "Listado de Depositos"
      End
      Begin VB.Menu LIstreci 
         Caption         =   "Listado de Recibos"
      End
      Begin VB.Menu Listord 
         Caption         =   "Listado de Ordenes de Pago"
      End
      Begin VB.Menu Listimpu 
         Caption         =   "Listado de Imputaciones de caja y Banco"
      End
      Begin VB.Menu Movban 
         Caption         =   "Movimientos de Bancos"
      End
      Begin VB.Menu Cheemi 
         Caption         =   "Listado de Cheques Emitidos"
      End
      Begin VB.Menu Cf 
         Caption         =   "Listado de Cuenta Corriente de Proveedores a Fecha"
      End
      Begin VB.Menu DifeI 
         Caption         =   "Listado de Diferencia de Cambio (Cobranzas)"
      End
      Begin VB.Menu DifeOtro 
         Caption         =   "Listado de Difefencia de Cambio (Acreditacion)"
      End
      Begin VB.Menu LIstIb 
         Caption         =   "Listado de Retenciones de Ingresos Brutos"
      End
      Begin VB.Menu LIstIbCiudad 
         Caption         =   "Listado de Retenciones de Ingresos Brutos Ciudad Bs. As."
      End
      Begin VB.Menu PagoImpo 
         Caption         =   "Listado de Proyeccion de Pagos de Importaciones"
      End
      Begin VB.Menu DifeExt 
         Caption         =   "Listado de Diferencia de Cambia de Facturas de Exportacion"
      End
      Begin VB.Menu Ccprvanalitico 
         Caption         =   "Listado de Cuentas Corrientes de Proveedores Analitico"
      End
      Begin VB.Menu ProyPrvAnalitico 
         Caption         =   "Listado de Proyeccion de Cuentas Corrientes de Proveedores Analitico"
      End
      Begin VB.Menu Agenda 
         Caption         =   "Listado de Agenda de Vencimentos de Letras y Despachos"
      End
      Begin VB.Menu AnalisispagoPrv 
         Caption         =   "Listado de Analisis de Pago de Facturas a Proveedores"
      End
      Begin VB.Menu LIstreciprovi 
         Caption         =   "Control de Recibos Provisorios"
      End
      Begin VB.Menu sgvsfds 
         Caption         =   "Listado de Deuda de Pyme Nacion "
      End
   End
   Begin VB.Menu ProcesoS 
      Caption         =   "Procesos"
      Begin VB.Menu Cierre 
         Caption         =   "Cierre del Mes"
      End
      Begin VB.Menu Citi 
         Caption         =   "Proceso de Citi"
      End
      Begin VB.Menu DepuraCtaCte 
         Caption         =   "Depuracion de Cuentas Corrientes (Clientes y Proveedores)"
      End
      Begin VB.Menu ProcesoReteib 
         Caption         =   "Proceso de Retenciones de Ingresos Brutos (O.P.)"
      End
      Begin VB.Menu ProcesoReteGanan 
         Caption         =   "Proceso de Retenciones de Ganacias Terceros"
      End
      Begin VB.Menu ProcesoReteIva 
         Caption         =   "Proceso de Retenciones de Iva "
      End
      Begin VB.Menu ListaReteGananII 
         Caption         =   "Proceso de Retenciones de Ganancias"
      End
      Begin VB.Menu ProcesoReteibrecibos 
         Caption         =   "Proceso de Retenciones de Ingresos Brutos"
      End
      Begin VB.Menu ProcesoPercepIb 
         Caption         =   "Proceso de Percepciones de Ingresos Brutos (Facturacion)"
      End
      Begin VB.Menu perceiva 
         Caption         =   "Proceso de Percepciones de Iva (Compras)"
      End
      Begin VB.Menu ProcesoSifere 
         Caption         =   "Proceso de Retenciones Sifere (No Aduana)"
      End
      Begin VB.Menu ProcesoSifereAduana 
         Caption         =   "Proceso de Retenciones Sifere (Aduana)"
      End
      Begin VB.Menu ProcesoPercepIbTucuman 
         Caption         =   "Proceso de Percepciones de Ingresos Brutos (Facturacion/Tucuman)"
      End
      Begin VB.Menu ProcesoCiudad 
         Caption         =   "Proceso Percepciones y Retenciones Ciudad de Bs. As."
      End
      Begin VB.Menu ProcesoCiudadNuevo 
         Caption         =   "Proceso Percepciones y Retenciones Ciudad de Bs. As. (nuevo)"
      End
      Begin VB.Menu ProcesoGananAduana 
         Caption         =   "Proceso Percepciones de ganancias aduana"
      End
      Begin VB.Menu ProcesoSiapre 
         Caption         =   "Proceso de Perciones Aduaneras SIAPRE"
      End
      Begin VB.Menu RecuperoIva 
         Caption         =   "Proceso de Recupero de Iva"
      End
      Begin VB.Menu Procesoperce 
         Caption         =   "Proceso de Recupero de Perce.Aduandas"
      End
      Begin VB.Menu Fin 
         Caption         =   "Fin del Sistema"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Agenda_Click()
    PrgAgenda.Show
End Sub

Private Sub AnalisispagoPrv_Click()
    PrgAnalisisPagoPrv.Show
End Sub

Private Sub Aplica_Click()
    PrgAplica.Show
End Sub

Private Sub AsiRes_Click()
    PrgImputa.Show
End Sub

Private Sub Bancos_Click()
    PrgBanco.Show
End Sub

Private Sub Cambio_Click()
    frmLogin1.Show
End Sub

Private Sub Ccprx_Click()
    PrgCcprvX.Show
End Sub


Private Sub CambioAdm_Click()
    PrgCambioAdm.Show
End Sub

Private Sub cargaInteres_Click()
    PrgCargaInteres.Show
End Sub

Private Sub Ccprvanalitico_Click()
    PrgCcprvAnalitico.Show
End Sub

Private Sub Cf_Click()
    PrgCcprvFec.Show
End Sub

Private Sub Cheemi_Click()
    PrgCheEmi.Show
End Sub

Private Sub Cierre_Click()
    PrgCierre.Show
End Sub

Private Sub Citi_Click()
    PrgCitinuevo.Show
End Sub

Private Sub Command1_Click()
    PrgRecibosProviNuevo.Show
End Sub

Private Sub Command2_Click()
    PrgRecibosnuevo.Show
End Sub

Private Sub Command3_Click()
        Rem
    Rem proceso los comodatos
    Rem

    Set appExcel = CreateObject("Excel.application")
    
    Rem modificar aca
    Rem Ruta = Nombre del archivo excel
    Rem
    
    ruta = "C:\david\sedro.xls"

    If Len(Dir(ruta)) > 0 Then
    
    
        Set objLibro = appExcel.workbooks.Open(ruta)
        
        Do
        
            LugarPlanilla = LugarPlanilla + 1
            
            ZZPartida = appExcel.cells(LugarPlanilla, 1).Value
            
            appExcel.cells(LugarPlanilla, 1).Value = LugarPlanilla
            
            If LugarPlanilla > 100 Then Exit Do
        Loop
            
        appExcel.Quit
        Set appExcel = Nothing
        
    End If
    

End Sub

Private Sub ConChe_Click()
    PrgConche.Show
End Sub

Private Sub ConsultaRemito_Click()
    PrgConsultaRemito.Show
End Sub

Private Sub CtaCtePrv_Click()
    PrgCcprv.Show
End Sub

Private Sub CtaCtePrv1_Click()
    PrgCcprv1.Show
End Sub

Private Sub CtaCtePrvSel_Click()
    PrgCcprvsel.Show
End Sub

Private Sub Cuenta_Click()
    PrgCuenta.Show
End Sub

Private Sub Depo_Click()
    PrgDeposito.Show
End Sub

Private Sub DepuraCtaCte_Click()
    PrgDepuraCtaCte.Show
End Sub

Private Sub DifeExt_Click()
    PrgDifeExt.Show
End Sub

Private Sub DifeI_Click()
    PrgDifeI.Show
End Sub

Private Sub DifeOtro_Click()
    Rem PrgDifeOtro.Show
    PrgDifeOtroNuevo.Show
End Sub

Private Sub EnvioEmailProve_Click()
    PrgEnvioEmailProve.Show
End Sub

Private Sub Fin_Click()
    Close
    End
End Sub

Private Sub Form_Activate()

    If Wempresa = "" Then
        Rem Empresa.Show
        Rem Empresa.SetFocus
        Wempresa = "0001"
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(Wempresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de Administracion : " + !Nombre
            End If
        End With
            Else
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(Wempresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de Administracion : " + !Nombre
            End If
        End With
    End If

End Sub

Private Sub FORM1_Click()
    PrgForm1.Show
End Sub



Private Sub IvaComp_Click()
    PrgIvacomp.Show
End Sub

Private Sub ListaReteGananII_Click()
    PrgProcesoReteGananII.Show
End Sub

Private Sub Listdepo_Click()
    PrgListdepo.Show
End Sub

Private Sub LIstIb_Click()
    PrgListIb.Show
End Sub

Private Sub LIstIbCiudad_Click()
    PrgListIbCiudad.Show
End Sub

Private Sub Listimpu_Click()
    PrgImpcyb.Show
End Sub

Private Sub Listord_Click()
    PrgListPago.Show
End Sub

Private Sub Listreci_Click()
    PrgListreci.Show
End Sub

Private Sub LIstreciprovi_Click()
    PrgListreciProvi.Show
End Sub

Private Sub Modificainteres_Click()
    PrgModificaInteres.Show
End Sub

Private Sub Movban_Click()
    PrgMovban.Show
End Sub

Private Sub movi_Click()
    PrgCompras.Show
End Sub

Private Sub PagoImpo_Click()
    PrgPagoImpo.Show
End Sub

Private Sub Pagoprv_Click()
    Rem  Menu.Hide
    PrgPagoNuevo.Show
End Sub

Private Sub perceiva_Click()
    PrgProcesoPerceIva.Show
End Sub

Private Sub Porveedores_Click()
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7
            PrgProve.Show
        Case Else
            PrgProvePelli.Show
    End Select
End Sub

Private Sub Posdat_Click()
    PrgPosdat.Show
End Sub

Private Sub PrgVarias_Click()
    Prgpagovar.Show
End Sub

Private Sub ProcesoCiudad_Click()
    PrgProcesoCiudad.Show
End Sub

Private Sub ProcesoCiudadNuevo_Click()
    PrgProcesoCiudadNuevo.Show
End Sub

Private Sub ProcesoGananAduana_Click()
    PrgProcesoGananciaAduana.Show
End Sub

Private Sub Procesoperce_Click()
    PrgRecuperoPerce.Show
End Sub

Private Sub ProcesoPercepIb_Click()
    PrgProcesoPerceIb.Show
End Sub

Private Sub ProcesoPercepIbTucuman_Click()
    PrgProcesoPerceIbTucuman.Show
End Sub

Private Sub ProcesoReteGanan_Click()
    PrgProcesoReteGanan.Show
End Sub

Private Sub ProcesoReteIb_Click()
    PrgProcesoReteIb.Show
End Sub

Private Sub ProcesoReteibrecibos_Click()
    PrgProcesoReteIbRecibos.Show
End Sub

Private Sub ProcesoReteIva_Click()
    PrgProcesoReteIva.Show
End Sub

Private Sub ProcesoSiapre_Click()
    PrgProcesoSiapre.Show
End Sub

Private Sub ProcesoSifere_Click()
    PrgProcesoSifere.Show
End Sub

Private Sub ProcesoSifereAduana_Click()
    PrgProcesoSifereAduana.Show
End Sub

Private Sub Proyecc_Click()
    PrgProyPrv.Show
End Sub

Private Sub ProyPrvAnalitico_Click()
    PrgProyPrvAnalitico.Show
End Sub

Private Sub Reci_Click()
    PrgRecibos.Show
End Sub

Private Sub RecibosProvi_Click()
    PrgRecibosProvi.Show
End Sub

Private Sub RecuperoIva_Click()
    PrgRecuperoIva.Show
End Sub

Private Sub SalCtaCtePrv_Click()
    PrgSalprv.Show
End Sub

Private Sub SeleccionaREcibo_Click()
    PrgSeleccionaRecibo.Show
End Sub

Private Sub sgvsfds_Click()
    PrgListaPyme.Show
End Sub

Private Sub Tipoprv_Click()
    PrgTipoPrv.Show
End Sub

Private Sub Valcar_Click()
    PrgValcar.Show
End Sub

Private Sub Valcarcuit_Click()
    PrgValcarcuit.Show
End Sub
