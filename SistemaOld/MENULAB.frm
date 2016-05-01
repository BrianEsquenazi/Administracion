VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Laboratorio"
   ClientHeight    =   7560
   ClientLeft      =   1830
   ClientTop       =   1050
   ClientWidth     =   8250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   8250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cambio 
      Caption         =   "Cambio de Empresa"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Menu ss 
      Caption         =   "Ingreso de Novedades"
      Begin VB.Menu Ensayo 
         Caption         =   "Ensayos"
      End
      Begin VB.Menu aa 
         Caption         =   "Materia Prima"
         Begin VB.Menu Especi1 
            Caption         =   "Consulta de Especificaciones (Historico)"
            Visible         =   0   'False
         End
         Begin VB.Menu Especi1UnificaVerison 
            Caption         =   "Consulta de Especificaciones por Version"
         End
         Begin VB.Menu Especi1Unifica 
            Caption         =   "Especificaciones (Unificado)"
         End
         Begin VB.Menu Control1 
            Caption         =   "Controles"
         End
         Begin VB.Menu ListaEnsayoMp 
            Caption         =   "Listado de Ensayos en Materia Prima"
         End
         Begin VB.Menu Homologaprove 
            Caption         =   "Homologacion de Muestras de Materias Primas"
         End
         Begin VB.Menu ListaVtoMp 
            Caption         =   "Verificacion de Vencimientos de Materia Prima"
         End
         Begin VB.Menu ListaEspecifMp 
            Caption         =   "Listado de Especificaciones de Materia Prima por Fecha"
         End
         Begin VB.Menu EtiContra 
            Caption         =   "Etiquetas de Muestra Simple"
         End
         Begin VB.Menu revalidady 
            Caption         =   "Revalida de DY"
         End
         Begin VB.Menu VerificaLoteArti 
            Caption         =   "Verificacion de Lotes"
         End
      End
      Begin VB.Menu sss 
         Caption         =   "Producto terminado"
         Begin VB.Menu Especifi2 
            Caption         =   "Consulta de Especificaciones (Historico)"
            Visible         =   0   'False
         End
         Begin VB.Menu Especi2UnificaVerison 
            Caption         =   "Consulta de Especificaciones por Version"
         End
         Begin VB.Menu Especifi2Unifica 
            Caption         =   "Especificaciones (Unificado)"
         End
         Begin VB.Menu Control2 
            Caption         =   "Controles"
         End
         Begin VB.Menu Pruedev 
            Caption         =   "Devolucion de NK o RE"
         End
         Begin VB.Menu ListaEnsayoPt 
            Caption         =   "Listado de Ensayos de Producto Terminado"
         End
         Begin VB.Menu AltaCeritificado 
            Caption         =   "Carga de Ensayos a Imprimir en los Certificados de Analisis"
         End
         Begin VB.Menu EmiteCerti 
            Caption         =   "Emision de Certificado de Analisis"
         End
         Begin VB.Menu ListaEspecifPt 
            Caption         =   "Listado de Especificaciones de Producto Terminado por Fecha"
         End
         Begin VB.Menu ListaPtVecido 
            Caption         =   "Listado de Productos Terminados Vencidos "
         End
         Begin VB.Menu CambioParametro 
            Caption         =   "Cambio de Valores Standard"
         End
      End
      Begin VB.Menu dfgdfg 
         Caption         =   "Producto terminado (Farma)"
         Begin VB.Menu CargaIIIFarma 
            Caption         =   "Especificaciones (Farma)"
         End
         Begin VB.Menu ControlFarma 
            Caption         =   "Controles (Farma)"
         End
         Begin VB.Menu actualizapedidofarma 
            Caption         =   "Liberacion de Pedidos"
         End
      End
      Begin VB.Menu ModHoja 
         Caption         =   "Ingreso y Actualizacion de  Hojas de Produccion"
      End
      Begin VB.Menu modhojaplantaii 
         Caption         =   "Ingreso Hoja de Produccion Planta III y V"
         Visible         =   0   'False
      End
      Begin VB.Menu MOvlab 
         Caption         =   "Movimientos Varios de Stock"
      End
      Begin VB.Menu LiberaTerminado 
         Caption         =   "Liberacion de Productos Devueltos a Verificar"
      End
      Begin VB.Menu ListaPendienteLiberar 
         Caption         =   "Listado de Productos Pendientes de Liberar"
      End
      Begin VB.Menu VerificaPedidoLabora 
         Caption         =   "Verificacion de Pedidos de Desarrollo"
      End
      Begin VB.Menu InformeLabo 
         Caption         =   "Ingreso de Informe de Recepcion de Drogas de Laboratorio"
      End
      Begin VB.Menu bajaLote 
         Caption         =   "Verificacion de Lotes Inactivos"
      End
      Begin VB.Menu ModHojaDY 
         Caption         =   "Traspaso de Pt a DY"
      End
      Begin VB.Menu FrasesH 
         Caption         =   "Ingresos de Frases H"
      End
      Begin VB.Menu FeasesdP 
         Caption         =   "Ingresos de Frases P"
      End
      Begin VB.Menu DatosEtiqueta 
         Caption         =   "Datos Adicionales de Etiquetas de PT"
      End
      Begin VB.Menu DatosEtiquetaMp 
         Caption         =   "Datos Adicionales de Etiquetas de MP"
      End
   End
   Begin VB.Menu x 
      Caption         =   "Procesos"
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
Dim rstAtributo As Recordset
Dim spAtributo As String
Dim Atri(10, 100) As Integer

Private Sub actualizapedidofarma_Click()
    PrgActualizaPedidoFarma.Show
End Sub

Private Sub AltaCeritificado_Click()
    ZZPasaCliente = ""
    ZZPasaTerminado = ""
    PrgAltaCertificado.Show
End Sub

Private Sub bajaLote_Click()
    If Val(Wempresa) = 3 Or Val(Wempresa) = 4 Then
        PrgBajaLote.Show
    End If
End Sub

Private Sub CambioParametro_Click()
    PrgCambiaParametro.Show
End Sub

Private Sub CargaIIIFarma_Click()
    If Val(Wempresa) = 5 Then
        PrgCargaIIILabo.Show
    End If
End Sub

Private Sub Command1_Click()
        PrgPrueArtRango.Show

End Sub

Private Sub Control1_Click()
    PrgPruart.Show
End Sub

Private Sub Control2_Click()
    PrgPruter.Show
End Sub

Private Sub ControlFarma_Click()
    If Val(Wempresa) = 5 Then
        PrgPruterFarma.Show
    End If
End Sub

Private Sub DatosEtiqueta_Click()
    PrgDatosEtiqueta.Show
End Sub

Private Sub DatosEtiquetaMp_Click()
    PrgDatosEtiquetaMp.Show
End Sub

Private Sub EmiteCerti_Click()
    PrgEmiteCertificado.Show
End Sub

Private Sub Ensayo_Click()
    PrgEnsayo.Show
End Sub

Private Sub Especi1_Click()
    PrgEspecifi.Show
End Sub

Private Sub Especi1Histo_Click()
    PrgEspeHistorico.Show
End Sub

Private Sub Especi1Unifica_Click()
    PrgEspecifiUnifica.Show
End Sub

Private Sub Especi1UnificaVerison_Click()
    PrgEspecifiUnificaVersion.Show
End Sub

Private Sub Especi2UnificaVerison_Click()
    PrgEspeUnificaVersion.Show
End Sub

Private Sub Especifi2_Click()
    PrgEspe.Show
End Sub

Private Sub Cambio_Click()
    frmLogin.Show
End Sub

Private Sub Especifi2Histo_Click()
    PrgEspecifiHistorico.Show
End Sub

Private Sub Especifi2Unifica_Click()
    PrgEspeUnifica.Show
End Sub

Private Sub EtiContra_Click()
    PrgEtiContra.Show
End Sub

Private Sub FeasesdP_Click()
    PrgFraseP.Show
End Sub

Private Sub Fin_Click()
    Close
    End
    Rem Menu.WindowState = 1
End Sub

Private Sub Form_Activate()

    If Wempresa = "" Then
        Wempresa = "0001"
    End If

    If Wempresa = "" Then
        Empresa.Show
        Empresa.SetFocus
        Wempresa = 1
            Else
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(Wempresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de laboratorio : " + !Nombre
            End If
        End With
    End If
    
    XOperador = Str$(WOperador)
    XProceso = "6"
    WAtributo1 = "00000000000000000000000000000000000000000000"
    WAtributo2 = "00000000000000000000000000000000000000000000"
    WAtributo3 = "00000000000000000000000000000000000000000000"
    WAtributo4 = "00000000000000000000000000000000000000000000"
    WAtributo5 = "00000000000000000000000000000000000000000000"
    WAtributo6 = "00000000000000000000000000000000000000000000"
    WAtributo7 = "00000000000000000000000000000000000000000000"
    WAtributo8 = "00000000000000000000000000000000000000000000"
    WAtributo9 = "00000000000000000000000000000000000000000000"
    WAtributo10 = "00000000000000000000000000000000000000000000"
    
    XParam = "'" + XOperador + "','" _
                 + XProceso + "'"
    spAtributo = "ConsultaAtributo " + XParam
    Set rstAtributo = db.OpenRecordset(spAtributo, dbOpenSnapshot, dbSQLPassThrough)
    If rstAtributo.RecordCount > 0 Then
        WAtributo1 = rstAtributo!Atributo1 + "00000000000000000000000000000"
        WAtributo2 = rstAtributo!Atributo2 + "00000000000000000000000000000"
        WAtributo3 = rstAtributo!Atributo3 + "00000000000000000000000000000"
        WAtributo4 = rstAtributo!Atributo4 + "00000000000000000000000000000"
        WAtributo5 = rstAtributo!Atributo5 + "00000000000000000000000000000"
        WAtributo6 = rstAtributo!Atributo6 + "00000000000000000000000000000"
        WAtributo7 = rstAtributo!Atributo7 + "00000000000000000000000000000"
        WAtributo8 = rstAtributo!Atributo8 + "00000000000000000000000000000"
        WAtributo9 = rstAtributo!Atributo9 + "00000000000000000000000000000"
        WAtributo10 = rstAtributo!Atributo10 + "00000000000000000000000000000"
        rstAtributo.Close
    End If
    
    For Ciclo = 1 To 10
        Select Case Ciclo
            Case 1
                Auxiliar = WAtributo1
            Case 2
                Auxiliar = WAtributo2
            Case 3
                Auxiliar = WAtributo3
            Case 4
                Auxiliar = WAtributo4
            Case 5
                Auxiliar = WAtributo5
            Case 6
                Auxiliar = WAtributo6
            Case 7
                Auxiliar = WAtributo7
            Case 8
                Auxiliar = WAtributo8
            Case 9
                Auxiliar = WAtributo9
            Case 10
                Auxiliar = WAtributo10
            Case Else
        End Select
        For Ciclo1 = 1 To 31
            Atri(Ciclo, Ciclo1) = Val(Mid$(Auxiliar, Ciclo1, 1))
        Next Ciclo1
    Next Ciclo
            
    Menu.Ensayo.Enabled = Atri(1, 1)
    Menu.Especi1.Enabled = Atri(1, 2)
    Menu.Especi1UnificaVerison.Enabled = Atri(1, 3)
    Menu.Especi1Unifica.Enabled = Atri(1, 4)
    Menu.Control1.Enabled = Atri(1, 5)
    Menu.ListaEnsayoMp.Enabled = Atri(1, 6)
    Menu.Homologaprove.Enabled = Atri(1, 7)
    Menu.ListaVtoMp.Enabled = Atri(1, 8)
    Menu.ListaEspecifMp.Enabled = Atri(1, 9)
    Menu.EtiContra.Enabled = Atri(1, 10)
    Menu.Especifi2.Enabled = Atri(1, 11)
    Menu.Especi2UnificaVerison.Enabled = Atri(1, 12)
    Menu.Especifi2Unifica.Enabled = Atri(1, 13)
    Menu.Control2.Enabled = Atri(1, 14)
    Menu.Pruedev.Enabled = Atri(1, 15)
    Menu.ListaEnsayoPt.Enabled = Atri(1, 16)
    Menu.AltaCeritificado.Enabled = Atri(1, 17)
    Menu.EmiteCerti.Enabled = Atri(1, 18)
    Menu.ListaEspecifPt.Enabled = Atri(1, 19)
    Menu.ListaPtVecido.Enabled = Atri(1, 20)
    Menu.CargaIIIFarma.Enabled = Atri(1, 21)
    Menu.ControlFarma.Enabled = Atri(1, 22)
    Menu.ModHoja.Enabled = Atri(1, 23)
    Menu.modhojaplantaii.Enabled = Atri(1, 24)
    Menu.MOvlab.Enabled = Atri(1, 25)
    Menu.LiberaTerminado.Enabled = Atri(1, 26)
    Menu.ListaPendienteLiberar.Enabled = Atri(1, 27)
    Menu.VerificaPedidoLabora.Enabled = Atri(1, 28)
    Menu.InformeLabo.Enabled = Atri(1, 29)
    Menu.bajaLote.Enabled = Atri(1, 30)
    Rem by nan habilito para carga especif etiquetas SGA
    Select Case Val(Wempresa)
                 Case 3, 5, 6, 7, 10, 11
                 
                  Menu.FrasesH.Enabled = False
                  Menu.FeasesdP.Enabled = False
                  Menu.DatosEtiqueta.Enabled = False
                  Menu.DatosEtiquetaMp.Enabled = False
                 
                 Case 2, 4, 9
                  Menu.FrasesH.Enabled = False
                  Menu.FeasesdP.Enabled = False
                  Menu.DatosEtiqueta.Enabled = False
                  Menu.DatosEtiquetaMp.Enabled = False
                 Case Else
                  Menu.FrasesH.Enabled = True
                  Menu.FeasesdP.Enabled = True
                  Menu.DatosEtiqueta.Enabled = True
                  Menu.DatosEtiquetaMp.Enabled = True
      
      End Select
   


   Rem fin by nan




End Sub

Private Sub FrasesH_Click()
    PrgFraseH.Show
End Sub

Private Sub Homologaprove_Click()
    PrgHomologaProve.Show
End Sub

Private Sub InformeLabo_Click()
    PrgInforme.Show
End Sub

Private Sub LiberaTerminado_Click()
    PrgLiberaTerminado.Show
End Sub

Private Sub ListaEnsayoMp_Click()
    PrgListaEnsayoMp.Show
End Sub

Private Sub ListaEnsayoPt_Click()
    PrgListaEnsayoPt.Show
End Sub

Private Sub ListaEspecifMp_Click()
    PrgListaEspecifMp.Show
End Sub

Private Sub ListaEspecifPt_Click()
    PrgListaEspecifPt.Show
End Sub

Private Sub ListaPendienteLiberar_Click()
    PrgListaPendienteLiberar.Show
End Sub

Private Sub ListaPtVecido_Click()
    PrgListaPtVencido.Show
End Sub

Private Sub ListaVtoMp_Click()
    PrgListaVto.Show
End Sub

Private Sub Modhoja_Click()
   Rem If Val(WEmpresa) <> 9 And Val(WEmpresa) <> 10 Then
        Prgmodhoja.Show
  Rem   End If
End Sub

Private Sub ModHojaDY_Click()
    Prgmodhojady.Show
End Sub

Private Sub modhojaplantaii_Click()
    If Val(Wempresa) <> 1 And Val(Wempresa) <> 2 And Val(Wempresa) <> 3 And Val(Wempresa) <> 4 Then
        If Val(Wempresa) <> 9 And Val(Wempresa) <> 10 Then
            PrgModHojaLaboraII.Show
        End If
    End If
End Sub

Private Sub MOvlab_Click()
    PrgMovlab.Show
End Sub

Private Sub Pruedev_Click()
    PrgPruedev.Show
End Sub

Private Sub revalidady_Click()
    PrgRevalidady.Show
End Sub

Private Sub VerificaLoteArti_Click()
    WEmpresaVerifica = Wempresa
    PrgVerificaLoteArti.Show
End Sub

Private Sub VerificaPedidoLabora_Click()
    PrgVerificaLabora.Show
End Sub
