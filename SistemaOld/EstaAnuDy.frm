VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEstaAnuDy 
   AutoRedraw      =   -1  'True
   Caption         =   "14.-Listado de Estadisticas Anuales"
   ClientHeight    =   4080
   ClientLeft      =   2175
   ClientTop       =   945
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4080
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Top             =   2280
         Width           =   2655
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1680
         TabIndex        =   12
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeFec 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   2760
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Listado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6840
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEsta7.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgEstaAnuDy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Uno As String
Private Dos As String
Private Costo As Double
Private Producto As String
Private Auxiliar(100, 7) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTermnado As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstLinea As Recordset
Dim spLinea As String
Dim XParam As String
Private Vecosto(5000, 2) As String
Dim Posi As Integer
Dim ImpreFecha(12, 2) As String


Private Sub Acepta_Click()

    On Error GoTo WError
    
    MesIni = Val(Mid$(DesdeFec.Text, 4, 2))
    AnoIni = Val(Right$(DesdeFec.Text, 4))
    
    MesCicla = MesIni
    AnoCicla = AnoIni
    Erase ImpreFecha
    
    For Ciclo = 1 To 12
    
        ImpreFecha(Ciclo, 1) = Str$(MesCicla)
        ImpreFecha(Ciclo, 2) = Str$(AnoCicla)
        
        MesCicla = MesCicla + 1
        If MesCicla > 12 Then
            MesCicla = 1
            AnoCicla = AnoCicla + 1
        End If
        
    Next Ciclo
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    Whasta = WAno + WMes + WDia
    
    WTitulo = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            Nomempresa = !Nombre
        End If
    End With
    
    With rstEstaAnu
        .Index = "Codigo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    WTitulo1 = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    WTitulo3 = Nomempresa
    
    Sql1 = "Select Estadistica.Tipo, Estadistica.Numero, Estadistica.Renglon, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Precio, Estadistica.PrecioUs, Estadistica.Importe, Estadistica.ImporteUs, Estadistica.Cliente, Estadistica.Linea, Estadistica.Costo1, Estadistica.Costo2, Estadistica.Coeficiente, Estadistica.Pedido, Estadistica.Fecha, Estadistica.OrdFecha, Estadistica.Articulo, Estadistica.Remito, Estadistica.Clave, Estadistica.WArticulo, Estadistica.Paridad, Estadistica.Importe1, Estadistica.Importe2, Estadistica.Importe3, Estadistica.Importe4, Estadistica.Vendedor, Estadistica.Rubro"
    Sql2 = " FROM Estadistica"
    Sql3 = " Where Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    Sql4 = " and Estadistica.OrdFecha <= " + "'" + Whasta + "'"
    spEstadistica = Sql1 + Sql2 + Sql3 + Sql4
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            Do
                
                WTipo = rstEstadistica!Tipo
                WNumero = rstEstadistica!numero
                WRenglon = rstEstadistica!Renglon
                WArticulo = rstEstadistica!Articulo
                WCantidad = rstEstadistica!Cantidad
                WPrecio = rstEstadistica!Precio
                WPrecioUs = rstEstadistica!PrecioUs
                WImporte = rstEstadistica!Importe
                WimporteUs = rstEstadistica!ImporteUs
                WCliente = rstEstadistica!Cliente
                WParidad = rstEstadistica!Paridad
                wvendedor = rstEstadistica!Vendedor
                WRubro = rstEstadistica!Rubro
                WLinea = rstEstadistica!Linea
                WCosto1 = rstEstadistica!Costo1
                WCosto2 = rstEstadistica!Costo2
                WCoeficiente = rstEstadistica!Coeficiente
                WPedido = rstEstadistica!Pedido
                WFecha = rstEstadistica!Fecha
                WImporte1 = rstEstadistica!Importe1
                WImporte2 = rstEstadistica!Importe2
                WImporte3 = rstEstadistica!Importe3
                WImporte4 = rstEstadistica!Importe4
                WOrdFecha = rstEstadistica!OrdFecha
                WWArticulo = rstEstadistica!WArticulo
                WRemito = rstEstadistica!Remito
                WClave = rstEstadistica!Clave
                
                If Desde.Text <= WArticulo And WArticulo <= Hasta.Text Then
                
                    Impo1 = 0
                    Impo2 = 0
                    Impo3 = 0
                    Impo4 = 0
                    Impo5 = 0
                    Impo6 = 0
                    Impo7 = 0
                    Impo8 = 0
                    Impo9 = 0
                    Impo10 = 0
                    Impo11 = 0
                    Impo12 = 0
                
                    MesCompara = Val(Mid$(WFecha, 4, 2))
                    AnoCompara = Val(Right$(WFecha, 4))
                    
                    Select Case Tipo.ListIndex
                        Case 0
                            XImpo = WCantidad
                            WTitulo2 = "Kilos Vendidos"
                        Case 1
                            XImpo = WImporte
                            WTitulo2 = "Monto Vendido expresado en  $"
                        Case Else
                            XImpo = WimporteUs
                            WTitulo2 = "Monto Vendido expresado en  U$S"
                    End Select
                    
                    If !Tipo = 2 Then
                        XImpo = Abs(XImpo) * -1
                    End If
                        
                    If MesCompara = Val(ImpreFecha(1, 1)) And AnoCompara = Val(ImpreFecha(1, 2)) Then
                        Impo1 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(2, 1)) And AnoCompara = Val(ImpreFecha(2, 2)) Then
                        Impo2 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(3, 1)) And AnoCompara = Val(ImpreFecha(3, 2)) Then
                        Impo3 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(4, 1)) And AnoCompara = Val(ImpreFecha(4, 2)) Then
                        Impo4 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(5, 1)) And AnoCompara = Val(ImpreFecha(5, 2)) Then
                        Impo5 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(6, 1)) And AnoCompara = Val(ImpreFecha(6, 2)) Then
                        Impo6 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(7, 1)) And AnoCompara = Val(ImpreFecha(7, 2)) Then
                        Impo7 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(8, 1)) And AnoCompara = Val(ImpreFecha(8, 2)) Then
                        Impo8 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(9, 1)) And AnoCompara = Val(ImpreFecha(9, 2)) Then
                        Impo9 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(10, 1)) And AnoCompara = Val(ImpreFecha(10, 2)) Then
                        Impo10 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(11, 1)) And AnoCompara = Val(ImpreFecha(11, 2)) Then
                        Impo11 = XImpo
                    End If
                    
                    If MesCompara = Val(ImpreFecha(12, 1)) And AnoCompara = Val(ImpreFecha(12, 2)) Then
                        Impo12 = XImpo
                    End If
                    
                    With rstEstaAnu
                        .Index = "Codigo"
                        .Seek "=", WArticulo
                        If .NoMatch = True Then
                            .AddNew
                            !Codigo = WArticulo
                            !Linea = 0
                            !Impo1 = Impo1
                            !Descri1 = ImpreFecha(1, 1) + "/" + ImpreFecha(1, 2)
                            !Impo2 = Impo2
                            !Descri2 = ImpreFecha(2, 1) + "/" + ImpreFecha(2, 2)
                            !Impo3 = Impo3
                            !Descri3 = ImpreFecha(3, 1) + "/" + ImpreFecha(3, 2)
                            !Impo4 = Impo4
                            !Descri4 = ImpreFecha(4, 1) + "/" + ImpreFecha(4, 2)
                            !Impo5 = Impo5
                            !Descri5 = ImpreFecha(5, 1) + "/" + ImpreFecha(5, 2)
                            !Impo6 = Impo6
                            !Descri6 = ImpreFecha(6, 1) + "/" + ImpreFecha(6, 2)
                            !Impo7 = Impo7
                            !Descri7 = ImpreFecha(7, 1) + "/" + ImpreFecha(7, 2)
                            !Impo8 = Impo8
                            !Descri8 = ImpreFecha(8, 1) + "/" + ImpreFecha(8, 2)
                            !Impo9 = Impo9
                            !Descri9 = ImpreFecha(9, 1) + "/" + ImpreFecha(9, 2)
                            !Impo10 = Impo10
                            !Descri10 = ImpreFecha(10, 1) + "/" + ImpreFecha(10, 2)
                            !Impo11 = Impo11
                            !Descri11 = ImpreFecha(11, 1) + "/" + ImpreFecha(11, 2)
                            !Impo12 = Impo12
                            !Descri12 = ImpreFecha(12, 1) + "/" + ImpreFecha(12, 2)
                            !Titulo1 = WTitulo1
                            !Titulo2 = WTitulo2
                            !Titulo3 = WTitulo3
                            .Update
                                Else
                            .Edit
                            !Impo1 = !Impo1 + Impo1
                            !Impo2 = !Impo2 + Impo2
                            !Impo3 = !Impo3 + Impo3
                            !Impo4 = !Impo4 + Impo4
                            !Impo5 = !Impo5 + Impo5
                            !Impo6 = !Impo6 + Impo6
                            !Impo7 = !Impo7 + Impo7
                            !Impo8 = !Impo8 + Impo8
                            !Impo9 = !Impo9 + Impo9
                            !Impo10 = !Impo10 + Impo10
                            !Impo11 = !Impo11 + Impo11
                            !Impo12 = !Impo12 + Impo12
                            .Update
                        End If
                    End With
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
    End If
    
    With rstEstaAnu
        .Index = "Codigo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WCodigo = !Codigo
                WLinea = !Linea
                WDescripcion = ""
                WDescriLinea = ""
                WImpreCodigo = ""
                        
                If Left$(WCodigo, 2) = "PT" Or Left$(WCodigo, 2) = "PE" Then
                    WImpreCodigo = WCodigo
                    spTerminado = "ConsultaTerminado" + "'" + WCodigo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WDescripcion = rstTerminado!Descripcion
                        WLinea = rstTerminado!Linea
                        rstTerminado.Close
                    End If
                        Else
                    XArti = Left$(WCodigo, 3) + Right$(WCodigo, 7)
                    WImpreCodigo = XArti
                    spArticulo = "ConsultaArticulo " + "'" + XArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WDescripcion = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                End If
                    
                spLinea = "ConsultaLinea" + "'" + Str$(WLinea) + "'"
                Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
                If rstLinea.RecordCount > 0 Then
                    WDescriLinea = rstLinea!Nombre
                    rstLinea.Close
                End If
                
                !Descripcion = WDescripcion
                !DescriLinea = WDescriLinea
                !ImpreCodigo = WImpreCodigo
                
                !Corte = Left$(WDescripcion, 5)
                    
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    WTitulo = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "14.-Listado de Estadistica de Ventas Anuales"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + Whasta + Chr$(34)
    Rem Dos = " and {Estadistica.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Rem Listado.GroupSelectionFormula = Uno + Dos
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 10 Then
        Listado.ReportFileName = "WEstaAnuDy.rpt"
            Else
        Listado.ReportFileName = "WEstaAnuDy.rpt"
    End If
    
    Listado.Action = 1
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEstaAnu
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Desde.SetFocus
    PrgEstaAnu.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = "  -     -   "
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        DesdeFec.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  -     -   "
    End If
End Sub

Private Sub DesdeFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFec.Text, Auxi)
        If Auxi = "S" Then
            HastaFec.SetFocus
                Else
            DesdeFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFec.Text = "  /  /    "
    End If
End Sub

Private Sub HastaFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Cantidad"
    Tipo.AddItem "Importe en $"
    Tipo.AddItem "Importe en U$S"
    
    Tipo.ListIndex = 0

    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Posi = 0
End Sub


Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_EstaAnu
End Sub






