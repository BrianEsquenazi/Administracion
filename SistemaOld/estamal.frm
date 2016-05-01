VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEstamal 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Ventas Fuera de Fecha"
   ClientHeight    =   3060
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3060
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   720
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
         Left            =   1680
         TabIndex        =   0
         Top             =   240
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1320
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   1320
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
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   240
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
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   720
         Width           =   1095
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
         TabIndex        =   7
         Top             =   840
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
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEstamal.rpt"
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
Attribute VB_Name = "PrgEstamal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Costo As Double
Private Producto As String
Private Auxiliar(100, 7) As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstFeriado As Recordset
Dim spFeriado As String
Dim XParam As String
Private Vecosto(5000, 2) As String
Private WPedido As String
Dim Posi As Integer
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim DiaFeriado(100) As String
Private TotalFeriado As Integer

Private Sub Acepta_Click()

    On Error GoTo WError
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    With rstEsta
        .Index = "Clave"
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
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spEstadistica = "ListaEstadisticaFecha" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            Do
                
                WTipo = rstEstadistica!Tipo
                WNumero = rstEstadistica!Numero
                WRenglon = rstEstadistica!Renglon
                WArticulo = rstEstadistica!Articulo
                WCantidad = rstEstadistica!Cantidad
                WPrecio = rstEstadistica!Precio
                WPrecioUs = rstEstadistica!PrecioUs
                WImporte = rstEstadistica!Importe
                WimporteUs = rstEstadistica!ImporteUs
                WCliente = rstEstadistica!Cliente
                WParidad = rstEstadistica!Paridad
                WVendedor = rstEstadistica!vendedor
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
                
                If Val(WPedido) <> 0 Then
                
                    With rstEsta
        
                        .Index = "Clave"
                        .AddNew
                        !Tipo = WTipo
                        !Numero = WNumero
                        !Renglon = WRenglon
                        !Articulo = WArticulo
                        !Cantidad = WCantidad
                        !Precio = WPrecio
                        !PrecioUs = WPrecioUs
                        !Importe = WImporte
                        !ImporteUs = WimporteUs
                        !Cliente = WCliente
                        !Paridad = WParidad
                        !vendedor = WVendedor
                        !Rubro = WRubro
                        !Linea = WLinea
                        !Costo1 = WCosto1
                        !Costo2 = WCosto2
                        !Coeficiente = WCoeficiente
                        !Pedido = WPedido
                        !Fecha = WFecha
                        !Importe1 = WImporte1
                        !Importe2 = WImporte2
                        !Importe3 = WImporte3
                        !Importe4 = WImporte4
                        !OrdFecha = WOrdFecha
                        !WArticulo = WWArticulo
                        !Remito = WRemito
                        !Clave = WClave
                        .Update
                    End With
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
    End If
    
    With rstEsta
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                !WCantidad = 0
                !WImporte = 0
                !WimporteUs = 0
                !Costo2 = 0
                
                WCliente = !Cliente
                WPedido = !Pedido
                WOrdFecha = !OrdFecha
                    
                spCliente = "ConsultaCliente" + "'" + WCliente + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    WDescriCliente = rstCliente!Razon
                    rstCliente.Close
                End If
                    
                !DescriCliente = WDescriCliente
                
                spPedido = "ListaPedido " + "'" + WPedido + "'"
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                If rstPedido.RecordCount > 0 Then
                    WAno = Right$(rstPedido!FecEntrega, 4)
                    WMes = Mid$(rstPedido!FecEntrega, 4, 2)
                    WDia = Left$(rstPedido!FecEntrega, 2)
                    WFecEntrega = WAno + WMes + WDia
                    WFecEntrega1 = rstPedido!FecEntrega
                   rstPedido.Close
                End If
                If WOrdFecha <= WFecEntrega Then
                    .Delete
                        Else
                    !FecEntrega = WFecEntrega1
                    
                    WDias = 0
                    WFechaHasta = !Fecha
                    WFechaDesde = !FecEntrega
    
                    Do
        
                        Feriado = "N"
                        For Ciclo = 1 To TotalFeriado
                            If DiaFeriado(Ciclo) = WFechaDesde Then
                                Feriado = "S"
                                Exit For
                            End If
                        Next Ciclo
                    
                        Rem 1 - DOMINGO
                        Rem 2 - LUNES
                        Rem 3 - MARTES
                        Rem 4 - MIERCOLES
                        Rem 5 - JUEVES
                        Rem 6 - VIERNES
                        Rem 7 - SABADO
                        XFec1 = WFechaDesde
                        strDia = Format$(XFec1, "dddd")
                        BDia = Format(XFec1, "w")
                        If BDia = 1 Or BDia = 7 Then
                            Feriado = "S"
                        End If
            
                        If Feriado = "N" Then
                            WDias = WDias + 1
                        End If
                        SumaDia = 2
                        Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                        WFechaDesde = XFec2
                        
                        If WFechaDesde = WFechaHasta Then
                            Exit Do
                        End If
        
                    Loop
                    
                    !Costo2 = WDias
                    
                    .Update
                End If
                
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
    
    Listado.WindowTitle = "Listado de Ventas fuera de fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Listado.GroupSelectionFormula = Uno
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEsta
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    DesdeFec.SetFocus
    PrgEstamal.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub DesdeFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFec.Text, Auxi)
        If Auxi = "S" Then
            HastaFec.Text = DesdeFec.Text
            HastaFec.SetFocus
                Else
            DesdeFec.SetFocus
        End If
    End If
End Sub

Private Sub HastaFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            DesdeFec.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
End Sub

Sub Form_Load()
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
    Erase DiaFeriado
    TotalFeriado = 0
    
    spFeriado = "ListaFeriado"
    Set rstFeriado = db.OpenRecordset(spFeriado, dbOpenSnapshot, dbSQLPassThrough)
    If rstFeriado.RecordCount > 0 Then
        With rstFeriado
            .MoveFirst
            Do
                If .EOF = False Then
                    TotalFeriado = TotalFeriado + 1
                    DiaFeriado(TotalFeriado) = rstFeriado!Fecha
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstFeriado.Close
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Esta
End Sub
