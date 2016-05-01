VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaDevol 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Analisis de Devoluciones"
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
      ReportFileName  =   "ListaDevol.rpt"
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
Attribute VB_Name = "PrgListaDevol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZVector(1000, 10) As String

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
    
    Erase ZVector
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Estadistica.Tipo = " + "'" + "2" + "'"
    ZSql = ZSql + " and Estadistica.Linea <> " + "'" + "50" + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            Do
                
                ZLugar = ZLugar + 1
                
                ZVector(ZLugar, 1) = rstEstadistica!Numero
                ZVector(ZLugar, 2) = rstEstadistica!Articulo
                ZVector(ZLugar, 3) = Str$(rstEstadistica!Cantidad)
                ZVector(ZLugar, 4) = rstEstadistica!Cliente
                ZVector(ZLugar, 5) = rstEstadistica!Fecha
                ZVector(ZLugar, 6) = rstEstadistica!Clave
                ZVector(ZLugar, 7) = rstEstadistica!lote1
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZNumero = ZVector(Ciclo, 1)
        ZArticulo = ZVector(Ciclo, 2)
        ZCantidad = ZVector(Ciclo, 3)
        ZCliente = ZVector(Ciclo, 4)
        ZFecha = ZVector(Ciclo, 5)
        ZClave = ZVector(Ciclo, 6)
        ZNroEntrada = ""
        ZFechaEntrada = ""
        ZNroPedido = ""
        ZFechaPedido = ""
        ZLoteOriginal = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.Numero = " + "'" + ZNumero + "'"
        ZSql = ZSql + " and CtaCte.Tipo = " + "'" + "02" + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            ZNroEntrada = rstCtacte!Remito
            rstCtacte.Close
        End If
        
        XEmpresa = WEmpresa
        
        For CiclaEmpre = 1 To 3
        
            Select Case CiclaEmpre
                Case 1
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    WEmpresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    WEmpresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
            
            ZZArticulo = ZArticulo
            
            If Left$(ZArticulo, 2) = "PT" Then
                ZZArticulo = "NK-" + Right$(ZArticulo, 9)
                    Else
                If Left$(ZArticulo, 2) = "DY" Then
                    ZZArticulo = "DK-" + Right$(ZArticulo, 9)
                        Else
                    If Left$(ZArticulo, 2) = "DQ" Then
                        ZZArticulo = "NQ-" + Right$(ZArticulo, 9)
                            Else
                        If Left$(ZArticulo, 2) = "DS" Then
                            ZZArticulo = "NS-" + Right$(ZArticulo, 9)
                        End If
                    End If
                End If
            End If
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM EntDev"
            ZSql = ZSql + " Where EntDev.Codigo = " + "'" + ZNroEntrada + "'"
            Rem ZSql = ZSql + " and EntDev.NroDev = " + "'" + ZNumero + "'"
            ZSql = ZSql + " and EntDev.Terminado = " + "'" + ZZArticulo + "'"
            spEntdev = ZSql
            Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            If rstEntdev.RecordCount > 0 Then
                ZFechaEntrada = rstEntdev!Fecha
                ZNroPedido = rstEntdev!Pedido
                ZLoteOriginal = rstEntdev!Lote
                rstEntdev.Close
                Exit For
            End If
            
        Next CiclaEmpre
        
        Call Conecta_Empresa
                
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PedidoDevol"
        ZSql = ZSql + " Where PedidoDevol.Pedido = " + "'" + Str$(ZNroPedido) + "'"
        ZSql = ZSql + " and PedidoDevol.NroDev = " + "'" + ZNroEntrada + "'"
        spPedidoDevol = ZSql
        Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedidoDevol.RecordCount > 0 Then
            ZFechaPedido = rstPedidoDevol!Fecha
            ZObservaDevol = rstPedidoDevol!Observaciones
            rstPedidoDevol.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Estadistica SET "
        ZSql = ZSql + " NroEntrada = " + "'" + ZNroEntrada + "',"
        ZSql = ZSql + " FechaEntrada = " + "'" + ZFechaEntrada + "',"
        ZSql = ZSql + " NroPedido = " + "'" + Str$(ZNroPedido) + "',"
        ZSql = ZSql + " FechaPedido = " + "'" + ZFechaPedido + "',"
        ZSql = ZSql + " ObservaDevol = " + "'" + ZObservaDevol + "',"
        ZSql = ZSql + " LoteOriginal = " + "'" + Str$(ZLoteOriginal) + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
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
    
    Listado.WindowTitle = "Listado de Analisis de Devoluciones"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Estadistica.Tipo, Estadistica.Numero, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Cliente, Estadistica.Linea, Estadistica.Fecha, Estadistica.OrdFecha, Estadistica.Lote1, Estadistica.Tipopro, Estadistica.NroEntrada, Estadistica.FechaEntrada, Estadistica.NroPedido, Estadistica.FechaPedido, Estadistica.LoteOriginal, Estadistica.ObservaDevol, Estadistica.Entrada, " _
                + "Cliente.Razon " _
                + "From " _
                + DSQ + ".dbo.Estadistica Estadistica, " _
                + DSQ + ".dbo.Cliente Cliente " _
                + "Where " _
                + "Estadistica.Cliente = Cliente.Cliente AND " _
                + "Estadistica.Tipo = 2 AND " _
                + "Estadistica.Linea <> 50 AND " _
                + "Estadistica.OrdFecha >= '" + WDesde + "' AND " _
                + "Estadistica.OrdFecha <= '" + WHasta + "'"
     
    Listado.Connect = Connect()
    
    Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Dos = " and {Estadistica.Tipo} = 2"
    Tres = " and {Estadistica.Linea} <> 50"
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    Listado.SelectionFormula = Uno + Dos + Tres
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    PrgListaDevol.Hide
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
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub





