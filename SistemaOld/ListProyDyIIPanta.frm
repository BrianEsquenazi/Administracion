VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListProyDyIIPanta 
   Caption         =   "Acutalizacion de Proyeccion de Ordenes de Compra"
   ClientHeight    =   8475
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8475
   ScaleWidth      =   11880
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   15
      Top             =   2280
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   3240
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3840
      TabIndex        =   13
      Top             =   2280
      Width           =   375
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4440
      TabIndex        =   16
      Top             =   2280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   6495
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11456
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton Actualiza 
         Caption         =   "Actualiza"
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
         Left            =   7440
         MaskColor       =   &H00000000&
         TabIndex        =   29
         Top             =   1080
         Width           =   1335
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   3840
         TabIndex        =   6
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
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
         Left            =   9840
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Top             =   960
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
         Left            =   9840
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hastafecha 
         Height          =   300
         Left            =   3840
         TabIndex        =   7
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   300
         Left            =   2280
         TabIndex        =   8
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox FechaOrden 
         Height          =   300
         Left            =   7440
         TabIndex        =   10
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin VB.Label Label2 
         Caption         =   "Desde Fecha de Proyeccion de Ordenes de Compra"
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
         Height          =   615
         Left            =   5400
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Desde-Hasta Periodo de Venta"
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
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Desde-Hasta Articulo"
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
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6600
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListProyDyIIPanta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vector(10000, 4) As String
Dim ZZVector(10000, 4) As String
Dim WFechaCierre As String
Dim WOrdFechaCierre As String
Private WArticulo As String
Private WInicial As Double
Private WEntradas As Double
Private WSalidas As Double
Private WSaldo As Double
Private XLote(100, 7) As String
Private WCodigo As String
Dim Empe(12, 10) As String

Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String

Private VentaMes(100, 2) As String
Private LugarMes As Integer
Private WFechaDesde As String
Private WFechaHasta As String
Dim WImpre(6, 3) As String
Dim WMes1 As String
Dim WAno1 As String
Dim WMes2 As String
Dim WAno2 As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub Acepta_Click()

    On Error GoTo WError
    
    WFechaDesde = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    WFechaHasta = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)
    
    If WFechaDesde > WFechaHasta Then
        Exit Sub
    End If
    
    WMes1 = Mid$(WFechaDesde, 5, 2)
    WAno1 = Mid$(WFechaDesde, 1, 4)
    
    WMes2 = Mid$(WFechaHasta, 5, 2)
    WAno2 = Mid$(WFechaHasta, 1, 4)
    
    Call Ceros(WMes1, 2)
    Call Ceros(WMes2, 2)
    Call Ceros(WAno1, 4)
    Call Ceros(WAno2, 4)
    
    Erase VentaMes
    LugarMes = 0
   
    Do
    
        LugarMes = LugarMes + 1
        VentaMes(LugarMes, 1) = WAno1 + WMes1
        
        WMes1 = Str$(Val(WMes1) + 1)
        If Val(WMes1) > 12 Then
            WAno1 = Str$(Val(WAno1) + 1)
            WMes1 = "1"
        End If
        
        Call Ceros(WMes1, 2)
        Call Ceros(WAno1, 4)
        
        If WAno1 + WMes1 > WAno2 + WMes2 Then
            Exit Do
        End If
        
    Loop
    
    WMes1 = Mid$(FechaOrden.Text, 4, 2)
    WAno1 = Right$(FechaOrden.Text, 4)
    
    Call Ceros(WMes1, 2)
    Call Ceros(WAno1, 4)
    
    Erase WImpre
    
    For Ciclo = 1 To 6
    
        If Val(WMes1) > 12 Then
            WAno1 = Str$(Val(WAno1) + 1)
            WMes1 = "1"
        End If
        
        Call Ceros(WMes1, 2)
        Call Ceros(WAno1, 4)
        WImpre(Ciclo, 1) = WMes1 + "/" + WAno1
        WImpre(Ciclo, 2) = WMes1
        WImpre(Ciclo, 3) = WAno1
        WMes1 = Str$(Val(WMes1) + 1)
        
    Next Ciclo
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    With rstStockDy
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
    
    
    XEmpresa = WEmpresa
    
    
    Rem
    Rem calcula los pedidos pendientess
    Rem
    
    spArticulo = "ModificaArticuloVenta0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spPedido = "ModificaPedpenDy " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spPedido = "ListaPedidoPendDesdeHasta " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    EntraVector = "S"
                    For Ciclo = 1 To Renglon
                        If Vector(Ciclo, 1) = rstPedido!Terminado Then
                            Vector(Ciclo, 2) = Str$(Val(Vector(Ciclo, 2)) + rstPedido!Importe)
                            EntraVector = "N"
                            Exit For
                        End If
                    Next Ciclo
                    If EntraVector = "S" Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = rstPedido!Terminado
                        Vector(Renglon, 2) = Str$(rstPedido!Importe)
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    For Ciclo = 1 To Renglon
        WProducto = Vector(Ciclo, 1)
        WImporte = Vector(Ciclo, 2)
        WArticulo = Left$(WProducto, 3) + Right$(WProducto, 7)
        XParam = "'" + WArticulo + "','" _
                     + WImporte + "','" _
                     + WDate + "'"
        spArticulo = "ModificaArticuloVenta " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo
    
    
    
    
    Rem
    Rem recalcula saldo de orden de compra
    Rem
    
    
    
    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        XHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        XHasta = 4
    End If
    
    For CicloEmpresa = 1 To XHasta
    
        WEmpresa = Empe(CicloEmpresa, 1)
        txtOdbc = Empe(CicloEmpresa, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    
    
        Erase Vector
        Renglon = 0
        
    
        XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
        spArticulo = "ModificaArticuloPedido0DesdeHasta" + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        
        ZSql = ""
        ZSql = ZSql + "Select Orden.Clave, Orden.Orden, Orden.Articulo, Orden.Cantidad, Orden.FechaOrd, Orden.Recibida"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Articulo >= " + "'" + Desde.Text + "'"
        ZSql = ZSql + " and Orden.Articulo <= " + "'" + Hasta.Text + "'"
        ZSql = ZSql + " Order by Orden.Orden"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                        If !FechaOrd > "20020101" Then
                            ZSaldo = rstOrden!Cantidad - rstOrden!Recibida
                            If (rstOrden!Orden < 900000) Or (rstOrden!Orden > 900000 And ZSaldo > 0) Then
                                Renglon = Renglon + 1
                                Vector(Renglon, 1) = rstOrden!Clave
                                Vector(Renglon, 2) = rstOrden!Orden
                                Vector(Renglon, 3) = rstOrden!Articulo
                                Vector(Renglon, 4) = rstOrden!Cantidad
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
        End If
    
    
        For Ciclo = 1 To Renglon
    
            WClave = Vector(Ciclo, 1)
            WOrden = Vector(Ciclo, 2)
            WArticulo = Vector(Ciclo, 3)
            WCantidad = Val(Vector(Ciclo, 4))
            WResta = 0
            
            XParam = "'" + WOrden + "','" _
                         + WArticulo + "'"
            spInforme = "ListaInformeOrdenArticulo" + XParam
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                With rstInforme
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                            WResta = WResta + rstInforme!Resta
                            .MoveNext
                            If .EOF = True Then
                                Exit Do
                            End If
                        Loop
                    End If
                End With
                rstInforme.Close
            End If
            
            WDife = WCantidad - WResta
            If WDife > 0 Then
                XParam = "'" + WArticulo + "','" _
                             + Str$(WDife) + "'"
                spArticulo = "ModificaArticuloPedido " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
        
        Next Ciclo
    
    
    Next CicloEmpresa
    
    Call Conecta_Empresa
    
    
    
    
    
    
    
    
    
    Rem
    Rem recalcula saldo de informes de recepcion
    Rem
    
    
    
    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        XHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        XHasta = 4
    End If
    
    For CicloEmpresa = 1 To XHasta
    
        WEmpresa = Empe(CicloEmpresa, 1)
        txtOdbc = Empe(CicloEmpresa, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        Erase Vector
        Renglon = 0
    
        WDesde = Right$(FechaOrden.Text, 4) + "0101"
        WHasta = Right$(FechaOrden.Text, 4) + "1231"
    
        spInforme = "ModificaInformeProcesoSaldo"
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
        spArticulo = "ModificaArticuloLaboratorio0"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
        XParam = "'" + "20020101" + "'"
        spInforme = "ModificaInformeProceso0 " + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
        XParam = "'" + WDesde + "','" _
                     + WHasta + "'"
        spInforme = "ListaInformeDesdeHastaFecha" + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            With rstInforme
                .MoveFirst
                Do
                    If .EOF = False Then
                        If !FechaOrd > "20020101" Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstInforme!Clave
                            Vector(Renglon, 2) = rstInforme!Informe
                            Vector(Renglon, 3) = rstInforme!Articulo
                            Vector(Renglon, 4) = rstInforme!Cantidad
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstInforme.Close
        End If
    
        For Ciclo = 1 To Renglon
    
            WClave = Vector(Ciclo, 1)
            WInforme = Vector(Ciclo, 2)
            WArticulo = Vector(Ciclo, 3)
            WCantidad = Val(Vector(Ciclo, 4))
            WResta = 0
            
            XParam = "'" + WInforme + "','" _
                         + WArticulo + "'"
            spLaudo = "ListaLaudoInforme " + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
    
                With rstLaudo
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                        Do
            
                            If .EOF = True Then
                                Exit Do
                            End If
                        
                            WLiberada = rstLaudo!Liberada
                            WDevuelta = rstLaudo!devuelta
                            WSuma = WLiberada + WDevuelta
                        
                            WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                            WDevueltaAnt = IIf(IsNull(rstLaudo!devueltaant), "0", rstLaudo!devueltaant)
                            WSumaAnt = WLiberadaAnt + WDevueltaAnt
                    
                            If WSumaAnt <> 0 Then
                                WResta = WResta + WSumaAnt
                                    Else
                                WResta = WResta + WSuma
                            End If
                
                            .MoveNext
                
                            If .EOF = True Then
                                Exit Do
                            End If
                
                        Loop
                    End If
                End With
                rstLaudo.Close
            End If
        
            XParam = "'" + WClave + "','" _
                         + Str$(WResta) + "'"
            spInforme = "ModificaInformeProceso " + XParam
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
            WDife = WCantidad - WResta
            If WDife > 0 Then
                XParam = "'" + WArticulo + "','" _
                             + Str$(WDife) + "'"
                spArticulo = "ModificaArticuloLaboratorio " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
        
        Next Ciclo
    
        spInforme = "ModificaInformeProcesoDife"
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
    Next CicloEmpresa
    
    Call Conecta_Empresa
    
    
    
    
    
    
    Rem
    Rem recalcula saldo del stock
    Rem
    
    
    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        XHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        XHasta = 4
    End If
    
    For CicloEmpresa = 1 To XHasta
    
        WEmpresa = Empe(CicloEmpresa, 1)
        txtOdbc = Empe(CicloEmpresa, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        
        
        Erase Vector
        Renglon = 0
        
    
        XParam = "'" + Desde.Text + "','" _
                     + Hasta.Text + "'"
        spArticulo = "ListaArticuloDesdeHasta " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        With rstArticulo
                .MoveFirst
                Do
                    If .EOF = True Then
                        Exit Do
                    End If
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstArticulo!Codigo
                    Vector(Renglon, 2) = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
                    Vector(Renglon, 3) = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
        End With
        rstArticulo.Close
    
        For Da = 1 To Renglon
    
            WEntradas = 0
            WSalidas = 0
            WArticulo = Vector(Da, 1)
            XCodigo = Vector(Da, 1)
            WFechaCierre = Vector(Da, 2)
            WOrdFechaCierre = Vector(Da, 3)
            XDate = Date$
        
            Call calcula_datos
        
            XEntradas = Str$(WEntradas)
            XSalidas = Str$(WSalidas)
        
            XParam = "'" + XCodigo + "','" _
                         + XEntradas + "','" _
                         + XSalidas + "','" _
                         + XDate + "'"
                                           
            spArticulo = "ModificaArticuloMovimientos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Next Da
    
    Next CicloEmpresa
    
    Call Conecta_Empresa
    
    
    Rem
    Rem reproceso de calculo de datos para el listado
    Rem
    
    
    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        XHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        XHasta = 4
    End If
    
    For CicloEmpresa = 1 To XHasta
    
        WEmpresa = Empe(CicloEmpresa, 1)
        txtOdbc = Empe(CicloEmpresa, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    
        XParam = "'" + Desde.Text + "','" _
                     + Hasta.Text + "'"
    
        spArticulo = "ListaArticuloDesdeHasta" + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            
            With rstArticulo
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        WArticulo = rstArticulo!Codigo
                        WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas + rstArticulo!Laboratorio
                        WDescripcion = rstArticulo!Descripcion
                        WCodigoDy = IIf(IsNull(rstArticulo!CodigoDy), "", rstArticulo!CodigoDy)
                        WOrden = rstArticulo!Pedido
                        WPedido = rstArticulo!Venta
                        If Left$(rstArticulo!Codigo, 2) = "DQ" Then
                            WFamilia = Val(Mid$(rstArticulo!Codigo, 4, 1)) + 1
                                Else
                            WFamilia = Mid$(WArticulo, 4, 3)
                        End If
                    
                        With rstStockDy
                            .Index = "Codigo"
                            .Seek "=", WArticulo
                            If .NoMatch = True Then
                                .AddNew
                                !Codigo = WArticulo
                                !Descripcion = WDescripcion
                                !Familia = WFamilia
                                !Desfamilia = WCodigoDy
                                !Stock1 = 0
                                !Stock2 = 0
                                !Stock3 = 0
                                !Stock4 = 0
                                !Stock5 = 0
                                !Orden1 = 0
                                !Orden2 = 0
                                !Orden3 = 0
                                !Orden4 = 0
                                !Orden5 = 0
                                !Orden6 = 0
                                !Venta1 = 0
                                !Venta2 = 0
                                !Venta3 = 0
                                !Impre1 = WImpre(1, 1)
                                !Impre2 = WImpre(2, 1)
                                !Impre3 = WImpre(3, 1)
                                !Impre4 = WImpre(4, 1)
                                !Impre5 = WImpre(5, 1)
                                !Impre6 = WImpre(6, 1)
                                !Stock1 = !Stock1 + WStock
                                !Titulo1 = "Surfactan S.A."
                                !Titulo2 = "desde el " + DesdeFecha.Text + " al " + HastaFecha.Text
                                !Orden = WOrden
                                !Pedido = WPedido
                                .Update
                                    Else
                                .Edit
                                !Stock1 = !Stock1 + WStock
                                !Orden = !Orden + WOrden
                                !Pedido = !Pedido + WPedido
                                .Update
                            End If
                        End With
                
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            
            rstArticulo.Close
            
        End If
        
    Next CicloEmpresa
    
    Call Conecta_Empresa
    
    
    With rstStockDy
        .Index = "Codigo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                
                WCodigo = !Codigo
                
                For Ciclo = 1 To 100
                    VentaMes(Ciclo, 2) = 0
                Next Ciclo
                
                Call Calcula_Ventas
                
                SumaVenta = 0
                VentaMayor = 0
                For Ciclo = 1 To LugarMes
                    SumaVenta = SumaVenta + Val(VentaMes(Ciclo, 2))
                    If Val(VentaMes(Ciclo, 2)) > Val(VentaMayor) Then
                        VentaMayor = VentaMes(Ciclo, 2)
                    End If
                Next Ciclo
                    
                .Edit
                
                !Venta1 = Val(VentaMes(LugarMes, 2))
                !Venta2 = SumaVenta / LugarMes
                !Venta3 = Val(VentaMayor)
                
                For Ciclo = 1 To 6
                
                    XMes = WImpre(Ciclo, 2)
                    XAno = WImpre(Ciclo, 3)
                    XOrden = 0
                    
                    Sql1 = "Select *"
                    Sql2 = " FROM AltaProyec"
                    Sql3 = " Where AltaProyec.Ano = " + "'" + XAno + "'"
                    Sql4 = " and AltaProyec.Codigo = " + "'" + WCodigo + "'"
                    spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                    Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    If rstAltaProyec.RecordCount > 0 Then
                        Select Case Val(XMes)
                            Case 1
                                XOrden = rstAltaProyec!Mes1
                            Case 2
                                XOrden = rstAltaProyec!Mes2
                            Case 3
                                XOrden = rstAltaProyec!Mes3
                            Case 4
                                XOrden = rstAltaProyec!Mes4
                            Case 5
                                XOrden = rstAltaProyec!Mes5
                            Case 6
                                XOrden = rstAltaProyec!Mes6
                            Case 7
                                XOrden = rstAltaProyec!Mes7
                            Case 8
                                XOrden = rstAltaProyec!Mes8
                            Case 9
                                XOrden = rstAltaProyec!Mes9
                            Case 10
                                XOrden = rstAltaProyec!Mes10
                            Case 11
                                XOrden = rstAltaProyec!Mes11
                            Case 12
                                XOrden = rstAltaProyec!Mes12
                            Case Else
                        End Select
                        rstAltaProyec.Close
                    End If
                        
                    Select Case Ciclo
                        Case 1
                            !Orden1 = XOrden
                        Case 2
                            !Orden2 = XOrden
                        Case 3
                            !Orden3 = XOrden
                        Case 4
                            !Orden4 = XOrden
                        Case 5
                            !Orden5 = XOrden
                        Case 6
                            !Orden6 = XOrden
                        Case Else
                    End Select
                    
                Next Ciclo
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    Call Limpia_Vector
    
    LugarVector = 0
    
    With rstStockDy
        .Index = "Codigo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                
                LugarVector = LugarVector + 1
                
                WVector1.TextMatrix(LugarVector, 1) = !Codigo
                WVector1.TextMatrix(LugarVector, 2) = !Descripcion
                WVector1.TextMatrix(LugarVector, 3) = Pusing("###,###", Str$(!Venta1))
                WVector1.TextMatrix(LugarVector, 4) = Pusing("###,###", Str$(!Venta2))
                WVector1.TextMatrix(LugarVector, 5) = Pusing("###,###", Str$(!Venta3))
                WVector1.TextMatrix(LugarVector, 6) = Pusing("###,###", Str$(!Stock1 + !Stock2 + !Stock3 + !Stock4 + !Stock5 + !Orden - !Pedido))
                WVector1.TextMatrix(LugarVector, 7) = Pusing("###,###", Str$(!Orden1))
                WVector1.TextMatrix(LugarVector, 8) = Pusing("###,###", Str$(!Orden2))
                WVector1.TextMatrix(LugarVector, 9) = Pusing("###,###", Str$(!Orden3))
                WVector1.TextMatrix(LugarVector, 10) = Pusing("###,###", Str$(!Orden4))
                WVector1.TextMatrix(LugarVector, 11) = Pusing("###,###", Str$(!Orden5))
                WVector1.TextMatrix(LugarVector, 12) = Pusing("###,###", Str$(!Orden6))
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    WVector1.Row = 1
    WVector1.Col = 7
    WVector1.TopRow = 1
    Call StartEdit
    
    Exit Sub

WError:

    Resume Next
    
End Sub


Private Sub Calcula_Ventas()

    Rem PROCESA LAS VENTAS
    
    XParam = "'" + WCodigo + "','" _
                 + WCodigo + "'"
    spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEstadistica!OrdFecha >= WFechaDesde And rstEstadistica!OrdFecha <= WFechaHasta Then
                
                    WCantidad = rstEstadistica!Cantidad
                    WTipo = rstEstadistica!Tipo
                    
                    WCompara = Left$(rstEstadistica!OrdFecha, 6)
                    For Ciclo = 1 To LugarMes
                        If VentaMes(Ciclo, 1) = WCompara Then
                            If WTipo = 1 Then
                                VentaMes(Ciclo, 2) = Str$(Val(VentaMes(Ciclo, 2)) + WCantidad)
                                    Else
                                VentaMes(Ciclo, 2) = Str$(Val(VentaMes(Ciclo, 2)) - WCantidad)
                            End If
                            Exit For
                        End If
                    Next Ciclo
                    
                End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
        rstEstadistica.Close
    End If
    
    WParametro = Left$(WCodigo, 3) + "00" + Right$(WCodigo, 7)
    
    With rstEsta8
        .Index = "Articulo"
        .Seek ">=", WParametro
        If .NoMatch = False Then
            Do
            
                If WFechaDesde <= !OrdFecha And !OrdFecha <= WFechaHasta Then
                
                If WParametro = !Articulo Then
                
                    WCantidad = !Cantidad
                    WTipo = !Tipo
                    WPedido = !Pedido
                    If WPedido = 0 Then
                        WCantidad = 0
                    End If
                    
                    WCompara = Left$(!OrdFecha, 6)
                    For Ciclo = 1 To LugarMes
                        If VentaMes(Ciclo, 1) = WCompara Then
                            If WTipo = 1 Then
                                VentaMes(Ciclo, 2) = Str$(Val(VentaMes(Ciclo, 2)) + WCantidad)
                                    Else
                                VentaMes(Ciclo, 2) = Str$(Val(VentaMes(Ciclo, 2)) - WCantidad)
                            End If
                            Exit For
                        End If
                    Next Ciclo
                    
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                If WParametro <> !Articulo Then
                    Exit Do
                End If
                
            Loop
        End If
    End With
    
    
    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        XHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        XHasta = 4
    End If
    
    For a = 1 To XHasta
    
        Erase Vector
        Suma = 0
    
        WEmpresa = Empe(a, 1)
        txtOdbc = Empe(a, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Articulo = " + "'" + WCodigo + "'"
        ZSql = ZSql + " and Hoja.FechaOrd >= " + "'" + WFechaDesde + "'"
        ZSql = ZSql + " and Hoja.FechaOrd <= " + "'" + WFechaHasta + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveFirst
                Do
                    If .EOF = False Then
                            
                        WCantidad = rstHoja!Cantidad
                        WCompara = Left$(rstHoja!FechaOrd, 6)
                        
                        For Ciclo = 1 To LugarMes
                            If VentaMes(Ciclo, 1) = WCompara Then
                                VentaMes(Ciclo, 2) = Str$(Val(VentaMes(Ciclo, 2)) + WCantidad)
                                Exit For
                            End If
                        Next Ciclo
                                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHoja.Close
        End If
                
    Next a
    
    Call Conecta_Empresa
    
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    With rstStockDy
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgListProyDyIIPanta.Hide
    Unload Me
    Menu.Show
 End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_StockDy
    OPEN_FILE_Esta8
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        DesdeFecha.SetFocus
    End If
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaOrden.SetFocus
    End If
End Sub

Private Sub FechaOrden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Acepta_Click
        Rem Desde.SetFocus
    End If
End Sub

Sub Form_Load()

    Call Limpia_Vector

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgListProyDyIIPanta.Caption = "Actualizacion de Proyeccion de Ordenes de Compra :  " + !Nombre
        End If
    End With
    
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    FechaOrden.Text = "  /  /    "
    
    Frame2.Visible = True
End Sub


Private Sub calcula_datos()


    Rem PROCESA LOS LAUDOS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spLaudo = "ListaLaudoRepro" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                If rstLaudo!Marca = "X" And WSaldo = 0 Then
                        Else
                    If rstLaudo!Articulo = WArticulo Then
                        WEntradas = WEntradas + rstLaudo!Liberada
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstLaudo.Close
    End If
    
    
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spHoja = "ListaHojaRepro" + XParam
    Rem spHoja = "ListaHojaArticuloDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                        Else
                    If rstHoja!Tipo = "M" And rstHoja!Articulo = WArticulo Then
                        XX = rstHoja!Clave
                        WSalidas = WSalidas + rstHoja!Cantidad
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                If rstHoja!Articulo > WArticulo Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstHoja.Close
    End If
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovvar = "ListaMovvarRepro1" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
        With rstMovvar
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovvar!Marca = "X" Then
                        Else
                    If rstMovvar!Tipo = "M" And rstMovvar!Articulo = WArticulo Then
                        If rstMovvar!Movi = "E" Then
                            WEntradas = WEntradas + rstMovvar!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovvar!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstMovvar.Close
    End If
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovguia = "ListaMovguiaRepro1" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                        Else
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                        If rstMovguia!Movi = "E" Then
                            WEntradas = WEntradas + rstMovguia!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovguia!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
            
        End With
        rstMovguia.Close
    End If
    
    
    
    Rem PROCESA LAS HOJAS DE LABORATORIO
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spMovlab = "ListaMovlabRepro1" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
        With rstMovlab
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovlab!Marca = "X" Then
                        Else
                    If rstMovlab!Tipo = "M" And rstMovlab!Articulo = WArticulo Then
                        If rstMovlab!Movi = "E" Then
                            WEntradas = WEntradas + rstMovlab!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovlab!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstMovlab.Close
    End If
    
    
    
    Rem PROCESA LAS VENTAS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spEstadistica = "ListaEstadisticaReproDy " + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
        With rstEstadistica
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstEstadistica!Marca = "X" Then
                
                        Else
                        
                    WTipo = rstEstadistica!Tipo
                    
                    XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                    XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                    XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                    XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                    XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                    XLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                    XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                    XLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                    XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                    XLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                        
                    WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                     
                    If Len(Trim(WLoteAdicional)) = 98 Then
                    
                         XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                         XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                         XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                         XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                         XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                         XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                         XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                         XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                         XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                         XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                         XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                         XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                         XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                         XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                         
                             Else
                             
                         XLote(6, 1) = "0"
                         XLote(6, 2) = "0"
                         XLote(7, 1) = "0"
                         XLote(7, 2) = "0"
                         XLote(8, 1) = "0"
                         XLote(8, 2) = "0"
                         XLote(9, 1) = "0"
                         XLote(9, 2) = "0"
                         XLote(10, 1) = "0"
                         XLote(10, 2) = "0"
                         XLote(11, 1) = "0"
                         XLote(11, 2) = "0"
                         XLote(12, 1) = "0"
                         XLote(12, 2) = "0"
                         
                    End If
                        
                    For Da = 1 To 12
                        WLote = XLote(Da, 1)
                        WCantidad = XLote(Da, 2)
                        If Val(WCantidad) <> 0 Then
                            If WTipo = 2 Then
                                WEntradas = WEntradas + Abs(Val(WCantidad))
                                    Else
                                WSalidas = WSalidas + WCantidad
                            End If
                        End If
                    Next Da
                    
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstEstadistica.Close
    End If
    
End Sub















Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 12
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 7
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If WAuxi1 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For Da = 1 To WVector1.Cols - 1
            WVector1.Col = Da
            WVector1.Text = WBorra(Ciclo, Da)
        Next Da
    Next Ciclo
    
    End If
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 13
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 10
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Codigo"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Ult.Mes"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###"
            Case 4
                WVector1.Text = "Promedio"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###"
            Case 5
                WVector1.Text = "Mes May."
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###"
            Case 6
                WVector1.Text = "Dispon."
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###"
            Case 7
                WVector1.Text = WImpre(1, 1)
                WVector1.ColWidth(Ciclo) = 850
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###"
            Case 8
                WVector1.Text = WImpre(2, 1)
                WVector1.ColWidth(Ciclo) = 850
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###"
            Case 9
                WVector1.Text = WImpre(3, 1)
                WVector1.ColWidth(Ciclo) = 850
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###"
            Case 10
                WVector1.Text = WImpre(4, 1)
                WVector1.ColWidth(Ciclo) = 850
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###"
            Case 11
                WVector1.Text = WImpre(5, 1)
                WVector1.ColWidth(Ciclo) = 850
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###"
            Case 12
                WVector1.Text = WImpre(6, 1)
                WVector1.ColWidth(Ciclo) = 850
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###"
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub Actualiza_Click()

    For Ciclo = 1 To 1000
    
        If WVector1.TextMatrix(Ciclo, 1) <> "" Then
        
            WCodigo = WVector1.TextMatrix(Ciclo, 1)
            
            For Ciclo1 = 7 To 12
                
                WMes = Left$(WVector1.TextMatrix(0, Ciclo1), 2)
                WAno = Right$(WVector1.TextMatrix(0, Ciclo1), 4)
                WCantidad = WVector1.TextMatrix(Ciclo, Ciclo1)
                
                Sql1 = "Select *"
                Sql2 = " FROM AltaProyec"
                Sql3 = " Where AltaProyec.Ano = " + "'" + WAno + "'"
                Sql4 = " and AltaProyec.Codigo = " + "'" + WCodigo + "'"
                spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                If rstAltaProyec.RecordCount > 0 Then
                    rstAltaProyec.Close
                            Else
                    WCanti = ""
                    Sql1 = "INSERT INTO AltaProyec ("
                    Sql2 = "Ano ,"
                    Sql3 = "Codigo ,"
                    Sql4 = "Mes1 ,"
                    Sql5 = "Mes2 ,"
                    Sql6 = "Mes3 ,"
                    Sql7 = "Mes4 ,"
                    Sql8 = "Mes5 ,"
                    Sql9 = "Mes6 ,"
                    Sql10 = "Mes7 ,"
                    Sql11 = "Mes8 ,"
                    Sql12 = "Mes9 ,"
                    Sql13 = "Mes10 ,"
                    Sql14 = "Mes11 ,"
                    Sql15 = "Mes12 )"
                    Sql16 = "Values ("
                    Sql17 = "'" + WAno + "',"
                    Sql18 = "'" + WCodigo + "',"
                    Sql19 = "'" + WCanti + "',"
                    Sql20 = "'" + WCanti + "',"
                    Sql21 = "'" + WCanti + "',"
                    Sql22 = "'" + WCanti + "',"
                    Sql23 = "'" + WCanti + "',"
                    Sql24 = "'" + WCanti + "',"
                    Sql25 = "'" + WCanti + "',"
                    Sql26 = "'" + WCanti + "',"
                    Sql27 = "'" + WCanti + "',"
                    Sql28 = "'" + WCanti + "',"
                    Sql29 = "'" + WCanti + "',"
                    Sql30 = "'" + WCanti + "')"
                    spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                                   Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                                   Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30
                    Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
                Select Case WMes
                    Case 1
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes1 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 2
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes2 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 3
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes3 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 4
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes4 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 5
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes5 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 6
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes6 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 7
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes7 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 8
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes8 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 9
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes9 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 10
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes10 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 11
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes11 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case 12
                        Sql1 = "UPDATE AltaProyec SET "
                        Sql2 = " Mes12 = " + "'" + WCantidad + "'"
                        Sql3 = " Where Ano = " + "'" + WAno + "'"
                        Sql4 = " and Codigo = " + "'" + WCodigo + "'"
                        spAltaProyec = Sql1 + Sql2 + Sql3 + Sql4
                        Set rstAltaProyec = db.OpenRecordset(spAltaProyec, dbOpenSnapshot, dbSQLPassThrough)
                    Case Else
                End Select
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Cancela_click

End Sub







