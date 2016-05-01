VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgModifTerminado 
   AutoRedraw      =   -1  'True
   Caption         =   "Asignacion de Pedidos de Pigmentos"
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   690
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11550
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11245
      _Version        =   327680
      Rows            =   4000
      Cols            =   7
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WImpreEtiDy.rpt"
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "PrgModifTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim XParam As String
Dim TotalPedidos As Integer
Dim WGraba As String
Dim WSaldo As Double
Dim xLote(100, 30) As String
Dim WLote(100, 5) As String
Dim WCanti(100, 5) As String
Dim WEti(100, 5) As String
Dim WTipo(100, 5) As String
Dim Vector(1000, 3) As String

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgModifTerminado.Hide
    Unload Me
    Close
    End
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 200
    Muestra.ColWidth(1) = 1200
    Muestra.ColWidth(2) = 1500
    Muestra.ColWidth(3) = 1200
    Muestra.ColWidth(4) = 3000
    Muestra.ColWidth(5) = 1500
    Muestra.ColWidth(6) = 2000
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Call Proceso_Click
    
End Sub

Private Sub Proceso_Click()

    WSalida = "N"
    
    Call Limpia_Vector
    
    Renglon = 0
    WSaldo = 0
    Lugar = 0
    
    Pasa = 0
    Pedido = ""
    Fecha = "  /  /    "
    Cliente = ""
    Razon = ""
    FEntrega = "  /  /    "
    Tipo = 0
    Importe = 0
    Estado = ""
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select *"
    Rem ZSql = ZSql + " FROM Pedido"
    Rem ZSql = ZSql + " Where Pedido.Pedido = " + "'" + "350635" + "'"
    Rem spPedido = ZSql
    
    Rem spPedido = "ListaPedidoTotalListado4"
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPedido.RecordCount > 0 Then
    
    ZSql = ""
    ZSql = ZSql + "Select Pedido.Proceso1, Pedido.TipoPedido, Pedido.Pedido, Pedido.TipoPed, Pedido.Impresion2, Pedido.autorizo, Pedido.fecha, Pedido.cliente, Pedido.fecentrega, Pedido.impresion "
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.Impresion2 <> 'S'"
    ZSql = ZSql + " and Pedido.Autorizo = 'X'"
    ZSql = ZSql + " and Pedido.TipoPedido = 5"
    ZSql = ZSql + " Order by Pedido.Pedido"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    If rstPedido!TipoPedido = 5 Then
                    
                        Entra = "S"
                                    
                        For XDa = 1 To Lugar
                            If Vector(Lugar, 1) = rstPedido!Pedido Then
                                Entra = "N"
                                Exit For
                            End If
                        Next XDa
                                        
                        If Entra = "S" Then
                        
                            Lugar = Lugar + 1
                            
                            Vector(Lugar, 1) = rstPedido!Pedido
                            Vector(Lugar, 2) = "1"
                            Vector(Lugar, 3) = rstPedido!Tipoped
                            
                            Corte = rstPedido!Pedido
                            Fecha = rstPedido!Fecha
                            Cliente = rstPedido!Cliente
                            FEntrega = rstPedido!FecEntrega
                            Tipo = rstPedido!Tipoped
                            Importe = 0
                            Estado = rstPedido!autorizo
                            Impresa = rstPedido!Impresion
                            
                            Muestra.Row = Lugar
                            
                            Muestra.Col = 1
                            Muestra.Text = Pusing("######", Str$(rstPedido!Pedido))
                            
                            Muestra.Col = 2
                            Muestra.Text = rstPedido!Fecha
                    
                            Muestra.Col = 3
                            Muestra.Text = rstPedido!Cliente
                            
                            Muestra.Col = 5
                            Muestra.Text = FEntrega
                            
                            Select Case Tipo
                                Case 0
                                    Muestra.Col = 6
                                    Muestra.Text = "Normal"
                                Case 1
                                    Muestra.Col = 6
                                    Muestra.Text = "A Fecha"
                                Case 2
                                    Muestra.Col = 6
                                    Muestra.Text = "Fecha LImite"
                                Case 3
                                    Muestra.Col = 6
                                    Muestra.Text = "Urgente"
                                Case 4
                                    Muestra.Col = 6
                                    Muestra.Text = "Retira Cliente"
                                Case 5
                                    Muestra.Col = 6
                                    Muestra.Text = "Muestra"
                                Case Else
                                    Muestra.Col = 6
                                    Muestra.Text = ""
                            End Select
                            
                        End If
                    
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
    
        rstPedido.Close
    
    End If
    
    For dada = 1 To Lugar
    
        Muestra.Row = dada
                        
        Muestra.Col = 3
        WCliente = Muestra.Text
    
        spCliente = "ConsultaCliente " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Muestra.Col = 4
            Muestra.Text = rstCliente!Razon
            rstCliente.Close
        End If
        
    Next dada
    
    TotalPedidos = Renglon
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    
    Muestra.TopRow = 1
    
    Rem Muestra.SetFocus

End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
End Sub

Private Sub Muestra_DblClick()

    Muestra.Col = 1
    WXPed = Muestra.Text
    
    PrgModifTerminado.Hide
    Unload Me
    PrgModPedTerminado.Show
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    If ProcesoActivate = 1 Then
        Call Proceso_Click
    End If
End Sub

