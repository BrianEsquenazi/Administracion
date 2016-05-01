VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPeddev 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Solicitud de Devolucion de Mercaderia"
   ClientHeight    =   8625
   ClientLeft      =   90
   ClientTop       =   330
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   11850
   Visible         =   0   'False
   Begin VB.TextBox Total 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6165
      _Version        =   327680
      Rows            =   100
      Cols            =   6
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10200
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Observaciones 
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
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   9
      Text            =   " "
      Top             =   840
      Width           =   7935
   End
   Begin VB.TextBox Cliente 
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
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   6
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe Solicitud"
      Height          =   615
      Left            =   10200
      TabIndex        =   13
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Observaciones"
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
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
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
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Solicitud"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgPeddev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxi As String
Private WVector(6, 3) As String
Private XLinea As Single
Private WDirentrega As String
Private WInicio As Integer
Private Auxiliar(100, 2) As String
Private XSaldo As Double
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPedidoDevol As Recordset
Dim spPedidoDevol As String
Dim XParam As String
Dim ClavePedido(100)

Private Sub cmdClose_Click()

    With rstEmpresa
        .Close
    End With
    PrgPeddev.Hide
    Unload Me
    PrgAutoriza.Show
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Form_Load()

    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 1800
    Muestra.ColWidth(2) = 4000
    Muestra.ColWidth(3) = 1200
    Muestra.ColWidth(4) = 1200
    Muestra.ColWidth(5) = 1200
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Producto"
    
    Muestra.Col = 2
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 3
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 4
    Muestra.Text = "Saldo"
    
    Muestra.Col = 5
    Muestra.Text = "Precio"
    
    Pedido.Text = WXPed
    
    spPedidoDevol = "ListaPedidoDevol " + "'" + Pedido.Text + "'"
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoDevol.RecordCount > 0 Then
            Fecha.Text = rstPedidoDevol!Fecha
            Cliente.Text = rstPedidoDevol!Cliente
            Observaciones.Text = rstPedidoDevol!Observaciones
            rstPedidoDevol.Close
            
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WDirentrega = rstCliente!DirEntrega
                Observaciones.Text = rstCliente!Observaciones
                rstCliente.Close
            End If
            Call Proceso_Click
                Else
            WPedido = Pedido.Text
            Pedido.Text = WPedido
    End If
    
End Sub

Private Sub Proceso_Click()

    Muestra.Clear
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Producto"
    
    Muestra.Col = 2
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 3
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 4
    Muestra.Text = "Saldo"
    
    Muestra.Col = 5
    Muestra.Text = "Precio"
    
    
    Erase Auxiliar
    Erase ClavePedido
    
    Renglon = 0
    WRenglon = 0

    spPedidoDevol = "ListaPedidoDevol " + "'" + Pedido.Text + "'"
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoDevol.RecordCount > 0 Then
            With rstPedidoDevol
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        Renglon = Renglon + 1
                        
                        Muestra.Row = Renglon
                
                        Muestra.Col = 1
                        Muestra.Text = rstPedidoDevol!Terminado
                        Auxi1 = rstPedidoDevol!Terminado
                
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", rstPedidoDevol!Cantidad)
                
                        Muestra.Col = 4
                        Muestra.Text = Pusing("###,###.##", rstPedidoDevol!Cantidad - rstPedidoDevol!Facturado)
                        
                        Muestra.Col = 5
                        Muestra.Text = Pusing("###,###.##", rstPedidoDevol!Precio)
                        
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstPedidoDevol!Cliente
                        Auxiliar(WRenglon, 2) = rstPedidoDevol!Terminado
                        
                        ClavePedido(WRenglon) = rstPedidoDevol!Clave
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedidoDevol.Close
    End If
    
    Renglon = 0
    Total = 0
    
    For DA = 1 To WRenglon
    
        Cliente = Auxiliar(DA, 1)
        Terminado = Auxiliar(DA, 2)
        
        If Left$(Terminado, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Terminado, 3) + Right$(Terminado, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
        
                    Renglon = Renglon + 1
                    
                    Muestra.Row = Renglon
                
                    Muestra.Col = 2
                    Muestra.Text = rstArticulo!Descripcion
            
                    Muestra.Col = 4
                    Canti = Val(Muestra.Text)
            
                    Muestra.Col = 5
                    Precio = Val(Muestra.Text)
            
                    Total = Total + (Canti * Precio)
            
                End If
            
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
        
                    Renglon = Renglon + 1
                    
                    Muestra.Row = Renglon
                
                    Muestra.Col = 2
                    Muestra.Text = rstPrecios!Descripcion
            
                    Muestra.Col = 4
                    Canti = Val(Muestra.Text)
            
                    Muestra.Col = 5
                    Precio = Val(Muestra.Text)
            
                    Total = Total + (Canti * Precio)
            
                End If
        End Select
        
    Next DA
    
    Total.Text = Pusing("###,###.##", Str$(Total))
    Muestra.Row = 1

End Sub

