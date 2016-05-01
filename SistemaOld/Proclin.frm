VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProclin 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Estadisitica de Ventas por Rubro y Clientes"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2775
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeFec 
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Rubro"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Rubro"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEsta1.rpt"
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "Proclin.frx":0000
      Left            =   840
      List            =   "Proclin.frx":0007
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgProclin"
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
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim rstLinea As Recordset
Dim spLinea As String
Dim rstRubro As Recordset
Dim spRubro As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Dim Vector(10000, 2) As String
Private Vecosto(5000, 2) As String
Dim Posi As Integer

Private Sub Acepta_Click()

    On Error GoTo WError
    
    WDesde = "00000000"
    Whasta = "99999999"
    
    XParam = "'" + WDesde + "','" _
                 + Whasta + "'"
    spEstadistica = "ListaEstadisticaFecha " + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            Do
            
                WClave = rstEstadistica!Clave
                WLinea = rstEstadistica!Linea
                WArticulo = rstEstadistica!Articulo
                
                If Val(WLinea) = 0 Then
                
                    Renglon% = Renglon% + 1
                    
                    Vector(Renglon%, 1) = WClave
                    Vector(Renglon%, 2) = WArticulo
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
        
        
        rstEstadistica.Close
        
    End If
    
    For da% = 1 To Renglon
    
        WClave = Vector(da%, 1)
        WArticulo = Vector(da%, 2)
        WLinea = "0"
    
        XParam = "'" + WArticulo + "'"
        spTerminado = "ConsultaTerminado" + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WLinea = Str$(rstTerminado!Linea)
            rstTerminado.Close
            
                XParam = "'" + WClave + "','" _
                            + WLinea + "'"
                    
                    spEstadistica = "ModificaEstadisticaLinea " + XParam
                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        End If
    Next da%
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
    Desde.SetFocus
    PrgEsta1.Hide
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeFec.SetFocus
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
End Sub

Sub Form_Load()
    Desde.Text = "0"
    Hasta.Text = "9999"
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub


Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    lugar = 1
    Cicla = 0
    
    For da = 1 To Posi
        If Producto = Vecosto(da, 1) Then
            Costo = Val(Vecosto(da, 2))
            Exit Sub
        End If
    Next da
    
    Do
        Cicla = Cicla + 1
        If Vector(Cicla, 1) <> "" Then
    
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    lugar = lugar + 1
                                    Vector(lugar, 1) = Articulo2
                                    Vector(lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = Cantidad
                                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                            Case Else
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For da = 1 To Renglon
        Articulo = Auxiliar(da, 1)
        Cantidad = Auxiliar(da, 2)
        XVector = Auxiliar(da, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = (Cantidad * rstArticulo!Costo2 * Val(XVector))
            Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(XVector))
            rstArticulo.Close
        End If
    Next da
    
    Posi = Posi + 1
    Vecosto(Posi, 1) = Producto
    Vecosto(Posi, 2) = Str$(Costo)
    
End Sub


Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Esta
End Sub

