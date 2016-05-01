VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCotprv 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cotizaciones por Proveedor"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1815
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Hasta 
         Height          =   285
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   12
         Text            =   "  "
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Desde 
         Height          =   285
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   11
         Text            =   " "
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Cotprv.rpt"
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
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "PrgCotprv.frx":0000
      Left            =   840
      List            =   "PrgCotprv.frx":0007
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCoTPRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Vector(3, 4) As String

Private Sub Acepta_Click()

    Da = 0
    With rstLiscot
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
    
    Pasa = 0
    Canti = 0
            
    With rstCotiza
    
            .Index = "Proveedor"
            .Seek ">=", Desde.Text
            
            Do
            
                If !Proveedor > Hasta.Text Then
                    Exit Do
                End If
                
                WArticulo = !Articulo
                WProveedor = !Proveedor
                WFecha = !Fecha
                WPrecio = !Precio
                WCondicion = !Condicion
                WObservaciones = !Observaciones
                
                If Pasa = 0 Then
                    Pasa = 1
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase Vector
                    Canti = 0
                End If
                
                If Corte1 <> !Proveedor Or Corte2 <> !Articulo Then
                
                    With rstLiscot
                
                        For Da = 1 To 3
                        
                            If Vector(Da, 1) <> "" Then
                                .AddNew
                                !Proveedor = Corte1
                                !Articulo = Corte2
                                !Fecha = Vector(Da, 1)
                                !Precio = Val(Vector(Da, 2))
                                !Condicion = Vector(Da, 3)
                                !Observaciones = Vector(Da, 4)
                                !Clave = !Proveedor + !Articulo
                                .Update
                            End If
                            
                        Next Da
                            
                    End With
                    
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase Vector
                    Canti = 0
                    
                End If
                
                Canti = Canti + 1
                
                If Canti > 3 Then
                    For Da = 1 To 2
                        Vector(Da, 1) = Vector(Da + 1, 1)
                        Vector(Da, 2) = Vector(Da + 1, 2)
                        Vector(Da, 3) = Vector(Da + 1, 3)
                        Vector(Da, 4) = Vector(Da + 1, 4)
                    Next Da
                    Canti = 3
                End If
                
                Vector(Canti, 1) = !Fecha
                Vector(Canti, 2) = Str$(!Precio)
                Vector(Canti, 3) = !Condicion
                Vector(Canti, 4) = !Observaciones
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                If !Proveedor > Hasta.Text Then
                    Exit Do
                End If
                
            Loop
    End With
    
    If Pasa <> 0 Then
        With rstLiscot
                
            For Da = 1 To 3
                    
                If Vector(Da, 1) <> "" Then
                    .AddNew
                    !Proveedor = Corte1
                    !Articulo = Corte2
                    !Fecha = Vector(Da, 1)
                    !Precio = Val(Vector(Da, 2))
                    !Condicion = Vector(Da, 3)
                    !Observaciones = Vector(Da, 4)
                    !Clave = !Proveedor + !Articulo
                    .Update
                End If
                
            Next Da
                        
        End With
    End If

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            .Update
        End If
    End With

    Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
    
    Listado.WindowTitle = "Listado de Cotizaciones por Proveedor"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Listcot.proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstArticulo
        .Close
    End With
    With rstProveedor
        .Close
    End With
    With rstCotiza
        .Close
    End With
    With rstLiscot
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    DbsVentas.Close
    Desde.SetFocus
    PrgCoTPRV.Hide
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    Desde.Text = "0"
    Hasta.Text = "99999999999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub


