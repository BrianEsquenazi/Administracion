VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgDifeInvMat 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Diferencias de Inventario de Materia Prima"
   ClientHeight    =   6405
   ClientLeft      =   1410
   ClientTop       =   1155
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   9585
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1815
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   4815
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Aceptar 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WDifeInvMat.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgDifeInvMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WClave As String
Private WArticulo As String
Private WInicial As Double
Private WEntradas As Double
Private WSalidas As Double
Private WSaldo As Double
Private Vector(10000, 2) As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstProveedor As Recordset
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
Dim XParam As String

Private Sub Cancela_click()

    PrgDifeInvMat.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Aceptar_Click()

    On Error GoTo WError

    Da = 0
    With rstInve
        .Index = "Codigo"
        .MoveFirst
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


    Erase Vector
    Renglon = 0

    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    With rstArticulo

            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Renglon = Renglon + 1
                
                Vector(Renglon, 1) = rstArticulo!Codigo
                Vector(Renglon, 2) = rstArticulo!Descripcion
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
        
        Call calcula_datos
        
        XEntradas = Str$(WEntradas)
        XSalidas = Str$(WSalidas)
        
        
        With rstInve
            .Index = "Codigo"
            .Seek "=", WArticulo
            If .NoMatch = False Then
                .Edit
                !Codigo = WArticulo
                !Descripcion = Vector(Da, 2)
                !Stock = WEntradas - WSalidas
                .Update
                    Else
                .AddNew
                !Codigo = WArticulo
                !Descripcion = Vector(Da, 2)
                !Stock = WEntradas - WSalidas
                !Inve = 0
                .Update
            End If
        End With
        
    Next Da
    
    spInventario = "ListaInventarioTotal"
    Set rstInventario = db.OpenRecordset(spInventario, dbOpenSnapshot, dbSQLPassThrough)
    If rstInventario.RecordCount > 0 Then
        
    With rstInventario
    
        .MoveFirst
        
        If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstInventario!Tipo = "M" Then
                
                    WArticulo = rstInventario!Articulo
                    WCantidad = rstInventario!Cantidad
                
                    With rstInve
                        .Index = "Codigo"
                        .Seek "=", WArticulo
                        If .NoMatch = False Then
                            .Edit
                            !Codigo = WArticulo
                            !Inve = !Inve + WCantidad
                            .Update
                        End If
                    End With
                
                End If
                
                .MoveNext
                        
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
    End With
            
    rstInventario.Close
            
    End If
    
    Listado.WindowTitle = "Listado de Diferencias de Inventario de Materia Prima"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{FichaMat.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    Listado.Action = 1
    
    Exit Sub

WError:
    Resume Next

End Sub

Private Sub calcula_datos()

    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        WFechaCierre = rstArticulo!FechaCierre
        WOrdFechaCierre = rstArticulo!OrdFechaCierre
        WEntradas = rstArticulo!entradas
        WSalidas = rstArticulo!salidas
        rstArticulo.Close
    End If
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgDifeInvMat.Caption = "Listado de Diferencias de Inventario de Materia Prima :  " + !Nombre
        End If
    End With

End Sub