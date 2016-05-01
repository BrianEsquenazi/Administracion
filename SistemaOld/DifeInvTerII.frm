VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgDifeInvTerII 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Diferencias de Inventario de Producto Terminado"
   ClientHeight    =   7170
   ClientLeft      =   225
   ClientTop       =   975
   ClientWidth     =   11655
   LinkTopic       =   "Form2"
   ScaleHeight     =   7170
   ScaleWidth      =   11655
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1815
      Left            =   2400
      TabIndex        =   0
      Top             =   1920
      Width           =   4815
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   0
      Top             =   0
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
Attribute VB_Name = "PrgDifeInvTerII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WTerminado As String
Private WEntradas As Double
Private WSalidas As Double
Private Vector(10000, 2) As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
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
Dim rstConsig As Recordset
Dim spConsig As String
Dim XParam As String


Private Sub Cancela_click()
    PrgDifeInvTer.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Acepta_Click()

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
        
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
    With rstTerminado
        .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If Left$(rstTerminado!Codigo, 2) = "PT" Then
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstTerminado!Codigo
                    Vector(Renglon, 2) = rstTerminado!Descripcion
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
    End With
    rstTerminado.Close
    
    End If
    
    For Da = 1 To Renglon
    
        Rem dada
        
    
        WEntradas = 0
        WSalidas = 0
        WTerminado = Vector(Da, 1)
        XCodigo = Vector(Da, 1)
        
        Call calcula_datos
        
        With rstInve
            .Index = "Codigo"
            .Seek "=", WTerminado
            If .NoMatch = False Then
                .Edit
                !Codigo = WTerminado
                !Descripcion = Vector(Da, 2)
                !Stock = WEntradas - WSalidas
                .Update
                    Else
                .AddNew
                !Codigo = WTerminado
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
                
                If rstInventario!Tipo = "T" Then
                
                    WTerminado = rstInventario!Terminado
                    WCantidad = rstInventario!Cantidad
                
                    With rstInve
                        .Index = "Codigo"
                        .Seek "=", WTerminado
                        If .NoMatch = False Then
                            .Edit
                            !Codigo = WTerminado
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
    
    Listado.WindowTitle = "Listado de Diferencias de Inventario de Producto Terminado"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
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

    Rem PROCESA LAS ESTADISTICAS
    
    
    Rem If Terminado = "PT-25012-100" Then Stop
    
    spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WFechaCierre = rstTerminado!FechaCierre
        WOrdFechaCierre = rstTerminado!OrdFechaCierre
        WEntradas = rstTerminado!entradas
        WSalidas = rstTerminado!salidas
        rstTerminado.Close
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
            PrgDifeInvTer.Caption = "Listado de Diferencias de Inventario de Producto Terminado :  " + !Nombre
        End If
    End With
End Sub
