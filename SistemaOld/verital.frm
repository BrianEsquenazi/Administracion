VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgVerital 
   AutoRedraw      =   -1  'True
   Caption         =   "Verificacion de Correlatividades de Talones"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   3135
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wverifica.rpt"
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
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5880
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgVerital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String
Dim rstInventario As Recordset
Dim spInventario As String
Dim XParam As String
Dim A1 As String
Dim A2 As String


Private Sub Acepta_Click()

    On Error GoTo WError
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            XEmpresa = !Nombre
        End If
    End With
    

    Da = 0
    With rstVerifica
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
                                        
                WNumero = rstInventario!Talon
                WTexto = rstInventario!Articulo + " / " + rstInventario!Terminado
                        
                With rstVerifica
                
                    .AddNew
                    !Numero = WNumero
                    !Descri = "Talones"
                    !Fecha = "  /  /    "
                    !Estado = ""
                    !Texto = WTexto
                    !Titulo = XEmpresa
                    .Update
                        
                End With
                
                .MoveNext
                        
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
    End With
            
    rstInventario.Close
            
    End If
            
    Pasa = 0

    With rstVerifica
    
        .Index = "CLAVE"
        .MoveFirst
            
        If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If Pasa = 0 Then
                    WNumero = !Numero
                    Pasa = 1
                End If
                
                If Val(WNumero) <> !Numero Then
                                        
                    With rstVerifica
                
                        .AddNew
                        !Numero = WNumero
                        !Descri = ""
                        !Fecha = "  /  /    "
                        !Estado = "FALTANTE"
                        !Titulo = XEmpresa
                        WTexto = ""
                        .Update
                            
                    End With
                    WNumero = WNumero + 1
                    
                        Else
                
                    WNumero = WNumero + 1
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                End If
                
            Loop
        End If
    End With

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    Listado.WindowTitle = "Listado de Verificacion de Correlatividades de Talones de Inventario"
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

Private Sub Cancela_click()

    With rstVerifica
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    PrgVerital.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Verifica
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgVerital.Caption = "Verificacion de Correlatividades de Talones de Inventario :  " + !Nombre
        End If
    End With
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
End Sub

