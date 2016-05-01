VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgMovcon 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Mercaderia en remitos a facturar por cliente"
   ClientHeight    =   6180
   ClientLeft      =   2085
   ClientTop       =   1500
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6180
   ScaleWidth      =   8085
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1815
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Hasta 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1200
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
      Begin VB.Label Label2 
         Caption         =   "Hasta Cliente"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Cliente"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5760
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WMovcon.rpt"
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
      Left            =   6240
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      ItemData        =   "movcon.frx":0000
      Left            =   120
      List            =   "movcon.frx":0007
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5880
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgMovcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstDevcon As Recordset
Dim spDevcon As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim XParam As String

Private Sub Acepta_Click()

    On Error GoTo WError


    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)

    DA = 0
    With rstFichaCon
        .Index = "CLiente"
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
            
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + "1" + "'"
    ZSql = ZSql + " and Estadistica.Numero >= " + "'" + "900000" + "'"
    ZSql = ZSql + " and Estadistica.Numero <= " + "'" + "999999" + "'"
    ZSql = ZSql + " and Estadistica.Cliente >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Estadistica.Cliente <= " + "'" + Hasta.Text + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WCliente = rstEstadistica!Cliente
                WCantidad = rstEstadistica!Cantidad
                WFecha = rstEstadistica!Fecha
                WNumero = rstEstadistica!Numero
                WTerminado = rstEstadistica!Terminado
                
                With rstFichaCon
                
                     .AddNew
                    !Cliente = WCliente
                    !Fecha = WFecha
                    !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    !Tipo = 0
                    !Numero = WNumero
                    !Inicial = 0
                    !Entrada = WCantidad
                    !Salida = 0
                    !Lista1 = "Remito"
                    !Observaciones = ""
                    !Lista2 = ""
                    !Terminado = WTerminado
                    !Saldo = !Entrada + !Salida
                    Rem WImpre1 + " " + Left$(WImpre2, 23)
                    .Update
                End With
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstEstadistica.Close
    End If
    
    
    
    DA = 0
    With rstFichaCon
        .Index = "Cliente"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WCliente = !Cliente
                WTerminado = !Terminado
                WDescripcion = ""
                WDescriter = ""
                
                spCliente = "ConsultaCliente " + "'" + WCliente + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    WDescripcion = rstCliente!Razon
                End If
                
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WDescriter = rstTerminado!Descripcion
                End If
                
                !Descripcion = WDescripcion
                !Descriter = WDescriter
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Mercaderia en Consignacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{FichaCon.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
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

    With rstEmpresa
        .Close
    End With
    With rstFichaCon
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgMovcon.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaCon
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgMovcon.Caption = "Listado de Mercaderia en remitos a facturar por cliente :  " + !Nombre
        End If
    End With
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstCliente
        .MoveFirst
            Do
            If .EOF = False Then
                IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                Pantalla.AddItem IngresaItem
                IngresaItem = rstCliente!Cliente
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstCliente.Close
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Desde.Text = rstCliente!Cliente
        Hasta.Text = rstCliente!Cliente
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub


