VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgGasimpo 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Conceptos de Gastos de Importacion"
   ClientHeight    =   4020
   ClientLeft      =   3255
   ClientTop       =   1755
   ClientWidth     =   5895
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   5895
   Begin VB.TextBox Codigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   25
      Text            =   " "
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   24
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   23
         Text            =   " "
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5160
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wGasimpo.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "Gasimpo.frx":0000
      Left            =   480
      List            =   "Gasimpo.frx":0007
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      Height          =   300
      Left            =   1800
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   600
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1335
      Left            =   4320
      TabIndex        =   7
      Top             =   840
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Nombre 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "PrgGasimpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstGasimpo As Recordset
Dim spGasimpo As String
Dim XParam As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Conceptos de Gastos de Importacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Gasimpo.Codigo} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Gasimpo.Codigo, Gasimpo.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Gasimpo Gasimpo " _
                    + "Where " _
                    + "Gasimpo.Codigo >= 0 AND Gasimpo.Codigo <= 9999"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Codigo.Text <> "" Then
    
        spGasimpo = "ConsultaGasimpo " + "'" + Codigo.Text + "'"
        Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
        If rstGasimpo.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        If WPasa = "N" Then
            XParam = "'" + Codigo.Text + "','" + Nombre.Text + "'"
            Set rstGasimpo = db.OpenRecordset("AltaGasimpo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Else
            XParam = "'" + Codigo.Text + "','" + Nombre.Text + "'"
            Set rstGasimpo = db.OpenRecordset("ModificaGasimpo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Call CmdLimpiar_Click
        Codigo.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
        
        spGasimpo = "ConsultaGasimpo " + "'" + Codigo.Text + "'"
        Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
        If rstGasimpo.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        If WPasa = "S" Then
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spGasimpo = "BorrarGasimpo " + "'" + Codigo.Text + "'"
                Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Codigo.Text = ""
    Nombre.Text = ""
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Rem With rstGasimpos
    Rem     .Close
    Rem End With
    Rem With rstEmpresa
    Rem     .Close
    Rem End With
    Rem With rstAuxiliar
    Rem     .Close
    Rem End With
    Rem DbsVentas.Close
    Codigo.SetFocus
    PrgGasimpo.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = "0"
    Hasta.Text = "9999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
    Codigo.SetFocus
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Codigo.SetFocus
    End If
End Sub

Sub Codigo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
            WCodigo = Codigo.Text
            spGasimpo = "ConsultaGasimpo " + "'" + Codigo.Text + "'"
            Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
            If rstGasimpo.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
                
            If WPasa = "S" Then
                Codigo.Text = rstGasimpo!Codigo
                Nombre.Text = rstGasimpo!Nombre
                    Else
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
        End If
        Nombre.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

Rem     Opcion.Clear
Rem
Rem     Opcion.AddItem "Productos"
Rem     Opcion.AddItem "Ensayos"
Rem
Rem     Opcion.Visible = True
Rem End Sub
Rem
Rem Private Sub Opcion_Click()
Rem
Rem     Opcion.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    Rem XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
            spGasimpo = "ListaGasimpo"
            Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstGasimpo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstGasimpo!Codigo) + " " + rstGasimpo!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstGasimpo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstGasimpo.Close
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WCodigo = WIndice.List(Indice)
            spGasimpo = "ConsultaGasimpo " + "'" + Str$(WCodigo) + "'"
            Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
            If rstGasimpo.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
                Codigo.Text = rstGasimpo!Codigo
                Nombre.Text = rstGasimpo!Nombre
                        Else
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
            
            Codigo.SetFocus
        
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    spGasimpo = "ListaGasimpo"
    Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstGasimpo
        .MoveFirst
        Codigo.Text = rstGasimpo!Codigo
        Nombre.Text = rstGasimpo!Nombre
    End With
    
    rstGasimpo.Close
    Codigo.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Conceptos de Gastos de Importacion", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spGasimpo = "ListaGasimpo"
    Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstGasimpo
        .MoveLast
        Codigo.Text = rstGasimpo!Codigo
        Nombre.Text = rstGasimpo!Nombre
        Codigo.SetFocus
    End With
    rstGasimpo.Close
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Conceptos de Gastos de Importacion", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spGasimpo = "AnteriorGasimpo " + "'" + Codigo.Text + "'"
    Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstGasimpo
        .MoveLast
        Codigo.Text = rstGasimpo!Codigo
        Nombre.Text = rstGasimpo!Nombre
    End With
    
    rstGasimpo.Close
    Codigo.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Conceptos de Gastos de Importacion", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spGasimpo = "PosteriorGasimpo " + "'" + Codigo.Text + "'"
    Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstGasimpo
        .MoveFirst
        Codigo.Text = rstGasimpo!Codigo
        Nombre.Text = rstGasimpo!Nombre
    End With
    
    rstGasimpo.Close
    Codigo.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Conceptos de Gastos de Importacion", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
    
End Sub


