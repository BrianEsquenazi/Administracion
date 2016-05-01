VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListStkFamDq 
   Caption         =   "Listado de Stock de Dq por Familia"
   ClientHeight    =   6495
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6495
   ScaleWidth      =   8145
   Begin VB.CommandButton AvisoError 
      Caption         =   "No se puede emitir el informe. Sistema sin Conexion con las otras plantas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2760
      Picture         =   "liststkfamdq.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   3495
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7440
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   6135
      Begin VB.ComboBox Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Hasta 
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
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Desde 
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
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   0
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
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
         Left            =   4680
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Left            =   2400
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         Left            =   960
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
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
         Left            =   4680
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   360
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
         Left            =   4680
         MaskColor       =   &H00000000&
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Listado"
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
         Left            =   480
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Familia"
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
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Familia"
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
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   1080
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
      Height          =   2760
      ItemData        =   "liststkfamdq.frx":0742
      Left            =   120
      List            =   "liststkfamdq.frx":0749
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgListStkFamDq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstLineaMp As Recordset
Dim spLineaMp As String
Dim Empe(12, 10) As String
Dim XDesde As String
Dim XHasta As String
Dim WVector(1000) As String
Dim CargaEmpresa(12, 2) As String

Private Sub Acepta_Click()

    Rem
    Rem verifica conexciones con las otras plantas
    Rem
    
    WSalidaError = ""
    On Error GoTo Control_error
    
    XEmpresa = WEmpresa
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0007"
    CargaEmpresa(2, 2) = "Empresa07"
                    
    For Cicla = 1 To 2
        If CargaEmpresa(Cicla, 1) <> "" Then
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    If WSalidaError = "N" Then Exit Sub

    Renglon = 0

    spLineaMp = "ListaLineaMp"
    Set rstLineaMp = db.OpenRecordset(spLineaMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineaMp.RecordCount > 0 Then
        With rstLineaMp
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    WVector(rstLineaMp!Linea) = rstLineaMp!Nombre
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstLineaMp.Close
    End If

    XDesde = Desde.Text
    XHasta = Hasta.Text
    
    Call Ceros(XDesde, 3)
    Call Ceros(XHasta, 3)


    WDesde = "DQ-" + XDesde + "-000"
    WHasta = "DQ-" + XHasta + "-999"
    
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
    
    Empe(1, 1) = "0001"
    Empe(1, 2) = "Empresa01"
    Empe(2, 1) = "0007"
    Empe(2, 2) = "Empresa07"
    XHasta = 2
    
    For Ciclo = 1 To XHasta
    
        WEmpresa = Empe(Ciclo, 1)
        txtOdbc = Empe(Ciclo, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    
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
                    WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    WDescripcion = rstArticulo!Descripcion
                
                    With rstStockDy
                        .Index = "Codigo"
                        .Seek "=", WArticulo
                        If .NoMatch = True Then
                            .AddNew
                            !Codigo = WArticulo
                            !Descripcion = WDescripcion
                            !Stock1 = 0
                            !Stock2 = 0
                            !Stock3 = 0
                            !Stock4 = 0
                            !Stock5 = 0
                            Select Case Ciclo
                                Case 1
                                    !Stock1 = !Stock1 + WStock
                                Case 2
                                    !Stock2 = !Stock2 + WStock
                                Case 3
                                    !Stock3 = !Stock3 + WStock
                                Case 4, 6, 7
                                    !Stock4 = !Stock4 + WStock
                                Case 5
                                    !Stock5 = !Stock5 + WStock
                                Case Else
                            End Select
                            !Titulo1 = "Surfactan S.A."
                            !Titulo2 = ""
                            !Familia = Val(Mid$(!Codigo, 4, 1)) + 1
                            !Desfamilia = WVector(!Familia)
                            .Update
                                Else
                            .Edit
                            Select Case Ciclo
                                Case 1, 6
                                    !Stock1 = !Stock1 + WStock
                                Case 2
                                    !Stock2 = !Stock2 + WStock
                                Case 3
                                    !Stock3 = !Stock3 + WStock
                                Case 4
                                    !Stock4 = !Stock4 + WStock
                                Case 5
                                    !Stock5 = !Stock5 + WStock
                                Case Else
                            End Select
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
    Next Ciclo
    
    Call Conecta_Empresa

    If Tipo.ListIndex = 0 Then
    
        With rstStockDy
            .Index = "Codigo"
            .Seek ">=", ""
            If .NoMatch = False Then
                Do
                    Suma = !Stock1 + !Stock2 + !Stock3 + !Stock4 + !Stock5
                    If Suma = 0 Then
                        .Delete
                    End If
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        
    End If

    Listado.WindowTitle = "Listado de Stock de DQ por Familias"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.ReportFileName = "WListStkFamDQ.rpt"
    
    Listado.Action = 1
    
    Exit Sub

Control_error:
    Rem MsgBox Err.Description
    Beep
    WSalidaError = "N"
    AvisoError.Visible = True
    Resume Next
    
End Sub

Private Sub AvisoError_Click()
    AvisoError.Visible = False
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
    PrgListStkFamDq.Hide
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
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgListStkFamDq.Caption = "Listado de Stock de DQ por Familia :  " + !Nombre
        End If
    End With
    
    Tipo.Clear
    
    Tipo.AddItem "C/Stock"
    Tipo.AddItem "Todos los Articulos"
    
    Tipo.ListIndex = 0
    
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

    XIndice = 0
    
    Select Case XIndice
        Case 0
            spLineaMp = "ListaLineaMp"
            Set rstLineaMp = db.OpenRecordset(spLineaMp, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstLineaMp
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstLineaMp!Linea) + " " + rstLineaMp!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstLineaMp!Linea
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstLineaMp.Close
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            With rstLineaMp
            
            Indice = Pantalla.ListIndex
            WLinea = WIndice.List(Indice)
            spLineaMp = "ConsultaLineaMp " + "'" + WLinea + "'"
            Set rstLineaMp = db.OpenRecordset(spLineaMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineaMp.RecordCount > 0 Then
                    Desde.Text = rstLineaMp!Linea
                    Hasta.Text = rstLineaMp!Linea
                End If
            End With
            Desde.SetFocus
        Case Else
    End Select
    
End Sub



Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spLineaMp = "ListaLineaMp"
    Set rstLineaMp = db.OpenRecordset(spLineaMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineaMp.RecordCount > 0 Then
    
    With rstLineaMp
        .MoveFirst
        Do
            If .EOF = False Then
            
                Da = Len(rstLineaMp!Nombre) - WEspacios
                For Aaa = 1 To Da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstLineaMp!Nombre, Aaa, WEspacios) Then
                        IngresaItem = Str$(rstLineaMp!Linea) + " " + rstLineaMp!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstLineaMp!Linea
                        WIndice.AddItem IngresaItem
                        Exit For
                    End If
                Next Aaa
                .MoveNext
                    
                        Else
                        
                Exit Do
                
            End If
        Loop
    End With
    
    rstLineaMp.Close
    
    End If
    
    End If

End Sub








