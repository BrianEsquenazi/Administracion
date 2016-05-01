VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgLinea 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Lineas de Ventas"
   ClientHeight    =   5055
   ClientLeft      =   2730
   ClientTop       =   1425
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   ScaleHeight     =   5055
   ScaleWidth      =   6720
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
      Left            =   360
      TabIndex        =   26
      Top             =   2400
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox Linea 
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
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   25
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   1440
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
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
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   24
         Text            =   " "
         Top             =   600
         Width           =   855
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
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   23
         Text            =   " "
         Top             =   240
         Width           =   855
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
         Left            =   2160
         TabIndex        =   22
         Top             =   1200
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
         Left            =   720
         TabIndex        =   21
         Top             =   1200
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
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   240
         Width           =   975
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
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
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
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
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
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5400
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlineas.rpt"
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
      Left            =   5400
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "Linea.frx":0000
      Left            =   360
      List            =   "Linea.frx":0007
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   13
      Top             =   1440
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
      Height          =   300
      Left            =   600
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   4320
      TabIndex        =   7
      Top             =   840
      Width           =   1935
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
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
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
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
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Nombre 
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
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion"
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
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo"
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
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "PrgLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstLinea As Recordset
Dim spLinea As String
Dim XParam As String

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Lineas de Ventas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Lineas.Linea} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Lineas.Linea , Lineas.Nombre " _
                        + "From " + DSQ + ".dbo.Lineas Lineas " _
                        + "Where Lineas.Linea >= 0 AND Lineas.Linea <= 9999"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Linea.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Linea.Text <> "" Then
    
        spLinea = "ConsultaLinea " + "'" + Linea.Text + "'"
        Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
        If rstLinea.RecordCount > 0 Then
            XParam = "'" + Linea.Text + "','" + Nombre.Text + "'"
            Set rstLinea = db.OpenRecordset("ModificaLinea " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Else
            XParam = "'" + Linea.Text + "','" + Nombre.Text + "'"
            Set rstLinea = db.OpenRecordset("AltaLinea " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Call CmdLimpiar_Click
        Linea.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Linea.Text <> "" Then
        spLinea = "ConsultaLinea " + "'" + Linea.Text + "'"
        Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
        If rstLinea.RecordCount > 0 Then
            T$ = "Lineas de Venta"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spLinea = "BorrarLinea " + "'" + Linea.Text + "'"
                Set rstLinea = db.OpenRecordset(spLinea, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    Linea.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Linea.Text = ""
    Nombre.Text = ""
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Linea.SetFocus
    PrgLinea.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
    Linea.SetFocus
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Linea.SetFocus
    End If
End Sub

Sub Linea_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Linea.Text) <> 0 Then
            WLinea = Linea.Text
            spLinea = "ConsultaLinea " + "'" + Linea.Text + "'"
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstLinea.RecordCount > 0 Then
                Linea.Text = rstLinea!Linea
                Nombre.Text = rstLinea!Nombre
                    Else
                WLinea = Linea.Text
                CmdLimpiar_Click
                Linea.Text = WLinea
            End If
        End If
        Nombre.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
            spLinea = "ListaLinea"
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstLinea
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstLinea!Linea) + " " + rstLinea!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstLinea!Linea
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstLinea.Close
        
        Case Else
    End Select
            
    Ayuda.Text = ""
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WLinea = WIndice.List(Indice)
            spLinea = "ConsultaLinea " + "'" + Str$(WLinea) + "'"
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstLinea.RecordCount > 0 Then
                Linea.Text = rstLinea!Linea
                Nombre.Text = rstLinea!Nombre
                        Else
                CmdLimpiar_Click
                Linea.Text = WLinea
            End If
            
            Linea.SetFocus
        
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    spLinea = "ListaLinea"
    Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstLinea.RecordCount > 0 Then
        With rstLinea
            .MoveFirst
            Linea.Text = rstLinea!Linea
            Nombre.Text = rstLinea!Nombre
        End With
        rstLinea.Close
    End If
    Linea.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Lineas de Venta", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Linea.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spLinea = "ListaLinea"
    Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstLinea.RecordCount > 0 Then
        With rstLinea
            .MoveLast
            Linea.Text = rstLinea!Linea
            Nombre.Text = rstLinea!Nombre
            Linea.SetFocus
        End With
        rstLinea.Close
    End If
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Lineas de Venta", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Linea.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spLinea = "AnteriorLinea " + "'" + Linea.Text + "'"
    Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstLinea.RecordCount > 0 Then
        With rstLinea
            .MoveLast
            Linea.Text = rstLinea!Linea
            Nombre.Text = rstLinea!Nombre
        End With
        rstLinea.Close
    End If
    
    Linea.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Lineas de Venta", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Linea.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spLinea = "PosteriorLinea " + "'" + Linea.Text + "'"
    Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstLinea.RecordCount > 0 Then
        With rstLinea
            .MoveFirst
            Linea.Text = rstLinea!Linea
            Nombre.Text = rstLinea!Nombre
        End With
        rstLinea.Close
    End If
    
    Linea.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Lineas de Venta", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Linea.SetFocus
    
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spLinea = "ListaLinea"
    Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstLinea.RecordCount > 0 Then
        With rstLinea
            .MoveFirst
            Do
                If .EOF = False Then
            
                    DA = Len(rstLinea!Nombre) - WEspacios
                
                    For aa = 1 To DA
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                            IngresaItem = Str$(rstLinea!Linea) + " " + rstLinea!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstLinea!Linea
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next aa
                    .MoveNext
                    
                        Else
                        
                    Exit Do
                
                End If
            Loop
        End With
        rstLinea.Close
    End If
    End If

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgLinea.Caption = "Ingreso de Lineas de Ventas :  " + !Nombre
        End If
    End With
End Sub


