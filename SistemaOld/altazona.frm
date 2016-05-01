VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAltaZona 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definicion de Zonas"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   7410
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Proceso"
      Height          =   615
      Left            =   8160
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   4
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   1320
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
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
      Left            =   1200
      TabIndex        =   10
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Ciudad 
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   6
      Text            =   " "
      Top             =   360
      Width           =   855
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
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   2280
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox Zona 
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
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   1455
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8520
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "altazona.frx":0000
      Left            =   120
      List            =   "altazona.frx":0007
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Top             =   1320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3615
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6376
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label DesZona 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   15
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   7200
      MouseIcon       =   "altazona.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "altazona.frx":031F
      ToolTipText     =   "Salida"
      Top             =   6480
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7200
      MouseIcon       =   "altazona.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "altazona.frx":0E6B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   7200
      MouseIcon       =   "altazona.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "altazona.frx":19B7
      ToolTipText     =   "Consulta de Datos"
      Top             =   5040
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   7200
      MouseIcon       =   "altazona.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "altazona.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5760
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Ciudad"
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
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Producto"
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
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "PrgAltaZona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Private WFecha As String
Private WAnterior As Integer
Private WDescri As String
Dim Renglon As Integer
Dim WCalle As String
Dim WZona As String
Dim WCampo1 As String
Dim WCampo2 As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Zonas"
     Opcion.AddItem "Calles"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    Select Case XIndice
        Case 0
            With rstZona
                .Index = "Codigo"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Codigo) + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 1
            With rstCalle
                .Index = "Codigo"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Val(Ciudad.Text) = !Ciudad Then
                            Auxi = Str$(!Codigo)
                            Call Ceros(Auxi, 4)
                            IngresaItem = Auxi + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose1_Click()

    Call Limpia_Click

    With rstAltaZona
        .Close
    End With
    With rstCalle
        .Close
    End With
    With rstCiudad
        .Close
    End With
    With rstZona
        .Close
    End With
    
    DbsAdminis.Close
    PrgAltaZona.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    For WRenglon = 1 To 1000
    
        With rstAltaZona
    
            Auxi = Zona.Text
            Call Ceros(Auxi, 6)
        
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 3)
        
            .Index = "Clave"
            .Seek "=", Auxi + Auxi1
            If .NoMatch = False Then
                .Delete
            End If
        
        End With
    
    Next WRenglon

    WRenglon = 0
    With rstAltaZona
        
        .Index = "Clave"
        For iRow = 1 To 1000
        
            WVector1.Row = iRow
            
            WVector1.Col = 1
            Calle = Val(WVector1.Text)
            
            WVector1.Col = 3
            Desde = Val(WVector1.Text)
            
            WVector1.Col = 4
            Hasta = Val(WVector1.Text)
            
            If Calle <> 0 Then
                    
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 3)
                        
                Auxi1 = Str$(Zona.Text)
                Call Ceros(Auxi1, 6)
                
                WZona = Zona.Text
                Call Ceros(WZona, 4)
                
                WCalle = Calle
                Call Ceros(WCalle, 4)
                    
                .AddNew
                !Clave = Auxi1 + Auxi
                !Zona = Val(Zona.Text)
                !Renglon = WRenglon
                !Ciudad = Val(Ciudad.Text)
                !Calle = Calle
                !Desde = Desde
                !Hasta = Hasta
                !Busqueda = WZona + WCalle
                .Update
                        
            End If
            
        Next iRow
            
    End With
    
    Call Limpia_Click

    WVector1.Col = 1
    WVector1.Row = 1
        
    Zona.SetFocus
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Zona.Text = ""
    DesZona.Caption = ""
    Ciudad.Text = ""
    DesCiudad.Caption = ""
    
    Renglon = 0
    Graba.Enabled = True
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    Zona.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            With rstZona
                Indice = Pantalla.ListIndex
                .Index = "Codigo"
                .Seek "=", WIndice.List(Indice)
                If .NoMatch = False Then
                    Zona.Text = !Codigo
                    DesZona.Caption = !Descripcion
                    Ciudad.Text = !Ciudad
                    With rstCiudad
                        .Index = "Codigo"
                        .Seek "=", Val(Ciudad.Text)
                        If .NoMatch = False Then
                            DesCiudad.Caption = !Descripcion
                        End If
                    End With
                End If
            End With
            Call Zona_KeyPress(13)
            Zona.SetFocus
            Rem Ayuda.Visible = False
            
        Case 1
            WTexto1.Visible = False
            WTexto2.Visible = False
            With rstCalle
                Indice = Pantalla.ListIndex
                .Index = "Codigo"
                .Seek "=", WIndice.List(Indice)
                If .NoMatch = False Then
                    WVector1.Col = 1
                    WVector1.Text = !Codigo
                    WVector1.Col = 2
                    WVector1.Text = !Descripcion
                    WVector1.Col = 3
                    Call StartEdit
                End If
            End With
            Rem Ayuda.Visible = False
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    WVector1.Col = 1
    WVector1.Row = 1

    Zona.Text = ""
    DesZona.Caption = ""
    Ciudad.Text = ""
    DesCiudad.Caption = ""
    
    Renglon = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    
    For WRenglon = 1 To 1000
    
    With rstAltaZona
    
        Auxi = Zona.Text
        Call Ceros(Auxi, 6)
        
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 3)
        
        .Index = "Clave"
        .Seek "=", Auxi + Auxi1
        If .NoMatch = False Then
        
            WVector1.Row = WRenglon
            Renglon = WRenglon
                
            WVector1.Col = 1
            WVector1.Text = !Calle
            WCalle = !Calle
            
            WVector1.Col = 3
            WVector1.Text = !Desde
            
            WVector1.Col = 4
            WVector1.Text = !Hasta
            
            With rstCalle
                .Index = "Codigo"
                .Seek "=", WCalle
                If .NoMatch = False Then
                    WVector1.Col = 2
                    WVector1.Text = !Descripcion
                End If
            End With
                
        End If
        
    End With
    
    Next WRenglon
    
    With rstZona
        .Index = "Codigo"
        .Seek "=", Val(Zona.Text)
        If .NoMatch = False Then
            DesZona.Caption = !Descripcion
            Ciudad.Text = !Ciudad
        End If
    End With
    
    With rstCiudad
        .Index = "Codigo"
        .Seek "=", Val(Ciudad.Text)
        If .NoMatch = False Then
            DesCiudad.Caption = !Descripcion
        End If
    End With
    
    Graba.Enabled = True

End Sub

Private Sub Zona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstZona
            .Index = "Codigo"
            .Seek "=", Val(Zona.Text)
            If .NoMatch = False Then
                DesZona.Caption = !Descripcion
                Ciudad.Text = !Ciudad
                With rstCiudad
                    .Index = "Codigo"
                    .Seek "=", Val(Ciudad.Text)
                    If .NoMatch = False Then
                        DesCiudad.Caption = !Descripcion
                    End If
                End With
                With rstAltaZona
                    Auxi = Zona.Text
                    Call Ceros(Auxi, 6)
                    .Index = "Clave"
            
                    Claveven$ = Auxi + "001"
                    .Seek "=", Claveven$
                    If .NoMatch = False Then
                        Ciudad.Text = !Ciudad
                        Call Proceso_Click
                        WVector1.Col = 1
                        WVector1.Row = 1
                        Call StartEdit
                            Else
                        Graba.Enabled = True
                        WZona = Zona.Text
                        Zona.Text = WZona
                        Call Limpia_Vector
                        WVector1.Col = 1
                        WVector1.Row = 1
                        Call StartEdit
                    End If
                End With
                    Else
                Zona.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    Busqueda = Left$(Ayuda.Text, WEspacios)
    
    Select Case XIndice
        Case 0
            With rstZona
                .Index = "Codigo"
                .MoveFirst
                Do
                    If .EOF = False Then
                        da = Len(!Descripcion) - WEspacios
                        For aa = 1 To da + 1
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                                IngresaItem = Str$(!Codigo) + " " + !Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Codigo
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
            
        Case 1
            With rstCalle
                .Index = "Codigo"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Val(Ciudad.Text) = !Ciudad Then
                            da = Len(!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(!Codigo) + " " + !Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = !Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case Else
    End Select
            
    End If

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Calle
    OPEN_FILE_Ciudad
    OPEN_FILE_Zona
    OPEN_FILE_AltaZona
End Sub

Private Sub Zona_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Zona"
    Opcion.AddItem "Calle"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 4
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            With rstCalle
                .Index = "Codigo"
                .Seek "=", Val(WVector1.Text)
                If .NoMatch = False Then
                    If !Ciudad <> Val(Ciudad.Text) Then
                        m$ = "La Calle Informada no pertenece a la ciudad"
                        A% = MsgBox(m$, 0, "Asignacion de Zonas")
                        WControl = "N"
                            Else
                        WVector1.Col = 2
                        WVector1.Text = !Descripcion
                    End If
                        Else
                    WControl = "N"
                End If
            End With
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If WAuxi1 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
    Opcion.AddItem "Zona"
    Opcion.AddItem "Calle"
    Opcion.AddItem "Ciudad"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Call Opcion_Click
    
    End If
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 5
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Calle"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Desde"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 5
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Hasta"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 5
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 340
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub Command1_Click()

    With rstAltaZona
        .Index = "Clave"
        .MoveFirst
        Do
            If .EOF = False Then
                WCampo1 = !Zona
                WCampo2 = !Renglon
                Call Ceros(WCampo1, 6)
                Call Ceros(WCampo2, 3)
                WClave = WCampo1 + WCampo2
                .Edit
                !Clave = WClave
                .Update
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

End Sub

