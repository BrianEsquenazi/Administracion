VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaLista 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Lista de Precios"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.TextBox Cliente 
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
      Left            =   6960
      MaxLength       =   6
      TabIndex        =   28
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Version 
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
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
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   27
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   26
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
      Index           =   1
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Observaciones 
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   3
      Top             =   840
      Width           =   7575
   End
   Begin VB.TextBox Titulo 
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   2
      Top             =   480
      Width           =   7575
   End
   Begin VB.Frame IngresaBase 
      Height          =   1215
      Left            =   4080
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox Listabase 
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
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton Base 
      Caption         =   "Lista de Precios Base"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   20
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3600
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton AgregaRenglon 
      Caption         =   "Agrega Renglon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   15
      Top             =   5880
      Width           =   1455
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
      Left            =   1680
      TabIndex        =   11
      Top             =   2280
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   10
      Top             =   2880
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
      Left            =   1080
      TabIndex        =   9
      Top             =   2280
      Width           =   375
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
      TabIndex        =   8
      Top             =   5760
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
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   6360
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
      Height          =   1980
      ItemData        =   "CargaLista.frx":0000
      Left            =   120
      List            =   "CargaLista.frx":0007
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Top             =   2280
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
      Height          =   4335
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7646
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   327680
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
   Begin VB.Label Label2 
      Caption         =   "Cliente"
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
      Height          =   285
      Left            =   5880
      TabIndex        =   30
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   8160
      TabIndex        =   29
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Observaciones"
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
      Left            =   240
      TabIndex        =   23
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Titulo"
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
      Left            =   240
      TabIndex        =   22
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   9360
      MouseIcon       =   "CargaLista.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "CargaLista.frx":031F
      ToolTipText     =   "Salida"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7200
      MouseIcon       =   "CargaLista.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "CargaLista.frx":0E6B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   7920
      MouseIcon       =   "CargaLista.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "CargaLista.frx":19B7
      ToolTipText     =   "Consulta de Datos"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   8640
      MouseIcon       =   "CargaLista.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "CargaLista.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Nro. Lista"
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
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgCargaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim spTerminado As String
Dim rstTerminado As Recordset
Dim rstCargaLista As Recordset
Dim rsCargaLista As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer

Dim ZTerminado As String
Dim ZDescripcion As String
Dim ZPrecio As String
Dim ZLinea As String

Dim WVersion As String
Dim WRenglon As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private WGraba As String
Private WGrabaII As String

Dim CargaEmpresa(12, 2) As String

Private Sub Base_Click()

    IngresaBase.Visible = True
    
    Listabase.Text = ""
    Listabase.SetFocus

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub ListaBase_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        Call Limpia_Vector
        WRenglon = 0
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaLista"
        ZSql = ZSql + " Where CargaLista.Lista = " + "'" + Listabase.Text + "'"
        ZSql = ZSql + " Order by CargaLista.Clave"
    
        rsCargaLista = ZSql
        Set rstCargaLista = db.OpenRecordset(rsCargaLista, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaLista.RecordCount > 0 Then
            With rstCargaLista
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        WRenglon = WRenglon + 1
                        WVector1.Row = WRenglon
                        Renglon = WRenglon
                
                        WVector1.Col = 1
                        WVector1.Text = rstCargaLista!Terminado
            
                        WVector1.Col = 2
                        WVector1.Text = ""
                    
                        WVector1.Col = 3
                        WVector1.Text = Str$(rstCargaLista!Precio)
            
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCargaLista.Close
        End If
        
        IngresaBase.Visible = False
        
    End If
    
    If KeyAscii = 27 Then
        Listabase.Text = ""
    End If
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "P.Terminados"
     Opcion.AddItem "Clientes"
     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Rem Ayuda.Visible = True
    Ayuda.Text = ""
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Terminado"
            ZSql = ZSql + " Order by Codigo"
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTerminado!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTerminado.Close
            End If
            
        Case 1
            spClientes = "ListaClienteConsulta"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                With rstClientes
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstClientes!Cliente + " " + rstClientes!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstClientes!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstClientes.Close
            End If
            
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
    PrgCargaLista.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    If WGraba <> "S" Then
    
        Call Ingresa_clave

               Else
               
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                ZEmpresa = !Nombre
            End If
        End With
               
        ZSql = ""
        ZSql = ZSql + "DELETE CargaLista"
        ZSql = ZSql + " Where Lista = " + "'" + Version.Text + "'"
        rsCargaLista = ZSql
        Set rstCargaLista = db.OpenRecordset(rsCargaLista, dbOpenSnapshot, dbSQLPassThrough)
    
        WRenglon = 0
        
        For iRow = 1 To 1000
    
            ZTerminado = WVector1.TextMatrix(iRow, 1)
            ZDescripcion = WVector1.TextMatrix(iRow, 2)
            ZPrecio = WVector1.TextMatrix(iRow, 3)
            ZLinea = ""
            ZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            
            If ZTerminado <> "" Or Val(WPrecio) <> 0 Then
            
                If Left$(ZTerminado, 2) = "PT" Then
            
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Terminado"
                    ZSql = ZSql + " Where Terminado.Codigo = " + "'" + ZTerminado + "'"
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZLinea = Trim(rstTerminado!Linea)
                        rstTerminado.Close
                    End If
                    
                        Else
                        
                    ZLinea = "16"
                    
                End If
                
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 3)
        
                Auxi1 = Version.Text
                Call Ceros(Auxi1, 6)
        
                WClave = Auxi1 + Auxi
        
                ZSql = ""
                ZSql = ZSql + "INSERT INTO CargaLista ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Lista ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "Titulo ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "Terminado ,"
                ZSql = ZSql + "Precio ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "Linea ,"
                ZSql = ZSql + "Cliente ,"
                ZSql = ZSql + "Empresa )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + Version.Text + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + Fecha.Text + "',"
                ZSql = ZSql + "'" + ZFechaOrd + "',"
                ZSql = ZSql + "'" + Titulo.Text + "',"
                ZSql = ZSql + "'" + Observaciones.Text + "',"
                ZSql = ZSql + "'" + ZTerminado + "',"
                ZSql = ZSql + "'" + ZPrecio + "',"
                ZSql = ZSql + "'" + ZDescripcion + "',"
                ZSql = ZSql + "'" + ZLinea + "',"
                ZSql = ZSql + "'" + Cliente.Text + "',"
                ZSql = ZSql + "'" + ZEmpresa + "')"
            
                rsCargaLista = ZSql
                Set rstCargaLista = db.OpenRecordset(rsCargaLista, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next iRow
    
        Call Limpia_Click

        WVector1.Col = 1
        WVector1.Row = 1
        
        Version.SetFocus
        
    End If
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Version.Text = ""
    Fecha.Text = "  /  /    "
    Titulo.Text = ""
    Observaciones.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    
    Renglon = 0
    
    WGraba = ""
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    Version.SetFocus

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WTexto1.Text = WIndice.List(Indice)
            WVector1.Col = 1
            WVector1.Text = WIndice.List(Indice)
            
        Case 1
            Indice = Pantalla.ListIndex
            Cliente.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    WVector1.Col = 1
    WVector1.Row = 1

    Version.Text = ""
    Fecha.Text = "  /  /    "
    Titulo.Text = ""
    Observaciones.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""

    WGraba = ""
    
    Renglon = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaLista"
    ZSql = ZSql + " Where CargaLista.Lista = " + "'" + Version.Text + "'"
    ZSql = ZSql + " Order by CargaLista.Terminado"
    
    rsCargaLista = ZSql
    Set rstCargaLista = db.OpenRecordset(rsCargaLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaLista.RecordCount > 0 Then
        With rstCargaLista
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Fecha.Text = rstCargaLista!Fecha
                    Titulo.Text = rstCargaLista!Titulo
                    Observaciones.Text = rstCargaLista!Observaciones
                    Cliente.Text = Trim(rstCargaLista!Cliente)
                
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = rstCargaLista!Terminado
                    
                    WVector1.Col = 2
                    WVector1.Text = ""
            
                    WVector1.Col = 3
                    WVector1.Text = Str$(rstCargaLista!Precio)
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaLista.Close
    End If
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = rstCliente!Razon
        rstCliente.Close
    End If
    
    For Ciclo = 1 To WRenglon
    
        If Left$(WVector1.TextMatrix(Ciclo, 1), 2) = "PT" Then
    
            WCliente = UCase(Cliente.Text)
            WTerminado = WVector1.TextMatrix(Ciclo, 1)
            WClave = WCliente + WTerminado
        
            spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                WVector1.TextMatrix(Ciclo, 2) = rstPrecios!Descripcion
                rstPrecios.Close
                    Else
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Terminado"
                ZSql = ZSql + " Where Terminado.Codigo = " + "'" + WVector1.TextMatrix(Ciclo, 1) + "'"
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WVector1.TextMatrix(Ciclo, 2) = Trim(rstTerminado!Descripcion)
                    rstTerminado.Close
                End If
            End If
            
                    Else
        
            WArticulo = Left$(WVector1.TextMatrix(Ciclo, 1), 3) + Right$(WVector1.TextMatrix(Ciclo, 1), 7)
                    
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.TextMatrix(Ciclo, 2) = rstArticulo!Descripcion
                rstArticulo.Close
            End If
            
        End If
        
    Next Ciclo
    
End Sub

Private Sub Version_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Call Limpia_Vector

        Sql1 = "Select *"
        Sql2 = " FROM CargaLista"
        Sql3 = " Where CargaLista.Lista = " + "'" + Version.Text + "'"
        rsCargaLista = Sql1 + Sql2 + Sql3
        Set rstCargaLista = db.OpenRecordset(rsCargaLista, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaLista.RecordCount > 0 Then
            rstCargaLista.Close
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Fecha.SetFocus
        End If
    End If
    
    If KeyAscii = 27 Then
        Version.Text = ""
    End If
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Cliente.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        If Cliente.Text <> "" Then
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Razon
                Titulo.Text = rstCliente!Razon
                rstCliente.Close
                Titulo.SetFocus
                    Else
                Cliente.SetFocus
            End If
                Else
            Titulo.SetFocus
        End If
    End If
End Sub

Private Sub Titulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Titulo.Text = ""
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
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
            
        Case 123
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Col > 1 Then
                WVector1.Col = WVector1.Col - 1
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
        Case 3
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
            If WVector1.Text <> "" Then
                WVector1.Text = UCase(WVector1.Text)
                
                If Left$(WVector1.Text, 2) = "PT" Then
                    
                    WCliente = UCase(Cliente.Text)
                    WTerminado = WVector1.Text
                    WClave = WCliente + WTerminado
        
                    spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPrecios.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstPrecios!Descripcion
                        rstPrecios.Close
                            Else
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Terminado"
                        ZSql = ZSql + " Where Terminado.Codigo = " + "'" + WVector1.Text + "'"
                        spTerminado = ZSql
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WVector1.Col = 2
                            WVector1.Text = rstTerminado!Descripcion
                                Else
                            WControl = "N"
                            rstTerminado.Close
                        End If
                    End If
                    
                        Else
                        
                    WArticulo = Left$(WVector1.Text, 3) + Right$(WVector1.Text, 7)
                    
                    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstArticulo!Descripcion
                        rstArticulo.Close
                            Else
                        WControl = "N"
                    End If
                    
                End If
                
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
        WTexto1.Visible = False
        WTexto2.Visible = False
        WTexto3.Visible = False
    
        RenglonAuxiliar = WVector1.Row

        For Ciclo = 1 To WVector1.Cols - 1
            WVector1.Col = Ciclo
            WVector1.Text = ""
        Next Ciclo
    
        Erase WBorra
        EntraVector = 0
    
        HastaRenglon = 0
        For iRow = 1000 To 1 Step -1
        
            Terminado = WVector1.TextMatrix(iRow, 1)
            Precio = WVector1.TextMatrix(iRow, 3)
            
            If Terminado <> "" Or Precio <> "" Then
                HastaRenglon = iRow
                Exit For
            End If
            
        Next iRow
    
        For Ciclo = 1 To HastaRenglon
            WVector1.Row = Ciclo
            WVector1.Col = 1
            WAuxi1 = WVector1.Text
            If Ciclo <> RenglonAuxiliar Then
                EntraVector = EntraVector + 1
                For Ciclo1 = 0 To WVector1.Cols - 1
                    WVector1.Col = Ciclo1
                    WBorra(EntraVector, Ciclo1) = WVector1.Text
                Next Ciclo1
            End If
        Next Ciclo
    
        Call Limpia_Vector
    
        For Ciclo = 1 To EntraVector
            WVector1.Row = Ciclo
            For DA = 0 To WVector1.Cols - 1
                WVector1.Col = DA
                WVector1.Text = WBorra(Ciclo, DA)
            Next DA
        Next Ciclo
    
    End If
    
End Sub

Private Sub AgregaRenglon_Click()

    Hasta = WVector1.Row

    For iRow = 1000 To Hasta Step -1
        WVector1.TextMatrix(iRow, 1) = WVector1.TextMatrix(iRow - 1, 1)
        WVector1.TextMatrix(iRow, 2) = WVector1.TextMatrix(iRow - 1, 2)
        WVector1.TextMatrix(iRow, 3) = WVector1.TextMatrix(iRow - 1, 3)
    Next iRow

    WVector1.TextMatrix(Hasta, 1) = ""
    WVector1.TextMatrix(Hasta, 2) = ""
    WVector1.TextMatrix(Hasta, 3) = ""
    
    WTexto1.Text = ""
    WTexto2.Text = ""

End Sub


Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

        Opcion.Clear
    
        Opcion.AddItem "Productos Terminados"
        Opcion.ListIndex = 1
    
        Rem Call Opcion_Click
    
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
    WVector1.Cols = 4
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
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Terminado"
                WVector1.ColWidth(Ciclo) = 1600
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Instrucciones"
                WVector1.ColWidth(Ciclo) = 6000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Precio"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
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
    
    WAncho = 400
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

Sub Ingresa_clave()
    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
End Sub

Private Sub CancelaGraba_Click()
    XClave.Visible = False
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
        If WClave = "PRECIO" Then
            WGraba = "S"
            XClave.Visible = False
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Ingreso de Listas de Precios")
            WClave.SetFocus
        End If
    End If
End Sub



