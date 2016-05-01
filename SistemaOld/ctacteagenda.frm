VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCtaCteAgenda 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Cuenta Corriente de Clientes"
   ClientHeight    =   7320
   ClientLeft      =   570
   ClientTop       =   1155
   ClientWidth     =   10995
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   10995
   Begin VB.CommandButton reclamo 
      Caption         =   "Reclamos"
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
      Left            =   2280
      TabIndex        =   28
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2280
      TabIndex        =   27
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   9
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   8
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3960
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4815
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8493
      _Version        =   327680
      BackColor       =   16776960
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
      Height          =   540
      Left            =   3480
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.CommandButton Listar 
      Caption         =   "Listar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Todos 
         Caption         =   "Total"
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
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Pendiente 
         Caption         =   "Pendiente"
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
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1335
      Left            =   6240
      TabIndex        =   4
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton Total 
         Caption         =   "Total"
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
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Documentos 
         Caption         =   "Documentos"
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
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton CtaCte 
         Caption         =   "Cta.Cte."
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
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Dolares 
         Caption         =   "Dolares"
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
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Pesos 
         Caption         =   "Pesos"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
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
      Height          =   375
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   1
      Text            =   " "
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Saldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
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
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
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
      Left            =   7560
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00FFFF00&
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
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "PrgCtaCteAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Importe1 As Double
Private Importe2 As Double
Private Importe3 As Double
Private WTipo As Integer
Private WSaldo As Double
Private Acumula As Double
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim XParam As String
Private WNume As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub cmdClose_Click()
    PrgCtaCteAgenda.Hide
    Unload Me
    PrgAltaAgenda.Show
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
 
    Cliente.Text = ""
    DesCliente.Caption = ""
    
    Pesos.Value = True
    CtaCte.Value = True
    Pendiente.Value = True
    
    Cliente.Text = PCliente
    Call Cliente_KeyPress(13)
    
End Sub

Private Sub Proceso_Click()

    Cliente.Text = UCase(Cliente.Text)
    
    WSalida = "N"
    
    Call Limpia_Vector

    Renglon = 0
    WSaldo = 0
    
    XParam = "'" + Cliente.Text + "'"
    spCtacte = "ListaCtacteCliente " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
        With rstCtacte
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    WPasa = "N"
                    
                    If CtaCte.Value = True Then
                        If !Tipo < 50 Then
                            WPasa = "S"
                        End If
                    End If
                    
                    If Documentos.Value = True Then
                        If !Tipo >= 50 Then
                            WPasa = "S"
                        End If
                    End If
                    
                    If Total.Value = True Then
                        WPasa = "S"
                    End If
                        
                    If WPasa = "S" Then
                        If Pesos.Value = True Then
                            If !Total > 0 Then
                                Importe1 = !Total
                                Importe2 = 0
                                    Else
                                Importe1 = 0
                                Importe2 = !Total
                            End If
                            Importe3 = !Saldo
                                Else
                            If !Totalus > 0 Then
                                Importe1 = !Totalus
                                Importe2 = 0
                                    Else
                                Importe1 = 0
                                Importe2 = !Totalus
                            End If
                            Importe3 = !Saldous
                        End If
                        
                        Call Redondeo(Importe3)
                    
                        If Importe3 <> 0 Or Todos.Value = True Then
                    
                            Renglon = Renglon + 1
                
                            Select Case !Tipo
                                Case 1
                                    WVector1.TextMatrix(Renglon, 1) = "Fac"
                                Case 2
                                    WVector1.TextMatrix(Renglon, 1) = "Dev"
                                Case 3
                                    WVector1.TextMatrix(Renglon, 1) = "Fac"
                                Case 4
                                    Select Case Left$(!Impre, 2)
                                        Case "DC"
                                            WVector1.TextMatrix(Renglon, 1) = "D.C"
                                        Case "CH"
                                            WVector1.TextMatrix(Renglon, 1) = "CHR"
                                        Case Else
                                            WVector1.TextMatrix(Renglon, 1) = "N/D"
                                    End Select
                                Case 5
                                    Select Case Left$(!Impre, 2)
                                        Case "DC"
                                            WVector1.TextMatrix(Renglon, 1) = "D.C"
                                        Case "CH"
                                            WVector1.TextMatrix(Renglon, 1) = "CHR"
                                        Case Else
                                            WVector1.TextMatrix(Renglon, 1) = "N/C"
                                    End Select
                                Case 6
                                    WVector1.TextMatrix(Renglon, 1) = "Rec"
                                Case 7
                                    WVector1.TextMatrix(Renglon, 1) = "Ant"
                                Case 10
                                    WVector1.TextMatrix(Renglon, 1) = "FCR"
                                Case 50
                                    WVector1.TextMatrix(Renglon, 1) = "Doc"
                                Case Else
                            End Select
                            
                            WVector1.TextMatrix(Renglon, 2) = Pusing("######", Str$(!Numero))
                            WVector1.TextMatrix(Renglon, 3) = !Fecha
                    
                            If Importe1 <> 0 Then
                                WVector1.TextMatrix(Renglon, 4) = Pusing("###,###,###.##", Str$(Importe1))
                                    Else
                                WVector1.TextMatrix(Renglon, 4) = ""
                            End If
                    
                            If Importe2 <> 0 Then
                                WVector1.TextMatrix(Renglon, 5) = Pusing("###,###,###.##", Str$(Importe2))
                                    Else
                                WVector1.TextMatrix(Renglon, 5) = ""
                            End If
                    
                            If Importe3 <> 0 Then
                                WVector1.TextMatrix(Renglon, 6) = Pusing("###,###,###.##", Str$(Importe3))
                                    Else
                                WVector1.TextMatrix(Renglon, 6) = ""
                            End If
                            
                            WSaldo = WSaldo + Importe3
                    
                            WVector1.TextMatrix(Renglon, 7) = !Vencimiento
                            WVector1.TextMatrix(Renglon, 8) = !Vencimiento1
                        
                        End If
                        
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
    
    End If
    
    Saldo.Caption = Pusing("###,###,###.##", Str$(WSaldo))
    
    WVector1.Col = 1
    WVector1.Row = 1
    WVector1.TopRow = 1

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        WCliente = Cliente.Text
        Cliente.Text = WCliente
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Razon
            rstCliente.Close
            Call Proceso_Click
                Else
            Cliente.SetFocus
        End If
    End If
End Sub

Private Sub reclamo_Click()
    cliente2 = Cliente.Text
    descliente2 = DesCliente.Caption
    PrgreclamoAgenda.Show
End Sub

Private Sub WVector1_DblClick()
    ZTipo = WVector1.TextMatrix(WVector1.Row, 1)
    ZRecibo = WVector1.TextMatrix(WVector1.Row, 2)
    
    If ZTipo = "Rec" Then
        WRecibo = ZRecibo
        PrgReciagenda.Show
    End If
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 10
    WVector1.FixedRows = 1
    WVector1.Rows = 5001
    
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
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WVector1.Text = "Debito"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector1.Text = "Credito"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector1.Text = "Saldo"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector1.Text = "Vencimiento"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector1.Text = "Vencimiento"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector1.Text = "Acumulado"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
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

