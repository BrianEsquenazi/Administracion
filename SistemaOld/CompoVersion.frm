VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCompoVersion 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Versiones de Composicion de Productos Terminados"
   ClientHeight    =   8085
   ClientLeft      =   840
   ClientTop       =   375
   ClientWidth     =   10170
   LinkTopic       =   "Form2"
   ScaleHeight     =   8085
   ScaleWidth      =   10170
   Visible         =   0   'False
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
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   15
      Text            =   " "
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Observaciones1 
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
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   13
      Text            =   " "
      Top             =   7320
      Width           =   6135
   End
   Begin VB.TextBox Observaciones2 
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
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   12
      Text            =   " "
      Top             =   7680
      Width           =   6135
   End
   Begin VB.TextBox Referencia 
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
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   10
      Text            =   " "
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "AA-#####-###"
      PromptChar      =   " "
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
      Left            =   1200
      TabIndex        =   2
      Top             =   120
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
      Height          =   420
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   5415
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9551
      _Version        =   327680
      BackColor       =   16777152
      ForeColor       =   4210752
   End
   Begin VB.Label FechaFinal 
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
      Left            =   5280
      TabIndex        =   14
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "OBservaciones"
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
      Left            =   240
      TabIndex        =   11
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Ref. Laboratorio"
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
      Left            =   240
      TabIndex        =   9
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Version"
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
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label FechaInicio 
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
      Left            =   3720
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha "
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
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label DesTerminado 
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
      Height          =   300
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "PrgCompoVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vector(100, 10) As String
Dim ZVector(100, 15) As String
Dim CargaEmpresa(12, 2) As String

Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstComposicionVersion As Recordset
Dim spComposicionVerion As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstCaratula As Recordset
Dim spCaratula As String
Dim XParam As String
Dim ZVersion As String
Dim ZRenglon As String

Private Lugar1 As Integer
Private Lugar2 As Integer
Private Auxi As String
Private Clave As String
Private Salva As String
Private WGraba As String
Private WVersion As Single
Private XVector(100, 20) As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click
    PrgCompoVersion.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Limpia_Click()

    Call Limpia_Vector

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Version.Text = ""
    FechaInicio.Caption = ""
    FechaFinal.Caption = ""
    Referencia.Text = ""
    Observaciones1.Text = ""
    Observaciones2.Text = ""

    Terminado.SetFocus

End Sub


Private Sub Form_Load()

    Call Limpia_Vector

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Observaciones1.Text = ""
    Observaciones2.Text = ""
    Referencia.Text = ""
    Version.Text = ""
    FechaInicio.Caption = ""
    FechaFinal.Caption = ""
    
    Rem Terminado.SetFocus
    
End Sub

Private Sub Proceso_Click()

    Rem On Error GoTo WError
    
    Call Limpia_Vector

    Renglon = 0
    Erase Vector
    
    Sql1 = "Select *"
    Sql2 = " FROM ComposicionVersion"
    Sql3 = " Where ComposicionVersion.Terminado = " + "'" + Terminado.Text + "'"
    Sql4 = " and ComposicionVersion.Version = " + "'" + Version.Text + "'"
    Sql5 = " Order by Clave"
    spComposicionVersion = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstComposicionVersion = db.OpenRecordset(spComposicionVersion, dbOpenSnapshot, dbSQLPassThrough)
    If rstComposicionVersion.RecordCount > 0 Then
    
        With rstComposicionVersion
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    Renglon = Renglon + 1
                
                    Vector(Renglon, 1) = rstComposicionVersion!Tipo
                    Vector(Renglon, 2) = rstComposicionVersion!Articulo1
                    Vector(Renglon, 3) = rstComposicionVersion!Articulo2
                    Vector(Renglon, 4) = ""
                    Vector(Renglon, 5) = rstComposicionVersion!Cantidad
                    
                    Referencia.Text = IIf(IsNull(rstComposicionVersion!Referencia), "", rstComposicionVersion!Referencia)
                    Observaciones1.Text = IIf(IsNull(rstComposicionVersion!Observaciones1), "", rstComposicionVersion!Observaciones1)
                    Observaciones2.Text = IIf(IsNull(rstComposicionVersion!Observaciones2), "", rstComposicionVersion!Observaciones2)
                    
                    Referencia.Text = RTrim(Referencia.Text)
                    Observaciones1.Text = RTrim(Observaciones1.Text)
                    Observaciones2.Text = RTrim(Observaciones2.Text)
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
        rstComposicionVersion.Close
    End If
    
    Renglon = 0
    
    For XX = 1 To 100
    
        If Vector(XX, 5) <> "" Then
            
            Renglon = Renglon + 1
        
            
            WVector1.TextMatrix(Renglon, 1) = Vector(XX, 1)
            WVector1.TextMatrix(Renglon, 2) = Vector(XX, 2)
            Articulo1 = Vector(XX, 2)
            WVector1.TextMatrix(Renglon, 3) = Vector(XX, 3)
            Articulo2 = Vector(XX, 3)
            WVector1.TextMatrix(Renglon, 5) = Pusing("###,###.#####", Vector(XX, 5))
            
            If Vector(XX, 1) = "M" Then
            
                spArticulo = "ConsultaArticulo " + "'" + Articulo1 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.TextMatrix(Renglon, 4) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                    
                    Else
                    
                WTerminado = Articulo2
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WVector1.TextMatrix(Renglon, 4) = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                
            End If
            
        End If
        
    Next XX
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    Terminado.SetFocus
    Exit Sub

WError:
    Resume Next

End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Terminado.Text = UCase(Terminado.Text)
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            Version.SetFocus
                Else
            Terminado.SetFocus
        End If
    End If
End Sub

Sub Version_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Version.Text <> "" Then
            Sql1 = "Select *"
            Sql2 = " FROM ComposicionVersion"
            Sql3 = " Where ComposicionVersion.Terminado = " + "'" + Terminado.Text + "'"
            Sql4 = " and ComposicionVersion.Version = " + "'" + Version.Text + "'"
            spComposicionVersion = Sql1 + Sql2 + Sql3 + Sql4
            Set rstComposicionVersion = db.OpenRecordset(spComposicionVersion, dbOpenSnapshot, dbSQLPassThrough)
            If rstComposicionVersion.RecordCount > 0 Then
                FechaInicio.Caption = rstComposicionVersion!FechaInicio
                FechaFinal.Caption = rstComposicionVersion!FechaFinal
                rstComposicionVersion.Close
                Call Proceso_Click
                    Else
                ZTerminado = Terminado.Text
                ZDesTerminado = DesTerminado.Caption
                ZVersion = Version.Text
                Call Limpia_Click
                Terminado.Text = ZTerminado
                DesTerminado.Caption = ZDesTerminado
                Version.Text = ZVersion
                Version.SetFocus
            End If
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 6
    WVector1.FixedRows = 1
    WVector1.Rows = 100
    
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
                WVector1.Text = "Materia Prima"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WVector1.Text = "Prod.Terminado"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 4200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    Rem WVector1.Row = 0
    Rem For Ciclo = 1 To WVector1.Cols - 1
    Rem      WVector1.Col = Ciclo
    Rem      WTituloVector(Ciclo).Text = WVector1.Text
    Rem      WTituloVector(Ciclo).Left = WVector1.CellLeft + WVector1.Left
    Rem      WTituloVector(Ciclo).Top = WVector1.CellTop + WVector1.Top
    Rem      WTituloVector(Ciclo).Width = WVector1.CellWidth
    Rem      WTituloVector(Ciclo).Height = WVector1.CellHeight
    Rem      WTituloVector(Ciclo).Visible = True
    Rem   Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = 11400
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







