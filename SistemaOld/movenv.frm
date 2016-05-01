VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovEnv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Movimiento de Entrada y Salida de Envases"
   ClientHeight    =   7605
   ClientLeft      =   390
   ClientTop       =   720
   ClientWidth     =   11490
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7605
   ScaleWidth      =   11490
   Visible         =   0   'False
   Begin VB.Frame Aviso 
      Height          =   4335
      Left            =   1320
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton FinAviso 
         Caption         =   "ACEPTA"
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
         Left            =   2520
         TabIndex        =   32
         Top             =   3840
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid Muestra 
         Height          =   2535
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4471
         _Version        =   393216
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "SE ENCUENTRANEN  NEGATIVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   5895
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "LOS SALDOS DE LOS ENVASES EN COMODATO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "ADVERTENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   500
      Left            =   1080
      TabIndex        =   21
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Cliente 
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   18
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   5400
      TabIndex        =   15
      Top             =   6480
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   8880
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Codigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   500
      Left            =   0
      TabIndex        =   10
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   500
      Left            =   4320
      TabIndex        =   9
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   500
      Left            =   2160
      TabIndex        =   7
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   7095
      Begin VB.TextBox WTipo 
         Height          =   300
         Left            =   6000
         MaxLength       =   1
         TabIndex        =   20
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox WEnvase 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   360
         MaxLength       =   6
         TabIndex        =   19
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox WCantidad 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   16
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   255
         Left            =   6000
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   4920
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Envase"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   500
      Left            =   3240
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4215
      Left            =   0
      OleObjectBlob   =   "movenv.frx":0000
      TabIndex        =   3
      Top             =   960
      Width           =   8535
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   5325
      ItemData        =   "movenv.frx":09FE
      Left            =   8640
      List            =   "movenv.frx":0A05
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   300
      Left            =   3000
      TabIndex        =   22
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Movimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "PrgMovEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 4 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxiliar(100, 5) As String
Dim rstMovenv As Recordset
Dim spMovenv As String
Dim rstClientes As Recordset
Dim spClientes As String
Dim rstEnvases As Recordset
Dim spEnvases As String
Dim XParam As String
Private Stk(19, 4) As String

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    WEnvase.Text = ""
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WTipo.Text = ""
    WLinea.Text = ""
    
    WEnvase.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    PrgMovEnv.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()

 Rem XParam = "'" + "19991018" + "','" _
 rem                + "E" + "','" _
 rem                + "X" + "'"
 Rem
 Rem
 Rem    spMovenv = "ModificaMovenvMovi " + XParam
 Rem    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
 Rem
 Rem XParam = "'" + "19991018" + "','" _
 rem                + "S" + "','" _
 rem                + "E" + "'"
 Rem
 Rem
 Rem    spMovenv = "ModificaMovenvMovi " + XParam
 Rem    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
 Rem
 Rem XParam = "'" + "19991018" + "','" _
 rem                + "X" + "','" _
 rem                + "S" + "'"
 Rem
 Rem
 Rem    spMovenv = "ModificaMovenvMovi " + XParam
 Rem    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)

 XParam = "'" + "20000101" + "'"
 spMovenv = "DepuraMovenv " + XParam
 Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Envases"
     Opcion.AddItem "Clientes"

     Opcion.Visible = True
     
 End Sub

Private Sub FinAviso_Click()
    Aviso.Visible = False
    Codigo.SetFocus
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    Rem OPEN_FILE_MovEnv
    Rem OPEN_FILE_Clientes
    Rem OPEN_FILE_ENVASES
End Sub

 Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Rem XIndice = 0
    
    Select Case XIndice
        Case 0
            spEnvases = "ListaEnvases"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstEnvases.RecordCount > 0 Then
                With rstEnvases
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEnvases!Envases) + " " + rstEnvases!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstEnvases!Envases
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnvases.Close
            End If
            
        Case 1
        
            spClientes = "ListaCliente"
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
    Pantalla.Visible = True
            
End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Val(DBGrid1.Text) <> 0 Then
        WLinea.Text = DBGrid1.Row + 1
        WEnvase.Text = DBGrid1.Text
            Else
        WEnvase.Text = ""
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    If Val(DBGrid1.Text) <> 0 Then
        WCantidad.Text = DBGrid1.Text
            Else
        WCantidad.Text = ""
    End If
    
    DBGrid1.Col = 3
    WTipo.Text = DBGrid1.Text
    
    WEnvase.SetFocus

End Sub

Private Sub Graba_Click()
                
    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "La fecha del movimiento de envases es incorrecta"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
        Exit Sub
    End If
                
                
    Cliente.Text = UCase(Cliente.Text)
        
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    spMovenv = "BorrarMovenv " + "'" + Codigo.Text + "'"
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenDynaset, dbSQLPassThrough)
    
    Erase Stk
    LugarEnvase = 0
    Renglon = 0
    DBGrid1.Refresh
        
                                        
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Envase = DBGrid1.Text
                    
            DBGrid1.Col = 2
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 3
            Tipo = DBGrid1.Text
                    
            If Val(Envase) <> 0 Then
                        
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Codigo.Text)
                Call Ceros(Auxi1, 6)
                
                WTipo = "1"
                WClave = WTipo + Auxi1 + Auxi
                WCodigo = Codigo.Text
                WRenglon = Str$(Renglon)
                WCliente = Cliente.Text
                WFecha = Fecha.Text
                WEnvase = Envase
                WCantidad = Val(Cantidad)
                Select Case Tipo
                    Case "E"
                        WMovimiento = "S"
                    Case "D"
                        WMovimiento = "E"
                    Case Else
                End Select
                        
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                
                XParam = "'" + WClave + "','" _
                         + WTipo + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WFechaord + "','" _
                         + WCliente + "','" _
                         + WEnvase + "','" _
                         + WMovimiento + "','" _
                         + WCantidad + "'"
                         
                spMovenv = "AltaMovenv " + XParam
                Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
                
                EntraEnvase = "S"
                For CiclaEnvase = 1 To LugarEnvase
                    If Val(Stk(CiclaEnvase, 1)) = Envase Then
                        EntraEnvase = "N"
                    End If
                Next CiclaEnvase
                
                If EntraEnvase = "S" Then
                    LugarEnvase = LugarEnvase + 1
                    Stk(LugarEnvase, 1) = Str$(Envase)
                End If
                
            End If
                
        Next iRow
            
    Next a
    
    Call Calcula_Saldo
            
    Call Limpia_Click
        
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
        
End Sub

Private Sub Calcula_Saldo()

    Rem On Error GoTo Error_saldo

    XParam = "'" + Cliente.Text + "','" _
                + Cliente.Text + "'"

    spMovenv = "ListaMovenvDesdeHastaCliente " + XParam
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovenv.RecordCount > 0 Then
    
        With rstMovenv
            .MoveFirst
            Do
                If .EOF = False Then

                    For Da = 1 To 9
                        If Val(Stk(Da, 1)) = !Envase Then
                            If !Movimiento = "S" Then
                                Stk(Da, 2) = Str$(Val(Stk(Da, 2)) + !Cantidad)
                                    Else
                                Stk(Da, 2) = Str$(Val(Stk(Da, 2)) - !Cantidad)
                            End If
                        End If
                    
                    Next Da
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovenv.Close
    End If
    
    ImpreNegativo = "N"
    For Da = 1 To 9
        If Val(Stk(Da, 2)) < 0 Then
            ImpreNegativo = "S"
        End If
    Next Da
    
    If ImpreNegativo = "S" Then
        
        Muestra.Clear

        Muestra.FixedCols = 1
        Muestra.Cols = 4
        Muestra.FixedRows = 1
        Muestra.Rows = 20
    
        Muestra.ColWidth(0) = 200
        Muestra.Row = 0
    
        For Ciclo = 1 To Muestra.Cols - 1
            Muestra.Col = Ciclo
            Select Case Ciclo
                Case 1
                    Muestra.Text = "Envase"
                    Muestra.ColWidth(Ciclo) = 1200
                    Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
                Case 2
                    Muestra.Text = "Descripcion"
                    Muestra.ColWidth(Ciclo) = 2500
                    Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
                Case 3
                    Muestra.Text = "Saldo"
                    Muestra.ColWidth(Ciclo) = 1500
                    Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            End Select
        Next Ciclo
    
        Rem Parametro que indica que el usuario puede
        Rem modificar el tamaño de las celdas
        Muestra.AllowUserResizing = flexResizeBoth
    
        Muestra.Col = 1
        Muestra.Row = 1
        LugarMuestra = 0
        
        For Da = 1 To 9
            If Val(Stk(Da, 2)) <> 0 Then
            
                LugarMuestra = LugarMuestra + 1
                
                Muestra.Row = LugarMuestra
                
                Muestra.Col = 1
                Muestra.Text = Stk(Da, 1)
                
                spEnvases = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
                Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvases.RecordCount > 0 Then
                    Muestra.Col = 2
                    Muestra.Text = rstEnvases!Descripcion
                    rstEnvases.Close
                End If
                
                Muestra.Col = 3
                Muestra.Text = Stk(Da, 2)
            End If
        Next Da
        
        Aviso.Visible = True
    
    End If

End Sub


Private Sub Ingresa_Click()

    WLinea.Text = ""
    WEnvase.Text = ""
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WTipo.Text = ""
    
    WEnvase.SetFocus
    
End Sub

Private Sub Limpia_Click()

    Pantalla.Visible = False
    
    WLinea.Text = ""
    WEnvase.Text = ""
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WTipo.Text = ""

    Codigo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cliente.Text = ""
    DesCliente.Caption = ""
    
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 3
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    Rem With rstMovenv
    Rem     .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Codigo.Text = !Codigo + 1
    Rem             Else
    Rem         Codigo.Text = ""
    Rem     End If
    Rem End With

    Codigo.Text = ""
    
    spMovenv = "ListaMovenvTotal"
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovenv.RecordCount > 0 Then
        With rstMovenv
            .MoveLast
            Codigo.Text = rstMovenv!Codigo + 1
        End With
        rstMovenv.Close
            Else
        Codigo.Text = "1"
    End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Codigo.SetFocus

End Sub

Private Sub WEnvase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + WEnvase.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            WDescripcion.Caption = rstEnvases!Descripcion
            WCantidad.SetFocus
                Else
            WEnvase.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WTipo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WTipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WTipo.Text = "E" Or WTipo.Text = "D" Then
            Call Alta_Vector
            Call Ingresa_Click
            WEnvase.SetFocus
                Else
            WTipo.SetFocus
        End If
    End If
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
        
            spEnvases = "ConsultaEnvases " + "'" + Claveven$ + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                    WEnvase.Text = rstEnvases!Envases
                    WDescripcion.Caption = rstEnvases!Descripcion
                    
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstEnvases!Envases
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstEnvases!Descripcion
                    
                    Call Alta_Vector
                    WLinea.Text = WAnterior + 1
                    If Val(WLinea.Text) > 0 Then
                        DBGrid1.Row = Val(WLinea.Text) - 1
                    End If
                    
                    Call DBGrid1.SetFocus
                    WCantidad.SetFocus
                    
            End If
        
            
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
        
            spClientes = "ConsultaCliente " + "'" + Claveven$ + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                    Cliente.Text = rstClientes!Cliente
                    DesCliente.Caption = rstClientes!Razon
                    Cliente.SetFocus
            End If

        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub


' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 3, 0 To 40)

mTotalRows& = 40

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 3
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Envase"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "(E)ntregado/(D)evolucion"
             DBGrid1.Columns(newcnt).Width = 2000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 0
             DBGrid1.Columns(newcnt).Locked = True
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Rem With rstMovenv
    Rem     .Index = "Clave"
    Rem Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Codigo.Text = !Codigo + 1
    Rem             Else
    Rem         Codigo.Text = ""
    Rem     End If
    Rem End With
    
    Codigo.Text = ""
    
    spMovenv = "ListaMovenvTotal"
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovenv.RecordCount > 0 Then
        With rstMovenv
            .MoveLast
            Codigo.Text = rstMovenv!Codigo + 1
        End With
        rstMovenv.Close
            Else
        Codigo.Text = "1"
    End If
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgMovEnv.Caption = "Listado de Movimiento de Entrada y Salida de Envases :  " + !Nombre
        End If
    End With
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 3
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    Erase Auxiliar
    
    spMovenv = "Listamovenv " + "'" + Codigo.Text + "'"
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovenv.RecordCount > 0 Then
    
        With rstMovenv
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Cliente.Text = rstMovenv!Cliente
        
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstMovenv!Envase
                    Auxi1 = rstMovenv!Envase
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", rstMovenv!Cantidad)
                
                    Select Case rstMovenv!Movimiento
                        Case "E"
                            DBGrid1.Col = 3
                            DBGrid1.Text = "D"
                        Case "S"
                            DBGrid1.Col = 3
                            DBGrid1.Text = "E"
                        Case Else
                    End Select
        
                    Auxiliar(Renglon, 1) = Auxi1
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovenv.Close
                
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Renglon = 1 To WRenglon
    
        Auxi1 = Auxiliar(Renglon, 1)
    
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        
        spEnvases = "ConsultaEnvases " + "'" + Auxi1 + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            DBGrid1.Col = 1
            DBGrid1.Text = rstEnvases!Descripcion
            WEnvase.SetFocus
        End If
        
    Next Renglon
    
    Cliente.Text = UCase(Cliente.Text)
    spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        DesCliente.Caption = rstClientes!Razon
    End If

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    WEnvase.SetFocus

End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            
            DBGrid1.Text = WEnvase.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = WTipo.Text
                
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WEnvase.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = WTipo.Text
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spMovenv = "ListaMovenv " + "'" + Codigo.Text + "'"
        Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovenv.RecordCount > 0 Then
            Fecha.Text = rstMovenv!Fecha
            rstMovenv.Close
            Call Proceso_Click
                Else
            WCodigo = Codigo.Text
            Call Limpia_Click
            Codigo.Text = WCodigo
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        If Cliente.Text <> "" Then
            spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                Cliente.Text = rstClientes!Cliente
                DesCliente.Caption = rstClientes!Razon
                    Else
                Cliente.Text = Claveven$
                Cliente.SetFocus
            End If
        End If
        WEnvase.SetFocus
    End If
End Sub
