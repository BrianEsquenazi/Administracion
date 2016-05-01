VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovlab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de E/S de Materia Prima y Productos"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   330
   ClientWidth     =   11805
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11805
   Visible         =   0   'False
   Begin VB.Frame Pass 
      Height          =   1575
      Left            =   3360
      TabIndex        =   34
      Top             =   1920
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton WCancela 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   36
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Ingrese su Password"
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
         Left            =   360
         TabIndex        =   35
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   480
      Width           =   5655
   End
   Begin VB.ComboBox Tipomov 
      Height          =   315
      Left            =   8880
      TabIndex        =   29
      Text            =   " "
      Top             =   120
      Width           =   2415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11040
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   2520
      TabIndex        =   17
      Top             =   6480
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   4680
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   5160
      TabIndex        =   15
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
      TabIndex        =   13
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   500
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   500
      Left            =   1320
      TabIndex        =   10
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   500
      Left            =   2520
      TabIndex        =   8
      Top             =   5880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   11535
      Begin VB.TextBox WLote 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9960
         MaxLength       =   6
         TabIndex        =   33
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox WMovi 
         Height          =   285
         Left            =   8760
         MaxLength       =   1
         TabIndex        =   21
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox WCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   20
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin MSMask.MaskEdBox WTerminado 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WTipo 
         Height          =   285
         Left            =   360
         MaxLength       =   1
         TabIndex        =   18
         Text            =   " "
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
         Height          =   255
         Left            =   9960
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E/S"
         Height          =   255
         Left            =   8760
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   7440
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materia Prima"
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto Terminado"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M/T"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   3840
         TabIndex        =   6
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.CommandButton Graba1 
      Caption         =   "Graba"
      Height          =   500
      Left            =   120
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3825
      Left            =   120
      OleObjectBlob   =   "Movlab.frx":0000
      TabIndex        =   3
      Top             =   840
      Width           =   11565
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10680
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2205
      ItemData        =   "Movlab.frx":0A12
      Left            =   3840
      List            =   "Movlab.frx":0A19
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   500
      Left            =   1320
      TabIndex        =   0
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Movimiento"
      Height          =   285
      Left            =   7080
      TabIndex        =   28
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro Movimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgMovlab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 7 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Tipo As String
Private Articulo As String
Private Terminado As String
Private WTipomov As String
Private Auxiliar(100, 6) As String
Dim rstMovlab As Recordset
Dim spmovlab As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Dim EmpresaActual As String

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    DBGrid1.Col = 6
    DBGrid1.Text = ""
    
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WMovi.Text = ""
    WLote.Text = ""
    WLinea.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub cmdClose_Click()


    Call Limpia_Click

    Rem With rstTerminado
    Rem     .Close
    Rem End With
    Rem With rstArticulo
    Rem     .Close
    Rem End With
    Rem With rstmovlab
    Rem     .Close
    Rem End With
    
    Rem DbsVentas.Close
    Rem DbsCotiza.Close
    Rem DbsEmpresa.Close
    
    Graba1.Enabled = True
    PrgMovlab.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima"

     Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    Select Case Val(EmpresaActual)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    Rem OPEN_FILE_movlab
    Rem OPEN_FILE_TERMINADO
    Rem OPEN_FILE_Articulo
End Sub

Private Sub Graba1_Click()
    If Tipomov.ListIndex = 0 Or Tipomov.ListIndex = 1 Then
        WClave.Text = ""
        Pass.Visible = True
        WClave.SetFocus
    End If
End Sub

Private Sub WCancela_Click()
    Pass.Visible = False
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WClave.Text = "AUTORIZO" Then
            Call Graba_Click
        End If
    End If
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
            spTerminado = "ListaTerminadoConsulta"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Rem IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                            IngresaItem = rstTerminado!Codigo
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
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 1 Then
        WLinea.Text = DBGrid1.Row + 1
        WTipo.Text = DBGrid1.Text
            Else
        WTipo.Text = ""
        WLinea.Text = ""
    End If

    DBGrid1.Col = 1
    If Len(DBGrid1.Text) = 12 Then
        WTerminado.Text = DBGrid1.Text
            Else
        WTerminado.Text = "  -     -   "
    End If

    DBGrid1.Col = 2
    If Len(DBGrid1.Text) = 10 Then
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -   -   "
    End If
    
    DBGrid1.Col = 3
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 4
    WCantidad.Text = DBGrid1.Text
    
    DBGrid1.Col = 5
    WMovi.Text = DBGrid1.Text
    
    DBGrid1.Col = 6
    WLote.Text = DBGrid1.Text
    
    WTipo.SetFocus

End Sub

Private Sub Graba_Click()

    Pass.Visible = False

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = 0
    Erase Auxiliar
    
    spmovlab = "Listamovlab " + "'" + Codigo.Text + "'"
    Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovlab.RecordCount > 0 Then
        With rstMovlab
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Auxiliar(Renglon, 1) = rstMovlab!Tipo
                    Auxiliar(Renglon, 2) = rstMovlab!Terminado
                    Auxiliar(Renglon, 3) = rstMovlab!Articulo
                    Auxiliar(Renglon, 4) = rstMovlab!Cantidad
                    Auxiliar(Renglon, 5) = rstMovlab!Movi
                    Auxiliar(Renglon, 6) = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovlab.Close
    End If
    
    For Da = 1 To Renglon
    
        Tipo = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Cantidad = Auxiliar(Da, 4)
        Movi = Auxiliar(Da, 5)
        Lote = Auxiliar(Da, 6)
        
        Select Case Tipo
            Case "M"
                WControla = 0
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
        
                    WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                    WCodigo = Articulo
                    If Movi = "E" Then
                        WEntradas = Str$(rstArticulo!Entradas - Val(Cantidad))
                        WSalidas = Str$(rstArticulo!Salidas)
                            Else
                        WSalidas = Str$(rstArticulo!Salidas - Val(Cantidad))
                        WEntradas = Str$(rstArticulo!Entradas)
                    End If
                    WDate = Date$
                    rstArticulo.Close
                
                    XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                    spArticulo = "ModificaArticuloMovimientos " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + Articulo + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            WClave = rstLaudo!Clave
                            If Movi = "S" Then
                                WSaldo = Str$(rstLaudo!Saldo + Val(Cantidad))
                                    Else
                                WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad))
                            End If
                            WDate = Date$
                            rstLaudo.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spLaudo = "ModificaLaudoSaldo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            XParam = "'" + Articulo + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                If Movi = "S" Then
                                    WSaldo = Str$(rstMovguia!Saldo + Val(Cantidad))
                                        Else
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                End If
                                WDate = Date$
                                rstMovguia.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                    
                End If
                
            Case "T"
                WControla = 0
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
        
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    WCodigo = Terminado
                    If Movi = "E" Then
                        WEntradas = Str$(rstTerminado!Entradas - Val(Cantidad))
                        WSalidas = Str$(rstTerminado!Salidas)
                            Else
                        WSalidas = Str$(rstTerminado!Salidas - Val(Cantidad))
                        WEntradas = Str$(rstTerminado!Entradas)
                    End If
                    WDate = Date$
                    rstTerminado.Close
                
                    XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                    spTerminado = "ModificaTerminadoMovimientos " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + Terminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WClave = rstHoja!Clave
                            If Movi = "S" Then
                                WSaldo = Str$(rstHoja!Saldo + Val(Cantidad))
                                    Else
                                WSaldo = Str$(rstHoja!Saldo - Val(Cantidad))
                            End If
                            WDate = Date$
                            rstHoja.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spHoja = "ModificaHojaSaldo " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            XParam = "'" + Terminado + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                If Movi = "S" Then
                                    WSaldo = Str$(rstMovguia!Saldo + Val(Cantidad))
                                        Else
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                End If
                                WDate = Date$
                                rstMovguia.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                    
                End If
            
            Case Else
        End Select
        
    Next Da
    
    spmovlab = "Borrarmovlab " + "'" + Codigo.Text + "'"
    Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenDynaset, dbSQLPassThrough)
    
    Renglon = 0
    Erase Auxiliar
    
    DBGrid1.Refresh
                
    For A = 0 To 3
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Tipo = DBGrid1.Text
                                       
            DBGrid1.Col = 1
            Terminado = DBGrid1.Text
                    
            DBGrid1.Col = 2
            Articulo = DBGrid1.Text
                    
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 5
            Movi = DBGrid1.Text
            
            DBGrid1.Col = 6
            Lote = DBGrid1.Text
                    
            If Tipo <> "" Then
                    
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Codigo.Text)
                Call Ceros(Auxi1, 6)
                
                WCodigo = Codigo.Text
                WRenglon = Str$(Renglon)
                WFecha = Fecha.Text
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WTipo = Tipo
                WArticulo = Articulo
                WTerminado = Terminado
                WCantidad = Cantidad
                WMovi = Movi
                WTipomov = Str$(Tipomov.ListIndex)
                Call Ceros(WTipomov, 1)
                WObservaciones = Observaciones.Text
                WClave = Auxi1 + Auxi
                WDate = Date$
                WMarca = ""
                WLote = Lote
                
                Auxiliar(Renglon, 1) = WTipo
                Auxiliar(Renglon, 2) = WTerminado
                Auxiliar(Renglon, 3) = WArticulo
                Auxiliar(Renglon, 4) = WCantidad
                Auxiliar(Renglon, 5) = WMovi
                Auxiliar(Renglon, 6) = WLote

                XParam = "'" + WClave + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WFechaord + "','" _
                         + WMovi + "','" _
                         + WTipomov + "','" _
                         + WObservaciones + "','" _
                         + WDate + "','" _
                         + WMarca + "','" _
                         + WLote + "'"
                         
                spmovlab = "Altamovlab " + XParam
                Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                
        Next iRow
            
    Next A
                
    For Da = 1 To Renglon
    
        Tipo = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Cantidad = Auxiliar(Da, 4)
        Movi = Auxiliar(Da, 5)
        Lote = Auxiliar(Da, 6)
        
        Select Case Tipo
            Case "M"
                WControla = 0
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
        
                    WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                    WCodigo = Articulo
                    If Movi = "E" Then
                        WEntradas = Str$(rstArticulo!Entradas + Val(Cantidad))
                        WSalidas = Str$(rstArticulo!Salidas)
                            Else
                        WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad))
                        WEntradas = Str$(rstArticulo!Entradas)
                    End If
                    WDate = Date$
                    rstArticulo.Close
                
                    XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                    spArticulo = "ModificaArticuloMovimientos " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + Articulo + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            WClave = rstLaudo!Clave
                            If Movi = "E" Then
                                WSaldo = Str$(rstLaudo!Saldo + Val(Cantidad))
                                    Else
                                WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad))
                            End If
                            WDate = Date$
                            rstLaudo.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spLaudo = "ModificaLaudoSaldo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            XParam = "'" + Articulo + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                If Movi = "E" Then
                                    WSaldo = Str$(rstMovguia!Saldo + Val(Cantidad))
                                        Else
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                End If
                                WDate = Date$
                                rstMovguia.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                End If
                
                
            Case "T"
                WControla = 0
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
        
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    WCodigo = Terminado
                    If Movi = "E" Then
                        WEntradas = Str$(rstTerminado!Entradas + Val(Cantidad))
                        WSalidas = Str$(rstTerminado!Salidas)
                            Else
                        WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad))
                        WEntradas = Str$(rstTerminado!Entradas)
                    End If
                    WDate = Date$
                    rstTerminado.Close
                
                    XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                    spTerminado = "ModificaTerminadoMovimientos " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + Terminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WClave = rstHoja!Clave
                            If Movi = "E" Then
                                WSaldo = Str$(rstHoja!Saldo + Val(Cantidad))
                                    Else
                                WSaldo = Str$(rstHoja!Saldo - Val(Cantidad))
                            End If
                            WDate = Date$
                            rstHoja.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spHoja = "ModificaHojaSaldo " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            XParam = "'" + Terminado + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                If Movi = "E" Then
                                    WSaldo = Str$(rstMovguia!Saldo + Val(Cantidad))
                                        Else
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                End If
                                WDate = Date$
                                rstMovguia.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                End If
            
            Case Else
        End Select
        
    Next Da
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WMovi.Text = ""
    WLote.Text = ""
    
    WTipo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WMovi.Text = ""
    WLote.Text = ""

    Codigo.Text = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Pass.Visible = False
    
    For A = 0 To 3
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 6
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    Rem With rstmovlab
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
    
    spmovlab = "ListamovlabNumero"
    Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
        With rstMovlab
            .MoveLast
            Codigo.Text = rstMovlab!Codigo + 1
        End With
        rstMovlab.Close
    End If
    
    Tipomov.ListIndex = 0
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Codigo.SetFocus

End Sub


Private Sub WTipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WTipo.Text = "M" Or WTipo.Text = "T" Then
            If WTipo.Text = "M" Then
                WArticulo.SetFocus
                    Else
                WTerminado.SetFocus
            End If
                Else
            WTipo.SetFocus
        End If
    End If
End Sub

Private Sub WTerminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            Rem WDescripcion.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            WCantidad.SetFocus
                Else
            WTerminado.SetFocus
        End If
    End If
End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDescripcion.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            WCantidad.SetFocus
                Else
            WArticulo.SetFocus
        End If
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        Select Case Tipomov.ListIndex
            Case 0, 2
                WMovi.Text = "E"
            Case 1, 3
                WMovi.Text = "S"
            Case Else
        End Select
        WLote.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Wlote_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WTipo.Text = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                WCanti = 0
                XParam = "'" + WLote.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WCanti = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra = "S" Then
                If WCanti >= Val(WCantidad.Text) Or WMovi.Text = "E" Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    WTipo.SetFocus
                        Else
                    m$ = WArticulo.Text + " Stock Insufucuente. Cantidad:" + Str$(WCanti)
                    G% = MsgBox(m$, 0, "Moviimentos varios de Stock")
                End If
                    Else
                m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote.Text + " inexistente"
                G% = MsgBox(m$, 0, "Movimientos Varios de Stock")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WCanti = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra = "S" Then
                If WCanti >= Val(WCantidad.Text) Or WMovi.Text = "E" Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    WTipo.SetFocus
                        Else
                    m$ = WTerminado.Text + " Stock Insufucuente. Cantidad:" + Str$(WCanti)
                    G% = MsgBox(m$, 0, "Movimientos Varios de Stock")
                End If
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                G% = MsgBox(m$, 0, "Movimientos Varios de Stock")
            End If
            
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
        
            spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WTipo.Text = "T"
                WTerminado.Text = Claveven$
                Rem WDescripcion.Caption = rstTerminado!Descripcion
                    
                DBGrid1.Col = 0
                DBGrid1.Text = "T"
                DBGrid1.Col = 1
                DBGrid1.Text = rstTerminado!Codigo
                DBGrid1.Col = 3
                Rem DBGrid1.Text = rstTerminado!Descripcion
                DBGrid1.Text = ""
                
                rstTerminado.Close
                    
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
        
            spArticulo = "ConsultaArticulo " + "'" + Claveven$ + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WTipo.Text = "M"
                WArticulo.Text = rstArticulo!Codigo
                WDescripcion.Caption = rstArticulo!Descripcion
                    
                DBGrid1.Col = 0
                DBGrid1.Text = "M"
                DBGrid1.Col = 2
                DBGrid1.Text = rstArticulo!Codigo
                DBGrid1.Col = 3
                DBGrid1.Text = rstArticulo!Descripcion
                
                rstArticulo.Close
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                                        
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
                    
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5
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
ReDim UserData(0 To 6, 0 To 40)

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
For i = 0 To 6
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Prod.Terminado"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Materia Prima"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3620
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Movimiento"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
             
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Lote"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Codigo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
 
    Rem With rstmovlab
    Rem     .Index = "Clave"
    Rem    Claveven$ = "99999999"
    Rem    .Seek "<=", Claveven$
    Rem    If .NoMatch = False Then
    Rem        Codigo.Text = !Codigo + 1
    Rem            Else
    Rem        Codigo.Text = ""
    Rem    End If
    Rem End With
    
    Codigo.Text = ""
    
    spmovlab = "ListamovlabNumero"
    Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
        With rstMovlab
            .MoveLast
            Codigo.Text = rstMovlab!Codigo + 1
        End With
        rstMovlab.Close
    End If
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgMovlab.Caption = "Listado de E/S de Materia Prima y Productos :  " + !Nombre
        End If
    End With
    EmpresaActual = WEmpresa
    
    Tipomov.Clear
    
    Tipomov.AddItem "Entrada"
    Tipomov.AddItem "Salida"
    
    Tipomov.ListIndex = 0
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Graba1.Enabled = True
    Codigo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For A = 0 To 3
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 6
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Erase Auxiliar
    Renglon = 0
    
    spmovlab = "Listamovlab " + "'" + Codigo.Text + "'"
    Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstMovlab!Tipo
                
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstMovlab!Terminado
                    Auxi1 = rstMovlab!Terminado
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = rstMovlab!Articulo
                    Auxi2 = rstMovlab!Articulo
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", rstMovlab!Cantidad)
                
                    DBGrid1.Col = 5
                    DBGrid1.Text = rstMovlab!Movi
                
                    DBGrid1.Col = 6
                    DBGrid1.Text = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                    
                    Tipomov.ListIndex = Val(rstMovlab!Tipomov)
                    Observaciones.Text = rstMovlab!Observaciones
                    
                    WMarca = IIf(IsNull(rstMovlab!Marca), "", rstMovlab!Marca)
                    If WMarca = "X" Then
                        Graba1.Enabled = False
                            Else
                        Graba1.Enabled = True
                    End If
                    
                    Auxiliar(Renglon, 1) = Auxi1
                    Auxiliar(Renglon, 2) = Auxi2
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovlab.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0

    For Da = 1 To WRenglon
    
        Auxi1 = Auxiliar(Da, 1)
        Auxi2 = Auxiliar(Da, 2)
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DBGrid1.Col = 3
            Rem DBGrid1.Text = rstTerminado!Descripcion
            DBGrid1.Text = ""
            rstTerminado.Close
        End If
        
        spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DBGrid1.Col = 3
            DBGrid1.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
    Next Da

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
    
    WTipo.SetFocus

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
            DBGrid1.Text = WTipo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WTerminado.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                
            DBGrid1.Col = 5
            DBGrid1.Text = WMovi.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WLote.Text
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WTipo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WTerminado.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                
            DBGrid1.Col = 5
            DBGrid1.Text = WMovi.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WLote.Text
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spmovlab = "Listamovlab " + "'" + Codigo.Text + "'"
        Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovlab.RecordCount > 0 Then
            Fecha.Text = rstMovlab!Fecha
            rstMovlab.Close
            Call Proceso_Click
                Else
            Graba1.Enabled = True
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
            Observaciones.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WTipo.SetFocus
    End If
End Sub

