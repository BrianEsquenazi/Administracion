VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCoti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Cotizaciones"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.Frame NombreComercial 
      Height          =   2655
      Left            =   1800
      TabIndex        =   33
      Top             =   1560
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton XAcepta 
         Caption         =   "Confirma Nombre"
         Height          =   375
         Left            =   2280
         TabIndex        =   43
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox XComercial 
         Height          =   285
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   41
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label13 
         Caption         =   "INGRESE EL NOMBRE COMERCIAL DEL PRODUCTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label XDesProve 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3360
         TabIndex        =   40
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label XProve 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   39
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label XDesArti 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3360
         TabIndex        =   38
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label XArti 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Nombre Comercial"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Articulo"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   6600
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Moneda 
      Height          =   315
      Left            =   8040
      TabIndex        =   31
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   3480
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.TextBox Proveedor 
      Height          =   285
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   23
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   120
      TabIndex        =   20
      Top             =   6840
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   5160
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   16
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Cotiza 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   14
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   500
      Left            =   120
      TabIndex        =   12
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   500
      Left            =   1200
      TabIndex        =   11
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   500
      Left            =   2280
      TabIndex        =   9
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   11775
      Begin VB.TextBox WObservaciones 
         Height          =   285
         Left            =   8280
         MaxLength       =   40
         TabIndex        =   22
         Text            =   " "
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox WCondicion 
         Height          =   285
         Left            =   5280
         MaxLength       =   40
         TabIndex        =   21
         Text            =   " "
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Text            =   " "
         Top             =   720
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WPrecio 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   6
         Text            =   " "
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   8280
         TabIndex        =   28
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condicion de Pago"
         Height          =   255
         Left            =   5280
         TabIndex        =   27
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materia  Prima"
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   500
      Left            =   2280
      TabIndex        =   4
      Top             =   6840
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "Coti.frx":0000
      TabIndex        =   3
      Top             =   840
      Width           =   11775
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8880
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1620
      ItemData        =   "Coti.frx":09DE
      Left            =   3480
      List            =   "Coti.frx":09E5
      TabIndex        =   1
      Top             =   6600
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   500
      Left            =   1200
      TabIndex        =   0
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Moneda"
      Height          =   255
      Left            =   6480
      TabIndex        =   30
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label DesProveedor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor"
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cotizacion"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgCoti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 5 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstCotiza As Recordset
Dim spCotiza As String
Dim rstProveedor As Recordset
Dim spProveedor As String
Dim rstMarcas As Recordset
Dim spMarcas As String
Dim XParam As String
Dim Vector(100) As String
Private WMoneda As String
Dim XProveedor As String


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
    
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WPrecio.Text = ""
    WCondicion.Text = ""
    WLinea.Text = ""
    WObservaciones.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    Pantalla.Visible = False
    Opcion.Visible = False
    
    PrgCoti.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()

    spCotiza = "ModificaCotizaMoneda "
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Articulos"

     Opcion.Visible = True
     
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
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
        
            Ayuda.Visible = True
            Ayuda.Text = ""
            
            spProveedor = "ListaProveedoresOrd"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = rstProveedor!Proveedor
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + " " + rstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstProveedor!Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstProveedor.Close
            Ayuda.SetFocus
            
        Case 1
        
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
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
            WArticulo.SetFocus
        
        Case Else
    End Select
                   
    Pantalla.Visible = True
    
End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 10 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -   -   "
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    If Val(DBGrid1.Text) = 0 Then
        WPrecio.Text = ""
            Else
        WPrecio.Text = DBGrid1.Text
    End If
    
    DBGrid1.Col = 3
    WCondicion.Text = DBGrid1.Text
    
    DBGrid1.Col = 4
    WObservaciones.Text = DBGrid1.Text
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    If Moneda.ListIndex = 2 Then
    
        spCambios = "ConsultaCambio  " + "'" + Fecha.Text + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
            ZParidad = rstCambios!Cambio
            ZParidadII = IIf(IsNull(rstCambios!CambioII), "0", rstCambios!CambioII)
            rstCambios.Close
            If ZParidadII <> 0 And ZParidad <> 0 Then
                ZCoeParidad = ZParidadII / ZParidad
                    Else
            End If
                    Else
            m$ = "Se debe informar la paridad"
            G% = MsgBox(m$, 0, "Actualizaion de Informe de Recepcion de Materia Prima")
            Exit Sub
        End If
        
    End If









    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""

    Rem Borra la cotizacion anterior
    
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
    
    
    
    
    

    spCotiza = "BorrarCotizaTotal " + "'" + Cotiza.Text + "'"
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenDynaset, dbSQLPassThrough)

    Renglon = 0
                
    DBGrid1.Refresh
        
    For A = 0 To 3
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = UCase(DBGrid1.Text)
                    
            DBGrid1.Col = 2
            Precio = DBGrid1.Text
                    
            DBGrid1.Col = 3
            Condicion = DBGrid1.Text
                    
            DBGrid1.Col = 4
            Observaciones = DBGrid1.Text
                    
            If Articulo <> "" Then
            
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Cotiza.Text)
                Call Ceros(Auxi1, 6)
                        
                WCotiza = Cotiza.Text
                WRenglon = Str$(Renglon)
                WFecha = Fecha.Text
                WProveedor = Proveedor.Text
                WArticulo = Articulo
                WPrecio = Precio
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WCondicion = Condicion
                WObservaciones = Observaciones
                WClave = Auxi1 + Auxi
                WDate = Date$
                WMoneda = Str$(Moneda.ListIndex)
        
                XParam = "'" + WClave + "','" _
                         + WCotiza + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WProveedor + "','" _
                         + WArticulo + "','" _
                         + WPrecio + "','" _
                         + WCondicion + "','" _
                         + WObservaciones + "','" _
                         + WFechaord + "','" _
                         + WDate + "','" _
                         + WMoneda + "'"
                         
                spCotiza = "AltaCotizaII " + XParam
                Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
                
                WPasaFactor = 0
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WPasaFactor = IIf(IsNull(rstArticulo!Factor), "0", rstArticulo!Factor)
                    rstArticulo.Close
                End If
                
                If WPasaFactor <> 0 Then
                
                    T$ = "Ingreso de Cotizaciones"
                    m$ = "Desea actualizar el costo de reposicion"
                    Respuesta% = MsgBox(m$, 32 + 4, T$)
                    If Respuesta% = 6 Then
                    
                        Recompra = Val(WPrecio) * WPasaFactor
                        
                        If Moneda.ListIndex = 2 Then
                            Recompra = Val(WPrecio) * WPasaFactor * ZCoeParidad
                        End If
                        
                        EmpresaAnterior = Wempresa
                        
            
                        Wempresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
            
                        Wempresa = "0002"
                        txtOdbc = "Empresa02"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
            
                        Wempresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
            
                        Wempresa = "0004"
                        txtOdbc = "Empresa04"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
            
                        Wempresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
            
                        Wempresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
            
                        Wempresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
            
                        Wempresa = "0008"
                        txtOdbc = "Empresa08"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            
                        Wempresa = "0009"
                        txtOdbc = "Empresa09"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            
                        Wempresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            
                        Wempresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo4 = " + "'" + Str$(Recompra) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            
            
                        Select Case Val(EmpresaAnterior)
                            Case 1
                                Wempresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 2
                                Wempresa = "0002"
                                txtOdbc = "Empresa02"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 3
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 4
                                Wempresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 5
                                Wempresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 6
                                Wempresa = "0006"
                                txtOdbc = "Empresa06"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 7
                                Wempresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 8
                                Wempresa = "0008"
                                txtOdbc = "Empresa08"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 9
                                Wempresa = "0009"
                                txtOdbc = "Empresa09"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 10
                                Wempresa = "0010"
                                txtOdbc = "Empresa10"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 11
                                Wempresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                        
                    End If
                
                End If
            
            End If
                        
        Next iRow
            
    Next A
    
    Call Conecta_Empresa
       
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Cotiza.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WPrecio.Text = ""
    WCondicion.Text = ""
    WObservaciones.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    On Error GoTo WError
    
    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WPrecio.Text = ""
    WCondicion.Text = ""
    WObservaciones.Text = ""
    
    Cotiza.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)

    Proveedor.Text = ""
    DesProveedor.Caption = ""
    
    Moneda.ListIndex = 0
    
    For A = 0 To 3
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 4
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    Cotiza.Text = ""
    
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    spCotiza = "ListaCotizaNumero"
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    If rstCotiza.RecordCount > 0 Then
        With rstCotiza
            .MoveLast
            Cotiza.Text = rstCotiza!Cotiza + 1
        End With
        rstCotiza.Close
    End If
    
    Call Conecta_Empresa
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Pantalla.Visible = False
    Opcion.Visible = False

    Cotiza.SetFocus
    
    Exit Sub
    
WError:
     coderr = Err
     Resume Next

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo.Text = UCase(WArticulo.Text)
        
        spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount <= 0 Then
            WArticulo.SetFocus
                    Else
            WDescripcion.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            
            XProveedor = Proveedor.Text
            Call Ceros(XProveedor, 11)
            ClaveMarcas = WArticulo.Text + XProveedor
            
            spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
            Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
            If rstMarcas.RecordCount <= 0 Then
                NombreComercial.Visible = True
                XArti.Caption = WArticulo.Text
                XDesArti.Caption = WDescripcion.Caption
                XProve.Caption = Proveedor.Text
                XDesProve.Caption = DesProveedor.Caption
                XComercial.Text = WDescripcion.Caption
                XComercial.SetFocus
                    Else
                rstMarcas.Close
                WPrecio.SetFocus
            End If
        End If
        
    End If
End Sub

Private Sub XAcepta_Click()

    NombreComercial.Visible = False
    
    XProveedor = Proveedor.Text
    Call Ceros(XProveedor, 11)
    ClaveMarcas = WArticulo.Text + XProveedor
    
    XParam = "'" + ClaveMarcas + "','" _
                + WArticulo.Text + "','" _
                + XProveedor + "','" _
                + XComercial.Text + "'"
                                         
                                         
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Wempresa = "0002"
    txtOdbc = "Empresa02"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Wempresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Wempresa = "0004"
    txtOdbc = "Empresa04"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Wempresa = "0005"
    txtOdbc = "Empresa05"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Wempresa = "0006"
    txtOdbc = "Empresa06"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Wempresa = "0007"
    txtOdbc = "Empresa07"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Wempresa = "0008"
    txtOdbc = "Empresa08"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Wempresa = "0009"
    txtOdbc = "Empresa09"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Wempresa = "0010"
    txtOdbc = "Empresa10"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Wempresa = "0011"
    txtOdbc = "Empresa11"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spMarcas = "AltaMarcas " + XParam
    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Conecta_Empresa
    
    WPrecio.SetFocus
    
End Sub

Private Sub WPrecio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPrecio.Text = Pusing("###,###.###", WPrecio.Text)
        WCondicion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCondicion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WObservaciones.SetFocus
    End If
End Sub

Private Sub WObservaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Alta_Vector
        Call Ingresa_Click
        WArticulo.SetFocus
    End If
End Sub

Private Sub pantalla_Click()

    If XIndice = 0 Then
        Pantalla.Visible = False
        Opcion.Visible = False
    End If
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WProveedor = WIndice.List(Indice)
            spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
                Proveedor.Text = WProveedor
                DesProveedor.Caption = rstProveedor!Nombre
            End If
            
            Ayuda.Visible = False
            Pantalla.Visible = False
            
            
        Case 1
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
                WArticulo.Text = rstArticulo!Codigo
                WDescripcion.Caption = rstArticulo!Descripcion
                    
                DBGrid1.Col = 0
                DBGrid1.Text = rstArticulo!Codigo
                DBGrid1.Col = 1
                DBGrid1.Text = rstArticulo!Descripcion
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                    
                Call DBGrid1.SetFocus
                WPrecio.SetFocus
                    
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4
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
ReDim UserData(0 To 4, 0 To 40)

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
For i = 0 To 4
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 2750
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Cond.Pago"
             DBGrid1.Columns(newcnt).Width = 2900
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Observaciones"
             DBGrid1.Columns(newcnt).Width = 3300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WPrecio.Text = ""
    WCondicion.Text = ""
    WObservaciones.Text = ""
    
    Cotiza.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)

    Proveedor.Text = ""
    DesProveedor.Caption = ""
    
    Moneda.Clear
    
    Moneda.AddItem "Dolares"
    Moneda.AddItem "Pesos"
    Moneda.AddItem "Euros"
    
    Moneda.ListIndex = 0
    
    XEmpresa = Wempresa
        
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    spCotiza = "ListaCotizaNumero"
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    If rstCotiza.RecordCount > 0 Then
        With rstCotiza
            .MoveLast
            Cotiza.Text = rstCotiza!Cotiza + 1
        End With
        rstCotiza.Close
    End If
    
    Call Conecta_Empresa
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgCoti.Caption = "Ingreso de Cotizaciones"
        End If
    End With
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Cotiza.SetFocus
    
End Sub

Private Sub Proceso_Click()

    spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        Proveedor.Text = rstProveedor!Proveedor
        DesProveedor.Caption = rstProveedor!Nombre
    End If

    For A = 0 To 3
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 4
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    Erase Vector
    
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    spCotiza = "ListaCotizaTotal " + "'" + Cotiza.Text + "'"
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    If rstCotiza.RecordCount > 0 Then
        With rstCotiza
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstCotiza!Articulo
                    Auxi1 = rstCotiza!Articulo
                    Vector(Renglon) = Auxi1
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.###", Str$(rstCotiza!Precio))
                
                    DBGrid1.Col = 3
                    DBGrid1.Text = rstCotiza!Condicion
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = rstCotiza!Observaciones
                    
                    Moneda.ListIndex = !Moneda

                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCotiza.Close
    End If
    
    Call Conecta_Empresa
    
    For Da = 1 To Renglon
    
        Lugar1 = Int((Da - 1) / 10) * 10
        Lugar2 = Da - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        Auxi1 = Vector(Da)
        spArticulo = "ConsultaArticulo " + "'" + Auxi1 + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DBGrid1.Col = 1
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
    
    WArticulo.SetFocus

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
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.###", WPrecio.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = WCondicion.Text
            
            DBGrid1.Col = 4
            DBGrid1.Text = WObservaciones.Text
            
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.###", WPrecio.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = WCondicion.Text
            
            DBGrid1.Col = 4
            DBGrid1.Text = WObservaciones.Text
            
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Cotiza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Proceso = "N"
    
        Auxi = Cotiza.Text
        Call Ceros(Auxi, 6)
        WClave = Auxi + "01"
        
        Wempresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        spCotiza = "ConsultaCotiza " + "'" + WClave + "'"
        Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
        If rstCotiza.RecordCount > 0 Then
                    
            Fecha.Text = rstCotiza!Fecha
            Proveedor.Text = rstCotiza!Proveedor
                
            Proceso = "S"
            rstCotiza.Close
                
                Else
                
            WCotiza = Cotiza.Text
            Call Limpia_Click
            Cotiza.Text = WCotiza
            Fecha.SetFocus
            
        End If
        
        Call Conecta_Empresa
        
        If Proceso = "S" Then
            Call Proceso_Click
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Proveedor.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Proveedor.Text) <> 0 Then
            spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                    Proveedor.Text = rstProveedor!Proveedor
                    DesProveedor.Caption = rstProveedor!Nombre
                        Else
                    Proveedor.SetFocus
            End If
            WArticulo.SetFocus
                Else
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 0
            Call Opcion_Click
            Ayuda.SetFocus
        End If
    End If
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spProveedor = "ListaProveedoresOrd"
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                Da = Len(rstProveedor!Nombre) - WEspacios
                
                For aa = 1 To Da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                        Auxi = Str$(rstProveedor!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + rstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
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
    rstProveedor.Close
    
    End If

End Sub

