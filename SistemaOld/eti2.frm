VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEti2 
   Caption         =   "Impresion de Etiquetas"
   ClientHeight    =   5910
   ClientLeft      =   1080
   ClientTop       =   1920
   ClientWidth     =   9900
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5910
   ScaleWidth      =   9900
   Begin VB.Frame PantaDirEntrega 
      Caption         =   "Seleccion de Lugar de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   8535
      Begin VB.ListBox ListaDirEntrega 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   8295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control de Etiquetas"
      Height          =   3975
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton Baja 
         Caption         =   "  Limpia Etiquetas"
         Height          =   495
         Left            =   6960
         TabIndex        =   24
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Tara 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   23
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   3720
         TabIndex        =   21
         Text            =   " "
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox Descripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   19
         Text            =   " "
         Top             =   1440
         Width           =   5775
      End
      Begin VB.TextBox Etiquetas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   16
         Text            =   " "
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Cantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   15
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Lote 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   14
         Text            =   "  "
         Top             =   1920
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Terminado 
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   327680
         Enabled         =   0   'False
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox Cliente 
         Height          =   285
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   6960
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   6960
         TabIndex        =   6
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Tara"
         Height          =   255
         Left            =   4560
         TabIndex        =   22
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label DesProducto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Label6 
         Caption         =   "Descripcion"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label DesCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad de Etiquetas"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Lote"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Producto Terminado"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6240
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "eti1.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1620
      ItemData        =   "eti2.frx":0000
      Left            =   720
      List            =   "eti2.frx":0007
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgEti2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLote As String
Private WCantidad As String
Private WImpreadi As String
Private WClase As String
Private WIntervencion As String
Private WNaciones As String
Private WEmbalaje As String
Private WTipoeti As String
Private WObservaciones As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim XParam As String
Dim WDirentrega As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String


Private Sub Acepta_Click()

    On Error GoTo WError
    
    Rem Listado.DataFiles(0) = WEmpresa + "coti.mdb"
    Rem Listado.DataFiles(1) = ""
    Rem Listado.DataFiles(0) = WEmpresa + "VENT.mdb"
    Rem Listado.DataFiles(2) = WEmpresa + "admi.mdb"
    
    Salida = "N"
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Imprsion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    If Salida <> "S" Then

    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Imprsion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Rem .Delete
                Rem .MoveNext
                Rem If .EOF = True Then
                Rem     Exit Do
                Rem End If
            Loop
        End If
    End With
    
    WTara = Val(Tara.Text)
    WNeto = Val(Cantidad.Text)
    
    If WTara = 0 Then
        WBruto = 0
            Else
        WBruto = WTara + WNeto
    End If

    With rstEtiqueta
        For Da = 1 To Val(Etiquetas)
            .Index = "Codigo"
            .AddNew
            !Codigo = Da
            WLote = Lote.Text
            Call Ceros(WLote, 5)
            WCantidad = Cantidad.Text
            Call Ceros(WCantidad, 4)
            !Terminado = Terminado.Text
            !Lote = WLote
            !Cliente = Cliente.Text
            !Cantidad = Val(Cantidad.Text)
            !Nombre = Descripcion.Text
            !Impre1 = Mid$(Terminado.Text, 4, 5) + Right$(Terminado.Text, 3) + Space$(1) + WLote + Space$(1) + WCantidad
            WRazon = ""
            Rem WDirEntrega = ""
            spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                WRazon = rstClientes!Razon
                Rem WDirEntrega = rstClientes!DirEntrega
                rstClientes.Close
            End If
            !Razon = WRazon
            !DirEntrega = WDirentrega
            !Clase = WClase
            !Intervencion = WIntervencion
            !Naciones = WNaciones
            !Embalaje = WEmbalaje
            !Bruto = WBruto
            !Tara = WTara
            !Neto = WNeto
            !Observaciones = WObservaciones
            
            .Update
        Next Da
    End With

    Listado.WindowTitle = "Emision de Etiquetas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Da = 0

    If Len(Descripcion.Text) > 20 Then
        For Da = 25 To 1 Step -1
            If Mid$(Descripcion.Text, Da, 1) <> Space$(1) Then
                Exit For
            End If
        Next Da
    End If
    
    If Tipo.ListIndex = 0 Then
        If WImpreadi <> "S" Then
            If Da > 20 Then
                Listado.ReportFileName = "eti10.rpt"
                    Else
                Listado.ReportFileName = "eti1.rpt"
            End If
                Else
            m$ = " Coloque la etiqueta que en su margen tengo el numero " + WTipoeti
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
            If Da > 20 Then
                Listado.ReportFileName = "eti110.rpt"
                    Else
                Listado.ReportFileName = "eti101.rpt"
            End If
        End If
            Else
        If WImpreadi <> "S" Then
            If Da > 20 Then
                Listado.ReportFileName = "weti20.rpt"
                    Else
                Listado.ReportFileName = "weti2.rpt"
            End If
                Else
            m$ = "Producto Peligrosos no se pueden imprimir en etiquetas Chicas"
            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
        End If
    End If
    
    Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
   
    Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
    Rem Listado.Connect = Connect()
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Listado.DataFiles(1) = ""
   
    Listado.Destination = 1
    Listado.PrinterCopies = 1
    Listado.Action = 1
    
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
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
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Baja_Click()
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
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
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    PrgEti2.Hide
    Unload Me
    PrgHoja.Show
End Sub

Sub Form_Load()

    On Error GoTo WError

    Tipo.Clear
    
    Tipo.AddItem "Grande"
    Tipo.AddItem "Chica"
    
    Select Case Val(WEmpresa)
        Case 1, 5
            Tipo.ListIndex = 1
        Case Else
            Tipo.ListIndex = 0
    End Select

    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Lote.Text = ""
    Descripcion.Text = ""
    Cantidad.Text = ""
    Etiquetas.Text = ""
    Tara.Text = ""
    
    DesCliente.Caption = ""
    DesProducto.Caption = ""
    
    Lote.Text = PLote
    Terminado.Text = PTerminado
    
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Terminado.Text = rstTerminado!Codigo
        If Val(WEmpresa) = 2 Or Val(WEmpresa) = 4 Or Val(WEmpresa) = 8 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 9 Then
            Descripcion.Text = rstTerminado!Descripcion
        End If
        WImpreadi = ""
        WClase = ""
        WIntervencion = ""
        WNaciones = ""
        WEmbalaje = ""
        WImpreadi = rstTerminado!Impreadi
        WClase = rstTerminado!Clase
        WIntervencion = rstTerminado!Intervencion
        WNaciones = rstTerminado!Naciones
        WEmbalaje = rstTerminado!Embalaje
        WTipoeti = IIf(IsNull(rstTerminado!Tipoeti), "", rstTerminado!Tipoeti)
        WObservaciones = IIf(IsNull(rstTerminado!Observaciones), "", rstTerminado!Observaciones)
    End If
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgEti2.Caption = "Impresion de Etiquetas :  " + !Nombre
        End If
    End With
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub




Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cliente.Text <> "" Then
        
            XEmpresa = WEmpresa
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2, 4, 8, 9
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
       
            spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                Cliente.Text = rstClientes!Cliente
                DesCliente.Caption = rstClientes!Razon
                
                Erase ZDirEntrega
                
                ZDirEntrega(1) = rstClientes!DirEntrega
                ZDirEntrega(2) = Trim(IIf(IsNull(rstClientes!DirEntregaII), "", rstClientes!DirEntregaII))
                ZDirEntrega(3) = Trim(IIf(IsNull(rstClientes!DirEntregaIII), "", rstClientes!DirEntregaIII))
                ZDirEntrega(4) = Trim(IIf(IsNull(rstClientes!DirEntregaIV), "", rstClientes!DirEntregaIV))
                ZDirEntrega(5) = Trim(IIf(IsNull(rstClientes!DirEntregaV), "", rstClientes!DirEntregaV))
                
                WDirentrega = ""
                CantiLugarEntrega = 0
                For CicloDirEntrega = 1 To 5
                    If ZDirEntrega(CicloDirEntrega) <> "" Then
                        WDirentrega = ZDirEntrega(CicloDirEntrega)
                        ZLugarDirEntrega = CicloDirEntrega
                        CantiLugarEntrega = CantiLugarEntrega + 1
                    End If
                Next CicloDirEntrega
                
                If CantiLugarEntrega > 1 Then
                    ListaDirEntrega.Clear
                    For CicloDirEntrega = 1 To 5
                        If ZDirEntrega(CicloDirEntrega) <> "" Then
                            ListaDirEntrega.AddItem ZDirEntrega(CicloDirEntrega)
                        End If
                    Next CicloDirEntrega
                    PantaDirEntrega.Top = 840
                    PantaDirEntrega.Visible = True
                    ListaDirEntrega.SetFocus
                        Else
                    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                End If
                
                rstClientes.Close
                
                    Else
                Cliente.Text = Claveven$
                Cliente.SetFocus
            End If
            
            XParam = "'" + Cliente.Text + _
                    Terminado.Text + "'"
            
            spPrecios = "ConsultaPrecios " + XParam
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Descripcion.Text = Left$(rstPrecios!Descripcion, 25)
                rstPrecios.Close
                    Else
                Rem Descripcion.Text = DesProducto.Caption
            End If
            
            Select Case Val(XEmpresa)
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
            
            Cantidad.SetFocus
        End If
    End If
End Sub

Private Sub ListaDirEntrega_Click()
    ZLugarDirEntrega = ListaDirEntrega.ListIndex + 1
    WDirentrega = ZDirEntrega(ZLugarDirEntrega)
    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
    PantaDirEntrega.Visible = False
    Cantidad.SetFocus
End Sub


Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Etiquetas.SetFocus
    End If
End Sub

Private Sub Etiquetas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.SetFocus
    End If
End Sub
