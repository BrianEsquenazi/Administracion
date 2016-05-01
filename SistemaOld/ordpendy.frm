VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgOrdPenDy 
   Caption         =   "Listado de Ordenes Pendientes de Dy por Planta"
   ClientHeight    =   3510
   ClientLeft      =   2640
   ClientTop       =   1020
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3510
   ScaleWidth      =   6630
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4815
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
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   1800
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
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   1800
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
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   1335
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
         Height          =   495
         Left            =   2640
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WOrdPendiente.rpt"
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
End
Attribute VB_Name = "PrgOrdPenDy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrden As Recordset
Dim spOrden As String
Dim XParam As String
Dim Vector(10000, 2) As String
Dim Empe(100, 10) As String

Private Sub Acepta_Click()

    


    XParam = "'" + "'"

    spOrden = "ModificaOrdenSaldo " + XParam
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase Vector
    Lugar = 0
    
    spOrden = "ListaOrdenTotal "
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
    
    With rstOrden
         .MoveFirst
         Do
             If .EOF = False Then
                 WClave = rstOrden!Clave
                 WOrden = rstOrden!Orden
                 WFecha2 = rstOrden!fecha2
                 WSaldo = Str$(rstOrden!Cantidad - rstOrden!Recibida)
                 If Val(WSaldo) > 0 Then
                    Entra = "S"
                    For XX = 1 To Lugar
                        If Val(Vector(XX, 1)) = WOrden Then
                            Entra = "N"
                            Exit For
                        End If
                    Next XX
                    
                    If Entra = "S" Then
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = WOrden
                        Vector(Lugar, 2) = Right$(WFecha2, 4) + Mid$(WFecha2, 4, 2) + Left$(WFecha2, 2)
                    End If
                    
                 End If
                 .MoveNext
                     Else
                 Exit Do
             End If
         Loop
    End With
    rstOrden.Close
    
    End If
    
    For XX = 1 To Lugar
        WOrden = Vector(XX, 1)
        WFecha2 = Vector(XX, 2)
        XParam = "'" + WOrden + "','" _
                     + WFecha2 + "'"
    
        spOrden = "ModificaOrdenFecha2 " + XParam
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    Next XX
    
    Listado.WindowTitle = "Listado de Ordenes Pendientes por Articulo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Uno = "{Orden.Saldo} > 0.00"
    Rem Dos = " and {Orden.OrdFecha2} in " + Chr$(34) + "00000000" + Chr$(34) + " to " + Chr$(34) + "99999999" + Chr$(34)
    Rem Tres = " and {Orden.Articulo} in " + Chr$(34) + "DW-000-000" + Chr$(34) + " to " + Chr$(34) + "DY-999-999" + Chr$(34)

    Rem Listado.GroupSelectionFormula = Uno + Dos + Tres
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
    
        Listado.SQLQuery = "SELECT Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Fecha2, Orden.Saldo, Orden.OrdFecha2, Orden.Carpeta, Orden.PedidoImpo, Orden.FechaImpo, Orden.TipoImpo, " _
                    + "Proveedor.Nombre, " _
                    + "Articulo.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Orden Orden, " _
                    + DSQ + ".dbo.Proveedor Proveedor, " _
                    + DSQ + ".dbo.Articulo Articulo " _
                    + "WHERE " _
                    + "Orden.Proveedor = Proveedor.Proveedor AND " _
                    + "Orden.Articulo = Articulo.Codigo AND " _
                    + "((Orden.Articulo >= 'DW-000-000' AND " _
                    + "Orden.Articulo <= 'DY-999-999') OR " _
                    + "(Orden.Articulo >= 'DQ-000-000' AND " _
                    + "Orden.Articulo <= 'DQ-999-999') OR " _
                    + "(Orden.Articulo >= 'CO-000-000' AND " _
                    + "Orden.Articulo <= 'CO-999-999')) AND " _
                    + "Orden.Saldo > 0"
                    
                        Else
    
        Listado.SQLQuery = "SELECT Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Fecha2, Orden.Saldo, Orden.OrdFecha2, Orden.Carpeta, Orden.PedidoImpo, Orden.FechaImpo, Orden.TipoImpo, " _
                    + "Proveedor.Nombre, " _
                    + "Articulo.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Orden Orden, " _
                    + DSQ + ".dbo.Proveedor Proveedor, " _
                    + DSQ + ".dbo.Articulo Articulo " _
                    + "WHERE " _
                    + "Orden.Proveedor = Proveedor.Proveedor AND " _
                    + "Orden.Articulo = Articulo.Codigo AND " _
                    + "((Orden.Articulo >= 'DS-000-000' AND " _
                    + "Orden.Articulo <= 'DY-999-999') OR " _
                    + "(Orden.Articulo >= 'DQ-000-000' AND " _
                    + "Orden.Articulo <= 'DQ-999-999') OR " _
                    + "(Orden.Articulo >= 'CO-000-000' AND " _
                    + "Orden.Articulo <= 'CO-999-999')) AND " _
                    + "Orden.Saldo > 0"
                    
    End If
    
    
    Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    PrgOrdPenDy.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgOrdPenDy.Caption = "Listado de Ordenes Pendientes de Materias Primas :  " + !Nombre
        End If
    End With
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

