VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDisponiblePt 
   Caption         =   "Listado de Stock Disponible de Producto Terminado"
   ClientHeight    =   3705
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3705
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   5295
      Begin VB.ComboBox Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   1440
         Width           =   2415
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1920
         TabIndex        =   8
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   2760
         TabIndex        =   7
         Top             =   2040
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
         Left            =   1320
         TabIndex        =   6
         Top             =   2040
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
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   1095
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
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Terminado"
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
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Terminado"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WDisponiblept.rpt"
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
Attribute VB_Name = "PrgDisponiblePt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstMinimo As Recordset
Dim spMinimo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String
Dim Empe(12, 10) As String
Dim Vector(50000, 7) As String
Dim WCodigo As String
Dim Articulo As String
Dim Terminado As String
Dim Descripcion As String
Dim Stock1 As String
Dim Stock2 As String
Dim Stock3 As String
Dim Stock4 As String
Dim Stock5 As String
Dim Stock As String
Dim Minimo As String
Dim WVector(1000, 5) As String
Dim LugarVector As Integer
Dim HojaPend(10000, 5) As String
Dim LugarHojaPend As Integer

Private Sub Acepta_Click()

    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        XHasta = 1
            Else
        Empe(1, 1) = "0008"
        Empe(1, 2) = "Empresa08"
        XHasta = 1
    End If
    
    For a = 1 To XHasta
    
        WEmpresa = Empe(a, 1)
        txtOdbc = Empe(a, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

        spPedido = "ModificaPedpen0"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
        spArticulo = "ModificaArticuloVenta0"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
        spTerminado = "ModificaTerminadoPedido0"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        Sql1 = "UPDATE Pedido SET "
        Sql2 = " Importe = Cantidad - Facturado"
        spPedido = Sql1 + Sql2
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

        Erase WVector
        LugarVector = 0
    
        spPedido = "ListaPedidoPend"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                        EntraVector = "S"
                        For Ciclo = 1 To LugarVector
                            If WVector(Ciclo, 1) = rstPedido!Terminado Then
                                WVector(Ciclo, 2) = Str$(Val(WVector(Ciclo, 2)) + rstPedido!Importe)
                                EntraVector = "N"
                                Exit For
                            End If
                        Next Ciclo
                        If EntraVector = "S" Then
                            LugarVector = LugarVector + 1
                            WVector(LugarVector, 1) = rstPedido!Terminado
                            WVector(Ciclo, 2) = Str$(rstPedido!Importe)
                        End If
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
        End If
    
        For Ciclo = 1 To LugarVector
            WProducto = WVector(Ciclo, 1)
            WTipopro = Left$(WProducto, 2)
            WImporte = WVector(Ciclo, 2)
            Select Case WTipopro
                Case "DY", "DS", "DQ"
                    WArticulo = Left$(WProducto, 3) + Right$(WProducto, 7)
                    XParam = "'" + WArticulo + "','" _
                             + WImporte + "','" _
                             + WDate + "'"
                    spArticulo = "ModificaArticuloVenta " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Case Else
                    WTerminado = WProducto
                    WDate = Date$
                    XParam = "'" + WTerminado + "','" _
                                + WImporte + "','" _
                                + WDate + "'"
                    spTerminado = "ModificaTerminadoPedido " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End Select
        Next Ciclo
    
    Next a

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

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    spMinimo = "BorrarMinimo "
    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)

    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        XHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        XHasta = 4
    End If
    
    Erase Vector
    Suma = 0
    
    For a = 1 To XHasta
    
        WEmpresa = Empe(a, 1)
        txtOdbc = Empe(a, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        Sql1 = "UPDATE Terminado SET "
        Sql2 = " Proceso = 0"
        spTerminado = Sql1 + Sql2
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        Erase HojaPend
        LugarHojaPend = 0
        
        Sql1 = "Select *"
        Sql2 = " FROM Hoja"
        Sql3 = " Where Hoja.Renglon = 1"
        Sql4 = " and Hoja.Teorico <> 0"
        Sql5 = " and Hoja.Real = 0"
        Sql6 = " and Hoja.Marca <> " + "'" + "X" + "'"
        spHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveFirst
                Do
                    If .EOF = False Then
                        LugarHojaPend = LugarHojaPend + 1
                        HojaPend(LugarHojaPend, 1) = rstHoja!Producto
                        HojaPend(LugarHojaPend, 2) = Str$(rstHoja!Teorico)
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHoja.Close
        End If
        
        For Ciclo = 1 To LugarHojaPend
            Sql1 = "UPDATE Terminado SET "
            Sql2 = " Proceso = Proceso + " + HojaPend(Ciclo, 2)
            Sql3 = " Where Codigo = " + "'" + HojaPend(Ciclo, 1) + "'"
            spHoja = Sql1 + Sql2 + Sql3
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        Next Ciclo
        
        Sql1 = "Select Codigo, Descripcion, Inicial, Salidas, Entradas, Minimo, Minimo1, Pedido, Proceso"
        Sql2 = " FROM Terminado"
        Sql3 = " Where Terminado.Codigo >= " + "'" + Desde.Text + "'"
        Sql4 = " and Terminado.Codigo <= " + "'" + Hasta.Text + "'"
        spTerminado = Sql1 + Sql2 + Sql3 + Sql4
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
    
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
    
                        If Desde.Text <= UCase(rstTerminado!Codigo) And Hasta.Text >= UCase(rstTerminado!Codigo) Then
                        
                            XSaldo = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                            
                            If XSaldo <> 0 Or rstTerminado!Pedido <> 0 Or rstTerminado!Proceso <> 0 Then
                        
                                Suma = Suma + 1
                            
                                Vector(Suma, 1) = UCase(rstTerminado!Codigo)
                                Vector(Suma, 2) = rstTerminado!Descripcion
                                Vector(Suma, 3) = Str$(rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas)
                                Vector(Suma, 4) = Str$(rstTerminado!Pedido)
                                Vector(Suma, 5) = Str$(rstTerminado!Proceso)
                                WMinimo = IIf(IsNull(rstTerminado!Minimo), "0", rstTerminado!Minimo)
                                Vector(Suma, 6) = Str$(WMinimo)
                                Vector(Suma, 7) = Str$(a)
                                
                            End If
                            
                        End If
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            rstTerminado.Close
        
        End If
        
    Next a
    
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
        
    For Ciclo = 1 To Suma
        
        Terminado = Vector(Ciclo, 1)
        Articulo = ""
        WCodigo = Terminado
        Descri1 = Left$(Vector(Ciclo, 2), 50)
        Stock = Vector(Ciclo, 3)
        Pedido = Vector(Ciclo, 4)
        Proceso = Vector(Ciclo, 5)
        Minimo = Vector(Ciclo, 6)
        a = Vector(Ciclo, 7)
        Descripcion = ""
            
        For Saca1 = 1 To 50
            cara = Mid$(Descri1, Saca1, 1)
            Ingre = "S"
            If Mid$(Descri1, Saca1, 1) <> "" Then
                If Asc(Mid$(Descri1, Saca1, 1)) = 39 Then
                    Ingre = "N"
                End If
            End If
            If Ingre = "S" Then
                Descripcion = Descripcion + cara
            End If
        Next Saca1
        
        spMinimo = "ConsultaMinimo " + "'" + Terminado + "'"
        Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
        If rstMinimo.RecordCount > 0 Then
            
            Stock1 = Str$(rstMinimo!Stock1)
            Stock2 = Str$(rstMinimo!Stock2)
            Stock3 = Str$(rstMinimo!Stock3)
            Stock4 = Str$(rstMinimo!Stock4)
            Stock5 = Str$(rstMinimo!Stock5)
            SumaPedido = Str$(rstMinimo!Minimo)
            SumaProceso = Str$(rstMinimo!Minimo1)
            SumaMinimo = Str$(rstMinimo!Minimo2)
            
            Select Case a
                Case 1
                    Stock1 = Str$(Val(Stock1) + Val(Stock))
                Case 2
                    Stock2 = Str$(Val(Stock2) + Val(Stock))
                Case 3
                    Stock3 = Str$(Val(Stock3) + Val(Stock))
                Case 4, 6, 7
                    Stock4 = Str$(Val(Stock4) + Val(Stock))
                Case 5
                    Stock5 = Str$(Val(Stock5) + Val(Stock))
                Case Else
            End Select
                    
            SumaPedido = Str$(Val(SumaPedido) + Val(Pedido))
            SumaProceso = Str$(Val(SumaProceso) + Val(Proceso))
            SumaMinimo = Str$(Val(SumaMinimo) + Val(Minimo))
                    
            rstMinimo.Close
                
            Sql1 = "UPDATE Minimo SET "
            Sql2 = " Stock1 = " + Stock1 + ","
            Sql3 = " Stock2 = " + Stock2 + ","
            Sql4 = " Stock3 = " + Stock3 + ","
            Sql5 = " Stock4 = " + Stock4 + ","
            Sql6 = " Stock5 = " + Stock5 + ","
            Sql7 = " Minimo = " + SumaPedido + ","
            Sql8 = " Minimo1 = " + SumaProceso + ","
            Sql9 = " Minimo2 = " + SumaMinimo
            Sql10 = " Where Codigo = " + "'" + WCodigo + "'"
            spMinimo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10
            Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
                    
                Else
                
            Stock1 = "0"
            Stock2 = "0"
            Stock3 = "0"
            Stock4 = "0"
            Stock5 = "0"
            
            Select Case a
                Case 1
                    Stock1 = Stock
                Case 2
                    Stock2 = Stock
                Case 3
                    Stock3 = Stock
                Case 4, 6, 7
                    Stock4 = Stock
                Case 5
                    Stock5 = Stock
                Case Else
            End Select
            
            SumaPedido = Pedido
            SumaProceso = Proceso
            SumaMinimo = Minimo
            
            Sql1 = "INSERT INTO Minimo ("
            Sql2 = "Codigo ,"
            Sql3 = "Articulo ,"
            Sql4 = "Terminado ,"
            Sql5 = "Descripcion ,"
            Sql6 = "Stock1 ,"
            Sql7 = "Stock2 ,"
            Sql8 = "Stock3 ,"
            Sql9 = "Stock4 ,"
            Sql10 = "Stock5 ,"
            Sql11 = "Minimo ,"
            Sql12 = "Minimo1,"
            Sql13 = "Minimo2 )"
            Sql14 = "Values ("
            Sql15 = "'" + WCodigo + "',"
            Sql16 = "'" + Articulo + "',"
            Sql17 = "'" + Terminado + "',"
            Sql18 = "'" + Descri1 + "',"
            Sql19 = "'" + Stock1 + "',"
            Sql20 = "'" + Stock2 + "',"
            Sql21 = "'" + Stock3 + "',"
            Sql22 = "'" + Stock4 + "',"
            Sql23 = "'" + Stock5 + "',"
            Sql24 = "'" + SumaPedido + "',"
            Sql25 = "'" + SumaProceso + "',"
            Sql26 = "'" + SumaMinimo + "')"
        
            spMinimo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                       Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                       Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26
            Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    Next Ciclo

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Minimo.Terminado, Minimo.Descripcion, Minimo.Stock1, Minimo.Stock2, Minimo.Stock3, Minimo.Stock4, Minimo.Stock5, Minimo.Minimo " _
                       + "From " _
                       + DSQ + ".dbo.Minimo Minimo"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    If Tipo.ListIndex = 0 Then
        Listado.ReportFileName = "WDisponiblePt.Rpt"
            Else
        Listado.ReportFileName = "WDisponiblePtResu.Rpt"
    End If
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    Desde.SetFocus
    PrgDisponiblePt.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgDisponiblePt.Caption = "Listado de Stock Disponible de Producto Terminado :  " + !Nombre
        End If
    End With
    
    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Punto Critico"
    
    Tipo.ListIndex = 0
    
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub


