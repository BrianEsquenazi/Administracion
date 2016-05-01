VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgArtiComple 
   ClientHeight    =   4515
   ClientLeft      =   3690
   ClientTop       =   2445
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   5595
   Begin Crystal.CrystalReport Listado 
      Left            =   4680
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.ComboBox UltimoTipo 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Calculo 
      Caption         =   "Ver Calculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox UltimoCosto 
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
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      Text            =   " "
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Factor 
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
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   7
      Text            =   " "
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Carpeta 
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
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      Text            =   " "
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox CostoStd 
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
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      Text            =   " "
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox UltimoFob 
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
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   0
      Text            =   " "
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Tipo Importacion"
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
      Left            =   600
      TabIndex        =   12
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Costo U$S"
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
      Left            =   600
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Factor"
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
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Carpeta"
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
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Costo Standard"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Costo Importacion"
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
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "PrgArtiComple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String
Dim Vector(100, 20) As String
Dim Gastos(100, 10) As String

Dim rstCarpeta As Recordset
Dim spCarpeta As String
Dim rstMovgas As Recordset
Dim spMovgas As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstCambios As Recordset
Dim spCambios As String

Dim WArancel As String
Dim WCarpeta As String
Dim WCosto As Double
Dim WLeyenda As Integer
Dim CargaEmpresa(12, 2) As String

Private Sub Calculo_Click()

    WCarpeta = Carpeta.Text
    Call Ceros(WCarpeta, 6)

    spCarpeta = "BorrarCarpeta"
    Set rstCarpeta = db.OpenRecordset(spCarpeta, dbOpenSnapshot, dbSQLPassThrough)

    Renglon = 0
    Erase Gastos
    ImpoGastos = 0
    ImpoSeguro = 0
    ImpoFlete = 0
    
    spMovgas = "ListaMovgas " + "'" + Carpeta.Text + "'"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovgas.RecordCount > 0 Then
    
        With rstMovgas
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstMovgas!concepto <> 10 Then
                        WArancel = Str$(rstMovgas!Derechos)
                        WOrden = Str$(rstMovgas!Orden)
                        WEmpresaOtro = rstMovgas!Empresa
                        Select Case rstMovgas!concepto
                            Case 2
                                ImpoSeguro = ImpoSeguro + rstMovgas!Importe
                            Case 4, 5
                                ImpoFlete = ImpoFlete + rstMovgas!Importe
                            Case Else
                                Renglon = Renglon + 1
                                Gastos(Renglon, 1) = Str$(rstMovgas!concepto)
                                Gastos(Renglon, 2) = Str$(rstMovgas!Importe)
                                ImpoGastos = ImpoGastos + rstMovgas!Importe
                        End Select
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovgas.Close
                
    End If
    
    WGastos = Renglon
    ImpoArancel = 0
    WTotal = 0
    
    Renglon = 0
    Erase Vector
    
    EmpresaAnterior = WEmpresa
    Select Case WEmpresaOtro
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
    
    ZCoeParidad = 1
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Informe"
    ZSql = ZSql + " Where Orden = " + "'" + WOrden + "'"
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        ZNroInforme = rstInforme!Informe
        ZFechaInforme = rstInforme!Fecha
        rstInforme.Close
        
        ZZZEmpresa = WEmpresa
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        spCambios = "ConsultaCambio  " + "'" + ZFechaInforme + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
            ZParidad = rstCambios!Cambio
            ZParidadII = IIf(IsNull(rstCambios!CambioII), "0", rstCambios!CambioII)
            If ZParidadII <> 0 And ZParidad <> 0 Then
                ZCoeParidad = ZParidadII / ZParidad
            End If
            rstCambios.Close
        End If
        
        Select Case ZZZEmpresa
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
        
    End If
    
    WTotalImpo = 0
    WTotalPeso = 0
    
    spOrden = "ListaOrden " + "'" + WOrden + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)

    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                    
                    WMoneda = rstOrden!Moneda
                    ZZPrecio = rstOrden!Precio
                    If WMoneda = 2 Then
                        ZZPrecio = ZZPrecio * ZCoeParidad
                    End If
                    
                    Vector(Renglon, 1) = Carpeta.Text
                    Vector(Renglon, 2) = rstOrden!Articulo
                    Vector(Renglon, 3) = Str$(rstOrden!Cantidad)
                    Vector(Renglon, 4) = Str$(ZZPrecio)
                    Vector(Renglon, 5) = Str$(rstOrden!Cantidad * ZZPrecio)
                  
                    If IsNull(rstOrden!Derechos) Then
                        Vector(Renglon, 6) = ""
                             Else
                        Vector(Renglon, 6) = Str$(rstOrden!Derechos)
                    End If
                    
                    Vector(Renglon, 7) = ""
                    Vector(Renglon, 8) = ""
                    Vector(Renglon, 9) = ""
                    Vector(Renglon, 10) = ""
                    Vector(Renglon, 11) = WCarpeta + "01"
                    
                    WTotalImpo = WTotalImpo + (rstOrden!Cantidad * ZZPrecio)
                    WTotalPeso = WTotalPeso + rstOrden!Cantidad
                    
                    WLeyenda = IIf(IsNull(rstOrden!Leyenda), "0", rstOrden!Leyenda)
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    For Ciclo = 1 To Renglon
        WSeguro = 0
        WFlete = 0
        If ImpoSeguro <> 0 Then
            If WTotalImpo <> 0 Then
                WSeguro = ((Val(Vector(Ciclo, 5)) / WTotalImpo) * ImpoSeguro) / Val(Vector(Ciclo, 3))
            End If
        End If
        If ImpoFlete <> 0 Then
            If WTotalPeso <> 0 Then
                WFlete = ((Val(Vector(Ciclo, 3)) / WTotalPeso) * ImpoFlete) / Val(Vector(Ciclo, 3))
            End If
        End If
        WCosto = Val(Vector(Ciclo, 4)) + WSeguro + WFlete
        Call Redondeo(WCosto)
        Vector(Ciclo, 4) = Str$(WCosto)
    Next Ciclo
    
    WTotal = 0
    
    For Ciclo = 1 To Renglon
        WArticulo = Vector(Ciclo, 2)
        
        spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Rem Vector(Ciclo, 4) = Str$(rstArticulo!Flete)
            Vector(Ciclo, 5) = Str$(Val(Vector(Ciclo, 3) * Val(Vector(Ciclo, 4))))
            XArancel = Val(Vector(Ciclo, 5)) * Val(Vector(Ciclo, 6)) / 100
            Vector(Ciclo, 7) = Str$(Val(Vector(Ciclo, 5)) + XArancel)
            WTotal = WTotal + Val(Vector(Ciclo, 5))
            ImpoArancel = ImpoArancel + XArancel
            rstArticulo.Close
        End If
    Next Ciclo
    
    For Ciclo = 1 To Renglon
        If WTotal <> 0 Then
            Vector(Ciclo, 8) = Str$((Val(Vector(Ciclo, 5)) / WTotal) * ImpoGastos)
        End If
        If Val(Vector(Ciclo, 3)) <> 0 Then
            Vector(Ciclo, 9) = Str$((Val(Vector(Ciclo, 7)) + Val(Vector(Ciclo, 8))) / Val(Vector(Ciclo, 3)))
        End If
    Next Ciclo
    
    If WTotal <> 0 Then
        WCoeficiente = Str$((ImpoGastos + ImpoArancel) / (WTotal / 100))
        WCoeficiente = Str$(1 + (Val(WCoeficiente) / 100))
            Else
        WCoeficiente = ""
    End If
    
    Select Case Val(EmpresaAnterior)
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
    
    XLeyenda = Str$(WLeyenda)
    
    For Ciclo = 1 To Renglon
        XParam = "'" + Vector(Ciclo, 1) + "','" _
                    + Vector(Ciclo, 2) + "','" _
                    + Vector(Ciclo, 3) + "','" _
                    + Vector(Ciclo, 4) + "','" _
                    + Vector(Ciclo, 5) + "','" _
                    + Vector(Ciclo, 6) + "','" _
                    + Vector(Ciclo, 7) + "','" _
                    + Vector(Ciclo, 8) + "','" _
                    + Vector(Ciclo, 9) + "','" _
                    + WCoeficiente + "','" _
                    + Vector(Ciclo, 11) + "','" _
                    + XLeyenda + "'"
        
        spCarpeta = "AltaCarpeta " + XParam
        Set rstCarpeta = db.OpenRecordset(spCarpeta, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo
    
    Listado.WindowTitle = "Listado de Calculo de Costo de Importacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Carpeta.Carpeta} in " + Carpeta.Text + " to " + Carpeta.Text
   
    Listado.Destination = 1
    Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Carpeta.Carpeta, Carpeta.Articulo, Carpeta.Cantidad, Carpeta.CostoFlete, Carpeta.Importe, Carpeta.Arancel, Carpeta.Costo, Carpeta.Gastos, Carpeta.Precio, Carpeta.Coeficiente, " _
                    + "Articulo.Descripcion, " _
                    + "Movgas.Fecha, Movgas.Orden, Movgas.Proveedor, Movgas.Origen, Movgas.Moneda, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Carpeta Carpeta, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Movgas Movgas, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Carpeta.Articulo = Articulo.Codigo AND " _
                    + "Carpeta.Clave = Movgas.Clave AND " _
                    + "Movgas.Proveedor = Proveedor.Proveedor AND " _
                    + "Carpeta.Carpeta >= 0 AND Carpeta.Carpeta <= 999999"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.ReportFileName = "WCarpeta.rpt"

    Listado.Action = 1
    

End Sub

Private Sub Cerrar_Click()
    PrgArtiComple.Hide
    Unload Me
    PrgArti.Show
End Sub

Private Sub Form_Load()

    UltimoTipo.Clear
    
    UltimoTipo.AddItem "FOB"
    UltimoTipo.AddItem "CIF"
    UltimoTipo.AddItem "CFR"
    UltimoTipo.AddItem "CPT"
    UltimoTipo.AddItem "EXW"
    UltimoTipo.AddItem "FCA"
    
    UltimoTipo.ListIndex = WPasaUltimoTipo
    UltimoFob.Text = Str$(WPasaUltimoFob)
    Factor.Text = Str$(WPasaFactor)
    UltimoCosto.Text = Str$(WPasaUltimoCosto)
    CostoStd.Text = Str$(WPasaUltimoCosto * 1.03)
    Carpeta.Text = Str$(WPasaCarpeta)

    UltimoFob.Text = Pusing("###,###.##", UltimoFob.Text)
    Factor.Text = Pusing("###.#####", Factor.Text)
    UltimoCosto.Text = Pusing("###,###.##", UltimoCosto.Text)
    CostoStd.Text = Pusing("###,###.##", CostoStd.Text)

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

