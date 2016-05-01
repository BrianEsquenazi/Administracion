VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgVerificaVentaDy 
   Caption         =   "Listado de Verirficacion de Materia Prima sin venta"
   ClientHeight    =   6075
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6075
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Porce 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   17
         Text            =   " "
         Top             =   2040
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2400
         TabIndex        =   12
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
         Mask            =   "AA-###-###"
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
         TabIndex        =   11
         Top             =   2760
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
         TabIndex        =   10
         Top             =   2760
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
         Left            =   4560
         TabIndex        =   9
         Top             =   600
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
         Left            =   4560
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   2400
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   300
         Left            =   2400
         TabIndex        =   14
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label5 
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Desde Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta MP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde MP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "verficaventa.rpt"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "verificaventady.frx":0000
      Left            =   1920
      List            =   "verificaventady.frx":0007
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgVerificaVentaDy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstMinimo As Recordset
Dim spMinimo As String

Dim XParam As String
Dim Empe(12, 10) As String
Dim Vector(50000, 4) As String
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
Dim MInimo As String

Dim ZZCodigo As String
Dim ZZVenta As Double
Dim ZZStock As Double
Dim ZZPorce As Double

Private Sub Acepta_Click()

    WAno = Right$(DesdeFecha.Text, 4)
    WMes = Mid$(DesdeFecha.Text, 4, 2)
    WDia = Left$(DesdeFecha.Text, 2)
    WDesdeFecha = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHastaFecha = WAno + WMes + WDia


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
    
    For a = 1 To XHasta
    
        Erase Vector
        Suma = 0
    
        WEmpresa = Empe(a, 1)
        txtOdbc = Empe(a, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo >= " + "'" + Desde.Text + "'"
        ZSql = ZSql + " and Articulo.Codigo <= " + "'" + Hasta.Text + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
    
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
    
                        ZZStock = rstArticulo!Entradas - rstArticulo!Salidas
                        Call Redondeo(ZZStock)
                        
                        If ZZStock > 0 Then
                            Suma = Suma + 1
                            Vector(Suma, 1) = UCase(rstArticulo!Codigo)
                            Vector(Suma, 2) = rstArticulo!Descripcion
                            Vector(Suma, 3) = Str$(ZZStock)
                        End If
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            rstArticulo.Close
        
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
        
        For Ciclo = 1 To Suma
        
            Terminado = Left$(Vector(Ciclo, 1), 3) + "00" + Right$(Vector(Ciclo, 1), 7)
            Articulo = ""
            WCodigo = Terminado
            Descri1 = Left$(Vector(Ciclo, 2), 50)
            Stock = Vector(Ciclo, 3)
            Descripcion = ""
            MInimo = ""
            
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
            
                    Select Case a
                        Case 1
                            Stock1 = Str$(Val(Stock1) + Val(Stock))
                        Case 2
                            Stock2 = Str$(Val(Stock2) + Val(Stock))
                        Case 3
                            Stock3 = Str$(Val(Stock3) + Val(Stock))
                        Case 4
                            Stock4 = Str$(Val(Stock4) + Val(Stock))
                        Case 5
                            Stock5 = Str$(Val(Stock5) + Val(Stock))
                        Case Else
                    End Select
                    
                    rstMinimo.Close
                
                    XParam = "'" + WCodigo + "','" _
                                + Stock1 + "','" _
                                + Stock2 + "','" _
                                + Stock3 + "','" _
                                + Stock4 + "','" _
                                + Stock5 + "'"
                                            
                    spMinimo = "ModificaMinimo " + XParam
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
                        Case 4
                            Stock4 = Stock
                        Case 5
                            Stock5 = Stock
                        Case Else
                    End Select
                
                    XParam = "'" + WCodigo + "','" _
                                + Articulo + "','" _
                                + Terminado + "','" _
                                + Descripcion + "','" _
                                + Stock1 + "','" _
                                + Stock2 + "','" _
                                + Stock3 + "','" _
                                + Stock4 + "','" _
                                + Stock5 + "','" _
                                + MInimo + "'"
                                            
                    spMinimo = "AltaMinimo " + XParam
                    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
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
    
    
    
    
    
    
    Erase Vector
    Suma = 0
                 
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Minimo"
    spMinimo = ZSql
    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
    If rstMinimo.RecordCount > 0 Then

        With rstMinimo
            .MoveFirst
            Do
                If .EOF = False Then

                    Suma = Suma + 1
                    Vector(Suma, 1) = UCase(rstMinimo!Codigo)
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstMinimo.Close
    
    End If
    
    For Ciclo = 1 To Suma
    
        ZZCodigo = Vector(Ciclo, 1)
        ZZVenta = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Estadistica"
        ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + ZZCodigo + "'"
        ZSql = ZSql + " and Estadistica.OrdFecha >= " + "'" + WDesdeFecha + "'"
        ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + WHastaFecha + "'"
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEstadistica.RecordCount > 0 Then
    
            With rstEstadistica
        
                .MoveFirst
                
                Do
                    
                    WCantidad = rstEstadistica!Cantidad
                    If !Tipo = 2 Then
                        WCantidad = Abs(WCantidad) * -1
                    End If
                    ZZVenta = ZZVenta + WCantidad
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End With
            
            rstEstadistica.Close
            
        End If
        
        spMinimo = "ConsultaMinimo " + "'" + ZZCodigo + "'"
        Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
        If rstMinimo.RecordCount > 0 Then
        
            Stock1 = Str$(rstMinimo!Stock1)
            Stock2 = Str$(rstMinimo!Stock2)
            Stock3 = Str$(rstMinimo!Stock3)
            Stock4 = Str$(rstMinimo!Stock4)
            Stock5 = Str$(rstMinimo!Stock5)
            
            ZZStock = Val(Stock1) + Val(Stock2) + Val(Stock3) + Val(Stock4) + Val(Stock5)
            ZZPorce = 0
            
            If ZZStock <= 0 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE Minimo"
                ZSql = ZSql + " Where Codigo = " + "'" + ZZCodigo + "'"
                rsminimo = ZSql
                Set rstMinimo = db.OpenRecordset(rsminimo, dbOpenSnapshot, dbSQLPassThrough)
                
                    Else
                    
                If ZZVenta <> 0 Then
                    ZZPorce = (ZZVenta / ZZStock) * 100
                    Call Redondeo(ZZPorce)
                End If
                
                If ZZPorce > Val(Porce.Text) Then
                
                    ZSql = ""
                    ZSql = ZSql + "DELETE Minimo"
                    ZSql = ZSql + " Where Codigo = " + "'" + ZZCodigo + "'"
                    rsminimo = ZSql
                    Set rstMinimo = db.OpenRecordset(rsminimo, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Minimo SET "
                    ZSql = ZSql + " Minimo = " + "'" + Str$(ZZPorce) + "',"
                    ZSql = ZSql + " Venta1 = " + "'" + Str$(ZZVenta) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + ZZCodigo + "'"
                    spMinimo = ZSql
                    Set rstMinimo = db.OpenRecordset(spMinimo, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                    
            End If
        End If
        
    Next Ciclo
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Minimo.Articulo, Minimo.Terminado, Minimo.Descripcion, Minimo.Stock1, Minimo.Stock2, Minimo.Stock3, Minimo.Stock4, Minimo.Stock5, Minimo.Minimo, Minimo.Dife, Minimo.Venta1  " _
                    + "From " _
                    + DSQ + ".dbo.Minimo Minimo " _
                    + "Where " _
                    + "Minimo.Terminado >= ' ' AND " _
                    + "Minimo.Terminado <= 'ZZZZZZZZZZ'"
                    
    Listado.ReportFileName = "VerificaVenta.rpt"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    PrgVerificaVentaDy.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        DesdeFecha.SetFocus
    End If
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Porce.SetFocus
    End If
End Sub

Private Sub Porce_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Porce.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub



