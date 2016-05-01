VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaCheckList 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Check List de Informes de Recepcion"
   ClientHeight    =   2400
   ClientLeft      =   1875
   ClientTop       =   2520
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   2400
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1800
         TabIndex        =   9
         Top             =   600
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1800
         TabIndex        =   0
         Top             =   240
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
         Left            =   2640
         TabIndex        =   8
         Top             =   1440
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
         Left            =   1080
         TabIndex        =   7
         Top             =   1440
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
         Left            =   3960
         TabIndex        =   6
         Top             =   360
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
         Left            =   3960
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6600
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlistinfpend.rpt"
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
      Left            =   6360
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaCheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim Vector(1000, 20) As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstInformeConsol As Recordset
Dim spInformeConsol As String
Dim Empe(10, 10) As String

Private Sub Acepta_Click()

    ZSql = "DELETE InformeConsol"
    spInformeConsol = ZSql
    Set rstInformeConsol = db.OpenRecordset(spInformeConsol, dbOpenSnapshot, dbSQLPassThrough)

    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Then
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
        XHasta = 5
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
    ZLugar = 0
    
    For A = 1 To XHasta
    
        WEmpresa = Empe(A, 1)
        txtOdbc = Empe(A, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Informe"
        ZSql = ZSql + " Where Informe.Expreso > 0 and Informe.Renglon = 1"
        ZSql = ZSql + " and Fechaord >= " + "'" + WDesde + "'"
        ZSql = ZSql + " and Fechaord <= " + "'" + WHasta + "'"
        spInforme = ZSql
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
    
            With rstInforme
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZLugar = ZLugar + 1
                        
                        Vector(ZLugar, 1) = rstInforme!Informe
                        Vector(ZLugar, 2) = rstInforme!Fecha
                        Vector(ZLugar, 3) = rstInforme!FechaOrd
                        Vector(ZLugar, 4) = rstInforme!Expreso
                        Vector(ZLugar, 5) = rstInforme!Chapa
                        Vector(ZLugar, 6) = rstInforme!Chofer
                        Vector(ZLugar, 7) = rstInforme!Item1
                        Vector(ZLugar, 8) = rstInforme!Item2
                        Vector(ZLugar, 9) = rstInforme!Item3
                        Vector(ZLugar, 10) = rstInforme!Item4
                        Vector(ZLugar, 11) = rstInforme!Item5
                        Vector(ZLugar, 12) = rstInforme!Item6
                        Vector(ZLugar, 13) = rstInforme!Item7
                        Vector(ZLugar, 14) = rstInforme!Item8
                        Vector(ZLugar, 15) = rstInforme!Placa
                        Vector(ZLugar, 16) = rstInforme!Rombo
                        Vector(ZLugar, 17) = IIf(IsNull(rstInforme!Observaciones), "", rstInforme!Observaciones)
                        Vector(ZLugar, 18) = IIf(IsNull(rstInforme!DesExpreso), "", rstInforme!DesExpreso)
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            rstInforme.Close
        
        End If
        
    Next A
    
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
        Case Else
    End Select
    
    For Ciclo = 1 To ZLugar
    
        ZZInforme = Vector(Ciclo, 1)
        ZZFecha = Vector(Ciclo, 2)
        ZZOrdFecha = Vector(Ciclo, 3)
        ZZExpreso = Vector(Ciclo, 4)
        ZZChapa = Vector(Ciclo, 5)
        ZZChofer = Vector(Ciclo, 6)
        ZZItem1 = Vector(Ciclo, 7)
        ZZItem2 = Vector(Ciclo, 8)
        ZZItem3 = Vector(Ciclo, 9)
        ZZItem4 = Vector(Ciclo, 10)
        ZZItem5 = Vector(Ciclo, 11)
        ZZItem6 = Vector(Ciclo, 12)
        ZZItem7 = Vector(Ciclo, 13)
        ZZItem8 = Vector(Ciclo, 14)
        ZZPlaca = Vector(Ciclo, 15)
        ZZRombo = Vector(Ciclo, 16)
        ZZObservaciones = Vector(Ciclo, 17)
        ZZDesExpreso = Vector(Ciclo, 18)
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO InformeConsol ("
        ZSql = ZSql + "Informe ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "ordFecha ,"
        ZSql = ZSql + "Expreso ,"
        ZSql = ZSql + "DesExpreso ,"
        ZSql = ZSql + "Chapa ,"
        ZSql = ZSql + "Chofer ,"
        ZSql = ZSql + "Item1 ,"
        ZSql = ZSql + "Item2 ,"
        ZSql = ZSql + "Item3 ,"
        ZSql = ZSql + "Item4 ,"
        ZSql = ZSql + "Item5 ,"
        ZSql = ZSql + "Item6 ,"
        ZSql = ZSql + "Item7 ,"
        ZSql = ZSql + "Item8 ,"
        ZSql = ZSql + "Placa ,"
        ZSql = ZSql + "Rombo ,"
        ZSql = ZSql + "Observaciones )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZInforme + "',"
        ZSql = ZSql + "'" + ZZFecha + "',"
        ZSql = ZSql + "'" + ZZFechaOrd + "',"
        ZSql = ZSql + "'" + ZZExpreso + "',"
        ZSql = ZSql + "'" + ZZDesExpreso + "',"
        ZSql = ZSql + "'" + ZZChapa + "',"
        ZSql = ZSql + "'" + ZZChofer + "',"
        ZSql = ZSql + "'" + ZZItem1 + "',"
        ZSql = ZSql + "'" + ZZItem2 + "',"
        ZSql = ZSql + "'" + ZZItem3 + "',"
        ZSql = ZSql + "'" + ZZItem4 + "',"
        ZSql = ZSql + "'" + ZZItem5 + "',"
        ZSql = ZSql + "'" + ZZItem6 + "',"
        ZSql = ZSql + "'" + ZZItem7 + "',"
        ZSql = ZSql + "'" + ZZItem8 + "',"
        ZSql = ZSql + "'" + ZZPlaca + "',"
        ZSql = ZSql + "'" + ZZRombo + "',"
        ZSql = ZSql + "'" + ZZObservaciones + "')"

        spInformeConsol = ZSql
        Set rstInformeConsol = db.OpenRecordset(spInformeConsol, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    Listado.WindowTitle = "Listado de Check-List de Informe de Recepcion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Informe.fechaord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.ReportFileName = "ListaCheckList.rpt"
    
    Listado.SQLQuery = "SELECT InformeConsol.Informe, InformeConsol.Fecha, InformeConsol.OrdFecha, InformeConsol.Expreso, InformeConsol.Chapa, InformeConsol.Chofer, InformeConsol.Item1, InformeConsol.Item2, InformeConsol.Item3, InformeConsol.Item4, InformeConsol.Item5, InformeConsol.Item6, InformeConsol.Item7, InformeConsol.Item8, InformeConsol.Placa, InformeConsol.Rombo, InformeConsol.Observaciones, InformeConsol.DesExpreso " _
            + "From " _
            + DSQ + ".dbo.InformeConsol InformeConsol " _
            + "Where " _
            + "InformeConsol.Informe >= 0 AND " _
            + "InformeConsol.Informe <= 999999"
        
    Listado.GroupSelectionFormula = "{InformeConsol.Informe} in 0 to 999999"
                        
    Listado.Connect = Connect()
    Listado.Action = 1
    
    If Val(XEmpresa) = 1 Then
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Else
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgListaCheckList.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    End Sub
