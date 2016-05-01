VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgAvisoDesvio 
   AutoRedraw      =   -1  'True
   Caption         =   "Aviso de Desvio de Produccion > 3%"
   ClientHeight    =   7365
   ClientLeft      =   225
   ClientTop       =   435
   ClientWidth     =   11655
   LinkTopic       =   "Form2"
   ScaleHeight     =   7365
   ScaleWidth      =   11655
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   7
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   4
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   5
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   6
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9128
      _Version        =   327680
      BackColor       =   16777088
   End
   Begin VB.CommandButton Confirma 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Aviso de Desvio de Produccion > al 3%"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   10935
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "PedpenII.rpt"
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
Attribute VB_Name = "PrgAvisoDesvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String
Dim WVector(10000) As String
Dim ZPlanta(100) As String
Dim LeeAviso(100, 3) As String
Dim CargaEmpresa(10, 2) As String
Dim CargaEmpresaII(10, 2) As String

Dim LugarVector As Integer

Private Sub Confirma_Click()

    For Cicla = 1 To 4
        If CargaEmpresa(Cicla, 1) <> "" Then
        
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + "ImpresionI =  " + "'" + "S" + "'"
            ZSql = ZSql + " Where (Hoja.PorceDife <= -3 OR Hoja.PorceDife >= 3)"
            ZSql = ZSql + " and Hoja.ImpresionI = 'N'"
            ZSql = ZSql + " and Hoja.Renglon = 1"
            
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    Next Cicla


    Call Cancela_click
End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Acepta_Click()

    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0003"
    CargaEmpresa(2, 2) = "Empresa03"
    CargaEmpresa(3, 1) = "0004"
    CargaEmpresa(3, 2) = "Empresa04"
    CargaEmpresa(4, 1) = "0005"
    CargaEmpresa(4, 2) = "Empresa05"
    
    ZPlanta(1) = "Surfactan Pta.I"
    ZPlanta(2) = "Surfactan Pta.II"
    ZPlanta(3) = "Pellital Pta.I"
    ZPlanta(4) = "Surfactan Pta.III"
                    
    Call Limpia_Vector
    LugarVector = 0
    
    For Cicla = 1 To 4
        If CargaEmpresa(Cicla, 1) <> "" Then
        
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Hoja.PorceDife = (Hoja.Real - Hoja.Teorico) / (Hoja.Teorico/100)"
            ZSql = ZSql + " Where Hoja.Teorico <> 0 and Hoja.Real <> 0"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            

            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where (Hoja.PorceDife <= -3 OR Hoja.PorceDife >= 3)"
            ZSql = ZSql + " and Hoja.ImpresionI = 'N'"
            ZSql = ZSql + " and Hoja.Renglon = 1"

            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
            
                With rstHoja
                    .MoveFirst
            
                        Do
            
                            If .EOF = True Then
                                Exit Do
                            End If
                
                            LugarVector = LugarVector + 1
                            WVector1.TextMatrix(LugarVector, 1) = ZPlanta(Cicla)
                            WVector1.TextMatrix(LugarVector, 2) = rstHoja!Hoja
                            WVector1.TextMatrix(LugarVector, 3) = rstHoja!producto
                            WVector1.TextMatrix(LugarVector, 4) = Pusing("###,###.##", Str$(rstHoja!Real))
                            WVector1.TextMatrix(LugarVector, 5) = Pusing("###,###.##", Str$(rstHoja!Teorico))
                            WVector1.TextMatrix(LugarVector, 6) = Pusing("###,###.##", Str$(rstHoja!POrceDife))
                            WVector1.TextMatrix(LugarVector, 7) = Str$(Cicla)
                
                            .MoveNext
                
                            If .EOF = True Then
                                Exit Do
                            End If
                
                        Loop
            
                End With
                rstHoja.Close
            End If
    
        End If
    Next Cicla
    
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    If LugarVector > 0 Then
    
        PrgAvisoDesvio.Refresh
        Aviso.Visible = True
        WVector1.Visible = True
        Confirma.Visible = True
        WTitulo(1).Visible = True
        WTitulo(2).Visible = True
        WTitulo(3).Visible = True
        WTitulo(4).Visible = True
        WTitulo(5).Visible = True
        WTitulo(6).Visible = True
        For A = 1 To 10
            Beep
        Next A
        PrgAvisoDesvio.Refresh
        
            Else
            
        Call Cancela_click
        
    End If
    
End Sub

Private Sub Cancela_click()
    PrgAvisoDesvio.Hide
    Unload Me
    Close
    End
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    WVector1.FixedCols = 1
    WVector1.Cols = 8
    WVector1.FixedRows = 1
    WVector1.Rows = 10001
    
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
                WVector1.Text = "Planta"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector1.Text = "Hoja"
                WVector1.ColWidth(Ciclo) = 1800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector1.Text = "Producto"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector1.Text = "Real"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector1.Text = "Teorico"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector1.Text = "% Dife"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
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

Private Sub WVector1_dblClick()
    T$ = "Aviso de Desvio"
    m$ = "Desea Bloqear el producto " + WVector1.TextMatrix(WVector1.Row, 3) + " para su produccion"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        ZEmpresa = Val(WVector1.TextMatrix(WVector1.Row, 7))
    
        Erase CargaEmpresaII
        Select Case ZEmpresa
            Case 1, 2, 4
                CargaEmpresaII(1, 1) = "0001"
                CargaEmpresaII(1, 2) = "Empresa01"
                CargaEmpresaII(2, 1) = "0003"
                CargaEmpresaII(2, 2) = "Empresa03"
                CargaEmpresaII(3, 1) = "0005"
                CargaEmpresaII(3, 2) = "Empresa05"
                CargaEmpresaII(4, 1) = "0006"
                CargaEmpresaII(4, 2) = "Empresa06"
                CargaEmpresaII(5, 1) = "0007"
                CargaEmpresaII(5, 2) = "Empresa07"
            Case Else
                CargaEmpresaII(1, 1) = "0002"
                CargaEmpresaII(1, 2) = "Empresa02"
                CargaEmpresaII(2, 1) = "0004"
                CargaEmpresaII(2, 2) = "Empresa04"
                CargaEmpresaII(3, 1) = "0008"
                CargaEmpresaII(3, 2) = "Empresa08"
                CargaEmpresaII(4, 1) = "0009"
                CargaEmpresaII(4, 2) = "Empresa09"
        End Select
                
        For CiclaEmpre = 1 To 5
            If CargaEmpresa(CiclaEmpre, 1) <> "" Then
            
                WEmpresa = CargaEmpresa(CiclaEmpre, 1)
                txtOdbc = CargaEmpresa(CiclaEmpre, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
                ZSql = ""
                ZSql = ZSql + "UPDATE Terminado SET "
                ZSql = ZSql + " Estado = " + "'" + "N" + "',"
                ZSql = ZSql + " EstadoI = " + "'" + "N" + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + WVector1.TextMatrix(WVector1.Row, 3) + "'"
                            
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
        Next CiclaEmpre
        
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        m$ = "El producto " + WVector1.TextMatrix(WVector1.Row, 3) + " ha sido bloqueado para la produccion en forma exitosa"
        G% = MsgBox(m$, 0, "Aviso de Desvio")
        
    End If
End Sub
