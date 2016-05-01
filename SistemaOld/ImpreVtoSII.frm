VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImpreVtoSII 
   AutoRedraw      =   -1  'True
   Caption         =   "Materias primas vencidas incluidas en hojas de Produccion"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   3135
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Imprime 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Materias primas vencidas a revalidar"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.CommandButton Acepta 
      Caption         =   "Aceptar"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   1200
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPedPen.rpt"
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
Attribute VB_Name = "PrgImpreVtoSiI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstVerificaVtoArti As Recordset
Dim spVerificaVtoArti As String

Dim XParam As String
Dim ZZLugar As Integer
Dim ZZVector(1000) As String
Dim WNumero As String


Private Sub Acepta_Click()

    WEmpresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    ZZLugar = 0
    ZZPasa = 0
    ZZCorte = ""

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM VerificaVtoArti"
    ZSql = ZSql + " Where VerificaVtoArti.Impre = " + "'" + "N" + "'"
    ZSql = ZSql + " Order by Numero"
    spVerificaVtoArti = ZSql
    Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
    If rstVerificaVtoArti.RecordCount > 0 Then
        With rstVerificaVtoArti
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    If ZZPasa <> 0 Then
                        ZZPasa = 1
                        ZZCorte = rstVerificaVtoArti!Numero
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar) = ZZCorte
                    End If
                    
                    If ZZCorte <> rstVerificaVtoArti!Numero Then
                        ZZCorte = rstVerificaVtoArti!Numero
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar) = ZZCorte
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
        rstVerificaVtoArti.Close
    
    End If
    
    If ZZLugar > 0 Then
        PrgImpreVtoSiI.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For a = 1 To 10
            Beep
        Next a
        PrgImpreVtoSiI.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()
    
    m$ = "Coloque  el papel para la impresion de los avisos de revalodia de materia prima"
    a% = MsgBox(m$, 0, "Impresion de Avisos de impresion de revalida de materia prima")

    For WWCicla = 1 To ZZLugar
    
        WNumero = ZZVector(WWCicla)
        
        Call Impresion
                
        ZSql = ""
        ZSql = ZSql + "UPDATE VerificaVtoArti SET "
        ZSql = ZSql + " Impre = " + "'" + "S" + "'"
        ZSql = ZSql + " Where Numero = " + "'" + WNumero + "'"
        spVerificaVtoArti = ZSql
        Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
    
    Next WWCicla
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImpreVtoSiI.Hide
    Unload Me
    End
End Sub

Private Sub Impresion()
    
    Listado.ReportFileName = "ImpreVto.rpt"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT VerificaVtoArti.Codigo, VerificaVtoArti.Articulo, VerificaVtoArti.EmpresaPartida, VerificaVtoArti.Partida, VerificaVtoArti.EmpresaTipo, VerificaVtoArti.Tipo, VerificaVtoArti.Numero, VerificaVtoArti.Fecha, VerificaVtoArti.Stock, VerificaVtoArti.StockI, VerificaVtoArti.StockII, VerificaVtoArti.StockIII, VerificaVtoArti.StockIV, VerificaVtoArti.StockV, VerificaVtoArti.StockVI, VerificaVtoArti.StockVII, VerificaVtoArti.Descripcion, VerificaVtoArti.TipoMov, VerificaVtoArti.Terminado " _
            + "From " _
            + DSQ + ".dbo.VerificaVtoArti VerificaVtoArti " _
            + "Where " _
            + "VerificaVtoArti.Numero >= " + WNumero + " AND " _
            + "VerificaVtoArti.Numero <= " + WNumero
                            
                            
    Listado.GroupSelectionFormula = "{VerificaVtoArti.Numero} in " + WNumero + " to " + WNumero
    Listado.SelectionFormula = "{VerificaVtoArti.Numero} in " + WNumero + " to " + WNumero

    Listado.Destination = 1
    Rem Listado.Destination = 0
                            
    Listado.Connect = Connect()
    Listado.Action = 1

End Sub


Private Sub Form_Load()
    Call Acepta_Click
End Sub


