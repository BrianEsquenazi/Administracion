VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgIvaven 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Iva Ventas"
   ClientHeight    =   3825
   ClientLeft      =   3315
   ClientTop       =   2175
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   3825
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   3735
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   282
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wIvaven.rpt"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "ivaven.frx":0000
      Left            =   840
      List            =   "ivaven.frx":0007
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgIvaven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZClientes(5000) As String
Dim WImpoIbTucu As Double



Private Sub Acepta_Click()

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    
    If dada = 9999 Then
        
        Erase ZZClientes
        ZZLugar = 0
        
        
        spCliente = "ListaClienteConsulta"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZLugar = ZZLugar + 1
                    ZZClientes(ZZLugar) = rstCliente!Cliente
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCliente.Close
        
        For Ciclo = 1 To ZZLugar
        
            ZZClie = ZZClientes(Ciclo)
            
            spCliente = "ConsultaCliente " + "'" + ZZClie + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
                rstCliente.Close
            End If
            If WPorceCm05Tucu = 0 Then
                WPorceCm05Tucu = 1
            End If
            Select Case WCodIbTucu
                Case 1, 2, 3
                    WImpoIbTucu = 0.0175 * WPorceCm05Tucu
                    Call Redondeo(WImpoIbTucu)
                    WImpoPorceIbTucu = 1.75
                Case 4
                    WImpoIbTucu = WNeto * 0.035
                    Call Redondeo(WImpoIbTucu)
                    WImpoPorceIbTucu = 3.5
                Case 5
                    WImpoIbTucu = WNeto * 0.025
                    Call Redondeo(WImpoIbTucu)
                    WImpoPorceIbTucu = 2.5
                Case Else
                    WImpoIbTucu = 0
            End Select
            
        Next Ciclo
        
    End If
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Ctacte SET "
    ZSql = ZSql + " ImpoIbTucu = 0"
    ZSql = ZSql + " Where ImpoIbTucu IS NULL"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Ctacte SET "
    ZSql = ZSql + " ImpoIbCiudad = 0"
    ZSql = ZSql + " Where ImpoIbCiudad IS NULL"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    spCtacte = "ModificaCtacteImporteIva0"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)

    XParam = "'" + WDesde + "','" _
            + WHasta + "'"
    spCtacte = "ModificaCtacteIva1 " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    XParam = "'" + WDesde + "','" _
            + WHasta + "'"
    
    spCtacte = "ModificaCtacteIva2 " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    WTitulo = "del " + Desde.Text + " al " + Hasta.Text
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Iva Ventas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem With rstCtacte
    Rem         .Index = "Clave"
    Rem         .MoveFirst
    Rem         Do
    Rem             .Edit
    Rem             !Importe4 = 0
    Rem             !Importe5 = 0
    Rem             !Importe6 = 0
    Rem             !Importe7 = 0
    Rem             If !OrdFecha >= WDesde And !OrdFecha <= WHasta Then
    Rem                 If !Iva1 > 0 Then
    Rem                     !Importe4 = !Neto
    Rem                     !Importe5 = !Iva1
    Rem                     !Importe6 = !Iva2
    Rem                     !Importe7 = 0
    Rem                         Else
    Rem                     !Importe4 = 0
    Rem                     !Importe5 = 0
    Rem                     !Importe6 = 0
    Rem                     !Importe7 = !Neto
    Rem                 End If
    Rem             End If
    Rem             .Update
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With
    
    Listado.GroupSelectionFormula = "{CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Listado.SelectionFormula = "{CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
     Listado.SQLQuery = "SELECT CtaCte.Tipo, CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.Paridad, CtaCte.Importe4, CtaCte.Importe5, CtaCte.Importe6, CtaCte.Importe7, CtaCte.Importe8, CtaCte.ImpoIbTucu, CtaCte.ImpoIbCiudad, CtaCte.ImpoIb, " _
                + "Cliente.Razon, Cliente.Provincia, Cliente.Cuit, Cliente.IbTucu " _
                + "From " _
                + DSQ + ".dbo.CtaCte CtaCte, " _
                + DSQ + ".dbo.Cliente Cliente " _
                + "Where " _
                + "CtaCte.Cliente = Cliente.Cliente AND " _
                + "CtaCte.Tipo >= '01' AND " _
                + "CtaCte.Tipo <= '05' AND " _
                + "CtaCte.OrdFecha >= '" + WDesde + "' AND " _
                + "CtaCte.OrdFecha <= '" + WHasta + "'"
    
    Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    If Val(WEmpresa) = 1 Then
        Listado.ReportFileName = "WIvavensurfa.rpt"
            Else
        Listado.ReportFileName = "WIvaven.rpt"
    End If
    
    Listado.Action = 1
End Sub

Private Sub Cancela_click()
    Desde.SetFocus
    PrgIvaven.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub
Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub


