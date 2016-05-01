VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListIbTucu 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Percepciones de Ingresos Brutos"
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
      ReportFileName  =   "wListIbVen.rpt"
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
      ItemData        =   "ListIbTucu.frx":0000
      Left            =   840
      List            =   "ListIbTucu.frx":0007
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
Attribute VB_Name = "PrgListIbTucu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstRecibo As Recordset
Dim spRecibo As String
Dim XParam As String
Dim Vector(10000, 4) As String
Dim WClave As String
Dim WFecha As String
Dim WTipo As String
Dim WNumero As String
Dim WImpoIb As Double

Private Sub Acepta_Click()

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    spCtacte = "ModificaCtacteImporteIva0"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem Procesa las cobranzas
    
    Renglon = 0
    Erase Vector
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spRecibo = "ListaRecibosDifeI" + XParam
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibo.RecordCount > 0 Then
        With rstRecibo
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstRecibo!Clave
                    Vector(Renglon, 2) = rstRecibo!Fecha
                    Vector(Renglon, 3) = rstRecibo!Tipo1
                    Vector(Renglon, 4) = rstRecibo!Numero1
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibo.Close
    End If
     
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        WFecha = Vector(Cicla, 2)
        WTipo = Vector(Cicla, 3)
        WNumero = Vector(Cicla, 4)
        
        ClaveCtacte = WTipo + WNumero + "01"
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            WImpoIb = IIf(IsNull(rstCtacte!ImpoIb), "0", rstCtacte!ImpoIb)
            If WImpoIb = 0 Then
                Vector(Cicla, 1) = ""
                Vector(Cicla, 2) = ""
                Vector(Cicla, 3) = ""
                Vector(Cicla, 4) = ""
            End If
            rstCtacte.Close
                Else
            Vector(Cicla, 1) = ""
            Vector(Cicla, 2) = ""
            Vector(Cicla, 3) = ""
            Vector(Cicla, 4) = ""
        End If
        
    Next Cicla
    
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        If WClave <> "" Then
        
            WTipo = Vector(Cicla, 3)
            WNumero = Vector(Cicla, 4)
            WRecibo = Val(Left$(WClave, 6))
            WSale = "N"
        
            XParam = "'" + WTipo + "','" _
                         + WNumero + "'"
            spRecibo = "ListaRecibosFactura " + XParam
            Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibo.RecordCount > 0 Then
                With rstRecibo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Val(rstRecibo!Recibo) < WRecibo Then
                                WSale = "S"
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibo.Close
            End If
            
            If WSale = "S" Then
                Vector(Cicla, 1) = ""
                Vector(Cicla, 2) = ""
                Vector(Cicla, 3) = ""
                Vector(Cicla, 4) = ""
            End If
            
        End If
        
    Next Cicla
    
    
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        If WClave <> "" Then
        
            WTipo = Vector(Cicla, 3)
            WNumero = Vector(Cicla, 4)
            
            ClaveCtacte = WTipo + WNumero + "01"
            XParam = "'" + ClaveCtacte + "','" _
                         + WClave + "'"
            spCtacte = "ModificaCtacteIb " + XParam
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Next Cicla
        
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
    
    Listado.WindowTitle = "Listado de Percepcion de Ingresos Brutos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CtaCte.Tipo, CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.Importe4, CtaCte.Importe8, " _
                    + "Cliente.Razon, Cliente.Cuit, " _
                    + "Recibos.Recibo " _
                    + "From " _
                    + DSQ + ".dbo.CtaCte CtaCte, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Recibos Recibos " _
                    + "Where " _
                    + "CtaCte.Cliente = Cliente.Cliente AND " _
                    + "CtaCte.ClaveRecibo = Recibos.Clave AND " _
                    + "CtaCte.Tipo >= '01' AND " _
                    + "CtaCte.Tipo <= '05' AND " _
                    + "CtaCte.OrdFecha >= '00000000' AND " _
                    + "CtaCte.OrdFecha <= '99999999' AND " _
                    + "CtaCte.Importe8 <> 0."
    
    Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
End Sub

Private Sub Cancela_click()
    Desde.SetFocus
    PrgListIbVen.Hide
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
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub


