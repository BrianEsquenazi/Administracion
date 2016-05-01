VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImpreAvisoEnvases 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Impresion de Pedidos"
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
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "EMAIL DE ENVASES"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   720
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
Attribute VB_Name = "PrgImpreAvisoEnvases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstMuestra As Recordset
Dim spMuestra As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim XParam As String
Dim Vector(1000, 10) As String
Dim Datos(100, 10) As String
Dim WPedido As String
Dim WCliente As String
Dim WVia As String
Dim WPago As String
Dim WDirentrega As String
Dim WObservaciones As String
Dim WDespago As String
Dim WFecha As String
Dim WFecEntrega As String
Dim WVersion As String
Dim WTipoped As String
Dim Lugar As Integer
Dim WEnvase(10) As String
Dim XEnvase(40, 6) As String
Dim Auxiliar(100, 2) As String
Dim WImpre(10) As String
Dim WEspecif(100) As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String

Dim WWLote As String
Dim WWTipo As String

Dim AuxiliarII(100, 5) As String

Dim WWCliente As String
Dim WWFecha  As String
Dim WWFecEntrega As String
Dim WWVersion As String
Dim WWTipoped As String
Dim WWObservaciones As String

Dim ZZLugarDirEntrega As Integer
Dim ZZDirEntrega(10) As String

Dim WWEspecif(100) As String
Dim WWRazon As String
Dim WWPago As String
Dim WWDirentrega As String
Dim WWArticulo As String
Dim WWDescripcion As String
Dim WWCantidad As Double
Dim WWPrecio As Double
Dim WWObserva As String
Dim WWOrdenCpa As String
Dim WWDesPago As String
Dim WWVia As String

Dim ZZRequiereCertificado As String
Dim ZZRequiereMsds As String
Dim ZZRequiereMsdsCada As String
Dim ZZRequiereHoja As String
Dim ZZPermiteParcial As String
Dim ZZPartidasVarias As String

Dim ZZEmailCertificado As String
Dim ZZEmailMsds As String
Dim ZZEmailHoja As String
Dim ZZDiasI As String
Dim ZZDiasII As String
Dim ZZDiasIII As String
Dim ZZEnvasesI As String
Dim ZZEnvasesII As String
Dim ZZEnvasesIII As String
Dim ZZEtiquetaI As String
Dim ZZEtiquetaII As String
Dim ZZEspecif1 As String
Dim ZZEspecif2 As String
Dim ZZEspecif3 As String
Dim ZZEspecif4 As String
Dim ZZEspecif5 As String
Dim ZZCantidadPartidas As String
Dim ImpreEnvase(10) As String

Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

Dim WDireccionEmail As String
Dim EmailAddress As String
Dim CopiaAddress As String
Dim WNombreEmail As String
Dim MAttach As String


Private Sub Acepta_Click()

    Rem WEmpresa = "0009"
    Rem txtOdbc = "Empresa09"
    
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Erase Vector
    Lugar = 0
    
     ZSql = ""
     ZSql = ZSql + "Select *"
     ZSql = ZSql + " FROM Pedido"
     ZSql = ZSql + " Where Pedido.MarcaEnvase = " + "'" + "N" + "'"
     ZSql = ZSql + " and Pedido.CantidadEnvase > " + "'" + "0" + "'"
     ZSql = ZSql + " and Pedido.Cliente <> " + "'" + "T00140" + "'"
     spPedido = ZSql
     Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
     If rstPedido.RecordCount > 0 Then
    
         With rstPedido
     
             .MoveFirst
             
             Do
             
                Lugar = Lugar + 1
             
                Vector(Lugar, 1) = rstPedido!Pedido
                Vector(Lugar, 2) = Str$(rstPedido!CantidadEnvase)
                Vector(Lugar, 3) = rstPedido!Cliente
                Vector(Lugar, 4) = rstPedido!fecentrega
                 
                 .MoveNext
                 
                 If .EOF = True Then
                     Exit Do
                 End If
                 
             Loop
         End With
         
         rstPedido.Close
         
     End If
    
    If Lugar > 0 Then
        PrgImpreAvisoEnvases.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For a = 1 To 10
            Beep
        Next a
        PrgImpreAvisoEnvases.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()
    
    ZZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    For WWCicla = 1 To Lugar
    
        WPedido = Vector(WWCicla, 1)
        WCantidadEnvase = Vector(WWCicla, 2)
        WCliente = Vector(WWCicla, 3)
        WFecEntrega = Vector(WWCicla, 4)
        WEmail = ""
        WEmailEnv = ""
        
        spCliente = "ConsultaCliente " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WEmail = Trim(rstCliente!EMail)
            WEmailEnv = Trim(IIf(IsNull(rstCliente!emailenv), "", rstCliente!emailenv))
            WFechaEmailEnvase = Trim(IIf(IsNull(rstCliente!FechaEmailEnvase), "", rstCliente!FechaEmailEnvase))
            rstCliente.Close
        End If
        
        If Trim(WEmail) <> "" Or Trim(WEmailEnv) <> "" Then
        
            If ZZFecha <> WFechaEmailEnvase Then
                
                If Trim(WEmailEnv) <> "" Then
                    WEmail = WEmailEnv
                End If
                
                sTo = WEmail
                sCC = "nsoto@surfactan.com.ar"
                sBCC = ""
                sSubject = "Aviso de Retiro de Contenedores"
                sBody = "Se informa que el dia " + WFecEntrega + " se le estara entregando mercaderia procedente de la firma SURFACTAN S.A. y rogamos tengan a bien preparar la cantidad de " + WCantidadEnvase + " contenedores de nuestra propiedad procedente de entregas anteriores, para su rapido retiro evitando asi demoras innecesarias. Por favor se ruega entregar los contenedores con las respectivas tapas superiores e inferiores. Muchas Gracias."
                SFile = ""
        
                EmailAddress = sTo
                CopiaAddress = sCC
                MSubject = sSubject
                MBody = sBody
                MAttach = ""
                MAttachI = ""
                MAttachII = ""
                MAttachIII = ""
                MAttachIV = ""
                MAttachVI = ""
                MAttachVII = ""
                MAttachVIII = ""
                
                SendEmail
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Cliente SET "
                ZSql = ZSql + "FechaEmailEnvase = " + "'" + ZZFecha + "'"
                ZSql = ZSql + " Where Cliente = " + "'" + WCliente + "'"
                spCliente = ZSql
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                        
            End If
            
                Else
                
            m$ = "El cliente " + WCliente + " tiene contenedores para reclamar y no posee email:"
            ca% = MsgBox(m$, 0, "reclamo de contenedores")
            
        End If
            
        ZSql = ""
        ZSql = ZSql + "UPDATE Pedido SET "
        ZSql = ZSql + "MarcaEnvase = " + "'" + "S" + "'"
        ZSql = ZSql + " Where Pedido = " + "'" + WPedido + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Next WWCicla
    
    Close #1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImpreAvisoEnvases.Hide
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub


Public Sub SendEmail()

    Dim objOutlook As Object
    Dim objMailItem

    Dim NumOfPath As Integer, i As Integer
    Dim AtachPath As String

    On Error GoTo 10

    NumOfPath = 0
    AllPath = ""
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMailItem = objOutlook.CreateItem(olMailItem)
    
    With objMailItem
        .To = EmailAddress
        .cc = CopiaAddress
        .Subject = MSubject
        .Body = MBody
        Rem .Attachments.Add MAttach
        Rem If MAttachI <> "" Then
        Rem     .Attachments.Add MAttachI
        Rem End If
        Rem If MAttachII <> "" Then
        Rem     .Attachments.Add MAttachII
        Rem End If
        Rem If MAttachIII > "" Then
        Rem     .Attachments.Add MAttachIII
        Rem End If
        Rem If MAttachIV <> "" Then
        Rem     .Attachments.Add MAttachIV
        Rem End If
        Rem If MAttachV <> "" Then
        Rem     .Attachments.Add MAttachV
        Rem End If
        Rem If MAttachVI <> "" Then
        Rem     .Attachments.Add MAttachVI
        Rem End If
        Rem If MAttachVII <> "" Then
        Rem     .Attachments.Add MAttachVII
        Rem End If
        Rem If MAttachVIII <> "" Then
        Rem     .Attachments.Add MAttachVIII
        Rem End If
        .Send
    End With

    Set objMailItem = Nothing
    Set objOutlook = Nothing
            
    Exit Sub

exit10:
    Exit Sub

10:
    If Err.Number = 429 Then
        MsgBox "Error on connecting with Outlook"
            Else
        MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    End If
    Set objMailItem = Nothing
    Set objOutlook = Nothing
    AllPath = ""

    Resume exit10

End Sub

