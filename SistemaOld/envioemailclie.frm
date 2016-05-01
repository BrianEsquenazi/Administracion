VERSION 5.00
Begin VB.Form PrgEnvioEmailClie 
   AutoRedraw      =   -1  'True
   Caption         =   "Envio de email a clientes"
   ClientHeight    =   7005
   ClientLeft      =   15
   ClientTop       =   480
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   7005
   ScaleWidth      =   11880
   Begin VB.Frame IngresaArchivo 
      BackColor       =   &H00C0FFFF&
      Height          =   6135
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   3975
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   3975
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   480
         TabIndex        =   2
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox Archivo 
         Height          =   285
         Left            =   480
         TabIndex        =   1
         Text            =   " "
         Top             =   4440
         Width           =   3975
      End
      Begin VB.Image AceptaFoto 
         Height          =   480
         Left            =   1920
         MouseIcon       =   "envioemailclie.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "envioemailclie.frx":030A
         ToolTipText     =   "Confirma el Proceso"
         Top             =   5160
         Width           =   480
      End
   End
   Begin VB.Frame EnvioPto 
      BackColor       =   &H00808080&
      Height          =   6495
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   10935
      Begin VB.TextBox ZMensajeIX 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Width           =   7815
      End
      Begin VB.TextBox ZMensajeVIII 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   7815
      End
      Begin VB.TextBox ZMensajeVII 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   3960
         Width           =   7815
      End
      Begin VB.TextBox ZMensajeI 
         Height          =   1725
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1800
         Width           =   7815
      End
      Begin VB.TextBox ZAsunto 
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1080
         Width           =   8055
      End
      Begin VB.TextBox ZMensajeVI 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   7815
      End
      Begin VB.CommandButton EnvioMensaje 
         Caption         =   "Enviar"
         Height          =   495
         Left            =   1680
         TabIndex        =   8
         Top             =   5160
         Width           =   2415
      End
      Begin VB.CommandButton CancelaMensaje 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   5040
         TabIndex        =   7
         Top             =   5160
         Width           =   2415
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar Archivo"
         Height          =   615
         Left            =   8760
         TabIndex        =   6
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label22 
         BackColor       =   &H00808080&
         Caption         =   "Asunto : "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackColor       =   &H00808080&
         Caption         =   "Mensaje"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
   End
End
Attribute VB_Name = "PrgEnvioEmailClie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZDireccion As String
Dim EmailAddress As String
Dim CopiaAddress As String
Dim MSubject As String
Dim MBody As String
Dim MAttach As String
Dim MAttachI As String
Dim MAttachII As String
Dim MAttachIII As String
Dim MAttachIV As String
Dim MAttachV As String
Dim MAttachVI As String
Dim MAttachVII As String
Dim MAttachVIII As String

Dim ZZCliente(10000, 3) As String

Dim ZZArchivoEnvio As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub
    
Private Sub AceptaFoto_Click()
    IngresaArchivo.Visible = False
End Sub

Private Sub CancelaMensaje_Click()
    PrgEnvioEmailClie.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub cmdAgregar_Click()
    IngresaArchivo.Visible = True
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path  ' Establece la ruta del archivo.
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
End Sub

Private Sub File1_Click()

    On Error GoTo WError
    
    WDrive = Drive1.Drive
    WDir = Dir1.Path
    WLon = Len(WDir)
    If Right$(WDir, 1) = "\" Then
        WDir = Mid(WDir, 1, WLon - 1)
    End If
    XNombre = WDir + "\"
    
    WPasoUnifica = XNombre + File1.filename
    Archivo.Text = XNombre + File1.filename
    
    Exit Sub
    
WError:
    Rem MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
    
End Sub

Private Sub EnvioMensaje_Click()

    ZDireccion = ""
    ZZArchivoEnvio = Archivo.Text
    
    ZLugar = 0
    Erase ZZCliente
    
    spCliente = "ListaCliente"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    With rstCliente
        .MoveFirst
        Do
            If .EOF = False Then
            
                If Val(!Provincia) < 24 Then
            
                    ZZRazon = !Razon
                    ZZEmail = IIf(IsNull(!email), "", !email)
                    ZZEmail = Trim(ZZEmail)
                    
                    If ZZEmail <> "" And Len(ZZEmail) > 8 Then
                        ZLugar = ZLugar + 1
                        
                        ZZCliente(ZLugar, 1) = !Cliente
                        ZZCliente(ZLugar, 2) = ZZRazon
                        ZZCliente(ZLugar, 3) = ZZEmail
                    
                    End If
                    
                End If
                
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    
    
    
    ZZMes = Left$(Date$, 2)
    ZZAno = Trim(Str$(Val(Right$(Date$, 4)) - 1))
    
    ZZOrdFecha = ZZAno + ZZMes + "01"

    For Ciclo = 1 To ZLugar
    
        ZZZZCliente = ZZCliente(Ciclo, 1)
        ZZRazon = ZZCliente(Ciclo, 2)
        ZZEmail = ZZCliente(Ciclo, 3)
        ZZEmail = Trim(ZZEmail)
        
        ZSql = "Select *"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.Cliente = " + "'" + ZZZZCliente + "'"
        ZSql = ZSql + " and CtaCte.Tipo = " + "'" + "01" + "'"
        ZSql = ZSql + " and CtaCte.OrdFecha >= " + "'" + ZZOrdFecha + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            rstCtacte.Close
        
            If Trim(ZDireccion) = "" Then
                ZDireccion = Trim(ZZEmail)
                    Else
                ZDireccion = ZDireccion + ";" + Trim(ZZEmail)
            End If
            
            If Len(ZDireccion) > 100 Then
                Call Envio_Email
                ZDireccion = ""
            End If
            
        End If
        
    Next Ciclo
    
    If ZDireccion <> "" Then
        Call Envio_Email
        ZDireccion = ""
    End If
                    
    
    m$ = "Proceso Finalizado"
    a% = MsgBox(m$, 0, "Envio de Email")
    
    Call CancelaMensaje_Click

End Sub

Public Sub Envio_Email()

    ZTexto1 = ZMensajeI.Text
    ZTexto6 = ZMensajeVI.Text
    ZTexto7 = ZMensajeVII.Text
    ZTexto8 = ZMensajeVIII.Text
    ZTexto9 = ZMensajeIX.Text

    sTo = "surfactan@surfactan.com.ar"
    sCC = ""
    sBCC = ZDireccion
    sSubject = ZAsunto.Text
    sBody = ZTexto1
    If Trim(ZTexto6) <> "" Then
        sBody = sBody + Chr$(13) + ZTexto6
    End If
    If Trim(ZTexto7) <> "" Then
        sBody = sBody + Chr$(13) + ZTexto7
    End If
    If Trim(ZTexto8) <> "" Then
        sBody = sBody + Chr$(13) + ZTexto8
    End If
    If Trim(ZTexto9) <> "" Then
        sBody = sBody + Chr$(13) + ZTexto9
    End If
    SFile = ZZArchivoEnvio
    
    EmailAddress = sTo
    CopiaAddress = sCC
    MSubject = sSubject
    MBody = sBody
    MAttach = SFile
    MAttachI = ""
    MAttachII = ""
    MAttachIII = ""
    MAttachIV = ""
    MAttachVI = ""
    MAttachVII = ""
    MAttachVIII = ""
    
    SendEmail
    
    Call CancelaMensaje_Click

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
        Rem .bcc = "d_esquenazi@yahoo.com"
        .bcc = ZDireccion
        .Subject = MSubject
        .Body = MBody
        .Attachments.Add MAttach
        If MAttachI <> "" Then
            .Attachments.Add MAttachI
        End If
        If MAttachII <> "" Then
            .Attachments.Add MAttachII
        End If
        If MAttachIII > "" Then
            .Attachments.Add MAttachIII
        End If
        If MAttachIV <> "" Then
            .Attachments.Add MAttachIV
        End If
        If MAttachV <> "" Then
            .Attachments.Add MAttachV
        End If
        If MAttachVI <> "" Then
            .Attachments.Add MAttachVI
        End If
        If MAttachVII <> "" Then
            .Attachments.Add MAttachVII
        End If
        If MAttachVIII <> "" Then
            .Attachments.Add MAttachVIII
        End If
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

