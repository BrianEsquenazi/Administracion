VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgVerifica 
   AutoRedraw      =   -1  'True
   Caption         =   "Verificacion de Correlatividades"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   3135
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox Desde1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Hasta1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wverifica.rpt"
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
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgVerifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim XParam As String
Dim A1 As String
Dim A2 As String


Private Sub Acepta_Click()

    On Error GoTo WError
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            XEmpresa = !Nombre
        End If
    End With
    
    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)

    Da = 0
    With rstVerifica
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Select Case Tipo.ListIndex
        Case 0
        
            ZSql = ""
            ZSql = ZSql + "Select Hoja.FechaIng, Hoja.Renglon, Hoja.Hoja, Hoja.Producto, Hoja.Fecha"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Hoja.Renglon = 1"
            ZSql = ZSql + " and Hoja.Hoja >= " + "'" + Desde1.Text + "'"
            ZSql = ZSql + " and Hoja.Hoja <= " + "'" + Hasta1.Text + "'"
            ZSql = ZSql + " and Hoja.FechaIngOrd >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and Hoja.FechaIngOrd <= " + "'" + WHasta + "'"
            ZSql = ZSql + " Order by Hoja.Hoja"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
        
            With rstHoja
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        Rem WFecha = Right$(rstHoja!fechaIng, 4) + Mid$(rstHoja!fechaIng, 4, 2) + Left$(rstHoja!fechaIng, 2)
                        Rem If WFecha >= WDesde And WFecha <= WHasta Then
                        
                            WHoja = rstHoja!Hoja
                            WFecha = rstHoja!Fecha
                            WProducto = rstHoja!Producto
                        
                            With rstVerifica
                
                                .AddNew
                                !Numero = WHoja
                                !Descri = "Hoja Prod."
                                !Fecha = WFecha
                                !Estado = ""
                                !Texto = WProducto
                                !Titulo = XEmpresa
                                .Update
                            
                            End With
                        
                        Rem End If
                
                        .MoveNext
                    
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            
            rstHoja.Close
            
            End If
            
            
            
            ZSql = ""
            ZSql = ZSql + "Select Hoja.FechaIng, Hoja.Renglon, Hoja.Hoja, Hoja.Producto, Hoja.Fecha"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Hoja.Cantidad = 0"
            ZSql = ZSql + " and Hoja.Renglon = 1"
            ZSql = ZSql + " and Hoja.Hoja >= " + "'" + Desde1.Text + "'"
            ZSql = ZSql + " and Hoja.Hoja <= " + "'" + Hasta1.Text + "'"
            ZSql = ZSql + " and Hoja.FechaIngOrd >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and Hoja.FechaIngOrd <= " + "'" + WHasta + "'"
            ZSql = ZSql + " Order by Hoja.Hoja"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
            
            With rstHoja
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        WHoja = rstHoja!Hoja
                        WFecha = rstHoja!Fecha
                        
                        With rstVerifica
                            .Index = "Clave"
                            .Seek "=", WHoja
                            If .NoMatch = False Then
                                .Delete
                            End If
                        End With
                
                        .MoveNext
                    
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            
            rstHoja.Close
            
            End If
            
        Case 1
        
            A1 = "      "
            A2 = "ZZZZZZ"
   
            XParam = "'" + A1 + "','" _
                    + A2 + "'"
            spCtacte = "ListaCtacteDesdeHasta " + XParam
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtacte.RecordCount > 0 Then
        
            With rstCtacte
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        If Val(rstCtacte!Tipo) < 6 Then
                
                        If rstCtacte!Numero >= Val(Desde1.Text) And rstCtacte!Numero <= Val(Hasta1.Text) Then
                    
                            WFecha = Right$(rstCtacte!Fecha, 4) + Mid$(rstCtacte!Fecha, 4, 2) + Left$(rstCtacte!Fecha, 2)
                            If WFecha >= WDesde And WFecha <= WHasta Then
                        
                            WHoja = rstCtacte!Numero
                            WFecha = rstCtacte!Fecha
                            WCliente = rstCtacte!Cliente
                        
                            With rstVerifica
                
                                .AddNew
                                !Numero = WHoja
                                !Descri = "Factura"
                                !Fecha = WFecha
                                !Estado = ""
                                !Texto = WCliente
                                !Titulo = XEmpresa
                                .Update
                            
                            End With
                            End If
                        End If
                        
                        End If
                
                        .MoveNext
                    
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            
            rstCtacte.Close
            
            End If
            
        Case 2
        
            A1 = "      "
            A2 = "ZZZZZZ"
   
            XParam = "'" + A1 + "','" _
                    + A2 + "'"
            spCtacte = "ListaCtacteDesdeHasta " + XParam
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtacte.RecordCount > 0 Then
        
            With rstCtacte
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        If Val(rstCtacte!Tipo) = 6 Then
                
                        If rstCtacte!Numero >= Val(Desde1.Text) And rstCtacte!Numero <= Val(Hasta1.Text) Then
                    
                            WFecha = Right$(rstCtacte!Fecha, 4) + Mid$(rstCtacte!Fecha, 4, 2) + Left$(rstCtacte!Fecha, 2)
                            If WFecha >= WDesde And WFecha <= WHasta Then
                        
                            WHoja = rstCtacte!Numero
                            WFecha = rstCtacte!Fecha
                            WCliente = rstCtacte!Cliente
                        
                            With rstVerifica
                
                                .AddNew
                                !Numero = WHoja
                                !Descri = "Recibos"
                                !Fecha = WFecha
                                !Estado = ""
                                !Texto = WCliente
                                !Titulo = XEmpresa
                                .Update
                            
                            End With
                            End If
                        End If
                        
                        End If
                
                        .MoveNext
                    
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            
            rstCtacte.Close
            
            End If
            
            
        Case 3
        
            spLaudo = "ListaLaudoTotal"
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
        
            With rstLaudo
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        If rstLaudo!Renglon = 1 Then
                
                        If rstLaudo!Laudo >= Val(Desde1.Text) And rstLaudo!Laudo <= Val(Hasta1.Text) Then
                    
                            WFecha = Right$(rstLaudo!Fecha, 4) + Mid$(rstLaudo!Fecha, 4, 2) + Left$(rstLaudo!Fecha, 2)
                            If WFecha >= WDesde And WFecha <= WHasta Then
                        
                            WHoja = rstLaudo!Laudo
                            WFecha = rstLaudo!Fecha
                        
                            With rstVerifica
                
                                .AddNew
                                !Numero = WHoja
                                !Descri = "Laudos"
                                !Fecha = WFecha
                                !Estado = ""
                                !Texto = ""
                                !Titulo = XEmpresa
                                .Update
                            
                            End With
                            End If
                        End If
                        
                        End If
                
                        .MoveNext
                    
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            
            rstLaudo.Close
            
            End If
            
        Case 4
        
            spMovguia = "ListaMovguiaTotal"
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
        
            With rstMovguia
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        If rstMovguia!Renglon = 1 And Val(rstMovguia!Codigo) < 900000 And Val(rstMovguia!Codigo) > 0 And rstMovguia!Tipomov = 0 Then
                
                            WFecha = Right$(rstMovguia!Fecha, 4) + Mid$(rstMovguia!Fecha, 4, 2) + Left$(rstMovguia!Fecha, 2)
                            If WFecha >= WDesde And WFecha <= WHasta Then
                        
                                WHoja = rstMovguia!Codigo
                                WFecha = rstMovguia!Fecha
                        
                                With rstVerifica
                
                                    .AddNew
                                    !Numero = WHoja
                                    !Descri = "Guias"
                                    !Fecha = WFecha
                                    !Estado = ""
                                    !Texto = ""
                                    !Titulo = XEmpresa
                                    .Update
                            
                                End With
                            End If
                        End If
                
                        .MoveNext
                    
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            
            rstMovguia.Close
            
            End If
            
        Case 5
        
            spMovguia = "ListaMovguiaTotal"
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
        
            With rstMovguia
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        If rstMovguia!Renglon = 1 And Val(rstMovguia!Codigo) >= 900000 And rstMovguia!Tipomov = 0 Then
                
                            WFecha = Right$(rstMovguia!Fecha, 4) + Mid$(rstMovguia!Fecha, 4, 2) + Left$(rstMovguia!Fecha, 2)
                            If WFecha >= WDesde And WFecha <= WHasta Then
                        
                                WHoja = rstMovguia!Codigo
                                WFecha = rstMovguia!Fecha
                        
                                With rstVerifica
                
                                    .AddNew
                                    !Numero = WHoja
                                    !Descri = "Prestamos"
                                    !Fecha = WFecha
                                    !Estado = ""
                                    !Texto = ""
                                    !Titulo = XEmpresa
                                    .Update
                            
                                End With
                            End If
                        End If
                
                        .MoveNext
                    
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            
            rstMovguia.Close
            
            End If
            
            
            
            
        Case Else
        
    End Select
    
    Pasa = 0

    With rstVerifica
    
        .Index = "CLAVE"
        .MoveFirst
            
        If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If Pasa = 0 Then
                    WNumero = !Numero
                    Pasa = 1
                End If
                
                If Val(WNumero) <> !Numero Then
                                        
                    With rstVerifica
                
                        .AddNew
                        !Numero = WNumero
                        !Descri = ""
                        !Fecha = "  /  /    "
                        !Estado = "FALTANTE"
                        !Titulo = XEmpresa
                        WTexto = ""
                        
                        If Tipo.ListIndex = 0 Then
                            spHoja = "ListaHoja " + "'" + WNumero + "'"
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                WTexto = rstHoja!Producto
                                rstHoja.Close
                            End If
                        End If
                        
                        !Texto = WTexto
                        
                        .Update
                            
                    End With
                    WNumero = WNumero + 1
                    
                        Else
                
                    WNumero = WNumero + 1
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                End If
                
            Loop
        End If
    End With

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    Listado.WindowTitle = "Listado de Verificacion de Correlatividades"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{FichaMat.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    Listado.Action = 1
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Cancela_click()

    With rstVerifica
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    Desde.SetFocus
    PrgVerifica.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Verifica
End Sub


Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgVerifica.Caption = "Verificacion de Correlatividades :  " + !Nombre
        End If
    End With
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Desde1.Text = 0
    Hasta1.Text = 999999
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
    Tipo.Clear
    
    Tipo.AddItem "Hoja de Producction"
    Tipo.AddItem "Facturas"
    Tipo.AddItem "Recibos"
    Tipo.AddItem "Laudos"
    Tipo.AddItem "Guias"
    Tipo.AddItem "Prestamos"
    
    Tipo.ListIndex = 0
    
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub
Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde1.SetFocus
    End If
End Sub
Private Sub Desde1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta1.SetFocus
    End If
End Sub
Private Sub Hasta1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

