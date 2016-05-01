VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCambios 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Cambios"
   ClientHeight    =   4560
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   7245
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   7245
   Begin VB.TextBox CambioVI 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   34
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox CambioV 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   32
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox CambioIV 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   30
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox CambioIII 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   28
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox CambioII 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   26
      Top             =   1000
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   360
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   3735
      Begin MSMask.MaskEdBox Hasta 
         Height          =   375
         Left            =   1320
         TabIndex        =   25
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   375
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4440
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCambios.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cambios"
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
      Left            =   4320
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "Cambios.frx":0000
      Left            =   480
      List            =   "Cambios.frx":0007
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      Height          =   300
      Left            =   1800
      TabIndex        =   13
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   600
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1335
      Left            =   4920
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3000
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Cambio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox Opcion 
      Height          =   1230
      Left            =   840
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Rofex 120"
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   35
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Rofex 90"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   33
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Rofex 60"
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   31
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Rofex 30"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   29
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Paridad del Euro"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   27
      Top             =   1000
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Paridad del Dolar"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "PrgCambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCambio As Recordset
Dim spCambio As String
Dim XParam As String
Dim ZCambioII As Double
Dim ZCambioIII As Double
Dim ZCambioIV As Double
Dim ZCambioV As Double
Dim ZCambioVI As Double


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


Sub Verifica_datos()
    If Val(Cambio.Text) = 0 Then
        Cambio.Text = "0"
    End If
    If Val(CambioII.Text) = 0 Then
        CambioII.Text = "0"
    End If
    If Val(CambioIII.Text) = 0 Then
        CambioIII.Text = "0"
    End If
    If Val(CambioIV.Text) = 0 Then
        CambioIV.Text = "0"
    End If
    If Val(CambioV.Text) = 0 Then
        CambioV.Text = "0"
    End If
    If Val(CambioVI.Text) = 0 Then
        CambioVI.Text = "0"
    End If
End Sub
Sub Format_datos()
    Cambio.Text = Pusing("###,###.###", Cambio.Text)
    CambioII.Text = Pusing("###,###.###", CambioII.Text)
    CambioIII.Text = Pusing("###,###.###", CambioIII.Text)
    CambioIV.Text = Pusing("###,###.###", CambioIV.Text)
    CambioV.Text = Pusing("###,###.###", CambioV.Text)
    CambioVI.Text = Pusing("###,###.###", CambioVI.Text)
End Sub

Sub Imprime_Datos()

    WFecha = Fecha.Text
    spCambio = "ConsultaCambio " + "'" + Fecha.Text + "'"
    Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambio.RecordCount > 0 Then
        Fecha.Text = rstCambio!Fecha
        Cambio.Text = rstCambio!Cambio
        
        ZCambioII = IIf(IsNull(rstCambio!CambioII), "0", rstCambio!CambioII)
        ZCambioIII = IIf(IsNull(rstCambio!CambioIII), "0", rstCambio!CambioIII)
        ZCambioIV = IIf(IsNull(rstCambio!CambioIV), "0", rstCambio!CambioIV)
        ZCambioV = IIf(IsNull(rstCambio!CambioV), "0", rstCambio!CambioV)
        ZCambioVI = IIf(IsNull(rstCambio!CambioVI), "0", rstCambio!CambioVI)
        
        CambioII.Text = Trim(Str$(ZCambioII))
        CambioIII.Text = Trim(Str$(ZCambioIII))
        CambioIV.Text = Trim(Str$(ZCambioIV))
        CambioV.Text = Trim(Str$(ZCambioV))
        CambioVI.Text = Trim(Str$(ZCambioVI))
        
        rstCambio.Close
        Call Format_datos
    End If

End Sub

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Cambios"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Cambios.OrdFecha} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Cambios.Fecha , Cambios.Cambio, Cambios.OrdFecha " _
                        + "From " + DSQ + ".dbo.Cambios Cambios " _
                        + "Where Cambios.OrdFecha >= '0' AND Cambios.OrdFecha <= '99999999'"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
      
    Fecha.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Fecha.Text <> "" Then
    
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cambios"
        ZSql = ZSql + " Where Cambios.Fecha = " + "'" + Fecha.Text + "'"
        spCambio = ZSql
        Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambio.RecordCount > 0 Then
            rstCambio.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Cambios SET "
            ZSql = ZSql + " Cambio = " + "'" + Cambio.Text + "',"
            ZSql = ZSql + " CambioII = " + "'" + CambioII.Text + "',"
            ZSql = ZSql + " CambioIII = " + "'" + CambioIII.Text + "',"
            ZSql = ZSql + " CambioIV = " + "'" + CambioIV.Text + "',"
            ZSql = ZSql + " CambioV = " + "'" + CambioV.Text + "',"
            ZSql = ZSql + " CambioVI = " + "'" + CambioVI.Text + "'"
            ZSql = ZSql + " Where Fecha = " + "'" + Fecha.Text + "'"
            spCambio = ZSql
            Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
                 Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Cambios ("
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Cambio ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "CambioII ,"
            ZSql = ZSql + "CambioIII ,"
            ZSql = ZSql + "CambioIV ,"
            ZSql = ZSql + "CambioV ,"
            ZSql = ZSql + "CambioVI )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + Cambio.Text + "',"
            ZSql = ZSql + "'" + WOrdFecha + "',"
            ZSql = ZSql + "'" + CambioII.Text + "',"
            ZSql = ZSql + "'" + CambioIII.Text + "',"
            ZSql = ZSql + "'" + CambioIV.Text + "',"
            ZSql = ZSql + "'" + CambioV.Text + "',"
            ZSql = ZSql + "'" + CambioVI.Text + "')"
            spCambio = ZSql
            Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        Fecha.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Fecha.Text <> "" Then
        
        spCambio = "ConsultaCambio " + "'" + Fecha.Text + "'"
        Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambio.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        If WPasa = "S" Then
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spCambio = "BorrarCambio " + "'" + Fecha.Text + "'"
                Set rstCambio = db.OpenRecordset(spCambio, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    Fecha.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cambio.Text = ""
    CambioII.Text = ""
    CambioIII.Text = ""
    CambioIV.Text = ""
    CambioV.Text = ""
    CambioVI.Text = ""
    Fecha.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Rem With rstCambios
    Rem     .Close
    Rem End With
    Rem With rstEmpresa
    Rem     .Close
    Rem End With
    Rem With rstAuxiliar
    Rem     .Close
    Rem End With
    Rem DbsVentas.Close
    PrgCambios.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Cambio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CambioII.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CambioII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CambioIII.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CambioIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CambioIV.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CambioIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CambioV.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CambioV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CambioVI.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CambioVI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cambio.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Fecha.Text <> "" Then
            WFecha = Fecha.Text
            spCambio = "ConsultaCambio " + "'" + Fecha.Text + "'"
            Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambio.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
                
            If WPasa = "S" Then
                Fecha.Text = rstCambio!Fecha
                Call Imprime_Datos
                    Else
                WFecha = Fecha.Text
                CmdLimpiar_Click
                Fecha.Text = WFecha
            End If
        End If
        Cambio.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     'Opcion.Clear
     '
     'Opcion.AddItem "Fecha"
     'Opcion.AddItem "Cuentas Contables"

     'Opcion.Visible = True
     
'End Sub

'' Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    'XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
            spCambio = "ListaCambio"
            Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstCambio
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCambio!Fecha + " " + Str$(rstCambio!Cambio)
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCambio!Fecha
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCambio.Close
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WFecha = WIndice.List(Indice)
            spCambio = "ConsultaCambio " + "'" + WFecha + "'"
            Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambio.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
                Fecha.Text = rstCambio!Fecha
                Call Imprime_Datos
                        Else
                CmdLimpiar_Click
                Fecha.Text = WFecha
            End If
        
            Fecha.SetFocus
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cambio.Text = ""
    CambioII.Text = ""
    CambioIII.Text = ""
    CambioIV.Text = ""
    CambioV.Text = ""
    CambioVI.Text = ""
End Sub


Private Sub Primer_Click()

    On Error GoTo WError
    
    spCambio = "ListaCambio"
    Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstCambio
        .MoveFirst
        Fecha.Text = rstCambio!Fecha
    End With
    
    rstCambio.Close
    Call Imprime_Datos
    Fecha.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cambio", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Fecha.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spCambio = "ListaCambio"
    Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstCambio
        .MoveLast
        Fecha.Text = rstCambio!Fecha
    End With
    
    rstCambio.Close
    Call Imprime_Datos
    Fecha.SetFocus
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Cambio", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Fecha.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spCambio = "AnteriorCambio " + "'" + Fecha.Text + "'"
    Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstCambio
        .MoveLast
        Fecha.Text = rstCambio!Fecha
    End With
    
    rstCambio.Close
    Call Imprime_Datos
    Fecha.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cambio", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Fecha.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spCambio = "PosteriorCambio " + "'" + Fecha.Text + "'"
    Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstCambio
        .MoveFirst
        Fecha.Text = rstCambio!Fecha
    End With
    
    rstCambio.Close
    Call Imprime_Datos
    Fecha.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cambio", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Fecha.SetFocus
    
End Sub


