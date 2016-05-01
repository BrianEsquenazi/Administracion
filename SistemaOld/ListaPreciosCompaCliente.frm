VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaPreciosCompaCliente 
   AutoRedraw      =   -1  'True
   Caption         =   "Emision de Lista de Precios Comparativa (Cliente)"
   ClientHeight    =   3060
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3060
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   6375
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
         Height          =   495
         Left            =   4920
         TabIndex        =   7
         Top             =   1560
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
         Height          =   495
         Left            =   4920
         TabIndex        =   6
         Top             =   840
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
         Left            =   960
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
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
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Cliente 
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
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   3
         Text            =   " "
         Top             =   1440
         Width           =   975
      End
      Begin MSMask.MaskEdBox FechaCompa 
         Height          =   300
         Left            =   2520
         TabIndex        =   1
         Top             =   960
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
      Begin MSMask.MaskEdBox Fecha 
         Height          =   300
         Left            =   2520
         TabIndex        =   0
         Top             =   480
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
      Begin VB.Label Label1 
         Caption         =   "Fecha Emision"
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
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Comparativa"
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
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde Hasta Cliente"
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
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WListaPrecios.rpt"
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
Attribute VB_Name = "PrgListaPreciosCompaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCargaLista As Recordset
Dim spCargaLista As String
Dim rstCargaListaII As Recordset
Dim spCargaListaII As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String

Private Auxiliar(100, 7) As String
Private ZVector(5000, 3) As String
Private Producto As String

Private ZCostoI As Double
Private ZPorceI As Double
Private ZDifeI As Double
Private ZFechaI As String

Private ZCostoII As Double
Private ZPorceII As Double
Private ZDifeII As Double
Private ZFechaII As String

Private ZDiferencia As Double
Private ZDife As Double

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Acepta_Click()

    On Error GoTo WError
    
    WFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WFechaCompa = Right$(FechaCompa.Text, 4) + Mid$(FechaCompa.Text, 4, 2) + Left$(FechaCompa.Text, 2)
    
    ZVersion = "9999"
    ZCliente = Cliente.Text
    
    ZSql = ""
    ZSql = ZSql + "DELETE CargaLista"
    ZSql = ZSql + " Where Lista = " + "'" + ZVersion + "'"
    rsCargaLista = ZSql
    Set rstCargaLista = db.OpenRecordset(rsCargaLista, dbOpenSnapshot, dbSQLPassThrough)
    
    OPEN_FILE_Empresa
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            ZEmpresa = !Nombre
        End If
    End With
    
    
    Erase ZVector
    LugarVector = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Precios"
    ZSql = ZSql + " Where Precios.Cliente = " + "'" + ZCliente + "'"
    spPrecios = ZSql
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveFirst
            Do
                If .EOF = False Then
                    If rstPrecios!Precio <> 0 Then
                        LugarVector = LugarVector + 1
                        ZVector(LugarVector, 1) = rstPrecios!Terminado
                        ZVector(LugarVector, 2) = rstPrecios!Fecha
                        ZVector(LugarVector, 3) = Str$(rstPrecios!Precio)
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPrecios.Close
    End If
    
    ZRazon = ""
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZCliente + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZRazon = rstCliente!Razon
        rstCliente.Close
    End If
    
    For Ciclo = 1 To LugarVector
    
        ZTerminado = ZVector(Ciclo, 1)
        ZFechaI = ZVector(Ciclo, 2)
        ZPrecio = Val(ZVector(Ciclo, 3))
        
        ZDifeI = 0
        ZPorceI = 0
        ZCostoI = 0
        
        ZDifeII = 0
        ZPorceII = 0
        ZCostoII = 0
        ZFechaII = ""
        
        ZDife = 0
        
        Producto = ZTerminado
        Call Calcula_Costo(Producto, ZCostoI)
        
        If ZCostoI <> 0 Then
            Rem ZDifeI = ZPrecio - ZCostoI
            ZPorceI = ZPrecio / ZCostoI
            Call Redondeo(ZPorceI)
        End If
            
        If FechaCompa.Text <> "  /  /    " Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaListaII"
            ZSql = ZSql + " Where CargaListaII.Terminado = " + "'" + ZTerminado + "'"
            ZSql = ZSql + " and CargaListaII.OrdFecha <= " + "'" + WFechaCompa + "'"
            ZSql = ZSql + " Order by OrdFecha"
            spCargaListaII = ZSql
            Set rstCargaListaII = db.OpenRecordset(spCargaListaII, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaListaII.RecordCount > 0 Then
                With rstCargaListaII
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            ZCostoII = rstCargaListaII!Costo
                            ZFechaII = rstCargaListaII!Fecha
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCargaListaII.Close
            End If
        End If
                
        If ZCostoII <> 0 Then
            Rem ZDifeII = ZPrecio - ZCostoII
            ZPorceII = ZPrecio / ZCostoII
            Call Redondeo(ZPorceII)
        End If
                    
        If ZCostoI <> 0 And ZCostoII <> 0 Then
            ZDiferencia = ZCostoI - ZCostoII
            ZDife = (ZDiferencia / ZCostoII) * 100
            Call Redondeo(ZDife)
        End If
        
        ZLinea = "0"
        ZDesTerminado = ""
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Terminado"
        ZSql = ZSql + " Where Terminado.Codigo = " + "'" + ZTerminado + "'"
        spTerminado = ZSql
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZLinea = Str$(rstTerminado!Linea)
            ZDescripcion = rstTerminado!Descripcion
            rstTerminado.Close
        End If
        
        Rem dar de alta
        
        ZLista = "9999"
        ZRenglon = Str$(Ciclo)
        ZFecha = ZFechaI
        ZOrdFecha = Right$(ZFechaI, 4) + Mid$(ZFechaI, 4, 2) + Left$(ZFechaI, 2)
        ZTitulo = "Cliente : " + ZCliente + "   " + ZRazon
        ZTitulo = Left$(ZTitulo, 50)
        ZObservaciones = ""
        ZTerminado = ZTerminado
        ZPrecio = ZPrecio
        ZDescripcion = ZDescripcion
        ZLinea = ZLinea
        ZEmpresa = ZEmpresa
        
        Auxi = Str$(Ciclo)
        Call Ceros(Auxi, 3)
        
        Auxi1 = ZVersion
        Call Ceros(Auxi1, 6)
        
        ZClave = Auxi1 + Auxi
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaLista ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Lista ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "Titulo ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Precio ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Linea ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "CostoI ,"
        ZSql = ZSql + "CostoII ,"
        ZSql = ZSql + "FactorI ,"
        ZSql = ZSql + "FactorII ,"
        ZSql = ZSql + "Porce ,"
        ZSql = ZSql + "FechaI ,"
        ZSql = ZSql + "FechaII )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZClave + "',"
        ZSql = ZSql + "'" + ZLista + "',"
        ZSql = ZSql + "'" + ZRenglon + "',"
        ZSql = ZSql + "'" + ZFecha + "',"
        ZSql = ZSql + "'" + ZOrdFecha + "',"
        ZSql = ZSql + "'" + ZTitulo + "',"
        ZSql = ZSql + "'" + ZObservaciones + "',"
        ZSql = ZSql + "'" + ZTerminado + "',"
        ZSql = ZSql + "'" + Str$(ZPrecio) + "',"
        ZSql = ZSql + "'" + ZDescripcion + "',"
        ZSql = ZSql + "'" + ZLinea + "',"
        ZSql = ZSql + "'" + ZEmpresa + "',"
        ZSql = ZSql + "'" + Str$(ZCostoI) + "',"
        ZSql = ZSql + "'" + Str$(ZCostoII) + "',"
        ZSql = ZSql + "'" + Str$(ZPorceI) + "',"
        ZSql = ZSql + "'" + Str$(ZPorceII) + "',"
        ZSql = ZSql + "'" + Str$(ZDife) + "',"
        ZSql = ZSql + "'" + ZFechaI + "',"
        ZSql = ZSql + "'" + ZFechaII + "')"
       
        spCargaLista = ZSql
        Set rstCargaLista = db.OpenRecordset(spCargaLista, dbOpenSnapshot, dbSQLPassThrough)
                        
        ZClave = WFecha + ZTerminado
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaListaII"
        ZSql = ZSql + " Where CargaListaII.Clave = " + "'" + ZClave + "'"
        spCargaListaII = ZSql
        Set rstCargaListaII = db.OpenRecordset(spCargaListaII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaListaII.RecordCount > 0 Then
            
            rstCargaListaII.Close
                
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaListaII SET "
            ZSql = ZSql + " Costo = " + "'" + Str$(ZCostoI) + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spCargaListaII = ZSql
            Set rstCargaListaII = db.OpenRecordset(spCargaListaII, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaListaII ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Costo )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + WFecha + "',"
            ZSql = ZSql + "'" + ZTerminado + "',"
            ZSql = ZSql + "'" + Str$(ZCostoI) + "')"
       
            spCargaListaII = ZSql
            Set rstCargaListaII = db.OpenRecordset(spCargaListaII, dbOpenSnapshot, dbSQLPassThrough)
                
        End If
        
    Next Ciclo
    
    Listado.WindowTitle = "Emision de Lista de Precios Comparativo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{CargaLista.Lista} in " + ZVersion + " to " + ZVersion
    Dos = " and {CargaLista.Linea} in " + "0" + " to " + "9999"
    Listado.GroupSelectionFormula = Uno + Dos
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CargaLista.Lista, CargaLista.Fecha, CargaLista.Titulo, CargaLista.Observaciones, CargaLista.Terminado, CargaLista.Precio, CargaLista.Linea, CargaLista.Empresa, CargaLista.CostoI, CargaLista.CostoII, CargaLista.FactorI, CargaLista.FactorII, CargaLista.Porce, CargaLista.FechaI, CargaLista.FechaII, " _
                + "Terminado.Descripcion, " _
                + "Lineas.Nombre " _
                + "From " _
                + DSQ + ".dbo.CargaLista CargaLista, " _
                + DSQ + ".dbo.Terminado Terminado, " _
                + DSQ + ".dbo.Lineas Lineas " _
                + "Where " _
                + "CargaLista.Terminado = Terminado.Codigo AND " _
                + "CargaLista.Linea = Lineas.Linea AND " _
                + "CargaLista.Lista >= 9999 AND " _
                + "CargaLista.Lista <= 9999 AND " _
                + "CargaLista.Linea >= 0 AND " _
                + "CargaLista.Linea <= 9999"
                
    If Val(WEmpresa) = 1 Then
        Listado.ReportFileName = "WListaPreciosCompaIII.rpt"
            Else
        Listado.ReportFileName = "WListaPreciosCompaIV.rpt"
    End If
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaPreciosCompaCliente.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaCompa.SetFocus
    End If
End Sub

Private Sub FechaCompa_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.SetFocus
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
End Sub

Sub Form_Load()
    Fecha.Text = "  /  /    "
    FechaCompa.Text = "  /  /    "
    Cliente.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
    Do
        Cicla = Cicla + 1
        If Vector(Cicla, 1) <> "" Then
    
            Entra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = Articulo2
                                    Vector(Lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = Cantidad
                                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                            Case Else
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
            End If
            
            If Entra = "S" Then
                If Left$(Vector(Cicla, 1), 2) <> "PT" Then
                    Renglon = Renglon + 1
                    Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
                    Auxiliar(Renglon, 2) = 1
                    Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                End If
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For DA = 1 To Renglon
        Articulo = Auxiliar(DA, 1)
        Cantidad = Auxiliar(DA, 2)
        WVector = Auxiliar(DA, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = (Cantidad * rstArticulo!Costo2 * Val(WVector))
            Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(WVector))
            rstArticulo.Close
        End If
    Next DA

End Sub


