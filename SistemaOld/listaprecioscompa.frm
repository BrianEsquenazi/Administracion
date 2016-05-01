VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaPreciosCompa 
   AutoRedraw      =   -1  'True
   Caption         =   "Emision de Lista de Precios Comparativa (Grupo)"
   ClientHeight    =   3060
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3060
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   6015
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
         Left            =   4440
         TabIndex        =   10
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
         Left            =   4440
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   2160
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
         Left            =   2520
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox DesdeLista 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   4
         TabIndex        =   2
         Text            =   " "
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox HastaLista 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   3
         Text            =   " "
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox HastaLinea 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   5
         Text            =   " "
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox DesdeLinea 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   4
         TabIndex        =   4
         Text            =   " "
         Top             =   1800
         Width           =   735
      End
      Begin MSMask.MaskEdBox FechaCompa 
         Height          =   300
         Left            =   2520
         TabIndex        =   1
         Top             =   840
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
         Left            =   480
         TabIndex        =   13
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
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Hasta Lista"
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
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Desde Hasta Linea"
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
         Left            =   480
         TabIndex        =   11
         Top             =   1800
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
Attribute VB_Name = "PrgListaPreciosCompa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCargaLista As Recordset
Dim spCargaLista As String
Dim rstCargaListaII As Recordset
Dim spCargaListaII As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim XParam As String

Private Auxiliar(100, 7) As String
Private ZVector(5000, 3) As String
Private Producto As String
Private Costo As Double
Private ZPorce As Double
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
    
    Erase ZVector
    LugarVector = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaLista"
    ZSql = ZSql + " Order by CargaLista.Clave"
    spCargaLista = ZSql
    Set rstCargaLista = db.OpenRecordset(spCargaLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaLista.RecordCount > 0 Then
        With rstCargaLista
            .MoveFirst
            Do
                If .EOF = False Then
                    LugarVector = LugarVector + 1
                    ZVector(LugarVector, 1) = rstCargaLista!Terminado
                    ZVector(LugarVector, 2) = rstCargaLista!Clave
                    ZVector(LugarVector, 3) = Str$(rstCargaLista!Precio)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaLista.Close
    End If
    
    For Ciclo = 1 To LugarVector
    
        ZTerminado = ZVector(Ciclo, 1)
        ZClave = ZVector(Ciclo, 2)
        ZPrecio = Val(ZVector(Ciclo, 3))
        
        Producto = ZTerminado
        Call Calcula_Costo(Producto, Costo)
        
        If Costo <> 0 Then
        
            Rem ZDife = ZPrecio - Costo
            ZPorce = ZPrecio / Costo
            Call Redondeo(ZPorce)
            CostoII = 0
            ZPorceII = 0
            FechaII = ""
            ZDife = 0
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaLista SET "
            ZSql = ZSql + " CostoI = " + "'" + Str$(Costo) + "',"
            ZSql = ZSql + " FactorI = " + "'" + Str$(ZPorce) + "',"
            ZSql = ZSql + " FechaI = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " CostoII = " + "'" + Str$(CostoII) + "',"
            ZSql = ZSql + " FactorII = " + "'" + Str$(ZPorceII) + "',"
            ZSql = ZSql + " FechaII = " + "'" + FechaII + "',"
            ZSql = ZSql + " Porce = " + "'" + Str$(ZDife) + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spCargaLista = ZSql
            Set rstCargaLista = db.OpenRecordset(spCargaLista, dbOpenSnapshot, dbSQLPassThrough)
            
            If FechaCompa.Text <> "  /  /    " Then
            
                CostoII = 0
                FechaII = ""
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CargaListaII"
                ZSql = ZSql + " Where CargaListaII.Terminado = " + "'" + ZTerminado + "'"
                ZSql = ZSql + " and CargaListaII.OrdFecha <= " + "'" + WFechaCompa + "'"
                spCargaListaII = ZSql
                Set rstCargaListaII = db.OpenRecordset(spCargaListaII, dbOpenSnapshot, dbSQLPassThrough)
                If rstCargaListaII.RecordCount > 0 Then
                    CostoII = rstCargaListaII!Costo
                    FechaII = rstCargaListaII!Fecha
                    rstCargaListaII.Close
                End If
                
                If CostoII <> 0 Then
                
                    Rem ZDife = ZPrecio - CostoII
                    ZPorce = ZPrecio / CostoII
                    Call Redondeo(ZPorce)
                    ZDife = 0
                    
                    If Costo <> 0 And CostoII <> 0 Then
                        ZDiferencia = Costo - CostoII
                        ZDife = (ZDiferencia / CostoII) * 100
                        Call Redondeo(ZDife)
                    End If
                        
                    ZSql = ""
                    ZSql = ZSql + "UPDATE CargaLista SET "
                    ZSql = ZSql + " CostoII = " + "'" + Str$(CostoII) + "',"
                    ZSql = ZSql + " FactorII = " + "'" + Str$(ZPorce) + "',"
                    ZSql = ZSql + " FechaII = " + "'" + FechaII + "',"
                    ZSql = ZSql + " Porce = " + "'" + Str$(ZDife) + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
                    spCargaLista = ZSql
                    Set rstCargaLista = db.OpenRecordset(spCargaLista, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
            End If
            
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
                ZSql = ZSql + " Costo = " + "'" + Str$(Costo) + "'"
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
                ZSql = ZSql + "'" + Str$(Costo) + "')"
       
                spCargaListaII = ZSql
                Set rstCargaListaII = db.OpenRecordset(spCargaListaII, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        End If
        
    Next Ciclo
    
    Listado.WindowTitle = "Emision de Lista de Precios Comparativo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{CargaLista.Lista} in " + DesdeLista.Text + " to " + HastaLista.Text
    Dos = " and {CargaLista.Linea} in " + DesdeLinea.Text + " to " + HastaLinea.Text
    Listado.GroupSelectionFormula = Uno + Dos
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CargaLista.Lista, CargaLista.Fecha, CargaLista.Titulo, CargaLista.Observaciones, CargaLista.Terminado, CargaLista.Precio, CargaLista.Descripcion, CargaLista.Linea, CargaLista.Empresa, CargaLista.CostoI, CargaLista.CostoII, CargaLista.FactorI, CargaLista.FactorII, CargaLista.Porce, CargaLista.FechaII, " _
                    + "Lineas.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.CargaLista CargaLista, " _
                    + DSQ + ".dbo.Lineas Lineas " _
                    + "Where " _
                    + "CargaLista.Linea = Lineas.Linea AND " _
                    + "CargaLista.Lista >= " + DesdeLista.Text + " AND " _
                    + "CargaLista.Lista <= " + HastaLista.Text + " AND " _
                    + "CargaLista.Linea >= " + DesdeLinea.Text + " AND " _
                    + "CargaLista.Linea <= " + HastaLinea.Text
                
    If Val(WEmpresa) = 1 Then
        Listado.ReportFileName = "WListaPreciosCompa.rpt"
            Else
        Listado.ReportFileName = "WListaPreciosCompaII.rpt"
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
    PrgListaPreciosCompa.Hide
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
        DesdeLista.SetFocus
    End If
End Sub

Private Sub DesdeLista_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaLista.SetFocus
    End If
End Sub

Private Sub HastaLista_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeLinea.SetFocus
    End If
End Sub

Private Sub DesdeLinea_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaLinea.SetFocus
    End If
End Sub

Private Sub HastaLinea_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
End Sub

Sub Form_Load()
    Fecha.Text = "  /  /    "
    FechaCompa.Text = "  /  /    "
    DesdeLista.Text = ""
    HastaLista.Text = ""
    DesdeLinea.Text = ""
    HastaLinea.Text = ""
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


