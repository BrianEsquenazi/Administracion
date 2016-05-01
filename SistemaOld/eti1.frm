VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEti1 
   Caption         =   "Impresion de Etiquetas"
   ClientHeight    =   5205
   ClientLeft      =   1080
   ClientTop       =   1920
   ClientWidth     =   9900
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   9900
   Begin VB.Frame PantaDirEntrega 
      Caption         =   "Seleccion de Lugar de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   8535
      Begin VB.ListBox ListaDirEntrega 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   8295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control de Etiquetas"
      Height          =   3375
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   8535
      Begin VB.TextBox Descripcion 
         Height          =   285
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   19
         Text            =   " "
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox Etiquetas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   16
         Text            =   " "
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Cantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   15
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Lote 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   14
         Text            =   "  "
         Top             =   1920
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Terminado 
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox Cliente 
         Height          =   285
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   4680
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label DesProducto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label6 
         Caption         =   "Descripcion"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label DesCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad de Etiquetas"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Lote"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Producto Terminado"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Weti1.rpt"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "eti1.frx":0000
      Left            =   840
      List            =   "eti1.frx":0007
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgEti1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLote As String
Private WCantidad As String
Private WImpreadi As String
Private WClase As String
Private WIntervencion As String
Private WNaciones As String
Private WEmbalaje As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim XParam As String

Dim WDirentrega As String
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String



Private Sub Acepta_Click()

    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
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



    WCodigo = 1
    
    With rstEtiqueta
        For Da = 1 To Val(Etiquetas)
            .Index = "Codigo"
            .AddNew
            !Codigo = Da
            WLote = Lote.Text
            Call Ceros(WLote, 5)
            WCantidad = Cantidad.Text
            Call Ceros(WCantidad, 4)
            !Terminado = Terminado.Text
            !Lote = WLote
            !Cliente = Cliente.Text
            !Cantidad = Val(Cantidad.Text)
            !Nombre = Descripcion.Text
            !Impre1 = Mid$(Terminado.Text, 4, 5) + Right$(Terminado.Text, 3) + Space$(1) + WLote + Space$(1) + WCantidad
            WRazon = ""
            Rem WDirEntrega = ""
            spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                WRazon = rstClientes!Razon
                Rem WDirEntrega = rstClientes!DirEntrega
                rstClientes.Close
            End If
            !Razon = WRazon
            !DirEntrega = WDirentrega
            Rem !Clase = ""
            Rem !Intervencion = ""
            Rem !Naciones = ""
            Rem !Embalaje = ""
            .Update
        Next Da
    End With

    Listado.WindowTitle = "Emision de Etiquetas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Da = 0

    If Len(Descripcion.Text) > 20 Then
        For Da = 25 To 1 Step -1
            If Mid$(Descripcion.Text, Da, 1) <> Space$(1) Then
                Exit For
            End If
        Next Da
    End If
    
        If WImpreadi <> "S" Then
            If Da > 20 Then
                Listado.ReportFileName = "eti10.rpt"
                    Else
                Listado.ReportFileName = "eti1.rpt"
            End If
                Else
            If Da > 20 Then
                Listado.ReportFileName = "eti110.rpt"
                    Else
                Listado.ReportFileName = "eti101.rpt"
            End If
        End If
    
    Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
   
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Listado.DataFiles(1) = ""
   
    Listado.Destination = 1
    Listado.PrinterCopies = 1
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstEtiqueta
        .Close
    End With
    PrgEti1.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Etiqueta
End Sub

Sub Form_Load()

    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Lote.Text = ""
    Descripcion.Text = ""
    Cantidad.Text = ""
    Etiquetas.Text = ""
    
    DesCliente.Caption = ""
    DesProducto.Caption = ""
    
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cliente.Text <> "" Then
            spClientes = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                Cliente.Text = rstClientes!Cliente
                DesCliente.Caption = rstClientes!Razon
                
                Erase ZDirEntrega
                
                ZDirEntrega(1) = rstClientes!DirEntrega
                ZDirEntrega(2) = Trim(IIf(IsNull(rstClientes!DirEntregaII), "", rstClientes!DirEntregaII))
                ZDirEntrega(3) = Trim(IIf(IsNull(rstClientes!DirEntregaIII), "", rstClientes!DirEntregaIII))
                ZDirEntrega(4) = Trim(IIf(IsNull(rstClientes!DirEntregaIV), "", rstClientes!DirEntregaIV))
                ZDirEntrega(5) = Trim(IIf(IsNull(rstClientes!DirEntregaV), "", rstClientes!DirEntregaV))
                
                WDirentrega = ""
                CantiLugarEntrega = 0
                For CicloDirEntrega = 1 To 5
                    If ZDirEntrega(CicloDirEntrega) <> "" Then
                        WDirentrega = ZDirEntrega(CicloDirEntrega)
                        ZLugarDirEntrega = CicloDirEntrega
                        CantiLugarEntrega = CantiLugarEntrega + 1
                    End If
                Next CicloDirEntrega
                
                If CantiLugarEntrega > 1 Then
                    ListaDirEntrega.Clear
                    For CicloDirEntrega = 1 To 5
                        If ZDirEntrega(CicloDirEntrega) <> "" Then
                            ListaDirEntrega.AddItem ZDirEntrega(CicloDirEntrega)
                        End If
                    Next CicloDirEntrega
                    PantaDirEntrega.Top = 840
                    PantaDirEntrega.Visible = True
                    ListaDirEntrega.SetFocus
                        Else
                    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
                End If
                
                rstClientes.Close
                
                    Else
                Cliente.Text = Claveven$
                Cliente.SetFocus
            End If
        End If
        Terminado.SetFocus
    End If
End Sub

Private Sub ListaDirEntrega_Click()
    ZLugarDirEntrega = ListaDirEntrega.ListIndex + 1
    WDirentrega = ZDirEntrega(ZLugarDirEntrega)
    ZDescriDirEntrega = ZDirEntrega(ZLugarDirEntrega)
    PantaDirEntrega.Visible = False
    Terminado.SetFocus
End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Terminado.Text <> "" Then
            spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Terminado.Text = rstTerminado!Codigo
                DesProducto.Caption = rstTerminado!Descripcion
                rstTerminado.Close
                    
                XParam = "'" + Cliente.Text + "','" _
                             + Terminado.Text + "'"
                spPrecios = "ConsultaPrecios " + "'" + Terminado.Text + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Descripcion.Text = Left$(rstPrecios!Descripcion, 25)
                        Else
                    Descripcion.Text = ""
                End If
                    Else
                Terminado.Text = Claveven$
                Terminado.SetFocus
            End If
        End If
        Descripcion.SetFocus
    End If
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Lote.SetFocus
    End If
End Sub

Private Sub Lote_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cantidad.SetFocus
    End If
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Etiquetas.SetFocus
    End If
End Sub

Private Sub Etiquetas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.SetFocus
    End If
End Sub

