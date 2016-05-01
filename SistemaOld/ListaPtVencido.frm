VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaPtVencido 
   Caption         =   "Listado de Productos Terminados Vencidos"
   ClientHeight    =   3600
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3600
   ScaleWidth      =   8145
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Archivos de programa\DevStudio\VB\listaptvencido.rpt"
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   6015
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2280
         TabIndex        =   0
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
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
         Left            =   2760
         TabIndex        =   9
         Top             =   2400
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
         Left            =   1200
         TabIndex        =   8
         Top             =   2400
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
         Height          =   375
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
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
         Height          =   375
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
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
         Left            =   600
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
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
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaPtVencido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VectorVto(5000, 9) As String
Dim XMes As String
Dim XAno As String
Dim Empe(12, 10) As String
Dim ZZMeses As String
Dim zzrevalida As String
Dim zzmesesrevalida2 As String


Dim XXEmpresa As String

Private Sub Acepta_Click()

    Erase Empe
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        ZHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        ZHasta = 4
    End If
    
    ZSql = "DELETE ListaPtVencido"
    spListaPtVencido = ZSql
    Set rstListaPtVencido = db.OpenRecordset(spListaPtVencido, dbOpenSnapshot, dbSQLPassThrough)
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    XEmpresa = WEmpresa
    
    Erase VectorVto
    Renglon = 0
    
    For A = 1 To ZHasta
        
        WEmpresa = Empe(A, 1)
        txtOdbc = Empe(A, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZSql = ""
        ZSql = ZSql + "Select Hoja.Hoja, Hoja.Marca, Hoja.Real, Hoja.Teorico, Hoja.Renglon, Hoja.Producto, Hoja.Revalida, Hoja.MesesRevalida, Hoja.Fecha, Hoja.FechaRevalida, Hoja.Saldo"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Marca <> 'X'"
        ZSql = ZSql + " and Hoja.Saldo <> 0"
        ZSql = ZSql + " and Hoja.Renglon = 1"
        ZSql = ZSql + " and Hoja.Producto >= " + "'" + Desde.Text + "'"
        ZSql = ZSql + " and Hoja.Producto <= " + "'" + Hasta.Text + "'"
        ZSql = ZSql + " and (Hoja.MarcaVencida = " + "'" + "S" + "'"
        ZSql = ZSql + " or Hoja.MarcaVencida = " + "'" + "V" + "'" + ")"
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
                
                    Renglon = Renglon + 1
                    VectorVto(Renglon, 1) = rstHoja!Producto
                    VectorVto(Renglon, 2) = rstHoja!Hoja
                    VectorVto(Renglon, 3) = "H"
                    VectorVto(Renglon, 4) = rstHoja!Fecha
                    VectorVto(Renglon, 5) = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                    VectorVto(Renglon, 6) = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                    VectorVto(Renglon, 7) = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                    VectorVto(Renglon, 8) = Str$(rstHoja!Saldo)
                    VectorVto(Renglon, 9) = WEmpresa
                    
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
        ZSql = ZSql + "Select Guia.Clave, Guia.Codigo, Guia.Marca, Guia.Saldo, Guia.Lote, Guia.Terminado"
        ZSql = ZSql + " FROM Guia"
        ZSql = ZSql + " Where Guia.Marca <> 'X'"
        ZSql = ZSql + " and Guia.Saldo <> 0"
        ZSql = ZSql + " and Guia.Tipo = 'T'"
        ZSql = ZSql + " and Guia.Terminado >= " + "'" + Desde.Text + "'"
        ZSql = ZSql + " and Guia.Terminado <= " + "'" + Hasta.Text + "'"
        ZSql = ZSql + " and (Guia.MarcaVencida = " + "'" + "S" + "'"
        ZSql = ZSql + " or Guia.MarcaVencida = " + "'" + "V" + "'" + ")"
        
        spGuia = ZSql
        Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
        If rstGuia.RecordCount > 0 Then

            With rstGuia
        
                .MoveFirst
                
                If .NoMatch = False Then
                
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                    Renglon = Renglon + 1
                    VectorVto(Renglon, 1) = rstGuia!Terminado
                    VectorVto(Renglon, 2) = rstGuia!Lote
                    VectorVto(Renglon, 3) = "G"
                    VectorVto(Renglon, 4) = ""
                    VectorVto(Renglon, 5) = ""
                    VectorVto(Renglon, 6) = ""
                    VectorVto(Renglon, 7) = ""
                    VectorVto(Renglon, 8) = Str$(rstGuia!Saldo)
                    VectorVto(Renglon, 9) = WEmpresa
                    
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                
                End If
                
            End With
            
            rstGuia.Close
    
        End If
        
    Next A
    
    Call Conecta_Empresa
        
    For Da = 1 To Renglon
    
        ZZTipoMov = VectorVto(Da, 3)
        
        Select Case ZZTipoMov
            Case "H"
                ZZProducto = VectorVto(Da, 1)
                ZZHoja = VectorVto(Da, 2)
                ZZTipoMov = VectorVto(Da, 3)
                ZZFecha = VectorVto(Da, 4)
                zzrevalida = VectorVto(Da, 5)
                zzmesesrevalida = VectorVto(Da, 6)
                ZZFechaRevalida = VectorVto(Da, 7)
                ZZSaldo = VectorVto(Da, 8)
                ZZEmpresa = VectorVto(Da, 9)
                ZZMeses = ""
                
                spTerminado = "ConsultaTerminado " + "'" + ZZProducto + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZZDescripcion = rstTerminado!Descripcion
                    ZZMeses = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                    rstTerminado.Close
                End If
                
                ZSql = ""
                ZSql = ZSql + ""
                ZSql = ZSql + "INSERT INTO ListaPtVencido ("
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "Lote ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Planta ,"
                ZSql = ZSql + "Meses ,"
                ZSql = ZSql + "Revalida ,"
                ZSql = ZSql + "FechaRevalida ,"
                ZSql = ZSql + "MesesRevalida ,"
                ZSql = ZSql + "DesEmpresa )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZZProducto + "',"
                ZSql = ZSql + "'" + ZZDescripcion + "',"
                ZSql = ZSql + "'" + ZZHoja + "',"
                ZSql = ZSql + "'" + ZZFecha + "',"
                ZSql = ZSql + "'" + ZZSaldo + "',"
                ZSql = ZSql + "'" + ZZEmpresa + "',"
                ZSql = ZSql + "'" + ZZMeses + "',"
                ZSql = ZSql + "'" + zzrevalida + "',"
                ZSql = ZSql + "'" + ZZFechaRevalida + "',"
                ZSql = ZSql + "'" + zzmesesrevalida + "',"
                ZSql = ZSql + "'" + "" + "')"

                spListaPtVencido = ZSql
                Set rstListaPtVencido = db.OpenRecordset(spListaPtVencido, dbOpenSnapshot, dbSQLPassThrough)
                
                
            Case "G"
                ZZProducto = VectorVto(Da, 1)
                ZZHoja = VectorVto(Da, 2)
                ZZTipoMov = VectorVto(Da, 3)
                zzrevalida = ""
                zzmesesrevalida = ""
                ZZFechaRevalida = ""
                ZZFecha = ""
                ZZSaldo = VectorVto(Da, 8)
                ZZEmpresa = VectorVto(Da, 9)
                ZZMeses = ""
                
                spTerminado = "ConsultaTerminado " + "'" + ZZProducto + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZZDescripcion = rstTerminado!Descripcion
                    ZZMeses = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                    rstTerminado.Close
                End If
                
                XEmpresa = WEmpresa
                
                For CiclaEmpresa = 1 To ZHasta
    
                    WEmpresa = Empe(CiclaEmpresa, 1)
                    txtOdbc = Empe(CiclaEmpresa, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Hoja"
                    ZSql = ZSql + " Where Hoja.Hoja = " + "'" + ZZHoja + "'"
                    ZSql = ZSql + " and Hoja.Producto = " + "'" + ZZProducto + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        zzrevalida = Int(IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida))
                        zzmesesrevalida2 = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                        ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                        ZZFecha = rstHoja!Fecha
                        rstHoja.Close
                        Exit For
                    End If
                    
                Next CiclaEmpresa
                
                Call Conecta_Empresa
                
                ZSql = ""
                ZSql = ZSql + ""
                ZSql = ZSql + "INSERT INTO ListaPtVencido ("
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "Lote ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Planta ,"
                ZSql = ZSql + "Meses ,"
                ZSql = ZSql + "Revalida ,"
                ZSql = ZSql + "FechaRevalida ,"
                ZSql = ZSql + "MesesRevalida ,"
                ZSql = ZSql + "DesEmpresa )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZZProducto + "',"
                ZSql = ZSql + "'" + ZZDescripcion + "',"
                ZSql = ZSql + "'" + ZZHoja + "',"
                ZSql = ZSql + "'" + ZZFecha + "',"
                ZSql = ZSql + "'" + ZZSaldo + "',"
                ZSql = ZSql + "'" + ZZEmpresa + "',"
                ZSql = ZSql + "'" + ZZMeses + "',"
                Rem    ZSql = ZSql + zzrevalida2 + ","
                ZSql = ZSql + "'" + zzrevalida + "',"
                ZSql = ZSql + "'" + ZZFechaRevalida + "',"
                ZSql = ZSql + "'" + zzmesesrevalida2 + "',"
                ZSql = ZSql + "'" + "" + "')"

                spListaPtVencido = ZSql
                Set rstListaPtVencido = db.OpenRecordset(spListaPtVencido, dbOpenSnapshot, dbSQLPassThrough)
            
            Case Else
            
        End Select
    Next Da

    
    
    
    
    
    
    
    

    Listado.WindowTitle = "Verificacion de Vencimientos de Producto Terminado"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT ListaPtVencido.Codigo, ListaPtVencido.Descripcion, ListaPtVencido.Lote, ListaPtVencido.Fecha, ListaPtVencido.Cantidad, ListaPtVencido.Planta, ListaPtVencido.Meses, ListaPtVencido.Revalida, ListaPtVencido.FechaRevalida, ListaPtVencido.MesesRevalida " _
                + "From " _
                + DSQ + ".dbo.ListaPtVencido ListaPtVencido " _
                + "Where " _
                + "ListaPtVencido.Codigo >= 'AA-00000-000' AND " _
                + "ListaPtVencido.Codigo <= 'ZZ-99999-999'"
    
    Listado.Connect = Connect()
    
    Rem Listado.GroupSelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.ReportFileName = "Listaptvencido.rpt"
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgListaEspecifPt.Hide
    Unload Me
    Menu.Show
End Sub


Private Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub





