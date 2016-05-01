VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsumoArt 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Consumo de Materia Prima"
   ClientHeight    =   6810
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   7545
   LinkTopic       =   "Form2"
   ScaleHeight     =   6810
   ScaleWidth      =   7545
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   18
         Top             =   1920
         Width           =   2775
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1920
         TabIndex        =   12
         Top             =   600
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Top             =   240
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
         Mask            =   "AA-###-###"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   2520
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
         Left            =   1080
         TabIndex        =   10
         Top             =   2520
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
         Left            =   3840
         TabIndex        =   9
         Top             =   600
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
         Left            =   3840
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   1920
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   300
         Left            =   1920
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label5 
         Caption         =   "Tipo"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta Fecha"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Desde Fecha"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
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
         Left            =   120
         TabIndex        =   7
         Top             =   600
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wConsumoArt.rpt"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "Consumoart.frx":0000
      Left            =   120
      List            =   "Consumoart.frx":0007
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   7215
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
Attribute VB_Name = "PrgConsumoArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim XParam As String
Dim Vector(10000, 7) As String
Dim ZVector(5000, 3) As String
Private XLote(100, 7) As String
Private WDescripcion As String
Private WSaldo As Double
Dim Empe(100, 10) As String

Private Sub Acepta_Click()

    On Error GoTo WError
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    WDesde = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    WHasta = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)

    Da = 0
    With rstFichaMat
        .Index = "Articulo"
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
        Case 0, 1
    
            If Tipo.ListIndex = 1 Then

                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " Lista = " + " '" + "N" + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
                ZLugar = 0
                Erase ZVector
    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Hoja"
                ZSql = ZSql + " Where Fechaingord >= " + "'" + WDesde + "'"
                ZSql = ZSql + " and Fechaingord <= " + "'" + WHasta + "'"
                ZSql = ZSql + " and Renglon = 2"
                ZSql = ZSql + " Order by Clave"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    With rstHoja
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                ZLugar = ZLugar + 1
                                ZVector(ZLugar, 1) = Str$(rstHoja!Hoja)
                                ZVector(ZLugar, 2) = ""
                                ZVector(ZLugar, 3) = rstHoja!Producto
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstHoja.Close
                End If

                For Ciclo = 1 To ZLugar
    
                    ZHoja = ZVector(Ciclo, 1)
                    ZReal = Val(ZVector(Ciclo, 2))
                    ZProducto = ZVector(Ciclo, 3)
        
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " Lista = " + " '" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
                Next Ciclo
    
            End If
    

            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Fechaingord >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and Fechaingord <= " + "'" + WHasta + "'"
            ZSql = ZSql + " and Articulo >= " + "'" + Desde.Text + "'"
            ZSql = ZSql + " and Articulo <= " + "'" + Hasta.Text + "'"
            ZSql = ZSql + " Order by Clave"
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
                
                        XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                        If XFec >= WDesde And XFec <= WHasta Then
                
                            If !Tipo = "M" Then
                    
                                If Tipo.ListIndex = 0 Or !lista = "S" Then
                                
                                    WArticulo = rstHoja!Articulo
                                    WCantidad = rstHoja!Cantidad
                                    WFecha = rstHoja!Fecha
                                    WHoja = rstHoja!Hoja
                                    WLote = ""
                                    WSaldo = "0"
                
                                    With rstFichaMat
                                        .AddNew
                                        !Articulo = WArticulo
                                        !Fecha = WFecha
                                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                        !Tipo = 0
                                        !Numero = WHoja
                                        !Inicial = 0
                                        !Entrada = 0
                                        !Salida = WCantidad
                                        !Observaciones = ""
                                        !Lista1 = "Hoja"
                                        !Lista2 = ""
                                        !Lote = 0
                                        !Saldo = WSaldo
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
                rstHoja.Close
            End If
            
        Case Else
            XEmpresa = WEmpresa
            If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
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
                XHasta = 7
                    Else
                Empe(1, 1) = "0002"
                Empe(1, 2) = "Empresa02"
                Empe(2, 1) = "0004"
                Empe(2, 2) = "Empresa04"
                Empe(3, 1) = "0008"
                Empe(3, 2) = "Empresa08"
                Empe(4, 1) = "0009"
                Empe(4, 2) = "Empresa09"
                XHasta = 4
            End If
    
            For a = 1 To XHasta
    
                WEmpresa = Empe(a, 1)
                txtOdbc = Empe(a, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
                LugarVector = 0
                Erase Vector
    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Hoja"
                ZSql = ZSql + " Where Fechaingord >= " + "'" + WDesde + "'"
                ZSql = ZSql + " and Fechaingord <= " + "'" + WHasta + "'"
                ZSql = ZSql + " and Articulo >= " + "'" + Desde.Text + "'"
                ZSql = ZSql + " and Articulo <= " + "'" + Hasta.Text + "'"
                ZSql = ZSql + " Order by Clave"
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
                
                                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                                If XFec >= WDesde And XFec <= WHasta Then
                
                                    If !Tipo = "M" Then
                    
                                        LugarVector = LugarVector + 1
                                    
                                        Vector(LugarVector, 1) = rstHoja!Articulo
                                        Vector(LugarVector, 2) = Str$(rstHoja!Cantidad)
                                        Vector(LugarVector, 3) = rstHoja!Fecha
                                        Vector(LugarVector, 4) = Str$(rstHoja!Hoja)
                                        Vector(LugarVector, 5) = ""
                                        Vector(LugarVector, 6) = "0"
                    
                                    End If
                    
                                End If
                
                                .MoveNext
                
                                If .EOF = True Then
                                    Exit Do
                                End If
                
                            Loop
                        End If
                    End With
                    rstHoja.Close
                End If
                
                Call Conecta_Empresa
                
                For Ciclo = 1 To LugarVector
                
                    WArticulo = Vector(Ciclo, 1)
                    WCantidad = Vector(Ciclo, 2)
                    WFecha = Vector(Ciclo, 3)
                    WHoja = Vector(Ciclo, 4)
                    WLote = Vector(Ciclo, 5)
                    WSaldo = Vector(Ciclo, 6)
                
                    With rstFichaMat
                        .AddNew
                        !Articulo = WArticulo
                        !Fecha = WFecha
                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        !Tipo = 0
                        !Numero = WHoja
                        !Inicial = 0
                        !Entrada = 0
                        !Salida = Val(WCantidad)
                        !Observaciones = ""
                        !Lista1 = "Hoja"
                        !Lista2 = ""
                        !Lote = 0
                        !Saldo = WSaldo
                        .Update
                    End With
                        
                Next Ciclo
                
            Next a
        
    End Select
    
    
    Rem XParam = "'" + Desde.Text + "','" _
    Rem              + Hasta.Text + "'"
    Rem
    Rem spMovvar = "ListaMovvarArticuloDesdeHasta" + XParam
    Rem Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstMovvar.RecordCount > 0 Then
    Rem
    Rem     With rstMovvar
    Rem
    Rem         .MoveFirst
    Rem
    Rem         If .NoMatch = False Then
    Rem         Do
    Rem
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem
    Rem             If rstMovvar!FechaOrd >= WDesde And rstMovvar!FechaOrd <= WHasta Then
    Rem
    Rem                 If rstMovvar!Movi = "S" Then
    Rem
    Rem                     If !Tipo = "M" Then
    Rem
    Rem                         WArticulo = rstMovvar!Articulo
    Rem                         WCantidad = rstMovvar!Cantidad
    Rem                         WFecha = rstMovvar!Fecha
    Rem                         WCodigo = rstMovvar!Codigo
    Rem                         WMovi = rstMovvar!Movi
    Rem                         WTipomov = Val(rstMovvar!Tipomov)
    Rem                         WObservaciones = rstMovvar!Observaciones
    Rem                         WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
    Rem                         WSaldo = "0"
    Rem
    Rem                         With rstFichaMat
    Rem
    Rem                             .AddNew
    Rem                             !Articulo = WArticulo
    Rem                             !Fecha = WFecha
    Rem                             !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    Rem                             !Tipo = 0
    Rem                             !Numero = WCodigo
    Rem                             !Inicial = 0
    Rem                             !Entrada = 0
    Rem                             !Salida = WCantidad
    Rem                             !Observaciones = WObservaciones
    Rem                             If WTipomov = 0 Or WTipomov = 1 Then
    Rem                                 !Lista1 = "Mov.Var"
    Rem                                     Else
    Rem                                !Lista1 = "Guia In"
    Rem                             End If
    Rem                             !Lista2 = ""
    Rem                             !Lote = WLote
    Rem                             !Saldo = WSaldo
    Rem                             .Update
    Rem                         End With
    Rem                     End If
    Rem                 End If
    Rem             End If
    Rem
    Rem             .MoveNext
    Rem
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem
    Rem         Loop
    Rem         End If
    Rem     End With
    Rem     rstMovvar.Close
    Rem End If
    
    
    
    
    
    
    
    
    
    Rem XParam = "'" + Desde.Text + "','" _
    rem              + Hasta.Text + "'"
    Rem
    Rem spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
    Rem Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstMovguia.RecordCount > 0 Then
    Rem
    Rem     With rstMovguia
    Rem
    Rem         .MoveFirst
    Rem
    Rem         If .NoMatch = False Then
    Rem         Do
    Rem
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem
    Rem             Rem If rstMovguia!Codigo = 20123 Then Stop
    Rem
    Rem             If rstMovguia!FechaOrd >= WDesde And rstMovguia!FechaOrd <= WHasta Then
    Rem
    Rem                 If rstMovguia!Tipo = "M" Then
    Rem
    Rem                     If rstMovguia!Movi = "S" Then
    Rem
    Rem                         WCanti = IIf(IsNull(rstMovguia!Cantidadant), "0", rstMovguia!Cantidadant)
    Rem
    Rem                         If WCanti <> 0 Then
    Rem                             WCantidad = WCanti
    Rem                                 Else
    Rem                             WCantidad = rstMovguia!Cantidad
    Rem                         End If
    Rem                         WArticulo = rstMovguia!Articulo
    Rem                         Rem WCantidad = rstMovguia!Cantidad
    Rem                         WFecha = rstMovguia!Fecha
    Rem                         WCodigo = rstMovguia!Codigo
    Rem                         WMovi = rstMovguia!Movi
    Rem                         WDestino = rstMovguia!Destino
    Rem                         WTipomov = rstMovguia!Tipomov
    Rem                         Rem WObservaciones = rstMovvar!Observaciones
    Rem                         WLote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
    Rem                         WSaldo = "0"
    Rem
    Rem                         Select Case WDestino
    Rem                             Case 1
    Rem                                 WObservaciones = "Envio a Surfactan"
    Rem                             Case 2
    Rem                                 WObservaciones = "Envio a Pellital"
    Rem                             Case 3
    Rem                                 WObservaciones = "Envio a Surfactan II"
    Rem                             Case 4
    Rem                                 WObservaciones = "Envio a Pellital II"
    Rem                             Case 5
    Rem                                 WObservaciones = "Envio a Surfactan III"
    Rem                             Case 6
    Rem                                 WObservaciones = "Envio a Surfactan IV"
    Rem                             Case 7
    Rem                                 WObservaciones = "Envio a Surfactan V"
    Rem                             Case 8
    Rem                                 WObservaciones = "Envio a Pellital V"
    Rem                             Case Else
    Rem                                 WObservaciones = ""
    Rem                         End Select
    Rem
    Rem                         With rstFichaMat
    Rem                             .AddNew
    Rem                             !Articulo = WArticulo
    Rem                             !Fecha = WFecha
    Rem                             !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    Rem                             !Tipo = 0
    Rem                             !Numero = WCodigo
    Rem                             !Inicial = 0
    Rem                             !Entrada = 0
    Rem                             !Salida = WCantidad
    Rem                             !Observaciones = WObservaciones
    Rem                             If !Numero > 900000 Then
    Rem                                 !Lista1 = "Prestamo"
    Rem                                 !Numero = !Numero - 900000
    Rem                                     Else
    Rem                                 !Lista1 = "Guia In"
    Rem                             End If
    Rem                             !Lista2 = ""
    Rem                             !Lote = WLote
    Rem                             !Saldo = WSaldo
    Rem                             .Update
    Rem                         End With
    Rem                     End If
    Rem
    Rem                 End If
    Rem
    Rem             End If
    Rem
    Rem             .MoveNext
    Rem
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem
    Rem         Loop
    Rem         End If
    Rem     End With
    Rem     rstMovguia.Close
    Rem End If
    
    
    
    
    
    
    
    Rem XParam = "'" + Desde.Text + "','" _
    Rem              + Hasta.Text + "'"
    Rem
    Rem spMovlab = "ListaMovlabArticuloDesdeHasta" + XParam
    Rem Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstMovlab.RecordCount > 0 Then
    Rem
    Rem     With rstMovlab
    Rem
    Rem         .MoveFirst
    Rem
    Rem         If .NoMatch = False Then
    Rem         Do
    Rem
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem
    Rem             If rstMovlab!FechaOrd >= WDesde And rstMovlab!FechaOrd <= WHasta Then
    Rem
    Rem                 If rstMovlab!Movi = "S" Then
    Rem
    Rem                     If rstMovlab!Tipo = "M" Then
    Rem
    Rem                         WArticulo = rstMovlab!Articulo
    Rem                         WCantidad = rstMovlab!Cantidad
    Rem                         WFecha = rstMovlab!Fecha
    Rem                         WCodigo = rstMovlab!Codigo
    Rem                        WMovi = rstMovlab!Movi
    Rem                         WTipomov = rstMovlab!Tipomov
    Rem                         WObservaciones = rstMovlab!Observaciones
    Rem                         WLote = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
    Rem                         WSaldo = "0"
    Rem
    Rem                         With rstFichaMat
    Rem
    Rem                             .AddNew
    Rem                             !Articulo = WArticulo
    Rem                             !Fecha = WFecha
    Rem                             !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    Rem                             !Tipo = 0
    Rem                             !Numero = WCodigo
    Rem                             !Inicial = 0
    Rem                             !Entrada = 0
    Rem                             !Salida = WCantidad
    Rem                             !Observaciones = WObservaciones
    Rem                             !Lista1 = "Mov.Lab"
    Rem                             !Lista2 = ""
    Rem                             !Lote = WLote
    Rem                             !Saldo = WSaldo
    Rem                             .Update
    Rem                         End With
    Rem                     End If
    Rem                 End If
    Rem             End If
    Rem
    Rem             .MoveNext
    Rem
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem
    Rem        Loop
    Rem         End If
    Rem     End With
    Rem     rstMovlab.Close
    Rem End If
    
    
    
    
    
    
    Da = 0
    With rstFichaMat
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WArticulo = !Articulo
                WDescripcion = ""
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                !Descripcion = WDescripcion
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    





    Listado.WindowTitle = "Listado de Consumo de Materias Primas"
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

    With rstEmpresa
        .Close
    End With
    With rstFichaMat
        .Close
    End With
    
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgConsumoArt.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaMat
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        DesdeFecha.SetFocus
    End If
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Produccion"
    Tipo.AddItem "Consolidado"
    
    Tipo.ListIndex = 0

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgConsumoArt.Caption = "Listado de Consumo de Materia Prima :  " + !Nombre
        End If
    End With
    
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstArticulo!Codigo
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstArticulo.Close
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WArticulo = WIndice.List(Indice)
    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
            Desde.Text = rstArticulo!Codigo
            Hasta.Text = rstArticulo!Codigo
                Else
            Desde.Text = WArticulo
            Hasta.Text = WArticulo
    End If
    Desde.SetFocus
    
    
End Sub






