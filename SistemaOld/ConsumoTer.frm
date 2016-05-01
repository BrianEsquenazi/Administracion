VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsumoTer 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Consumo de Productos"
   ClientHeight    =   6180
   ClientLeft      =   2085
   ClientTop       =   1500
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6180
   ScaleWidth      =   8085
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   5295
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
         Left            =   1800
         TabIndex        =   17
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3840
         TabIndex        =   16
         Top             =   450
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1800
         TabIndex        =   11
         Top             =   600
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
         Left            =   1800
         TabIndex        =   0
         Top             =   240
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
         Left            =   2640
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   960
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
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   1800
         TabIndex        =   12
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
         Left            =   1800
         TabIndex        =   13
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
         TabIndex        =   18
         Top             =   1920
         Width           =   1215
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
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Prod.Term."
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
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Prod.Term."
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
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5760
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WConsumoTer.rpt"
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
      Left            =   6240
      TabIndex        =   3
      Top             =   120
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
      Height          =   2160
      ItemData        =   "ConsumoTer.frx":0000
      Left            =   240
      List            =   "ConsumoTer.frx":0007
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgConsumoTer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovguia As String
Dim rstMovguia As Recordset
Dim spMovvar As String
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstEntdev As Recordset
Dim spEntdev As String
Dim XParam As String
Dim Vector(10000, 7) As String
Dim ZVector(5000, 3) As String
Private XLote(100, 7) As String
Private WCantidad As Double
Private WSaldo As Double

Private Sub Acepta_Click()

    On Error GoTo WError
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    WDesde = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    WHasta = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)

    Da = 0
    With rstFichaTer
        .Index = "Terminado"
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
    ZSql = ZSql + " and Terminado >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Terminado <= " + "'" + Hasta.Text + "'"
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
                
                    If rstHoja!Tipo = "T" Then
                    
                    If Tipo.ListIndex = 0 Or !lista = "S" Then
                
                        WTerminado = rstHoja!Terminado
                        WCantidad = rstHoja!Cantidad * -1
                        WFecha = rstHoja!Fecha
                        WHoja = rstHoja!Hoja
                        WLote = ""
                
                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
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
                            !Saldo = 0
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
    

GoTo dada
    
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spHoja = "ListaHojaProductoDesdeHasta" + XParam
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
                
                    If Val(rstHoja!Renglon) = 1 Then
                    
                        WCantAnt = IIf(IsNull(rstHoja!realant), "0", rstHoja!realant)
                        WCanti = IIf(IsNull(rstHoja!Real), "0", rstHoja!Real)
                        If WCantAnt > 0 Then
                            WCanti = WCantAnt
                        End If
                
                        WTerminado = rstHoja!Producto
                        WCantidad = WCanti
                        WFecha = rstHoja!Fecha
                        WHoja = rstHoja!Hoja
                        WLote = ""
                
                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
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
                            !Saldo = 0
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
        rstHoja.Close
    End If
    
    
dada:
    
    
    Rem XParam = "'" + Desde.Text + "','" _
    Rem              + Hasta.Text + "'"
    Rem spMovvar = "ListaMovvarTerminadoDesdeHasta" + XParam
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
    Rem                     If rstMovvar!Tipo = "T" Then
    Rem
    Rem                         WTerminado = rstMovvar!Terminado
    Rem                         WCantidad = rstMovvar!Cantidad
    Rem                         WFecha = rstMovvar!Fecha
    Rem                         WCodigo = rstMovvar!Codigo
    Rem                         WMovi = rstMovvar!Movi
    Rem                         WTipomov = Val(rstMovvar!Tipomov)
    Rem                         WObservaciones = rstMovvar!Observaciones
    Rem                         WLote = rstMovvar!Lote
    Rem
    Rem                         With rstFichaTer
    Rem
    Rem                             .AddNew
    Rem                             !Terminado = WTerminado
    Rem                             !Fecha = WFecha
    Rem                             !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    Rem                             !Tipo = 0
    Rem                             !Numero = WCodigo
    Rem                             !Inicial = 0
    Rem                             If WMovi = "E" Then
    Rem                                 !Entrada = WCantidad
    Rem                                 !Salida = 0
    Rem                                     Else
    Rem                                 !Entrada = 0
    Rem                                 !Salida = WCantidad
    Rem                             End If
    Rem                             !Observaciones = ""
    Rem                             If WTipomov = 1 Or WTipomov = 2 Then
    Rem                                 !Lista1 = "Mov.Var"
    Rem                                     Else
    Rem                                 !Lista1 = "Guia In"
    Rem                             End If
    Rem                             !Lista2 = Left$(WObservaciones, 30)
    Rem                             !Lote = WLote
    Rem                             !Saldo = 0
    Rem                             .Update
    Rem                         End With
    Rem
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
    Rem              + Hasta.Text + "'"
    Rem spMovlab = "ListaMovlabTerminadoDesdeHasta" + XParam
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
    Rem                     If rstMovlab!Tipo = "T" Then
    Rem
    Rem                         WTerminado = rstMovlab!Terminado
    Rem                         WCantidad = rstMovlab!Cantidad
    Rem                         WFecha = rstMovlab!Fecha
    Rem                         WCodigo = rstMovlab!Codigo
    Rem                         WMovi = rstMovlab!Movi
    Rem                         WTipomov = rstMovlab!Tipomov
    Rem                         WObservaciones = rstMovlab!Observaciones
    Rem                         WLote = rstMovlab!Lote
    Rem
    Rem                         With rstFichaTer
    Rem
    Rem                             .AddNew
    Rem                             !Terminado = WTerminado
    Rem                             !Fecha = WFecha
    Rem                             !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    Rem                             !Tipo = 0
    Rem                             !Numero = WCodigo
    Rem                             !Inicial = 0
    Rem                             If WMovi = "E" Then
    Rem                                 !Entrada = WCantidad
    Rem                                 !Salida = 0
    Rem                                     Else
    Rem                                 !Entrada = 0
    Rem                                 !Salida = WCantidad
    Rem                             End If
    Rem                             !Observaciones = ""
    Rem                             !Lista1 = "Mov.Lab"
    Rem                             !Lista2 = Left$(WObservaciones, 30)
    Rem                             !Lote = WLote
    Rem                             !Saldo = 0
    Rem                             .Update
    Rem                         End With
    Rem                     End If
    Rem                 End If
    Rem             End If
    Rem
    Rem             .MoveNext
    Rem
    Rem            If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem
    Rem         Loop
    Rem         End If
    Rem
    Rem     End With
    Rem     rstMovlab.Close
    Rem End If
    
    Da = 0
    With rstFichaTer
        .Index = "Terminado"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WTerminado = !Terminado
                WDescripcion = ""
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WDescripcion = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                !Descripcion = WDescripcion
                
                If Left$(!Lista1, 8) = "Rem.Con." Then
                    spCliente = "ConsultaCliente " + "'" + Left$(!Observaciones, 6) + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        !Lista2 = Left$(rstCliente!Razon, 30)
                        rstCliente.Close
                    End If
                End If
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Consumo de Producto Terminado "
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{FichaTer.Terminado} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    
    Listado.Action = 1
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    With rstFichaTer
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgConsumoTer.Hide
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
    OPEN_FILE_FichaTer
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
    
    Tipo.ListIndex = 0
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgConsumoTer.Caption = "Listado de Consumo de Productos :  " + !Nombre
        End If
    End With
    
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstTerminado
        .MoveFirst
            Do
            If .EOF = False Then
                IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstTerminado!Codigo
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstTerminado.Close
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Desde.Text = rstTerminado!Codigo
        Hasta.Text = rstTerminado!Codigo
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub


