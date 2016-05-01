VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgModifColor 
   AutoRedraw      =   -1  'True
   Caption         =   "Modificacion de Pedidos de Colorantes / DY / DW  / DS"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11805
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11805
   Begin VB.CommandButton ImprePdfIII 
      Caption         =   "Cert. Ana"
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
      Left            =   10440
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton ImprePdfII 
      Caption         =   "Hoja Seg."
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
      Left            =   9240
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton ImprePdf 
      Caption         =   "Hoja y Cert."
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
      Left            =   7920
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton PlantaIV 
      Caption         =   "Planta IV"
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
      Left            =   5520
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton ImpreEti 
      Caption         =   "Etiquetas"
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
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin MSMask.MaskEdBox HastaFecha 
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
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
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
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
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6375
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11245
      _Version        =   327680
      Rows            =   4000
      Cols            =   11
      BackColor       =   16777215
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11160
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WImpreEtiDy.rpt"
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancela"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "PrgModifColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim XParam As String
Dim TotalPedidos As Integer
Dim WGraba As String
Dim WSaldo As Double

Dim ZZRuta As String
Dim ZZEstado As String

Dim ZZLote1 As String
Dim ZZLote2 As String
Dim ZZLote3 As String
Dim ZZLote4 As String
Dim ZZLote5 As String
Dim ZLote(100) As String
Dim ZZLote(100) As String

Dim ZZProceso As Integer

Dim WCantiLote1 As Double
Dim WCantiLote2  As Double
Dim WCantiLote3  As Double
Dim WCantiLote4  As Double
Dim WCantiLote5  As Double
Dim WCantiLote  As Double

Dim XLote(100, 30) As String
Dim WLote(100, 5) As String
Dim WCanti(100, 5) As String
Dim WEti(100, 5) As String
Dim WTipo(100, 5) As String

Dim ZImpreConcepto(200) As String

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgModifColor.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()

    ZImpreConcepto(0) = ""
    ZImpreConcepto(1) = "Error del Sistema"
    ZImpreConcepto(2) = "Varios"
    ZImpreConcepto(3) = "Problemas Vehiculos"
    ZImpreConcepto(4) = "Problemas Logistica"
    ZImpreConcepto(5) = "Problemas Recepcion Cliente"
    ZImpreConcepto(6) = "Varios"
    ZImpreConcepto(7) = "Corte de Luz"
    ZImpreConcepto(8) = "Pedido por el Cliente"
    ZImpreConcepto(9) = "Falta de Pago"
    ZImpreConcepto(10) = "Confirmacion Pedido Parcial"
    ZImpreConcepto(11) = "Envase"


    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 200
    Muestra.ColWidth(1) = 800
    Muestra.ColWidth(2) = 1200
    Muestra.ColWidth(3) = 800
    Muestra.ColWidth(4) = 2000
    Muestra.ColWidth(5) = 1200
    Muestra.ColWidth(6) = 800
    Muestra.ColWidth(7) = 1000
    Muestra.ColWidth(8) = 700
    Muestra.ColWidth(9) = 700
    Muestra.ColWidth(10) = 2000
    
    Muestra.ColAlignment(10) = flexAlignLeftCenter
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Muestra.Col = 7
    Muestra.Text = "Estado"
    
    Muestra.Col = 8
    Muestra.Text = "$ Pendiente"
    
    Rem DesdeFecha.Text = "  /  /    "
    Rem HastaFecha.Text = "  /  /    "
    
    DesdeFecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    HastaFecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Call DesdeFecha_Keypress(13)
    
    Rem DesdeFecha.SetFocus
    
End Sub

Private Sub ImpreEti_Click()

    On Error GoTo WError

    Da = 0
    With rstImpreEtiDy
        .Index = "Renglon"
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
    
    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    WLugar = 0
    
    For Ciclo = RowIni To Rowfin
    
        Muestra.Row = Ciclo
        Muestra.Col = 1
        WPedido = Muestra.Text
        
        spPedido = "ConsultaPedido1 " + "'" + WPedido + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
    
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Canti = !Cantidad
                    
                        If Canti > 0 Then
                        
                            WLugar = WLugar + 1
                        
                            WLote(WLugar, 1) = IIf(IsNull(rstPedido!lote1), "", rstPedido!lote1)
                            WLote(WLugar, 2) = IIf(IsNull(rstPedido!lote2), "", rstPedido!lote2)
                            WLote(WLugar, 3) = IIf(IsNull(rstPedido!lote3), "", rstPedido!lote3)
                            WLote(WLugar, 4) = IIf(IsNull(rstPedido!lote4), "", rstPedido!lote4)
                            WLote(WLugar, 5) = IIf(IsNull(rstPedido!lote5), "", rstPedido!lote5)
                            
                            WCanti(WLugar, 1) = IIf(IsNull(rstPedido!CantiLote1), "", rstPedido!CantiLote1)
                            WCanti(WLugar, 2) = IIf(IsNull(rstPedido!CantiLote2), "", rstPedido!CantiLote2)
                            WCanti(WLugar, 3) = IIf(IsNull(rstPedido!CantiLote3), "", rstPedido!CantiLote3)
                            WCanti(WLugar, 4) = IIf(IsNull(rstPedido!CantiLote4), "", rstPedido!CantiLote4)
                            WCanti(WLugar, 5) = IIf(IsNull(rstPedido!CantiLote5), "", rstPedido!CantiLote5)
                            
                            WEti(WLugar, 1) = IIf(IsNull(rstPedido!Eti1), "", rstPedido!Eti1)
                            WEti(WLugar, 2) = IIf(IsNull(rstPedido!Eti2), "", rstPedido!Eti2)
                            WEti(WLugar, 3) = IIf(IsNull(rstPedido!Eti3), "", rstPedido!Eti3)
                            WEti(WLugar, 4) = IIf(IsNull(rstPedido!Eti4), "", rstPedido!Eti4)
                            WEti(WLugar, 5) = IIf(IsNull(rstPedido!Eti5), "", rstPedido!Eti5)
                            
                            WTipo(WLugar, 1) = IIf(IsNull(rstPedido!tipo1), "", rstPedido!tipo1)
                            WTipo(WLugar, 2) = IIf(IsNull(rstPedido!tipo2), "", rstPedido!tipo2)
                            WTipo(WLugar, 3) = IIf(IsNull(rstPedido!tipo3), "", rstPedido!tipo3)
                            WTipo(WLugar, 4) = IIf(IsNull(rstPedido!tipo4), "", rstPedido!tipo4)
                            WTipo(WLugar, 5) = IIf(IsNull(rstPedido!tipo5), "", rstPedido!tipo5)
                            
                            XLote(WLugar, 1) = WPedido
                            XLote(WLugar, 2) = !Terminado
                            XLote(WLugar, 3) = !Cliente
                        
                        End If
        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
        End If
    
    Next Ciclo
    
    
    
    
    ZZPasa = 0
    ZZClave = ""
    
    For Da = 1 To WLugar
    
        WPedido = XLote(Da, 1)
        WTerminado = XLote(Da, 2)
        WCliente = XLote(Da, 3)
        
        If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DS" Or Left$(WTerminado, 2) = "DQ" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(WTerminado, 3) + Right$(WTerminado, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZZClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
                    ZZZClase = Trim(ZZZClase)
                    If ZZZClase <> "" Then
                        ZPasa = 1
                    End If
                    rstArticulo.Close
                End If
            
            Case Else
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZZZClase = ""
                    ZZZClase = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
                    ZZZClase = Trim(ZZZClase)
                    If ZZZClase <> "" Then
                        ZPasa = 1
                    End If
                    rstTerminado.Close
                End If
                
        End Select
        
        If ZPasa = 1 Then
            m$ = "Hay productos peligrosos, use la opcion de emision de etiquetas general"
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            Exit Sub
        End If
        
    Next Da
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Renglon = 0
    
    For Da = 1 To WLugar
    
        WPedido = XLote(Da, 1)
        WTerminado = XLote(Da, 2)
        WCliente = XLote(Da, 3)
        
        If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DS" Or Left$(WTerminado, 2) = "DQ" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        WDescripcion = ""
        WRazon = ""
        
        spCliente = "ConsultaCliente " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WRazon = rstCliente!Razon
            rstCliente.Close
        End If
        
        If Len(WRazon) > 25 Then
            For Cicla = 25 To 1 Step -1
                If Mid(WRazon, Cicla, 1) = Space(1) Then
                    WRazonII = Mid(WRazon, Cicla + 1, 25)
                    WRazon = Mid(WRazon, 1, Cicla)
                    Exit For
                End If
            Next Cicla
                Else
            WRazonII = ""
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(WTerminado, 3) + Right$(WTerminado, 7)
                
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                For Ciclo1 = 1 To 5
                    If Val(WLote(Da, Ciclo1)) = 0 Then
                        WLote(Da, Ciclo1) = ""
                            Else
                            
                        ZEntra = "N"
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Laudo = " + "'" + WLote(Da, Ciclo1) + "'"
                        ZSql = ZSql + " and Laudo.Articulo = " + "'" + WArti + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            WLote(Da, Ciclo1) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                            ZEntra = "S"
                            rstLaudo.Close
                        End If
                        
                        If ZEntra = "N" Then
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Lote = " + "'" + WLote(Da, Ciclo1) + "'"
                            ZSql = ZSql + " and Guia.Articulo = " + "'" + WArti + "'"
                            spMovguia = ZSql
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WLote(Da, Ciclo1) = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                                ZEntra = "S"
                                rstMovguia.Close
                            End If
                        End If
                        
                    End If
                Next Ciclo1
            
            Case Else
                ClavePrecios = WCliente + WTerminado
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    WDescripcion = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
                
        End Select
        
        For Ciclo1 = 1 To 5
            If Val(WCanti(Da, Ciclo1)) <> 0 Then
                WHasta = Val(WEti(Da, Ciclo1))
                WTipoeti = WTipo(Da, Ciclo1)
                For Ciclo2 = 1 To WHasta
                    Renglon = Renglon + 1
                    If WTipoeti = "T" Then
                        With rstImpreEtiDy
                            .Index = "Renglon"
                            .AddNew
                            !Renglon = Renglon
                            !pedido = WPedido
                            If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DS" Or Left$(WTerminado, 2) = "DQ" Then
                                !Codigo = Left$(Mid$(WTerminado, 6, 3) + Right$(WTerminado, 3) + WLote(Da, Ciclo1), 20)
                                    Else
                                !Codigo = Left$(Mid$(WTerminado, 4, 5) + Right$(WTerminado, 3) + WLote(Da, Ciclo1), 20)
                            End If
                            !Cliente = WCliente
                            !Descripcion = WDescripcion
                            !Razon = WRazon
                            !RazonII = WRazonII
                            !Lote = 0
                            !lote1 = WLote(Da, Ciclo1)
                            !Cantidad = Val(WCanti(Da, Ciclo1)) / Val(WEti(Da, Ciclo1))
                            .Update
                        End With
                            Else
                        With rstImpreEtiDy
                            .Index = "Renglon"
                            .AddNew
                            !Renglon = Renglon
                            !pedido = WPedido
                            !Codigo = ""
                            !Cliente = WCliente
                            !Descripcion = ""
                            !Razon = WRazon
                            !RazonII = WRazonII
                            !Lote = 0
                            !lote1 = ""
                            !Cantidad = 0
                            .Update
                        End With
                    End If
                Next Ciclo2
            End If
        Next Ciclo1
        
    Next Da

    Listado.WindowTitle = "Emision de Etiquetas de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.PrinterCopies = 1
    Listado.Action = 1
    
    Exit Sub

WError:
     Resume Next
    

End Sub

Private Sub ImprePdf_Click()
    ZZProceso = 1
    Call Proceso_ImprePdf
End Sub

Private Sub ImprePdfII_Click()
    ZZProceso = 2
    Call Proceso_ImprePdf
End Sub

Private Sub ImprePdfIII_Click()
    ZZProceso = 3
    Call Proceso_ImprePdf
End Sub

Private Sub Proceso_ImprePdf()

    On Error GoTo WError
    
    Dim ZZBusca(10000) As String
    Dim ZZLugarBusca As Integer

    ' Muestra los nombres en C:\ que representan directorios.
    ZZCodigoExe = "AcroRd32.exe"
    ZZPasaExe = ""
    
    Erase ZZBusca
    ZZLugarBusca = 1
    ZZBusca(ZZLugarBusca) = "c:\Archivos de programa\Adobe\"
    CicloBusca = 1
    ZZSalida = "N"
    
    Do
    
        MiRuta = ZZBusca(CicloBusca)
        MiNombre = Dir(MiRuta, vbDirectory) ' Recupera la primera entrada.
        Do While MiNombre <> "" ' Inicia el bucle.
                
            If MiNombre <> "." And MiNombre <> ".." Then
        
                If (GetAttr(MiRuta & MiNombre) And vbDirectory) = vbDirectory Then
                    
                    ZZLugarBusca = ZZLugarBusca + 1
                    ZZBusca(ZZLugarBusca) = MiRuta & MiNombre + "\"
                    
                        Else
                        
                    WEspacios = Len(ZZCodigoExe)
                    Da = Len(MiNombre) - WEspacios
                    If UCase(Trim(ZZCodigoExe)) = UCase(Trim(MiNombre)) Then
                        ZZPasaExe = MiRuta & MiNombre
                        ZZSalida = "S"
                        Exit Do
                    End If
                    
                End If
            
            End If
            MiNombre = Trim(UCase(Dir))  ' Obtiene siguiente entrada.
            
        Loop

        If CicloBusca = ZZLugarBusca Or ZZSalida = "S" Then
            Exit Do
                Else
            CicloBusca = CicloBusca + 1
        End If

    Loop
    
    
    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    WLugar = 0
    Erase WLote
    Erase XLote
    
    For Ciclo = RowIni To Rowfin
    
        Muestra.Row = Ciclo
        Muestra.Col = 1
        WPedido = Muestra.Text
        
        spPedido = "ConsultaPedido1 " + "'" + WPedido + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
    
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Canti = !Cantidad
                    
                        If Canti > 0 Then
                        
                            WCantiLote1 = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                            WCantiLote2 = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                            WCantiLote3 = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                            WCantiLote4 = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                            WCantiLote5 = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                            
                            WCantiLote = WCantiLote1 + WCantiLote2 + WCantiLote3 + WCantiLote4 + WCantiLote5
                            
                            If WCantiLote > 0 Then
                        
                                WLugar = WLugar + 1
                            
                                WLote(WLugar, 1) = IIf(IsNull(rstPedido!lote1), "", rstPedido!lote1)
                                WLote(WLugar, 2) = IIf(IsNull(rstPedido!lote2), "", rstPedido!lote2)
                                WLote(WLugar, 3) = IIf(IsNull(rstPedido!lote3), "", rstPedido!lote3)
                                WLote(WLugar, 4) = IIf(IsNull(rstPedido!lote4), "", rstPedido!lote4)
                                WLote(WLugar, 5) = IIf(IsNull(rstPedido!lote5), "", rstPedido!lote5)
                                
                                XLote(WLugar, 1) = WPedido
                                XLote(WLugar, 2) = !Terminado
                                XLote(WLugar, 3) = !Cliente
                            
                                    Else
                                                            
                                WCantiLote1 = IIf(IsNull(rstPedido!UltimoCantiLote1), "0", rstPedido!UltimoCantiLote1)
                                WCantiLote2 = IIf(IsNull(rstPedido!UltimoCantiLote2), "0", rstPedido!UltimoCantiLote2)
                                WCantiLote3 = IIf(IsNull(rstPedido!UltimoCantiLote3), "0", rstPedido!UltimoCantiLote3)
                                WCantiLote4 = IIf(IsNull(rstPedido!UltimoCantiLote4), "0", rstPedido!UltimoCantiLote4)
                                WCantiLote5 = IIf(IsNull(rstPedido!UltimoCantiLote5), "0", rstPedido!UltimoCantiLote5)
    
                                WCantiLote = WCantiLote1 + WCantiLote2 + WCantiLote3 + WCantiLote4 + WCantiLote5
                                
                                If WCantiLote > 0 Then
                            
                                    WLugar = WLugar + 1
                                
                                    WLote(WLugar, 1) = IIf(IsNull(rstPedido!Ultimolote1), "", rstPedido!Ultimolote1)
                                    WLote(WLugar, 2) = IIf(IsNull(rstPedido!Ultimolote2), "", rstPedido!Ultimolote2)
                                    WLote(WLugar, 3) = IIf(IsNull(rstPedido!Ultimolote3), "", rstPedido!Ultimolote3)
                                    WLote(WLugar, 4) = IIf(IsNull(rstPedido!Ultimolote4), "", rstPedido!Ultimolote4)
                                    WLote(WLugar, 5) = IIf(IsNull(rstPedido!Ultimolote5), "", rstPedido!Ultimolote5)
                                    
                                    XLote(WLugar, 1) = WPedido
                                    XLote(WLugar, 2) = !Terminado
                                    XLote(WLugar, 3) = !Cliente
                                    
                                End If
                                
                            End If
                        
                        End If
        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
        End If
    Next Ciclo
    
    Renglon = 0
    
    For Da = 1 To WLugar
    
        ZZPedido = XLote(Da, 1)
        ZZTerminado = Trim(XLote(Da, 2))
        ZZArti = Left$(ZZTerminado, 3) + Right$(ZZTerminado, 7)
        ZZCliente = Trim(XLote(Da, 3))
        ZLote(1) = WLote(Da, 1)
        ZLote(2) = WLote(Da, 2)
        ZLote(3) = WLote(Da, 3)
        ZLote(4) = WLote(Da, 4)
        ZLote(5) = WLote(Da, 5)
        
        If Left$(ZZTerminado, 2) = "DY" Or Left$(ZZTerminado, 2) = "DS" Or Left$(ZZTerminado, 2) = "DQ" Then
        
        
            ZZLugar = 0
            Erase ZZLote
            
            For Ciclo = 1 To 9 Step 2
                ZZLugar = ZZLugar + 1
                If Val(ZLote(Ciclo)) = 0 Then
                    ZZLote(ZZLugar) = ""
                        Else
                    ZEntra = "N"
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Laudo = " + "'" + ZLote(Ciclo) + "'"
                    ZSql = ZSql + " and Laudo.Articulo = " + "'" + ZZArti + "'"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        ZZLote(ZZLugar) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                        ZEntra = "S"
                        rstLaudo.Close
                    End If
                    
                    If ZEntra = "N" Then
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Guia"
                        ZSql = ZSql + " Where Guia.Lote = " + "'" + ZLote(Ciclo) + "'"
                        ZSql = ZSql + " and Guia.Articulo = " + "'" + ZZArti + "'"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            ZZLote(ZZLugar) = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                            ZEntra = "S"
                            rstMovguia.Close
                        End If
                    End If
                    
                End If
            Next Ciclo
            
            If ZZProceso = 1 Or ZZProceso = 2 Then
                ZZRuta = "Z:\MSDS" + ZZArti + ".PDF"
                ZZEstado = Dir(ZZRuta)
                ZZEstado = Trim(ZZEstado)
                If ZZEstado <> "" Then
                    RetVal = Shell(ZZPasaExe + " " + ZZRuta + " ", 3)
                End If
            End If
            
            If ZZProceso = 1 Or ZZProceso = 3 Then
                For CicloII = 1 To 5
                    If ZZLote(CicloII) <> "" Then
                        ZZRuta = "Z:\" + Trim(ZZLote(CicloII)) + ".PDF"
                        ZZEstado = Dir(ZZRuta)
                        ZZEstado = Trim(ZZEstado)
                        If ZZEstado <> "" Then
                            RetVal = Shell(ZZPasaExe + " " + ZZRuta + " ", 3)
                        End If
                    End If
                Next CicloII
            End If
            
            
                Else
                
                
            If ZZProceso = 1 Or ZZProceso = 2 Then
                ZZRuta = "Z:\" + ZZTerminado + ".PDF"
                ZZEstado = Dir(ZZRuta)
                ZZEstado = Trim(ZZEstado)
                If ZZEstado <> "" Then
                    RetVal = Shell(ZZPasaExe + " " + ZZRuta + " ", 3)
                End If
            End If
            
            If ZZProceso = 1 Or ZZProceso = 3 Then
                For CicloII = 1 To 5
                    If ZLote(CicloII) <> "" Then
                        ZZRuta = "Z:\" + Trim(ZLote(CicloII)) + ".PDF"
                        ZZEstado = Dir(ZZRuta)
                        ZZEstado = Trim(ZZEstado)
                        If ZZEstado <> "" Then
                            RetVal = Shell(ZZPasaExe + " " + ZZRuta + " ", 3)
                        End If
                    End If
                Next CicloII
            End If
                
            
        End If
        
    Next Da
    
    Exit Sub

WError:
     Resume Next

End Sub

Private Sub PlantaIV_Click()
    Muestra.Col = 1
    WXPed = Muestra.Text
    
    PrgModifColor.Hide
    Unload Me
    PrgModPedZona.Show
End Sub

Private Sub Proceso_Click()

    WSalida = "N"
    
    Call Limpia_Vector
    
    Renglon = 0
    WSaldo = 0
    
    WAno = Right$(DesdeFecha.Text, 4)
    WMes = Mid$(DesdeFecha.Text, 4, 2)
    WDia = Left$(DesdeFecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia
    
    Pasa = 0
    pedido = ""
    Fecha = "  /  /    "
    Cliente = ""
    Razon = ""
    FEntrega = "  /  /    "
    Tipo = 0
    Importe = 0
    Estado = ""
    
    If Val(WEmpresa) <> 8 Then
    
        ZSql = ""
        ZSql = ZSql + "Select Pedido.TipoPedido, Pedido.FechaOrd, Pedido.Autorizo, Pedido.Clave, Pedido.Cantidad, Pedido.Facturado, Pedido.Terminado, Pedido.Clave, Pedido.Pedido, Pedido.Fecha, Pedido.Cliente, Pedido.FecEntrega, Pedido.TipoPed, Pedido.Impresion, Pedido.MarcaFactura, Pedido.Precio, Pedido.FechaInicial, Pedido.MarcaAutorizacion"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.TipoPedido = 1"
        ZSql = ZSql + " and Pedido.FechaOrd < " + "'" + WDesde + "'"
        ZSql = ZSql + " and ((Pedido.Autorizo <> " + "'" + "X" + "') or (Pedido.Cantidad-Pedido.Facturado) <> 0)"
        ZSql = ZSql + " Order by Clave"
        spPedido = ZSql
        
            Else
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + "Select Pedido.TipoPedido, Pedido.FechaOrd, Pedido.Autorizo, Pedido.Clave, Pedido.Cantidad, Pedido.Facturado, Pedido.Terminado, Pedido.Clave, Pedido.Pedido, Pedido.Fecha, Pedido.Cliente, Pedido.FecEntrega, Pedido.TipoPed, Pedido.Impresion, Pedido.MarcaFactura, Pedido.Precio, Pedido.FechaInicial, Pedido.MarcaAutorizacion"
        ZSql = ZSql + " Where Pedido.Terminado >= 'DS-00000-000'"
        ZSql = ZSql + " and Pedido.Terminado <= 'DS-99999-999'"
        ZSql = ZSql + " and Pedido.FechaOrd < " + "'" + WDesde + "'"
        ZSql = ZSql + " and ((Pedido.Autorizo <> " + "'" + "X" + "') or (Pedido.Cantidad-Pedido.Facturado) <> 0)"
        ZSql = ZSql + " Order by Clave"
        spPedido = ZSql
        
    End If
        
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        
        With rstPedido
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                        
                    Rem If rstPedido!Pedido = 352216 Then Stop
                    
                    
                    If WDesde > rstPedido!FechaOrd Then
                    
                    WSaldo = rstPedido!Cantidad - rstPedido!Facturado
                    Call Redondeo(WSaldo)
                    
                    If rstPedido!Autorizo <> "X" Or WSaldo <> 0 Then
                
                        If Pasa = 0 Then
                            Corte = rstPedido!pedido
                            Fecha = rstPedido!Fecha
                            FechaInicial = rstPedido!FechaInicial
                            Cliente = rstPedido!Cliente
                            FEntrega = rstPedido!FecEntrega
                            Tipo = rstPedido!Tipoped
                            Importe = 0
                            Estado = rstPedido!Autorizo
                            Impresa = rstPedido!Impresion
                            ZZMarca = IIf(IsNull(rstPedido!MarcaFactura), "0", rstPedido!MarcaFactura)
                            ZZMarcaAutorizacion = IIf(IsNull(rstPedido!MarcaAutorizacion), "", rstPedido!MarcaAutorizacion)
                            Autoriza = rstPedido!Autorizo
                            Pasa = 1
                        End If
                        
                        If Corte <> rstPedido!pedido Then
                        
                            Renglon = Renglon + 1
                            Muestra.Row = Renglon
                        
                            If Autoriza <> "X" Then
                 
                                Muestra.Col = 1
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 2
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 3
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 4
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 5
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 6
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 7
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 8
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 9
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 10
                                Muestra.CellBackColor = &HFF&
                                
                                
                            End If
                            
                            
                            Rem ZZFechaDia = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                            Rem If ZZFechaDia = Fecha Then
                            Rem     If Fecha <> FechaInicial Then
                                
                            If ZZMarcaAutorizacion = "S" Then
                
                               Muestra.Col = 1
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 2
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 3
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 4
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 5
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 6
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 7
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 8
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 9
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 10
                               Muestra.CellBackColor = &H80FF80

                            
                            End If
                            
                            If ZZMarca = "1" Then
                                Muestra.Col = 0
                                Muestra.Text = "F"
                                    Else
                                Muestra.Col = 0
                                Muestra.Text = ""
                            End If
    
                            Muestra.Col = 1
                            Muestra.Text = Pusing("######", Str$(Corte))
                            
                            Muestra.Col = 2
                            Muestra.Text = Fecha
                    
                            Muestra.Col = 3
                            Muestra.Text = Cliente
                            
                            Muestra.Col = 5
                            Muestra.Text = FEntrega
                            
                            Select Case Tipo
                                Case 0
                                    Muestra.Col = 6
                                    Muestra.Text = "Normal"
                                Case 1
                                    Muestra.Col = 6
                                    Muestra.Text = "A Fecha"
                                Case 2
                                    Muestra.Col = 6
                                    Muestra.Text = "Fec.Lim."
                                Case 3
                                    Muestra.Col = 6
                                    Muestra.Text = "Urgente"
                                Case 4
                                    Muestra.Col = 6
                                    Muestra.Text = "Ret.Cli"
                                Case 5
                                    Muestra.Col = 6
                                    Muestra.Text = "Muestra"
                                Case Else
                                    Muestra.Col = 6
                                    Muestra.Text = ""
                            End Select
                            
                            If Impresa = "N" Then
                                Muestra.Col = 7
                                Muestra.Text = "A verificar"
                                    Else
                                Muestra.Col = 7
                                Muestra.Text = "Impreso"
                            End If
                            
                            Muestra.Col = 8
                            If Importe <> 0 Then
                                Muestra.Text = "Pend."
                                    Else
                                Muestra.Text = ""
                            End If
                            
                            Corte = rstPedido!pedido
                            Fecha = rstPedido!Fecha
                            Cliente = rstPedido!Cliente
                            FEntrega = rstPedido!FecEntrega
                            Tipo = rstPedido!Tipoped
                            Importe = 0
                            Estado = rstPedido!Autorizo
                            Impresa = rstPedido!Impresion
                            ZZMarca = IIf(IsNull(rstPedido!MarcaFactura), "0", rstPedido!MarcaFactura)
                            ZZMarcaAutorizacion = IIf(IsNull(rstPedido!MarcaAutorizacion), "", rstPedido!MarcaAutorizacion)
                            
                            Autoriza = rstPedido!Autorizo
                            Pasa = 1
                        
                        End If
                        
                        Importe = Importe + ((rstPedido!Cantidad - rstPedido!Facturado) * rstPedido!Precio)
                        
                        
                    End If
                    
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
        
        If Pasa <> 0 Then
                        
            Renglon = Renglon + 1
        
            Muestra.Row = Renglon
                        
            If Autoriza <> "X" Then
                 
                Muestra.Col = 1
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 2
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 3
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 4
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 5
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 6
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 7
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 8
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 9
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 10
                Muestra.CellBackColor = &HFF&
                                
            End If
                
            If ZZMarcaAutorizacion = "S" Then

               Muestra.Col = 1
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 2
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 3
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 4
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 5
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 6
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 7
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 8
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 9
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 10
               Muestra.CellBackColor = &H80FF80

            
            End If
            
            If ZZMarca = "1" Then
                Muestra.Col = 0
                Muestra.Text = "F"
                    Else
                Muestra.Col = 0
                Muestra.Text = ""
            End If
                            
            Muestra.Col = 1
            Muestra.Text = Pusing("######", Str$(Corte))
                            
            Muestra.Col = 2
            Muestra.Text = Fecha
                    
            Muestra.Col = 3
            Muestra.Text = Cliente
                            
            Muestra.Col = 5
            Muestra.Text = FEntrega
                            
            Select Case Tipo
                Case 0
                    Muestra.Col = 6
                    Muestra.Text = "Normal"
                Case 1
                    Muestra.Col = 6
                    Muestra.Text = "A Fecha"
                Case 2
                    Muestra.Col = 6
                    Muestra.Text = "Fec.Lim"
                Case 3
                    Muestra.Col = 6
                    Muestra.Text = "Urgente"
                Case 4
                    Muestra.Col = 6
                    Muestra.Text = "Ret.Cli"
                Case 5
                    Muestra.Col = 6
                    Muestra.Text = "Muestra"
                Case Else
                    Muestra.Col = 6
                    Muestra.Text = ""
            End Select
                            
            If Impresa = "N" Then
                Muestra.Col = 7
                Muestra.Text = "A verificar"
                    Else
                Muestra.Col = 7
                Muestra.Text = "Impreso"
            End If
            
            Muestra.Col = 8
            If Importe <> 0 Then
                Muestra.Text = "Pend."
                    Else
                Muestra.Text = ""
            End If
                            
        End If
        
        rstPedido.Close
    
    End If
    
    
    
    
    Pasa = 0
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select Pedido.TipoPedido, Pedido.FechaOrd, Pedido.Autorizo, Pedido.Clave, Pedido.Cantidad, Pedido.Facturado, Pedido.Terminado, Pedido.Clave, Pedido.Pedido, Pedido.Fecha, Pedido.Cliente, Pedido.FecEntrega, Pedido.TipoPed, Pedido.Impresion, Pedido.MarcaFactura, Pedido.Precio, Pedido.FechaInicial, Pedido.MarcaAutorizacion"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Pedido.FechaOrd <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Pedido.Autorizo = " + "'" + "X" + "'"
    ZSql = ZSql + " Order by Clave"
    spPedido = ZSql
    
    Rem     XParam = "'" + WDesde + "','" _
    rem         + WHasta + "'"
    Rem spPedido = "ListaPedidoFecha " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        
        With rstPedido
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                    
                    If rstPedido!Autorizo = "X" Then
                    
                    If WDesde <= rstPedido!FechaOrd And WHasta >= rstPedido!FechaOrd Then
                    
                    WPasa = "N"
                    If Val(WEmpresa) <> 8 Then
                        If rstPedido!TipoPedido = 1 Then
                            WPasa = "S"
                        End If
                                Else
                        If Left$(rstPedido!Terminado, 2) = "DS" Then
                            WPasa = "S"
                        End If
                    End If
                    
                    If WPasa = "S" Then
                    
                        If Pasa = 0 Then
                            Corte = rstPedido!pedido
                            Fecha = rstPedido!Fecha
                            FechaInicial = rstPedido!FechaInicial
                            Cliente = rstPedido!Cliente
                            FEntrega = rstPedido!FecEntrega
                            Tipo = rstPedido!Tipoped
                            Importe = 0
                            Estado = rstPedido!Autorizo
                            Impresa = rstPedido!Impresion
                            ZZMarca = IIf(IsNull(rstPedido!MarcaFactura), "0", rstPedido!MarcaFactura)
                            ZZMarcaAutorizacion = IIf(IsNull(rstPedido!MarcaAutorizacion), "", rstPedido!MarcaAutorizacion)
                            Autoriza = rstPedido!Autorizo
                            Pasa = 1
                        End If
                        
                        If Corte <> rstPedido!pedido Then
                        
                            Renglon = Renglon + 1
                
                            Muestra.Row = Renglon
                        
                            If Autoriza <> "X" Then
                 
                                Muestra.Col = 1
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 2
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 3
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 4
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 5
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 6
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 7
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 8
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 9
                                Muestra.CellBackColor = &HFF&
                 
                                Muestra.Col = 10
                                Muestra.CellBackColor = &HFF&
                                
                            End If
                                
                            If ZZMarcaAutorizacion = "S" Then
                
                               Muestra.Col = 1
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 2
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 3
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 4
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 5
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 6
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 7
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 8
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 9
                               Muestra.CellBackColor = &H80FF80
                
                               Muestra.Col = 10
                               Muestra.CellBackColor = &H80FF80

                            
                            End If
                            
                            If ZZMarca = "1" Then
                                Muestra.Col = 0
                                Muestra.Text = "F"
                                    Else
                                Muestra.Col = 0
                                Muestra.Text = ""
                            End If
                            
                            Muestra.Col = 1
                            Muestra.Text = Pusing("######", Str$(Corte))
                            
                            Muestra.Col = 2
                            Muestra.Text = Fecha
                    
                            Muestra.Col = 3
                            Muestra.Text = Cliente
                            
                            Muestra.Col = 5
                            Muestra.Text = FEntrega
                            
                            Select Case Tipo
                                Case 0
                                    Muestra.Col = 6
                                    Muestra.Text = "Normal"
                                Case 1
                                    Muestra.Col = 6
                                    Muestra.Text = "A Fecha"
                                Case 2
                                    Muestra.Col = 6
                                    Muestra.Text = "Fec.Lim"
                                Case 3
                                    Muestra.Col = 6
                                    Muestra.Text = "Urgente"
                                Case 4
                                    Muestra.Col = 6
                                    Muestra.Text = "Ret.Cli."
                                Case 5
                                    Muestra.Col = 6
                                    Muestra.Text = "Muestra"
                                Case Else
                                    Muestra.Col = 6
                                    Muestra.Text = ""
                            End Select
                            
                            If Impresa = "N" Then
                                Muestra.Col = 7
                                Muestra.Text = "A verificar"
                                    Else
                                Muestra.Col = 7
                                Muestra.Text = "Impreso"
                            End If
                            
                            Muestra.Col = 8
                            If Importe <> 0 Then
                                Muestra.Text = "Pend."
                                    Else
                                Muestra.Text = ""
                            End If
                            
                            Corte = rstPedido!pedido
                            Fecha = rstPedido!Fecha
                            FechaInicial = rstPedido!FechaInicial
                            Cliente = rstPedido!Cliente
                            FEntrega = rstPedido!FecEntrega
                            Tipo = rstPedido!Tipoped
                            Importe = 0
                            Estado = rstPedido!Autorizo
                            Impresa = rstPedido!Impresion
                            ZZMarca = IIf(IsNull(rstPedido!MarcaFactura), "0", rstPedido!MarcaFactura)
                            ZZMarcaAutorizacion = IIf(IsNull(rstPedido!MarcaAutorizacion), "", rstPedido!MarcaAutorizacion)
                            Autoriza = rstPedido!Autorizo
                            Pasa = 1
                            
                        
                        End If
                        
                        Importe = Importe + ((rstPedido!Cantidad - rstPedido!Facturado) * rstPedido!Precio)
                        
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
        
        If Pasa <> 0 Then
                        
            Renglon = Renglon + 1
    
            Muestra.Row = Renglon
                        
            If Autoriza <> "X" Then
    
                Muestra.Col = 1
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 2
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 3
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 4
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 5
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 6
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 7
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 8
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 9
                Muestra.CellBackColor = &HFF&
    
                Muestra.Col = 10
                Muestra.CellBackColor = &HFF&
                
            End If
                                
            If ZZMarcaAutorizacion = "S" Then

               Muestra.Col = 1
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 2
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 3
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 4
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 5
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 6
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 7
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 8
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 9
               Muestra.CellBackColor = &H80FF80

               Muestra.Col = 10
               Muestra.CellBackColor = &H80FF80

            
            End If
                            
            
            If ZZMarca = "1" Then
                Muestra.Col = 0
                Muestra.Text = "F"
                    Else
                Muestra.Col = 0
                Muestra.Text = ""
            End If
                            
            Muestra.Col = 1
            Muestra.Text = Pusing("######", Str$(Corte))
                            
            Muestra.Col = 2
            Muestra.Text = Fecha
                    
            Muestra.Col = 3
            Muestra.Text = Cliente
                            
            Muestra.Col = 5
            Muestra.Text = FEntrega
                            
            Select Case Tipo
                Case 0
                    Muestra.Col = 6
                    Muestra.Text = "Normal"
                Case 1
                    Muestra.Col = 6
                    Muestra.Text = "A Fecha"
                Case 2
                    Muestra.Col = 6
                    Muestra.Text = "Fec.Lim."
                Case 3
                    Muestra.Col = 6
                    Muestra.Text = "Urgente"
                Case 4
                    Muestra.Col = 6
                    Muestra.Text = "Ret.Cli"
                Case 5
                    Muestra.Col = 6
                    Muestra.Text = "Muestra"
                Case Else
                    Muestra.Col = 6
                    Muestra.Text = ""
            End Select
                            
            If Impresa = "N" Then
                Muestra.Col = 7
                Muestra.Text = "A verificar"
                    Else
                Muestra.Col = 7
                Muestra.Text = "Impreso"
            End If
            
            Muestra.Col = 8
            If Importe <> 0 Then
                Muestra.Text = "Pend."
                    Else
                Muestra.Text = ""
            End If
                            
        End If
        
        rstPedido.Close
    
    End If
    
    For dada = 1 To Renglon
    
        WCliente = Muestra.TextMatrix(dada, 3)
    
        spCliente = "ConsultaCliente " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Muestra.TextMatrix(dada, 4) = rstCliente!Razon
            rstCliente.Close
        End If
        
        WPedido = Muestra.TextMatrix(dada, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select Pedido.Pedido, Pedido.Cantidad, Pedido.Cantidad"
        ZSql = ZSql + " FROM Atraso"
        ZSql = ZSql + " Where Atraso.Pedido = " + "'" + WPedido + "'"
        spAtraso = ZSql
        Rem spPedido = "ConsultaPedido1 " + "'" + WPedido + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
    
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                        Canti = !Cantidad - !Facturado
                        If Canti <= 0 Then
                            Muestra.TextMatrix(dada, 9) = "Parcial"
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select Atraso.Pedido, Atraso.Fecha, Atraso.Problema, Atraso.Concepto, Atraso.Articulo"
        ZSql = ZSql + " FROM Atraso"
        ZSql = ZSql + " Where Atraso.Pedido = " + "'" + WPedido + "'"
        spAtraso = ZSql
        Set rstatraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
        If rstatraso.RecordCount > 0 Then
            ZZImpre = Mid$(rstatraso!Fecha, 1, 5) + " " + Trim(rstatraso!problema) + " " + Trim(rstatraso!Articulo) + " " + ZImpreConcepto(rstatraso!concepto)
            Muestra.TextMatrix(dada, 10) = ZZImpre
            rstatraso.Close
        End If
        
    Next dada
    
    TotalPedidos = Renglon
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Muestra.Row = 1
    Muestra.Col = 1
    Muestra.TopRow = 1
    
    Rem Muestra.SetFocus

End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Muestra.Col = 7
    Muestra.Text = "Estado"
    
    Muestra.Col = 8
    Muestra.Text = ""
    
    Muestra.Col = 9
    Muestra.Text = ""
    
    Muestra.Col = 10
    Muestra.Text = "Aviso"
    
End Sub

Private Sub Muestra_DblClick()

    Muestra.Col = 1
    WXPed = Muestra.Text
    
    PrgModifColor.Hide
    Unload Me
    PrgModPedCol.Show
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_ImpreEtiDy
    Rem If ProcesoActivate = 1 Then
    Rem     Call DesdeFecha_Keypress(13)
    Rem End If
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            If HastaFecha.Text = "  /  /    " Or HastaFecha.Text = "00/00/0000" Then
                HastaFecha.Text = DesdeFecha.Text
            End If
            Call Proceso_Click
                Else
            DesdeFecha.SetFocus
        End If
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            Call Proceso_Click
                Else
            HastaFecha.SetFocus
        End If
    End If
End Sub

