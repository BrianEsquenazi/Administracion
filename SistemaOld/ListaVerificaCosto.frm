VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaVerificaCosto 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Verificacion de Cambios de Costo Standard"
   ClientHeight    =   3165
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   4815
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox DesdeFec 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Left            =   3360
         TabIndex        =   4
         Top             =   360
         Width           =   1215
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
         Left            =   3360
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6720
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPedpen.rpt"
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
      Left            =   6480
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaVerificaCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String

Dim XParam As String
Dim EmpresaActual As String
Dim XIndice As Integer
Dim WPedido As String
Dim WGraba As String
Dim ZMateria As String
Dim ZLugar As Integer

Dim ZTerminado(10000) As String
Dim ZArticulo(10000) As String
Dim Auxiliar(10000, 3) As String
Dim ZProducto As String
Dim ZCostoI As Double
Dim ZCostoII As Double
Dim Otro(1000) As String

Dim ZVector(1000) As String


Dim WDireccionEmail As String
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
Dim AllPath As String

Dim ZZPasaTo As String
Dim ZZPasaCC As String
Dim ZZPasaBody As String
Dim ZZPasaFile As String



Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Acepta_Click()

    ver = WEmpresa


    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    ZZLugar = 0
    Erase ZVector
    
    Rem by nan
    If WEmpresa = "0001" Then
        cic = 2
        Opcion = 1
            Else
        cic = 1
        Opcion = 0
    End If
    Rem fin by nan
    
    Rem by nan
    Rem agrg by nan
    For i = 1 To cic
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.OrdFechaCosto >= " + "'" + WDesde + "'"
        ZSql = ZSql + " and Articulo.OrdFechaCosto <= " + "'" + WHasta + "'"
        spArticulo = ZSql
                
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            With rstArticulo
        
                .MoveFirst
                If .NoMatch = False Then
                    Do
                    
                        ZZLugar = ZZLugar + 1
                        
                        ZVector(ZZLugar) = rstArticulo!Codigo
                        
                        .MoveNext
                    
                        If .EOF = True Then
                            Exit Do
                        End If
                    
                    Loop
                End If
            
            End With
            rstArticulo.Close
        End If
        
        
        
        
        
        
        
        
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Terminado SET "
        ZSql = ZSql + "CodigoEmpresa = " + "'" + "1" + "',"
        ZSql = ZSql + "Precio1 = " + "'" + "0" + "',"
        ZSql = ZSql + "Precio2 = " + "'" + "0" + "',"
        ZSql = ZSql + "MarcaPrecio = " + "'" + "" + "'"
                         
        spTerminado = ZSql
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
        Erase ZTerminado
        ZLugar = 0
    
        For Ciclo = 1 To ZZLugar
            
            ZMateria = ZVector(Ciclo)
            ZArticulo(Ciclo) = ZVector(Ciclo)
            
            If ZMateria <> "" Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Composicion"
                ZSql = ZSql + " Where Composicion.Articulo1 = " + "'" + ZMateria + "'"
                spComposicion = ZSql
                Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                If rstComposicion.RecordCount > 0 Then
                    With rstComposicion
                        .MoveFirst
                        Do
                            If .EOF = False Then
                            
                                Entra = "S"
                                
                                For CicloII = 1 To ZLugar
                                    If ZTerminado(CicloII) = rstComposicion!Terminado Then
                                        Entra = "N"
                                        Exit For
                                    End If
                                Next CicloII
                                
                                If Entra = "S" Then
                                    ZLugar = ZLugar + 1
                                    ZTerminado(ZLugar) = rstComposicion!Terminado
                                End If
                                    
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstComposicion.Close
                End If
                
            End If
    
        Next Ciclo
        
        XLugar = ZLugar
        
        For Ciclo = 1 To XLugar
        
            ZZTerminado = ZTerminado(Ciclo)
            
            If ZZTerminado <> "" Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Composicion"
                ZSql = ZSql + " Where Composicion.Articulo2 = " + "'" + ZZTerminado + "'"
                spComposicion = ZSql
                Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                If rstComposicion.RecordCount > 0 Then
                    With rstComposicion
                        .MoveFirst
                        Do
                            If .EOF = False Then
                            
                                Entra = "S"
                                
                                For CicloII = 1 To ZLugar
                                    If ZTerminado(CicloII) = rstComposicion!Terminado Then
                                        Entra = "N"
                                        Exit For
                                    End If
                                Next CicloII
                                
                                If Entra = "S" Then
                                    ZLugar = ZLugar + 1
                                    XLugar = XLugar + 1
                                    ZTerminado(ZLugar) = rstComposicion!Terminado
                                End If
                                    
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstComposicion.Close
                End If
                
            End If
            
        Next Ciclo
            
            
            
        
        For Ciclo = 1 To ZLugar
        
            ZProducto = ZTerminado(Ciclo)
            
            Call Calcula_Costo(ZProducto, ZCostoI)
            
            Call Calcula_CostoII(ZProducto, ZCostoII)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Terminado SET "
            ZSql = ZSql + "Precio1 = " + "'" + Str$(ZCostoI) + "',"
            ZSql = ZSql + "Precio2 = " + "'" + Str$(ZCostoII) + "',"
            ZSql = ZSql + "MarcaPrecio = " + "'" + "S" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + ZProducto + "'"
                         
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        Next Ciclo
        
        
        
        
        
        
        
        
        
        
        
        
        Listado.WindowTitle = "Verificacion de Costo de Productos Terminados"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
    
        Listado.GroupSelectionFormula = "{Terminado.MarcaPrecio} in " + Chr$(34) + "S" + Chr$(34)
        Listado.SelectionFormula = "{Terminado.MarcaPrecio} in " + Chr$(34) + "S" + Chr$(34)
        
        Listado.Destination = 0
        Listado.ReportFileName = "VerificaCosto.rpt"
        
        DbConnect = db.Connect()
        DSQ = getDatabase(DbConnect)
        
        Listado.SQLQuery = "SELECT Terminado.Codigo, Terminado.Descripcion, Terminado.Linea, Terminado.Precio1, Terminado.Precio2, Terminado.MarcaPrecio, " _
                + "Lineas.Nombre, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.Terminado Terminado, " _
                + DSQ + ".dbo.Lineas Lineas, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "Terminado.Linea = Lineas.Linea AND " _
                + "Terminado.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Terminado.MarcaPrecio = 'S'"
        
        
        ZZEstado = Dir("c:\VerificaCosto\VerificaCosto.xls")
        If ZZEstado <> "" Then
            Kill "c:\VerificaCosto\VerificaCosto.xls"
        End If
        
        ZZEstado = Dir("c:\VerificaCosto\VerificaCostopelli.xls")
        If ZZEstado <> "" Then
            Kill "c:\VerificaCosto\VerificaCostopelli.xls"
        End If
        
        Listado.Destination = 2
        Listado.PrintFileType = crptExcel50
        
        If WEmpresa = "0001" Then
            Listado.PrintFileName = "c:\VerificaCosto\VerificaCosto.xls"
                Else
            Listado.PrintFileName = "c:\VerificaCosto\VerificaCostopelli.xls"
        End If
        
        Listado.Connect = Connect()
        Listado.Action = 1
        
        
        
        
        
    
        
        
        
        
        
        
        
        
        
        
        Listado.WindowTitle = "Cambio de Costo Std. de Materias Primas"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
    
        Listado.GroupSelectionFormula = "{Articulo.OrdFechaCosto} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
        Listado.SelectionFormula = "{Articulo.OrdFechaCosto} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
        
        Listado.Destination = 0
        Listado.ReportFileName = "ListaCambioCosto.rpt"
        
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        
        Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Costo2, Articulo.FechaCosto, Articulo.OrdFechaCosto " _
                + "From " _
                + DSQ + ".dbo.Articulo Articulo " _
                + "Where " _
                + "Articulo.OrdFechaCosto >= '" + WDesde + "' AND " _
                + "Articulo.OrdFechaCosto <= '" + WHasta + "'"
        
        
        
        ZZEstado = Dir("c:\VerificaCosto\ListaCambioCosto.xls")
        If ZZEstado <> "" Then
            Kill "c:\VerificaCosto\ListaCambioCosto.xls"
        End If
        
        ZZEstado = Dir("c:\VerificaCosto\ListaCambioCostopelli.xls")
        If ZZEstado <> "" Then
            Kill "c:\VerificaCosto\ListaCambioCostopelli.xls"
        End If
        
        Listado.Destination = 2
        Listado.PrintFileType = crptExcel50
        
        If WEmpresa = "0001" Then
            Listado.PrintFileName = "c:\VerificaCosto\ListaCambioCosto.xls"
                Else
            Listado.PrintFileName = "c:\VerificaCosto\ListaCambioCostopelli.xls"
        End If
        Listado.Connect = Connect()
        Listado.Action = 1
        
        
            
       If WEmpresa = "0001" Then
            EmailAddress = "juanfs@surfactan.com.ar;amenta@surfactan.com.ar; lsantos@surfactan.com.ar"
                Else
            EmailAddress = "juanfs@surfactan.com.ar;amenta@surfactan.com.ar;hferral@pellital.com.ar;argenta@pellital.com.ar;hgutierrez@pellital.com.ar"
       End If
        
        
        CopiaAddress = ""
        MSubject = "Cambios de Costo Std."
        MBody = "Se envia una lista e las materias primas que cambiaron el costo std. y lo productos que afectan."
        MAttach = ""
        If WEmpresa = "0001" Then
            MAttachI = "c:\VerificaCosto\ListaCambioCosto.xls"
            MAttachII = "c:\VerificaCosto\VerificaCosto.xls"
                Else
            MAttachI = "c:\VerificaCosto\ListaCambioCostopelli.xls"
            MAttachII = "c:\VerificaCosto\VerificaCostopelli.xls"
        End If
        MAttachIII = ""
        MAttachIV = ""
        MAttachV = ""
    
        SendEmail
     
        WEmpresa = "0008"
        Rem bynn
        XEmpresa = "08"
        Call Conecta_Empresa
        Rem agrego el otro cliclo
           
        Rem agregado by na
    Next i
    
    Rem fin by nan
    WEmpresa = ver
    Call Conecta_Empresa
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgListaVerificaCosto.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub DesdeFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFec.Text, Auxi)
        If Auxi = "S" Then
            HastaFec.SetFocus
                Else
            DesdeFec.SetFocus
        End If
    End If
End Sub


Private Sub HastaFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            DesdeFec.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
End Sub

Sub Form_Load()
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
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
                                Auxiliar(Renglon, 2) = Str$(Cantidad)
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
                    Auxiliar(Renglon, 2) = "1"
                    Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                End If
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For DA = 1 To Renglon
        Articulo = Auxiliar(DA, 1)
        Cantidad = Val(Auxiliar(DA, 2))
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




Private Sub Calcula_CostoII(Producto As String, Costo As Double)

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
                                Auxiliar(Renglon, 2) = Str$(Cantidad)
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
        Cantidad = Val(Auxiliar(DA, 2))
        WVector = Auxiliar(DA, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            Entra = "N"
            For Ciclo = 1 To 100
                If ZArticulo(Ciclo) = Articulo Then
                    Entra = "S"
                    Exit For
                End If
            Next Ciclo
            
            If Entra = "N" Then
                XCosto = rstArticulo!Costo2
                    Else
                XCosto = IIf(IsNull(rstArticulo!Costo2Anterior), "0", rstArticulo!Costo2Anterior)
                If XCosto = 0 Then
                    XCosto = rstArticulo!Costo2
                End If
            End If
                
            WCosto = (Cantidad * XCosto * Val(WVector))
            Costo = Costo + (Cantidad * XCosto * Val(WVector))
            rstArticulo.Close
            
        End If
    Next DA

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
        .Subject = MSubject
        .Body = MBody
        Rem .Attachments.Add MAttach
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
    
    
    
    
    




