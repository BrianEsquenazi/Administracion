VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDifeOtroNuevo 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Diferencia de Cambio Acreditacion"
   ClientHeight    =   4785
   ClientLeft      =   2925
   ClientTop       =   2415
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   ScaleHeight     =   4785
   ScaleWidth      =   6240
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   3015
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   4815
      Begin VB.CheckBox OPcion 
         Caption         =   "Recalcula Anticipo"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox Desde 
         Height          =   285
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   16
         Text            =   " "
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Hasta 
         Height          =   285
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   15
         Text            =   " "
         Top             =   1920
         Width           =   1215
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   2040
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desdefecha 
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
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
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   300
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Caption         =   "Desde Cliente"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Cliente"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Emision"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Desde fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5160
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WDifeOtro.rpt"
      Destination     =   1
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
      Left            =   4680
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "DifeOtronuevo.frx":0000
      Left            =   0
      List            =   "DifeOtronuevo.frx":0007
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4920
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgDifeOtroNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstRecibo As Recordset
Dim spRecibo As String
Dim rstReciboListado As Recordset
Dim spReciboListado As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim XParam As String
Dim Vector(10000, 4) As String
Dim WClave As String
Dim WFecha As String
Dim WTipo As String
Dim WNumero As String
Dim Paridad1 As String
Dim Paridad2 As String
Dim WFechaFactura As String
Dim WGraba(10000) As String
Dim WCambiaFecha(10000, 3) As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer

Dim ZZImporte2 As Double


Dim ZZPasa(100, 100) As String
Dim ZZPasaII(100, 10) As String


Private Sub Acepta_Click()

    If Opcion.Value = 1 Then
        Call AceptaAnticipo_Click
        Exit Sub
    End If
    

    WFechaDia = Fecha.Text
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    FechaDia = WAno + WMes + WDia
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE Recibos SET "
    Rem ZSql = ZSql + "FechaLista = " + "'" + "" + "',"
    Rem ZSql = ZSql + "OrdFechaLista = " + "'" + "" + "',"
    Rem ZSql = ZSql + "MarcaII = " + "'" + "" + "'"
    Rem ZSql = ZSql + "Where Recibos.FechaOrd >= " + "'" + "20140101" + "'"
    Rem spRecibos = ZSql
    Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    
        
    ZSql = ""
    ZSql = ZSql + "UPDATE Recibos SET "
    ZSql = ZSql + " MarcaII = " + "'" + "" + "'"
    ZSql = ZSql + " Where MarcaII IS NULL"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        
    
    
    
    Erase WGraba
    Renglon = 0
        
        
    ZSql = ""
    ZSql = ZSql + "Select Recibos.Tiporeg, Recibos.Marca, Recibos.FechaOrd, Recibos.Clave, Recibos.Fechaord2, Recibos.Recibo, Recibos.Cliente, Recibos.MarcaII"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.TipoReg = " + "'" + "2" + "'"
    ZSql = ZSql + " and Recibos.MarcaII <> " + "'" + "X" + "'"
    ZSql = ZSql + " and Recibos.FechaOrd >= " + "'" + "20141201" + "'"
    ZSql = ZSql + " Order by Recibos.Clave"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then

        With rstRecibos
            .MoveFirst
            
            
            Do
                
                WMarca = IIf(IsNull(!MarcaII), "", !MarcaII)
                If WMarca <> "X" Then
                
                    Entra = "S"
                    
                    For Cicla = 1 To Renglon
                        If WGraba(Cicla) = rstRecibos!Recibo Then
                            Entra = "N"
                            Exit For
                        End If
                    Next Cicla
                    
                    If Entra = "S" Then
                        Renglon = Renglon + 1
                        WGraba(Renglon) = rstRecibos!Recibo
                    End If
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    
    End If

    For Cicla = 1 To Renglon
    
        WRecibo = WGraba(Cicla)
        XGraba = "S"
        ZFecha = "00000000"
        ZFechaI = "00/00/0000"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Recibo = " + "'" + WRecibo + "'"
        ZSql = ZSql + " Order by Recibos.Clave"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
        
            With rstRecibos
                .MoveFirst
                Do
                    If .EOF = True Then
                        Exit Do
                    End If
                    If rstRecibos!TipoReg = 2 Then
                        If Val(!Tipo2) = 2 Then
                            If !FechaOrd2 > ZFecha Then
                                ZFechaI = !Fecha2
                                ZFecha = !FechaOrd2
                            End If
                                Else
                            If !FechaOrd > ZFecha Then
                                ZFechaI = !Fecha
                                ZFecha = !FechaOrd
                            End If
                        End If
                    End If
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End With
            rstRecibos.Close
        End If
            
        If XGraba = "S" Then
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Recibos SET "
            ZSql = ZSql + "FechaLista = " + "'" + ZFechaI + "',"
            ZSql = ZSql + "OrdFechaLista = " + "'" + ZFecha + "',"
            ZSql = ZSql + "MarcaII = " + "'" + "X" + "'"
            ZSql = ZSql + " Where Recibo = " + "'" + WRecibo + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Cicla

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Wempresa
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WTitulo
            !varios = "Desde el " + Desdefecha.Text + " hasta el " + HastaFecha.Text
            .Update
        End If
    End With

    Listado.WindowTitle = "Listado de Difefrencias de Cambio (Acreditacion)"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia

    Rem spRecibo = "ModificaReciboImpolista0"
    Rem Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE Recibos SET "
    Rem ZSql = ZSql + "ImpoList = " + "'" + "0" + "',"
    Rem ZSql = ZSql + "Impo1List = " + "'" + "0" + "'"
    Rem ZSql = ZSql + " Where Recibos.FechaOrd >= " + "'" + "20141201" + "'"
    Rem spRecibos = ZSql
    Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Erase Vector
    Renglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.TipoReg = " + "'" + "1" + "'"
    ZSql = ZSql + " and Recibos.MarcaII = " + "'" + "X" + "'"
    ZSql = ZSql + " and Recibos.OrdFechaLista >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Recibos.OrdFechaLista <= " + "'" + WHasta + "'"
    ZSql = ZSql + " Order by Recibos.Clave"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstRecibos!Clave
                    Vector(Renglon, 2) = rstRecibos!Fecha
                    Vector(Renglon, 3) = rstRecibos!Tipo1
                    Vector(Renglon, 4) = rstRecibos!Numero1
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    End If
    
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        WFecha = Vector(Cicla, 2)
        WTipo = Vector(Cicla, 3)
        WNumero = Vector(Cicla, 4)
        
        ClaveCtacte = WTipo + WNumero + "01"
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            Paridad = Str$(rstCtacte!Paridad)
            If Val(Paridad) = 0 Then
                Paridad = "1"
            End If
            WFechaFactura = rstCtacte!Fecha
            rstCtacte.Close
            
            If Val(Paridad) = 1 Then
            
                WAno = Right$(WFechaFactura, 4)
                WMes = Mid$(WFechaFactura, 4, 2)
                WDia = Left$(WFechaFactura, 2)
                XClave = WAno + WMes + WDia

                spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
                Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                If rstCambios.RecordCount > 0 Then
                    With rstCambios
                        .MoveLast
                        aa1 = rstCambios!Fecha
                        aa2 = rstCambios!OrdFecha
                        Paridad = Str$(rstCambios!Cambio)
                        rstCambios.Close
                    End With
                        Else
                    Paridad = "1"
                End If
                
            End If
            
                Else
                
            WFechaFactura = "00/00/0000"
            Paridad = "1"
            
        End If
        
        XParam = "'" + WClave + "','" _
                    + Paridad + "','" _
                    + WFechaFactura + "'"
        spRecibo = "ModificaReciboDifeOtroI " + XParam
        Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Cicla
    
    

    Erase Vector
    Renglon = 0
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.TipoReg = " + "'" + "2" + "'"
    ZSql = ZSql + " and Recibos.MarcaII = " + "'" + "X" + "'"
    ZSql = ZSql + " and Recibos.OrdFechaLista >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Recibos.OrdFechaLista <= " + "'" + WHasta + "'"
    ZSql = ZSql + " Order by Recibos.Clave"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    XFecha = rstRecibos!Fecha2
                    If rstRecibos!Cliente = "G00065" And Val(rstRecibos!Tipo2) = 2 Then
                        Rem 1 - DOMINGO
                        Rem 2 - LUNES
                        Rem 3 - MARTES
                        Rem 4 - MIERCOLES
                        Rem 5 - JUEVES
                        Rem 6 - VIERNES
                        Rem 7 - SABADO
                        XFec1 = XFecha
                        strDia = Format$(XFec1, "dddd")
                        BDia = Format(XFec1, "w")
                        Select Case BDia
                            Case 2, 3, 4
                                SumaDia = 2
                            Case 5, 6, 7
                                SumaDia = 4
                            Case 1
                                SumaDia = 3
                            Case Else
                        End Select
                        SumaDia = SumaDia + 1
                        Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                        XFecha = XFec2
                    End If
                    
                    Vector(Renglon, 1) = rstRecibos!Clave
                    Vector(Renglon, 2) = rstRecibos!Fecha
                    Vector(Renglon, 3) = rstRecibos!Tipo2
                    Vector(Renglon, 4) = XFecha
                    WAno = Right$(XFecha, 4)
                    WMes = Mid$(XFecha, 4, 2)
                    WDia = Left$(XFecha, 2)
                    XFechaOrd = WAno + WMes + WDia
                    If rstRecibos!FechaOrd > XFechaOrd Then
                        Vector(Renglon, 4) = rstRecibos!Fecha
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    End If
    
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        WFecha = Vector(Cicla, 2)
        WTipo = Vector(Cicla, 3)
        WFechaord = Vector(Cicla, 4)
        
        If Val(WTipo) = 2 Or Val(WTipo) = 3 Then
        
            WAno = Right$(WFechaord, 4)
            WMes = Mid$(WFechaord, 4, 2)
            WDia = Left$(WFechaord, 2)
            XClave = WAno + WMes + WDia

            spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                With rstCambios
                    .MoveLast
                    aa1 = rstCambios!Fecha
                    aa2 = rstCambios!OrdFecha
                    Paridad = Str$(rstCambios!Cambio)
                    rstCambios.Close
                End With
                    Else
                Paridad = "1"
            End If
            
                Else
                
            WAno = Right$(WFecha, 4)
            WMes = Mid$(WFecha, 4, 2)
            WDia = Left$(WFecha, 2)
            XClave = WAno + WMes + WDia
            
            spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                With rstCambios
                    .MoveLast
                    aa1 = rstCambios!Fecha
                    aa2 = rstCambios!OrdFecha
                    Paridad = Str$(rstCambios!Cambio)
                    rstCambios.Close
                End With
                    Else
                Paridad = "1"
            End If
            
        End If
        
        XParam = "'" + WClave + "','" _
                    + Paridad + "'"
        spRecibo = "ModificaReciboDifeOtroII " + XParam
        Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Cicla
    
    
    

    Erase Vector
    Renglon = 0
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.OrdFechaLista >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Recibos.OrdFechaLista <= " + "'" + WHasta + "'"
    spRecibo = ZSql
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibo.RecordCount > 0 Then
    
        With rstRecibo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Retencion = 0
                    If Val(rstRecibo!Renglon) = 1 Then
                        Retencion = rstRecibo!Retganancias + rstRecibo!RetIva + rstRecibo!RetOtra + rstRecibo!RetSuss
                        
                        If Retencion > 0 Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstRecibo!Clave
                            Vector(Renglon, 2) = rstRecibo!Fecha
                            Vector(Renglon, 3) = Str$(Retencion)
                        End If
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibo.Close
    End If
    
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        WFecha = Vector(Cicla, 2)
        WRete = Vector(Cicla, 3)
        
        WAno = Right$(WFecha, 4)
        WMes = Mid$(WFecha, 4, 2)
        WDia = Left$(WFecha, 2)
        XClave = WAno + WMes + WDia

        spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
            With rstCambios
                .MoveLast
                aa1 = rstCambios!Fecha
                aa2 = rstCambios!OrdFecha
                Paridad = Str$(rstCambios!Cambio)
                rstCambios.Close
            End With
                Else
            Paridad = "1"
        End If
        
        WBanco = "Retenciones"
        XParam = "'" + WClave + "','" _
                     + WRete + "','" _
                     + Paridad + "'"
        spRecibo = "ModificaReciboDifeOtroV " + XParam
        Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Cicla
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    XClave = WAno + WMes + WDia

    spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        With rstCambios
            .MoveLast
            aa1 = rstCambios!Fecha
            aa2 = rstCambios!OrdFecha
            Paridad = Str$(rstCambios!Cambio)
            rstCambios.Close
        End With
            Else
        Paridad = "1"
    End If
    
    XParam = "'" + Paridad + "'"
    spRecibo = "ModificaReciboParidad " + XParam
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)

    Uno = "{Recibos.OrdFechaLista} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Dos = " and {Recibos.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT Recibos.Recibo, Recibos.Cliente, Recibos.Fecha, Recibos.Fechaord, Recibos.Tipo1, Recibos.Numero1, Recibos.Importe1, Recibos.Numero2, Recibos.Fecha2, Recibos.banco2, Recibos.Importe2, Recibos.Impolist, Recibos.Impo1list, Recibos.MarcaII, Recibos.OrdFechaLista, Recibos.Paridad, " _
            + "Cliente.Razon " _
            + "From " _
            + DSQ + ".dbo.Recibos Recibos, " _
            + DSQ + ".dbo.Cliente Cliente " _
            + "Where " _
            + "Recibos.Cliente = Cliente.Cliente And " _
            + "Recibos.Cliente >= '" + Desde.Text + "' AND " _
            + "Recibos.Cliente <= '" + Hasta.Text + "' AND " _
            + "Recibos.MarcaII = 'X' AND " _
            + "Recibos.OrdFechaLista >= '" + WDesde + "' AND " _
            + "Recibos.OrdFechaLista <= '" + WHasta + "'"
            
    Listado.ReportFileName = "WDifeOtroNuevo.rpt"
    
                        
    Rem Listado.DataFiles(1) = WEmpresa + "Auxi.mdb"
    Listado.DataFiles(2) = Wempresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub



Private Sub AceptaAnticipo_Click()



    WFechaDia = Fecha.Text
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    FechaDia = WAno + WMes + WDia
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE RecibosListado"
    spRecibosListado = ZSql
    Set rstReciboslistado = db.OpenRecordset(spRecibosListado, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Erase WGraba
    Renglon = 0
        
        
    ZSql = ""
    ZSql = ZSql + "Select Recibos.Tiporeg, Recibos.Marca, Recibos.FechaOrd, Recibos.Clave, Recibos.Fechaord2, Recibos.Recibo, Recibos.Cliente, Recibos.MarcaII"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.Tipo1 = " + "'" + "07" + "'"
    ZSql = ZSql + " and Recibos.Importe1 < 0"
    ZSql = ZSql + " and Recibos.FechaOrd >= " + "'" + "20151210" + "'"
    ZSql = ZSql + " and Recibos.Cliente >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Recibos.Cliente <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " Order by Recibos.Clave"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then

        With rstRecibos
            .MoveFirst
            
            
            Do
                
                Entra = "S"
                
                For Cicla = 1 To Renglon
                    If WGraba(Cicla) = rstRecibos!Recibo Then
                        Entra = "N"
                        Exit For
                    End If
                Next Cicla
                
                If Entra = "S" Then
                    Renglon = Renglon + 1
                    WGraba(Renglon) = rstRecibos!Recibo
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    
    End If

    For Cicla = 1 To Renglon
    
        Erase ZZPasa
        Erase ZZPasaII
        
        ZLugarI = 0
        ZLugarII = 0
    
        WRecibo = WGraba(Cicla)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Recibo = " + "'" + WRecibo + "'"
        ZSql = ZSql + " Order by Recibos.Clave"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
        
            With rstRecibos
                .MoveFirst
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If Val(!Tipo1) = 7 Then
                    
                        ZLugarII = ZLugarII + 1
                        ZZPasaII(ZLugarII, 1) = !Numero1
                        ZZPasaII(ZLugarII, 2) = !Recibo
                        ZZPasaII(ZLugarII, 3) = Str$(!Importe1 * -1)
                        
                        ZZZZClave = !Clave
                        ZZZZRecibo = !Recibo
                        ZZZZRenglon = !Renglon
                        ZZZZCliente = !Cliente
                        ZZZZFecha = !Fecha
                        ZZZZFechaOrd = !FechaOrd
                        ZZZZTipoRec = !TipoRec
                        ZZZZRetganancias = Str$(!Retganancias)
                        ZZZZRetIva = Str$(!RetIva)
                        ZZZZRetotra = Str$(!RetOtra)
                        ZZZZRetencion = Str$(!Retencion)
                        ZZZZTiporeg = !TipoReg
                        ZZZZTipo1 = !Tipo1
                        ZZZZLetra1 = !Letra1
                        ZZZZPunto1 = !Punto1
                        ZZZZNumero1 = !Numero1
                        ZZZZImporte1 = Str$(!Importe1)
                        ZZZZTipo2 = !Tipo2
                        ZZZZNumero2 = !Numero2
                        ZZZZFecha2 = !Fecha2
                        ZZZZBanco2 = !Banco2
                        ZZZZImporte2 = Str$(!Importe2)
                        ZZZZEstado2 = !Estado2
                        ZZZZEmpresa = Str$(!Empresa)
                        ZZZZFechaOrd2 = !FechaOrd2
                        ZZZZImporte = Str$(!Importe)
                        ZZZZObservaciones = !Observaciones
                        ZZZZImpolist = Str$(!Impolist)
                        ZZZZImpo1list = Str$(!impo1list)
                        ZZZZDestino = !Destino
                        ZZZZCuenta = !Cuenta
                        ZZZZMarca = !Marca
                        ZZZZFechaDepo = !fechadepo
                        ZZZZFechaDepoOrd = !FechaDepoOrd
                        
                        
                            Else
                            
                        ZLugarI = ZLugarI + 1
                        
                        XXRecibo = !Recibo
                        
                        ZZPasa(ZLugarI, 1) = !Clave
                        ZZPasa(ZLugarI, 2) = !Recibo
                        ZZPasa(ZLugarI, 3) = !Renglon
                        ZZPasa(ZLugarI, 4) = !Cliente
                        ZZPasa(ZLugarI, 5) = !Fecha
                        ZZPasa(ZLugarI, 6) = !FechaOrd
                        ZZPasa(ZLugarI, 7) = !TipoRec
                        ZZPasa(ZLugarI, 8) = Str$(!Retganancias)
                        ZZPasa(ZLugarI, 9) = Str$(!RetIva)
                        ZZPasa(ZLugarI, 10) = Str$(!RetOtra)
                        ZZPasa(ZLugarI, 11) = Str$(!Retencion)
                        ZZPasa(ZLugarI, 12) = !TipoReg
                        ZZPasa(ZLugarI, 13) = !Tipo1
                        ZZPasa(ZLugarI, 14) = !Letra1
                        ZZPasa(ZLugarI, 15) = !Punto1
                        ZZPasa(ZLugarI, 16) = !Numero1
                        ZZPasa(ZLugarI, 17) = Str$(!Importe1)
                        ZZPasa(ZLugarI, 18) = !Tipo2
                        ZZPasa(ZLugarI, 19) = !Numero2
                        ZZPasa(ZLugarI, 20) = !Fecha2
                        ZZPasa(ZLugarI, 21) = !Banco2
                        ZZPasa(ZLugarI, 22) = Str$(!Importe2)
                        ZZPasa(ZLugarI, 23) = !Estado2
                        ZZPasa(ZLugarI, 24) = Str$(!Empresa)
                        ZZPasa(ZLugarI, 25) = !FechaOrd2
                        ZZPasa(ZLugarI, 26) = Str$(!Importe)
                        ZZPasa(ZLugarI, 27) = !Observaciones
                        ZZPasa(ZLugarI, 28) = Str$(!Impolist)
                        ZZPasa(ZLugarI, 29) = Str$(!impo1list)
                        ZZPasa(ZLugarI, 30) = !Destino
                        ZZPasa(ZLugarI, 31) = !Cuenta
                        ZZPasa(ZLugarI, 32) = !Marca
                        ZZPasa(ZLugarI, 33) = !fechadepo
                        ZZPasa(ZLugarI, 34) = !FechaDepoOrd
                        
                    End If
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End With
            rstRecibos.Close
        End If
        
        
        For CiclaII = 1 To ZLugarII
            
            ZZAnticipo = ZZPasaII(CiclaII, 1)
            ZZRecibo = ZZPasaII(CiclaII, 2)
            ZZImporte = Val(ZZPasaII(CiclaII, 3))
            
            ZZSuma = 0
            ZZResta = 0
            ZDada = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Recibos"
            ZSql = ZSql + " Where Recibos.Tipo1 = " + "'" + "07" + "'"
            ZSql = ZSql + " and Recibos.Importe1 < 0"
            ZSql = ZSql + " and Recibos.Numero1 = " + "'" + ZZAnticipo + "'"
            ZSql = ZSql + " Order by Recibos.Clave"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
        
                With rstRecibos
                    .MoveFirst
                    
                    Do
                        
                        If Val(!Recibo) <> Val(ZZRecibo) Then
                            If Val(!Recibo) < Val(ZZRecibo) Then
                                ZZResta = ZZResta + (!Importe1 * -1)
                            End If
                        End If
                        
                        ZDada = ZDada + 1
                        
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibos.Close
            
            End If
            
            ZZZZPasa = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Recibos"
            ZSql = ZSql + " Where Recibos.Recibo = " + "'" + Mid$(ZZAnticipo, 3, 6) + "'"
            ZSql = ZSql + " and Recibos.TipoReg = " + "'" + "2" + "'"
            ZSql = ZSql + " Order by Recibos.Clave"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
        
                With rstRecibos
                    .MoveFirst
                    
                    Do
                    
                    
                        If ZZZZPasa = 0 Then
                            ZZZZPasa = 1
                            Retencion = !Retganancias + !RetIva + !RetOtra + !RetSuss
                            If Retencion > 0 Then
                            
                                ZZTipo2 = "04"
                                ZZNumero2 = ""
                                ZZFecha2 = !Fecha
                                ZZBanco2 = ""
                                ZZImporte2 = Retencion
                                ZZFechaOrd2 = ""
                            
                                If ZZResta > 0 Then
                                    If ZZResta > ZZImporte2 Then
                                        ZZResta = ZZResta - ZZImporte2
                                        ZZImporte2 = 0
                                            Else
                                        ZZImporte2 = ZZImporte2 - ZZResta
                                        ZZResta = 0
                                    End If
                                End If
                                
                                Call Redondeo(ZZImporte2)
                                
                                If ZZImporte2 > 0 Then
                                    
                                    ZZCompara = ZZSuma + ZZImporte2
                                    If ZZCompara > ZZImporte Then
                                        ZZImporte2 = ZZImporte - (ZZSuma + ZZImporte2)
                                    End If
                                    
                                    ZZSuma = ZZSuma + ZZImporte2
                                            
                                    ZLugarI = ZLugarI + 1
                                    
                                    Auxi = Str$(ZLugarI + 20)
                                    Call Ceros(Auxi, 2)
                                    
                                    ZZPasa(ZLugarI, 1) = XXRecibo + Auxi
                                    ZZPasa(ZLugarI, 2) = XXRecibo
                                    ZZPasa(ZLugarI, 3) = !Renglon
                                    ZZPasa(ZLugarI, 4) = !Cliente
                                    ZZPasa(ZLugarI, 5) = !Fecha
                                    ZZPasa(ZLugarI, 6) = !FechaOrd
                                    ZZPasa(ZLugarI, 7) = !TipoRec
                                    ZZPasa(ZLugarI, 8) = Str$(!Retganancias)
                                    ZZPasa(ZLugarI, 9) = Str$(!RetIva)
                                    ZZPasa(ZLugarI, 10) = Str$(!RetOtra)
                                    ZZPasa(ZLugarI, 11) = Str$(!Retencion)
                                    ZZPasa(ZLugarI, 12) = !TipoReg
                                    ZZPasa(ZLugarI, 13) = !Tipo1
                                    ZZPasa(ZLugarI, 14) = !Letra1
                                    ZZPasa(ZLugarI, 15) = !Punto1
                                    ZZPasa(ZLugarI, 16) = !Numero1
                                    ZZPasa(ZLugarI, 17) = Str$(!Importe1)
                                    ZZPasa(ZLugarI, 18) = ZZTipo2
                                    ZZPasa(ZLugarI, 19) = ZZNumero2
                                    ZZPasa(ZLugarI, 20) = ZZFecha2
                                    ZZPasa(ZLugarI, 21) = ZZBanco2
                                    ZZPasa(ZLugarI, 22) = Str$(ZZImporte2)
                                    ZZPasa(ZLugarI, 23) = Str$(ZZEstado2)
                                    ZZPasa(ZLugarI, 24) = Str$(!Empresa)
                                    ZZPasa(ZLugarI, 25) = !FechaOrd2
                                    ZZPasa(ZLugarI, 26) = Str$(!Importe)
                                    ZZPasa(ZLugarI, 27) = !Observaciones
                                    ZZPasa(ZLugarI, 28) = Str$(!Impolist)
                                    ZZPasa(ZLugarI, 29) = Str$(!impo1list)
                                    ZZPasa(ZLugarI, 30) = !Destino
                                    ZZPasa(ZLugarI, 31) = !Cuenta
                                    ZZPasa(ZLugarI, 32) = !Marca
                                    ZZPasa(ZLugarI, 33) = !fechadepo
                                    ZZPasa(ZLugarI, 34) = !FechaDepoOrd
                                            
                                End If
                                    
                                    
                            
                            
                            
                            End If
                        End If
                    
                    
                    
                    
                    
                        ZZTipo2 = !Tipo2
                        ZZNumero2 = !Numero2
                        ZZFecha2 = !Fecha2
                        ZZBanco2 = !Banco2
                        ZZImporte2 = !Importe2
                        ZZFechaOrd2 = !FechaOrd2
                    
                        If ZZResta > 0 Then
                            If ZZResta > ZZImporte2 Then
                                ZZResta = ZZResta - ZZImporte2
                                ZZImporte2 = 0
                                    Else
                                ZZImporte2 = ZZImporte2 - ZZResta
                                ZZResta = 0
                            End If
                        End If
                        
                        Call Redondeo(ZZImporte2)
                        
                        If ZZImporte2 > 0 Then
                            
                            ZZCompara = ZZSuma + ZZImporte2
                            If ZZCompara > ZZImporte Then
                                ZZImporte2 = ZZImporte - ZZSuma
                            End If
                            
                            ZZSuma = ZZSuma + ZZImporte2
                                    
                            ZLugarI = ZLugarI + 1
                            
                            Auxi = Str$(ZLugarI + 20)
                            Call Ceros(Auxi, 2)
                            
                            ZZPasa(ZLugarI, 1) = XXRecibo + Auxi
                            ZZPasa(ZLugarI, 2) = XXRecibo
                            ZZPasa(ZLugarI, 3) = !Renglon
                            ZZPasa(ZLugarI, 4) = !Cliente
                            ZZPasa(ZLugarI, 5) = !Fecha
                            ZZPasa(ZLugarI, 6) = !FechaOrd
                            ZZPasa(ZLugarI, 7) = !TipoRec
                            ZZPasa(ZLugarI, 8) = Str$(!Retganancias)
                            ZZPasa(ZLugarI, 9) = Str$(!RetIva)
                            ZZPasa(ZLugarI, 10) = Str$(!RetOtra)
                            ZZPasa(ZLugarI, 11) = Str$(!Retencion)
                            ZZPasa(ZLugarI, 12) = !TipoReg
                            ZZPasa(ZLugarI, 13) = !Tipo1
                            ZZPasa(ZLugarI, 14) = !Letra1
                            ZZPasa(ZLugarI, 15) = !Punto1
                            ZZPasa(ZLugarI, 16) = !Numero1
                            ZZPasa(ZLugarI, 17) = Str$(!Importe1)
                            ZZPasa(ZLugarI, 18) = ZZTipo2
                            ZZPasa(ZLugarI, 19) = ZZNumero2
                            ZZPasa(ZLugarI, 20) = ZZFecha2
                            ZZPasa(ZLugarI, 21) = ZZBanco2
                            ZZPasa(ZLugarI, 22) = Str$(ZZImporte2)
                            ZZPasa(ZLugarI, 23) = Str$(ZZEstado2)
                            ZZPasa(ZLugarI, 24) = Str$(!Empresa)
                            ZZPasa(ZLugarI, 25) = !FechaOrd2
                            ZZPasa(ZLugarI, 26) = Str$(!Importe)
                            ZZPasa(ZLugarI, 27) = !Observaciones
                            ZZPasa(ZLugarI, 28) = Str$(!Impolist)
                            ZZPasa(ZLugarI, 29) = Str$(!impo1list)
                            ZZPasa(ZLugarI, 30) = !Destino
                            ZZPasa(ZLugarI, 31) = !Cuenta
                            ZZPasa(ZLugarI, 32) = !Marca
                            ZZPasa(ZLugarI, 33) = !fechadepo
                            ZZPasa(ZLugarI, 34) = !FechaDepoOrd
                                    
                        End If
                        
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibos.Close
            
            End If
            
        Next CiclaII
        
        
        
        
        
        For CiclaII = 1 To ZLugarI
        
            ZZZZClave = ZZPasa(CiclaII, 1)
            ZZZZRecibo = ZZPasa(CiclaII, 2)
            ZZZZRenglon = ZZPasa(CiclaII, 3)
            ZZZZCliente = ZZPasa(CiclaII, 4)
            ZZZZFecha = ZZPasa(CiclaII, 5)
            ZZZZFechaOrd = ZZPasa(CiclaII, 6)
            ZZZZTipoRec = ZZPasa(CiclaII, 7)
            ZZZZRetganancias = ZZPasa(CiclaII, 8)
            ZZZZRetIva = ZZPasa(CiclaII, 9)
            ZZZZRetotra = ZZPasa(CiclaII, 10)
            ZZZZRetencion = ZZPasa(CiclaII, 11)
            ZZZZTiporeg = ZZPasa(CiclaII, 12)
            ZZZZTipo1 = ZZPasa(CiclaII, 13)
            ZZZZLetra1 = ZZPasa(CiclaII, 14)
            ZZZZPunto1 = ZZPasa(CiclaII, 15)
            ZZZZNumero1 = ZZPasa(CiclaII, 16)
            ZZZZImporte1 = ZZPasa(CiclaII, 17)
            ZZZZTipo2 = ZZPasa(CiclaII, 18)
            ZZZZNumero2 = ZZPasa(CiclaII, 19)
            ZZZZFecha2 = ZZPasa(CiclaII, 20)
            ZZZZBanco2 = Trim(ZZPasa(CiclaII, 21))
            ZZZZImporte2 = ZZPasa(CiclaII, 22)
            Rem ZZZZEstado2 = ZZPasa(CiclaII, 23)
            ZZZZEstado2 = ""
            ZZZZEmpresa = ZZPasa(CiclaII, 24)
            ZZZZFechaOd2 = ZZPasa(CiclaII, 25)
            ZZZZImpore = ZZPasa(CiclaII, 26)
            ZZZZObservaciones = ZZPasa(CiclaII, 27)
            ZZZZImpolist = ZZPasa(CiclaII, 28)
            ZZZZImpo1list = ZZPasa(CiclaII, 29)
            ZZZZDestino = ZZPasa(CiclaII, 30)
            ZZZZCuenta = ZZPasa(CiclaII, 31)
            ZZZZMarca = ZZPasa(CiclaII, 32)
            ZZZZFechaDepo = ZZPasa(CiclaII, 33)
            ZZZZFechaDepoOrd = ZZPasa(CiclaII, 34)
            
            ZSql = "INSERT INTO RecibosListado ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Recibo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Fechaord ,"
            ZSql = ZSql + "TipoRec ,"
            ZSql = ZSql + "RetGanancias ,"
            ZSql = ZSql + "RetIva ,"
            ZSql = ZSql + "RetOtra ,"
            ZSql = ZSql + "Retencion ,"
            ZSql = ZSql + "TipoReg ,"
            ZSql = ZSql + "Tipo1 ,"
            ZSql = ZSql + "Letra1 ,"
            ZSql = ZSql + "Punto1 ,"
            ZSql = ZSql + "Numero1 ,"
            ZSql = ZSql + "Importe1 ,"
            ZSql = ZSql + "Tipo2 ,"
            ZSql = ZSql + "Numero2 ,"
            ZSql = ZSql + "Fecha2 ,"
            ZSql = ZSql + "banco2 ,"
            ZSql = ZSql + "Importe2 ,"
            ZSql = ZSql + "Estado2 ,"
            ZSql = ZSql + "Empresa ,"
            ZSql = ZSql + "FechaOrd2 ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Impolist ,"
            ZSql = ZSql + "Impo1list ,"
            ZSql = ZSql + "Destino ,"
            ZSql = ZSql + "Cuenta ,"
            ZSql = ZSql + "Marca ,"
            ZSql = ZSql + "FechaDepo ,"
            ZSql = ZSql + "FechaDepoOrd)"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZZZClave + "',"
            ZSql = ZSql + "'" + ZZZZRecibo + "',"
            ZSql = ZSql + "'" + ZZZZRenglon + "',"
            ZSql = ZSql + "'" + ZZZZCliente + "',"
            ZSql = ZSql + "'" + ZZZZFecha + "',"
            ZSql = ZSql + "'" + ZZZZFechaOrd + "',"
            ZSql = ZSql + "'" + ZZZZTipoRec + "',"
            ZSql = ZSql + "'" + ZZZZRetganancias + "',"
            ZSql = ZSql + "'" + ZZZZRetIva + "',"
            ZSql = ZSql + "'" + ZZZZRetotra + "',"
            ZSql = ZSql + "'" + ZZZZRetencion + "',"
            ZSql = ZSql + "'" + ZZZZTiporeg + "',"
            ZSql = ZSql + "'" + ZZZZTipo1 + "',"
            ZSql = ZSql + "'" + ZZZZLetra1 + "',"
            ZSql = ZSql + "'" + ZZZZPunto1 + "',"
            ZSql = ZSql + "'" + ZZZZNumero1 + "',"
            ZSql = ZSql + "'" + ZZZZImporte1 + "',"
            ZSql = ZSql + "'" + ZZZZTipo2 + "',"
            ZSql = ZSql + "'" + ZZZZNumero2 + "',"
            ZSql = ZSql + "'" + ZZZZFecha2 + "',"
            ZSql = ZSql + "'" + ZZZZBanco2 + "',"
            ZSql = ZSql + "'" + ZZZZImporte2 + "',"
            ZSql = ZSql + "'" + ZZZZEstado2 + "',"
            ZSql = ZSql + "'" + ZZZZEmpresa + "',"
            ZSql = ZSql + "'" + ZZZZFechaOrd2 + "',"
            ZSql = ZSql + "'" + ZZZZImporte + "',"
            ZSql = ZSql + "'" + ZZZZObservaciones + "',"
            ZSql = ZSql + "'" + ZZZZImpolist + "',"
            ZSql = ZSql + "'" + ZZZZImpo1list + "',"
            ZSql = ZSql + "'" + ZZZZDestino + "',"
            ZSql = ZSql + "'" + ZZZZCuenta + "',"
            ZSql = ZSql + "'" + ZZZZMarca + "',"
            ZSql = ZSql + "'" + ZZZZFechaDepo + "',"
            ZSql = ZSql + "'" + ZZZZFechaDepoOrd + "')"
            spRecibosListado = ZSql
            Set rstReciboslistado = db.OpenRecordset(spRecibosListado, dbOpenSnapshot, dbSQLPassThrough)
        
        
            
        Next CiclaII
        
    Next Cicla










    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Wempresa
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WTitulo
            !varios = "Desde el " + Desdefecha.Text + " hasta el " + HastaFecha.Text
            .Update
        End If
    End With

    Listado.WindowTitle = "Listado de Difefrencias de Cambio (Acreditacion)"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesde = WAno + WMes + WDia
    
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia

    Rem spRecibo = "ModificaReciboImpolista0"
    Rem Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE RecibosListado SET "
    ZSql = ZSql + "ImpoList = " + "'" + "0" + "',"
    ZSql = ZSql + "Impo1List = " + "'" + "0" + "'"
    spRecibosListado = ZSql
    Set rstReciboslistado = db.OpenRecordset(spRecibosListado, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Erase Vector
    Renglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM RecibosListado"
    ZSql = ZSql + " Where RecibosListado.TipoReg = " + "'" + "1" + "'"
    ZSql = ZSql + " Order by RecibosListado.Clave"
    spRecibosListado = ZSql
    Set rstReciboslistado = db.OpenRecordset(spRecibosListado, dbOpenSnapshot, dbSQLPassThrough)
    If rstReciboslistado.RecordCount > 0 Then
        With rstReciboslistado
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstReciboslistado!Clave
                    Vector(Renglon, 2) = rstReciboslistado!Fecha
                    Vector(Renglon, 3) = rstReciboslistado!Tipo1
                    Vector(Renglon, 4) = rstReciboslistado!Numero1
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstReciboslistado.Close
    End If
    
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        WFecha = Vector(Cicla, 2)
        WTipo = Vector(Cicla, 3)
        WNumero = Vector(Cicla, 4)
        
        ClaveCtacte = WTipo + WNumero + "01"
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            Paridad = Str$(rstCtacte!Paridad)
            If Val(Paridad) = 0 Then
                Paridad = "1"
            End If
            WFechaFactura = rstCtacte!Fecha
            rstCtacte.Close
            
            If Val(Paridad) = 1 Then
            
                WAno = Right$(WFechaFactura, 4)
                WMes = Mid$(WFechaFactura, 4, 2)
                WDia = Left$(WFechaFactura, 2)
                XClave = WAno + WMes + WDia

                spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
                Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                If rstCambios.RecordCount > 0 Then
                    With rstCambios
                        .MoveLast
                        aa1 = rstCambios!Fecha
                        aa2 = rstCambios!OrdFecha
                        Paridad = Str$(rstCambios!Cambio)
                        rstCambios.Close
                    End With
                        Else
                    Paridad = "1"
                End If
                
            End If
            
                Else
                
            WFechaFactura = "00/00/0000"
            Paridad = "1"
            
        End If
        
        Rem XParam = "'" + WClave + "','" _
        rem             + Paridad + "','" _
        rem             + WFechaFactura + "'"
        Rem spRecibo = "ModificaReciboDifeOtroI " + XParam
        Rem Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE RecibosListado SET "
        ZSql = ZSql + "ImpoList = " + "'" + Paridad + "',"
        ZSql = ZSql + "Banco2 = " + "'" + WFechaFactura + "'"
        ZSql = ZSql + "Where RecibosListado.Clave = " + "'" + WClave + "'"
        spRecibosListado = ZSql
        Set rstReciboslistado = db.OpenRecordset(spRecibosListado, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Cicla
    
    

    Erase Vector
    Renglon = 0
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM RecibosListado"
    ZSql = ZSql + " Where RecibosListado.TipoReg = " + "'" + "2" + "'"
    ZSql = ZSql + " Order by RecibosListado.Clave"
    spRecibosListado = ZSql
    Set rstReciboslistado = db.OpenRecordset(spRecibosListado, dbOpenSnapshot, dbSQLPassThrough)
    If rstReciboslistado.RecordCount > 0 Then
        With rstReciboslistado
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    XFecha = rstReciboslistado!Fecha2
                    If XFecha = "" Then
                        XFecha = rstReciboslistado!Fecha
                    End If
                    If rstReciboslistado!Cliente = "G00065" And Val(rstReciboslistado!Tipo2) = 2 Then
                        Rem 1 - DOMINGO
                        Rem 2 - LUNES
                        Rem 3 - MARTES
                        Rem 4 - MIERCOLES
                        Rem 5 - JUEVES
                        Rem 6 - VIERNES
                        Rem 7 - SABADO
                        XFec1 = XFecha
                        strDia = Format$(XFec1, "dddd")
                        BDia = Format(XFec1, "w")
                        Select Case BDia
                            Case 2, 3, 4
                                SumaDia = 2
                            Case 5, 6, 7
                                SumaDia = 4
                            Case 1
                                SumaDia = 3
                            Case Else
                        End Select
                        SumaDia = SumaDia + 1
                        Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                        XFecha = XFec2
                    End If
                    
                    Vector(Renglon, 1) = rstReciboslistado!Clave
                    Vector(Renglon, 2) = rstReciboslistado!Fecha
                    Vector(Renglon, 3) = rstReciboslistado!Tipo2
                    Vector(Renglon, 4) = XFecha
                    WAno = Right$(XFecha, 4)
                    WMes = Mid$(XFecha, 4, 2)
                    WDia = Left$(XFecha, 2)
                    XFechaOrd = WAno + WMes + WDia
                    If rstReciboslistado!FechaOrd > XFechaOrd Then
                        Vector(Renglon, 4) = rstReciboslistado!Fecha
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstReciboslistado.Close
    End If
    
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        WFecha = Vector(Cicla, 2)
        WTipo = Vector(Cicla, 3)
        WFechaord = Vector(Cicla, 4)
        
        If Val(WTipo) = 2 Or Val(WTipo) = 3 Then
        
            WAno = Right$(WFechaord, 4)
            WMes = Mid$(WFechaord, 4, 2)
            WDia = Left$(WFechaord, 2)
            XClave = WAno + WMes + WDia

            spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                With rstCambios
                    .MoveLast
                    aa1 = rstCambios!Fecha
                    aa2 = rstCambios!OrdFecha
                    Paridad = Str$(rstCambios!Cambio)
                    rstCambios.Close
                End With
                    Else
                Paridad = "1"
            End If
            
                Else
                
            WAno = Right$(WFecha, 4)
            WMes = Mid$(WFecha, 4, 2)
            WDia = Left$(WFecha, 2)
            XClave = WAno + WMes + WDia
            
            spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                With rstCambios
                    .MoveLast
                    aa1 = rstCambios!Fecha
                    aa2 = rstCambios!OrdFecha
                    Paridad = Str$(rstCambios!Cambio)
                    rstCambios.Close
                End With
                    Else
                Paridad = "1"
            End If
            
        End If
        
        Rem XParam = "'" + WClave + "','" _
        rem             + Paridad + "'"
        Rem spRecibo = "ModificaReciboDifeOtroII " + XParam
        Rem Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE RecibosListado SET "
        ZSql = ZSql + "Impo1List = " + "'" + Paridad + "'"
        ZSql = ZSql + "Where RecibosListado.Clave = " + "'" + WClave + "'"
        spRecibosListado = ZSql
        Set rstReciboslistado = db.OpenRecordset(spRecibosListado, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Next Cicla
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFechaDia = WAno + WMes + WDia
    

    Erase Vector
    Renglon = 0
    ZZZZPasa = 0
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM RecibosListado"
    ZSql = ZSql + " Where RecibosListado.TipoReg = " + "'" + "2" + "'"
    ZSql = ZSql + " Order by RecibosListado.Clave"
    spRecibosListado = ZSql
    Set rstReciboslistado = db.OpenRecordset(spRecibosListado, dbOpenSnapshot, dbSQLPassThrough)
    If rstReciboslistado.RecordCount > 0 Then
        With rstReciboslistado
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If ZZZZPasa = 0 Then
                        ZZZZPasa = 1
                        ZZZZCorte = !Recibo
                        ZZZZFecha = ""
                    End If
                    
                    If ZZZZCorte <> !Recibo Then
                        WAno = Right$(ZZZZFecha, 4)
                        WMes = Mid$(ZZZZFecha, 4, 2)
                        WDia = Left$(ZZZZFecha, 2)
                        XFechaOrd = WAno + WMes + WDia
                        If XFechaOrd < WDesde Or XFechaOrd > WHasta Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = ZZZZCorte
                        End If
                        ZZZZCorte = !Recibo
                        ZZZZFecha = ""
                    End If
                        
                    If !TipoReg = 2 And !Tipo2 = "02" Then
                    
                        WAno = Right$(!Fecha2, 4)
                        WMes = Mid$(!Fecha2, 4, 2)
                        WDia = Left$(!Fecha2, 2)
                        XFechaOrd = WAno + WMes + WDia
                        
                        WAno = Right$(ZZZZFecha, 4)
                        WMes = Mid$(ZZZZFecha, 4, 2)
                        WDia = Left$(ZZZZFecha, 2)
                        XFechaOrdII = WAno + WMes + WDia
                        
                        If XFechaOrd > XFechaOrdII Then
                            ZZZZFecha = !Fecha2
                        End If
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstReciboslistado.Close
    End If
    
    If ZZZZPasa <> 0 Then
        WAno = Right$(ZZZZFecha, 4)
        WMes = Mid$(ZZZZFecha, 4, 2)
        WDia = Left$(ZZZZFecha, 2)
        XFechaOrd = WAno + WMes + WDia
        If XFechaOrd < WDesde Or XFechaOrd > WHasta Then
            Renglon = Renglon + 1
            Vector(Renglon, 1) = ZZZZCorte
        End If
    End If
    
    
    For Cicla = 1 To Renglon
    
        ZZZZRecibo = Vector(Cicla, 1)
    
        ZSql = ""
        ZSql = ZSql + "DELETE RecibosListado"
        ZSql = ZSql + " Where RecibosListado.Recibo = " + "'" + ZZZZRecibo + "'"
        spRecibosListado = ZSql
        Set rstReciboslistado = db.OpenRecordset(spRecibosListado, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Cicla
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    XClave = WAno + WMes + WDia

    spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        With rstCambios
            .MoveLast
            aa1 = rstCambios!Fecha
            aa2 = rstCambios!OrdFecha
            Paridad = Str$(rstCambios!Cambio)
            rstCambios.Close
        End With
            Else
        Paridad = "1"
    End If
    
    Rem XParam = "'" + Paridad + "'"
    Rem spRecibo = "ModificaReciboParidad " + XParam
    Rem Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)

    ZSql = ""
    ZSql = ZSql + "UPDATE RecibosListado SET "
    ZSql = ZSql + "Paridad = " + "'" + Paridad + "'"
    spRecibosListado = ZSql
    Set rstReciboslistado = db.OpenRecordset(spRecibosListado, dbOpenSnapshot, dbSQLPassThrough)






    Uno = "{RecibosListado.Recibo} in " + Chr$(34) + "000000" + Chr$(34) + " to " + Chr$(34) + "999999" + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT RecibosListado.Recibo, RecibosListado.Renglon, RecibosListado.Tipo1, RecibosListado.Numero1, RecibosListado.Importe1, RecibosListado.Numero2, RecibosListado.Fecha2, RecibosListado.banco2, RecibosListado.Importe2, RecibosListado.Impolist, RecibosListado.Impo1list, RecibosListado.Paridad, " _
            + "Cliente.Razon " _
            + "From " _
            + DSQ + ".dbo.RecibosListado RecibosListado, " _
            + DSQ + ".dbo.Cliente Cliente " _
            + "Where " _
            + "RecibosListado.Cliente = Cliente.Cliente AND " _
            + "RecibosListado.Recibo >= '000000' AND " _
            + "RecibosListado.Recibo <= '999999'"
            
    Listado.ReportFileName = "WDifeOtroNuevoAnticipo.rpt"
                        
    Rem Listado.DataFiles(1) = WEmpresa + "Auxi.mdb"
    Listado.DataFiles(2) = Wempresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub



Private Sub Cancela_Click()

    With rstEmpresa
        .Close
    End With
    PrgDifeOtroNuevo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()

    spRecibos = "ModificaReciboImpoLista0"
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    
    spRecibos = "ActualizaRecibosSalvaMarca"
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    
    spRecibos = "ActualizaRecibosSalvaMarcaII"
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.Text = Desdefecha.Text
        HastaFecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
End Sub

Sub Form_Load()

    Fecha.Text = "  /  /    "
    Desdefecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Desde.Text = "A00000"
    Hasta.Text = "Z99999"
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

