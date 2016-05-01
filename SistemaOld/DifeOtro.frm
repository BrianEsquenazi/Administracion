VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDifeOtro 
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
         Left            =   2280
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   1080
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
      ItemData        =   "DifeOtro.frx":0000
      Left            =   0
      List            =   "DifeOtro.frx":0007
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
Attribute VB_Name = "PrgDifeOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstRecibo As Recordset
Dim spRecibo As String
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

Private Sub Acepta_Click()

    WFECHADIA = Fecha.Text
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    FechaDia = WAno + WMes + WDia
    
    Renglon = 0
    Erase WGraba
    
    ZSql = ""
    ZSql = ZSql + "Select Recibos.Recibo, Recibos.Tiporeg, Recibos.Estado2, Recibos.FechaOrd, Recibos.Clave, Recibos.Fechaord2, Recibos.Tipo2 "
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.TipoReg = " + "'" + "2" + "'"
    ZSql = ZSql + " and Recibos.Estado2 <> " + "'" + "X" + "'"
    ZSql = ZSql + " and Recibos.FechaOrd >= " + "'" + "20080901" + "'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then

        With rstRecibos
            .MoveFirst
            Do
            
                Select Case Val(!Tipo2)
                    Case 1, 4
                        Renglon = Renglon + 1
                        WGraba(Renglon) = rstRecibos!Clave
                    Case 3
                        If FechaDia >= rstRecibos!FechaOrd2 Then
                            Renglon = Renglon + 1
                            WGraba(Renglon) = rstRecibos!Clave
                        End If
                    Case Else
                End Select
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    
    End If
    
    For Cicla = 1 To Renglon
    
        Recibo = WGraba(Cicla)
        
        XParam = "'" + Recibo + "'"
        spRecibos = "ActualizaRecibosOtro " + XParam
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Cicla
    
    Erase WGraba
    Renglon = 0
        
        
    ZSql = ""
    ZSql = ZSql + "Select Recibos.Tiporeg, Recibos.Marca, Recibos.FechaOrd, Recibos.Clave, Recibos.Fechaord2, Recibos.Recibo, Recibos.Cliente"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.TipoReg = " + "'" + "2" + "'"
    ZSql = ZSql + " and Recibos.Marca <> " + "'" + "X" + "'"
    ZSql = ZSql + " and Recibos.FechaOrd >= " + "'" + "20080901" + "'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then

        With rstRecibos
            .MoveFirst
            
            
            Do
                
                WMarca = IIf(IsNull(!Marca), "", !Marca)
                Rem If WMarca <> "X" Then
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
        
        If Val(WRecibo) = 85258 Then Stop
                
        spRecibos = "ConsultaRecibos " + "'" + WRecibo + "'"
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
            With rstRecibos
                .MoveFirst
                Do
                    If .EOF = True Then
                        Exit Do
                    End If
                    If rstRecibos!Tiporeg = 2 Then
                        If rstRecibos!Estado2 <> "X" Then
                            XGraba = "N"
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
            XParam = "'" + WRecibo + "','" _
                        + WFECHADIA + "','" _
                        + FechaDia + "','" _
                        + "X" + " '"
            spRecibos = "ActualizaRecibosMarca " + XParam
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next Cicla
    
    If Wempresa = "0008" Then
    
        Erase WCambiaFecha
        Renglon = 0
            
        spRecibos = "ListaRecibosDifeOtroVI "
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
    
        With rstRecibos
                .MoveFirst
                Do
                     Entra = "S"
                    
                    aa = rstRecibos!fechadepo
                    aa1 = rstRecibos!Fecha2
                    aa2 = rstRecibos!FechaOrd
                    
                    Entra = "S"
                        
                    For Cicla = 1 To Renglon
                        If WCambiaFecha(Cicla, 1) = rstRecibos!Recibo Then
                            If rstRecibos!FechaOrd2 > WCambiaFecha(Cicla, 3) Then
                                WCambiaFecha(Cicla, 2) = rstRecibos!Fecha2
                                WCambiaFecha(Cicla, 3) = rstRecibos!FechaOrd2
                            End If
                            Entra = "N"
                            Exit For
                        End If
                    Next Cicla
                        
                    If Entra = "S" Then
                        Renglon = Renglon + 1
                        WCambiaFecha(Renglon, 1) = rstRecibos!Recibo
                        WCambiaFecha(Renglon, 2) = rstRecibos!Fecha2
                        WCambiaFecha(Renglon, 3) = rstRecibos!FechaOrd2
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
        
            WRecibo = WCambiaFecha(Cicla, 1)
            WFechadepo = WCambiaFecha(Cicla, 2)
            WFechadepoord = WCambiaFecha(Cicla, 3)
            
            XParam = "'" + WRecibo + "','" _
                         + WFechadepo + "','" _
                         + WFechadepoord + " '"
            spRecibos = "ActualizaRecibosOtroVI " + XParam
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
        Next Cicla
    
    End If

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

    spRecibo = "ModificaReciboImpolista0"
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spRecibo = "ListaRecibosDifeOtroI" + XParam
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibo.RecordCount > 0 Then
        With rstRecibo
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstRecibo!Clave
                    Vector(Renglon, 2) = rstRecibo!Fecha
                    Vector(Renglon, 3) = rstRecibo!Tipo1
                    Vector(Renglon, 4) = rstRecibo!Numero1
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
        WTipo = Vector(Cicla, 3)
        WNumero = Vector(Cicla, 4)
        
        ClaveCtacte = WTipo + WNumero + "01"
        spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            Paridad = Str$(rstCtaCte!Paridad)
            If Val(Paridad) = 0 Then
                Paridad = "1"
            End If
            WFechaFactura = rstCtaCte!Fecha
            rstCtaCte.Close
            
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
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spRecibo = "ListaRecibosDifeOtroII" + XParam
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibo.RecordCount > 0 Then
        With rstRecibo
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    XFecha = rstRecibo!Fecha2
                    If !Cliente = "G00065" And Val(rstRecibo!Tipo2) = 2 Then
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
                    
                    Vector(Renglon, 1) = rstRecibo!Clave
                    Vector(Renglon, 2) = rstRecibo!Fecha
                    Vector(Renglon, 3) = rstRecibo!Tipo2
                    Vector(Renglon, 4) = XFecha
                    WAno = Right$(XFecha, 4)
                    WMes = Mid$(XFecha, 4, 2)
                    WDia = Left$(XFecha, 2)
                    XFechaOrd = WAno + WMes + WDia
                    If rstRecibo!FechaOrd > XFechaOrd Then
                        Vector(Renglon, 4) = rstRecibo!Fecha
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
    ZSql = ZSql + " Where Recibos.FechaDepoOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Recibos.FechaDepoOrd <= " + "'" + WHasta + "'"
    
    spRecibo = ZSql
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibo.RecordCount > 0 Then
    
    Rem XParam = "'" + WDesde + "','" _
    rem              + WHasta + "'"
    Rem spRecibo = "ListaRecibosDifeOtroV" + XParam
    Rem Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstRecibo.RecordCount > 0 Then
    
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

    Uno = "{Recibos.Fechadepoord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Dos = " and {Recibos.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT Recibos.Recibo, Recibos.Cliente, Recibos.Fecha, Recibos.Fechaord, Recibos.Tipo1, Recibos.Numero1, Recibos.Importe1, Recibos.Numero2, Recibos.Fecha2, Recibos.banco2, Recibos.Importe2, Recibos.Impolist, Recibos.Impo1list, Recibos.Marca, Recibos.FechaDepoOrd, Recibos.Paridad, " _
            + "Cliente.Razon " _
            + "From " _
            + DSQ + ".dbo.Recibos Recibos, " _
            + DSQ + ".dbo.Cliente Cliente " _
            + "Where " _
            + "Recibos.Cliente = Cliente.Cliente And " _
            + "Recibos.Cliente >= '" + Desde.Text + "' AND " _
            + "Recibos.Cliente <= '" + Hasta.Text + "' AND " _
            + "Recibos.Fechaord > '20080901' AND " _
            + "Recibos.Marca = 'X' AND " _
            + "Recibos.FechaDepoOrd >= '" + WDesde + "' AND " _
            + "Recibos.FechaDepoOrd <= '" + WHasta + "'"
    
                        
    Rem Listado.DataFiles(1) = WEmpresa + "Auxi.mdb"
    Listado.DataFiles(2) = Wempresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_Click()

    With rstEmpresa
        .Close
    End With
    Desdefecha.SetFocus
    PrgDifeOtro.Hide
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
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

