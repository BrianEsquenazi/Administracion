VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgSeleccionaRecibo 
   AutoRedraw      =   -1  'True
   Caption         =   "Seleccion de Recibos a Aplicar la diferencia de cambio"
   ClientHeight    =   7890
   ClientLeft      =   1260
   ClientTop       =   1260
   ClientWidth     =   15180
   LinkTopic       =   "Form2"
   ScaleHeight     =   7890
   ScaleWidth      =   15180
   Begin VB.CommandButton ImpreAnexo 
      Caption         =   "Consulta Calculo"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancela"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
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
      Left            =   1320
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1800
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   6255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   11033
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   1800
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
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
   Begin Crystal.CrystalReport Listado 
      Left            =   8160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
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
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "PrgSeleccionaRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Auxi As String
Dim XParam As String
Dim WGraba(10000) As String
Dim Vector(10000, 4) As String

Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim ZNroRecibo As String


Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub cmdClose_Click()
    PrgSeleccionaRecibo.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Rem Call Proceso_Click
    Fecha.Text = "  /  /    "
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Limpia_Vector
        Call Proceso_Click
    End If
End Sub


Private Sub Proceso_Click()

    On Error GoTo WError

    Call Limpia_Vector


    WFECHADIA = Fecha.Text
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    FechaDia = WAno + WMes + WDia
    
    
    
    
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
    ZSql = ZSql + " and Recibos.FechaOrd >= " + "'" + "20151126" + "'"
    ZSql = ZSql + "Order by Recibos.Clave"
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
        ZSql = ZSql + "Order by Recibos.Clave"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
        
            With rstRecibos
                .MoveFirst
                Do
                    If .EOF = True Then
                        Exit Do
                    End If
                    If rstRecibos!Tiporeg = 2 Then
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

    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WDesde = WAno + WMes + WDia
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WHasta = WAno + WMes + WDia

    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE Recibos SET "
    Rem ZSql = ZSql + "ImpoList = " + "'" + "0" + "',"
    Rem ZSql = ZSql + "Impo1List = " + "'" + "0" + "'"
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
    ZSql = ZSql + " and Recibos.FechaOrd >= " + "'" + "20151126" + "'"
    ZSql = ZSql + "Order by Recibos.Clave"
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
    ZSql = ZSql + " and Recibos.FechaOrd >= " + "'" + "20151126" + "'"
    ZSql = ZSql + "Order by Recibos.Clave"
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
    ZSql = ZSql + " and Recibos.FechaOrd >= " + "'" + "20151126" + "'"
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

    
    
    
    
    
        
    
    

    Erase Vector
    Renglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.OrdFechaLista >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Recibos.OrdFechaLista <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Recibos.Renglon = " + "'" + "01" + "'"
    ZSql = ZSql + " and Recibos.FechaOrd >= " + "'" + "20151126" + "'"
    ZSql = ZSql + " Order by Recibos.Clave"
    spRecibo = ZSql
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibo.RecordCount > 0 Then
    
        With rstRecibo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WMarca = IIf(IsNull(!MarcaDebito), "", !MarcaDebito)
                    If Val(rstRecibo!TipoRec) = 1 And Trim(WMarca) = "" Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = rstRecibo!Recibo
                        Vector(Renglon, 2) = rstRecibo!Cliente
                        Vector(Renglon, 3) = rstRecibo!Fecha
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibo.Close
    End If
    
    
    
    For Ciclo = 1 To Renglon
    
        ZRecibo = Vector(Ciclo, 1)
        ZCliente = Vector(Ciclo, 2)
        ZFecha = Vector(Ciclo, 3)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Zrazon = rstCliente!Razon
            rstCliente.Close
        End If
    
        WVector1.TextMatrix(Ciclo, 1) = ZRecibo
        WVector1.TextMatrix(Ciclo, 2) = ZFecha
        WVector1.TextMatrix(Ciclo, 3) = ZCliente
        WVector1.TextMatrix(Ciclo, 4) = Zrazon
        
        
        ZSuma1 = 0
        ZSuma2 = 0
        ZSuma3 = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Recibo = " + "'" + ZRecibo + "'"
        ZSql = ZSql + "Order by Recibos.Clave"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
        
            With rstRecibos
                .MoveFirst
                Do
                    If .EOF = True Then
                        Exit Do
                    End If
                    If rstRecibos!Impolist <> 0 And rstRecibos!Importe1 <> 0 Then
                        ZSuma = rstRecibos!Importe1 / rstRecibos!Impolist
                        ZSuma1 = ZSuma1 + ZSuma
                    End If
                
                    If rstRecibos!impo1list <> 0 And rstRecibos!Importe2 <> 0 Then
                        ZSuma = rstRecibos!Importe2 / rstRecibos!impo1list
                        ZSuma2 = ZSuma2 + ZSuma
                    End If
                    
                    ZSuma3 = ZSuma3 + rstRecibos!Importe2
                    
                    ZParidad = rstRecibos!Paridad
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End With
            rstRecibos.Close
        End If
        
        WVector1.TextMatrix(Ciclo, 5) = Pusing("###,###,###.##", Str$(ZSuma3))
        ZDife = ZSuma1 - ZSuma2
        WVector1.TextMatrix(Ciclo, 6) = Pusing("###,###,###.##", Str$(ZDife))
        WVector1.TextMatrix(Ciclo, 7) = Pusing("###,###,###.##", Str$(ZDife * ZParidad))
        
        
    Next Ciclo
    
    WVector1.TopRow = 1
    WVector1.Row = 1
    WVector1.Col = 8
    
    WVector1.SetFocus
    
    Exit Sub
    
WError:

    WChequeo = "N"
    Resume Next

End Sub



Private Sub Graba_Click()

    For Ciclo = 1 To 1000
    
        If UCase(Trim(WVector1.TextMatrix(Ciclo, 8))) = "S" Or UCase(Trim(WVector1.TextMatrix(Ciclo, 8))) = "N" Then
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Recibos SET "
            ZSql = ZSql + " MarcaDebito = " + "'" + UCase(Trim(WVector1.TextMatrix(Ciclo, 8))) + "'"
            ZSql = ZSql + " Where Recibo = " + "'" + WVector1.TextMatrix(Ciclo, 1) + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    Call cmdClose_Click
            
End Sub

Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub



Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 123
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Col > 1 Then
                WVector1.Col = WVector1.Col - 1
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 8
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 8
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 9
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Recibo"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cliente"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Razon"
                WVector1.ColWidth(Ciclo) = 4000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 6
                WVector1.Text = "Dife U$S"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 7
                WVector1.Text = "Dife $"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 8
                WVector1.Text = "Marca"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector1.Text
        Rem WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        Rem WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        Rem WTitulo(Ciclo).Width = WVector1.CellWidth
        Rem WTitulo(Ciclo).Height = WVector1.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 340
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub


Private Sub WVector1_DblClick()

    WVector1.Col = 1
    WRecibo = WVector1.Text
    
    PrgReciOtro.Show
    
End Sub

Private Sub ImpreAnexo_Click()

    WVector1.Col = 1
    ZNroRecibo = WVector1.Text
    Call Ceros(ZNroRecibo, 6)

    Uno = "{Recibos.Recibo} in " + Chr$(34) + ZNroRecibo + Chr$(34) + " to " + Chr$(34) + ZNroRecibo + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno
    
    Listado.Destination = 1
    Listado.Destination = 0
        
    Listado.WindowTitle = "Consulta de Calculo de Diferencia de Cambio"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT Recibos.Recibo, Recibos.Fecha, Recibos.Fechaord, Recibos.Tipo1, Recibos.Numero1, Recibos.Importe1, Recibos.Numero2, Recibos.Fecha2, Recibos.banco2, Recibos.Importe2, Recibos.Impolist, Recibos.Impo1list, Recibos.Paridad, Recibos.OrdFechaLista, " _
            + "Cliente.Razon " _
            + "From " _
            + DSQ + ".dbo.Recibos Recibos, " _
            + DSQ + ".dbo.Cliente Cliente " _
            + "Where " _
            + "Recibos.Cliente = Cliente.Cliente AND " _
            + "Recibos.Recibo >= '" + ZNroRecibo + "' AND " _
            + "Recibos.Recibo <= '" + ZNroRecibo + "'"

    Listado.ReportFileName = "WDifeOtroAnexo.rpt"
    Listado.Action = 1

End Sub


