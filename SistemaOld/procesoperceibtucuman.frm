VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProcesoPerceIbTucuman 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Traspaso de Percepcion de Ingresos Brutos (Tucuman)"
   ClientHeight    =   5775
   ClientLeft      =   3060
   ClientTop       =   1425
   ClientWidth     =   7290
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   7290
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   4815
      Begin VB.CommandButton MOratoria 
         Caption         =   "Moratoria"
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
         Left            =   2040
         TabIndex        =   10
         Top             =   4560
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
         Height          =   495
         Left            =   1080
         TabIndex        =   6
         Top             =   3960
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
         Left            =   2640
         TabIndex        =   5
         Top             =   3960
         Width           =   1215
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   1920
         TabIndex        =   3
         Top             =   1920
         Width           =   2055
      End
      Begin VB.DriveListBox Drive 
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
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1920
         TabIndex        =   4
         Top             =   720
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Top             =   360
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
      Begin VB.Label Label3 
         Caption         =   "Desde fecha"
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
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta fecha"
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
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Destino"
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
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgProcesoPerceIbTucuman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstClientes As Recordset
Dim spClientes As String
Dim rstCtaCte As Recordset
Dim spCtaCte As String
Dim rstRecibo As Recordset
Dim spRecibo As String
Dim XParam As String
Dim Vector(10000, 10) As String
Dim WClave As String
Dim WFecha As String
Dim WTipo As String
Dim WNumero As String
Dim WNeto As Double
Dim WImpoIb As Double
Dim WPorceIb As Double
Dim XNeto As String
Dim XImpoIb As String
Dim XPorceIb As String
Dim WCuit As String
Dim WIbTucu As Integer
Dim ZEntra(1000) As String
Dim WNroIbTucu As String
Dim WNombre As String
Dim WDomicilio As String
Dim WPuerta As String
Dim WLocalidad As String
Dim WProvincia As String
Dim WPostal As String
Dim Provincia(100) As String

Private Sub Drive_Change()
    Dir1.Path = Drive.Drive
End Sub

Private Sub Acepta_Click()

    WDrive = Drive.Drive
    WDir = Dir1.Path
    
    XNombre = WDir + "\" + "Datos.Txt"
    Open XNombre For Output As #1
    
    XNombre = WDir + "\" + "RetPer.Txt"
    Open XNombre For Output As #2
    
    XNombre = WDir + "\" + "NcFact.Txt"
    Open XNombre For Output As #3

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia

    Renglon = 0
    ZLugarEntra = 0
    Erase Vector
    Erase ZEntra

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.ImpoIbTucu > 0"
    Rem ZSql = ZSql + " Where CtaCte.ImpoIb <> 0"
    ZSql = ZSql + " and CtaCte.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and CtaCte.OrdFecha <= " + "'" + WHasta + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
        With rstCtaCte
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstCtaCte!OrdFecha
                    Vector(Renglon, 2) = rstCtaCte!Cliente
                    Vector(Renglon, 3) = rstCtaCte!Tipo
                    Vector(Renglon, 4) = Str$(rstCtaCte!Numero)
                    Vector(Renglon, 5) = Str$(Abs(rstCtaCte!Neto))
                    Vector(Renglon, 6) = Str$(Abs(rstCtaCte!ImpoIbTucu))
                    Rem Vector(Renglon, 6) = Str$(rstCtaCte!ImpoIb)
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        End With
    End If


    For Cicla = 1 To Renglon
    
        WFecha = Vector(Cicla, 1)
        WCliente = Vector(Cicla, 2)
        WTipo = Vector(Cicla, 3)
        WNumero = Vector(Cicla, 4)
        WNeto = Val(Vector(Cicla, 5))
        WImpoIb = Val(Vector(Cicla, 6))
        WPorceIb = 1.75
        
        Call Redondeo(WPorceIb)
        XPorceIb = Str$(WPorceIb)
        Call Ceros(XPorceIb, 6)
        
        Call Redondeo(WImpoIb)
        XImpoIb = Str$(WImpoIb)
        Call Ceros(XImpoIb, 15)
        
        Call Redondeo(WNeto)
        XNeto = Str$(WNeto)
        Call Ceros(XNeto, 15)
            
        WCuit = ""
        WNroIbTucu = ""
        WNombre = ""
        WDomicilio = ""
        WPuerta = "00000"
        WLocalidad = ""
        WProvincia = "0"
        WPostal = ""
        
        spClientes = "ConsultaClientes " + "'" + WCliente + "'"
        Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
        If rstClientes.RecordCount > 0 Then
            WCuit = Left$(rstClientes!Cuit, 13)
            WNroIbTucu = IIf(IsNull(rstClientes!NroIbTucu), "", rstClientes!NroIbTucu)
            WNombre = rstClientes!Razon
            WDomicilio = rstClientes!Direccion
            WLocalidad = rstClientes!Localidad
            WPuerta = "00000"
            WProvincia = rstClientes!Provincia
            WPostal = rstClientes!Postal
            rstClientes.Close
        End If
        
        spCliente = "ConsultaCliente " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
            WPorceCm05Tucu = IIf(IsNull(rstCliente!PorceCm05Tucu), "0", rstCliente!PorceCm05Tucu)
            rstCliente.Close
        End If
        If WPorceCm05Tucu = 0 Then
            WPorceCm05Tucu = 1
        End If
    
        
        Select Case WCodIbTucu
            Case 1, 2, 3
                If WPorceCm05Tucu <> 1 Then
                    WNeto = WNeto * WPorceCm05Tucu
                    Call Redondeo(WNeto)
                    XNeto = Str$(WNeto)
                    Call Ceros(XNeto, 15)
                End If
            Case Else
        End Select
        
        
        Rem If Val(WNroIbTucu) <> 0 Then
        Rem     WNroIbTucu = Left$(WNroIbTucu, 3) + "-" + Mid$(WNroIbTucu, 4, 6) + "-" + Mid$(WNroIbTucu, 10, 1)
        Rem         Else
        Rem     WNroIbTucu = "99999999999"
        Rem End If
        If Val(WNroIbTucu) = 0 Then
            WNroIbTucu = "99999999999"
        End If
        
        Call Ceros(WNumero, 8)
        
        Call Eval
        Call Ceros(WCuit, 11)
        Call Ceros(WPostal, 4)
        
        
        Call Ceros(WNroIbTucu, 11)
        
        Rem fecha
        Campo1 = WFecha
        
        Rem tipo de documento
        Campo2 = "80"
        
        Rem documento
        Campo3 = WCuit
        
        Rem tipo de comprobante
        Select Case Val(WTipo)
            Case 1, 3
                Campo4 = "01"
            Case 4
                Campo4 = "02"
            Case Else
                Campo4 = "03"
        End Select
        
        Rem Letra de comprobante
        Campo5 = "A"
        
        Rem Punto de Venta
        Campo6 = "0001"
        
        Rem Numero del Comprobante
        Campo7 = WNumero
        
        Rem Numero del Comprobante
        Campo8 = XNeto
        
        Rem alicutoa
        Campo9 = XPorceIb
        
        Rem importe de la retencion
        Campo10 = XImpoIb
        
        Rem nor de ingresos brutos
        Rem Campo11 = Left$(WNroIbTucu + Space$(11), 11)
        Campo11 = ""
        
        WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5 + Campo6 + Campo7 + Campo8 + Campo9 + Campo10 + Campo11
        
        Print #1, WImpre
        
        WEntra = "S"
        
        For CicloII = 1 To ZLugarEntra
            If ZEntra(CicloII) = WCliente Then
                WEntra = "N"
                Exit For
            End If
        Next CicloII
        
        If WEntra = "S" Then
        
            Rem tipo de documento
            Campo1 = "80"
        
            Rem documento
            Campo2 = WCuit
        
            Rem razon Social
            Campo3 = Left$(WNombre + Space$(40), 40)
        
            Rem domicilio
            Campo4 = Left$(WDomicilio + Space$(40), 40)
        
            Rem altura
            Campo5 = WPuerta
        
            Rem Localidad
            Campo6 = Left$(WLocalidad + Space$(15), 15)
        
            Rem provincia
            WImpreProvincia = Provincia(Val(WProvincia))
            Campo7 = Left$(WImpreProvincia + Space$(15), 15)
            
            Rem nor de ingresos brutos
            Campo8 = Left$(WNroIbTucu + Space$(11), 11)
            
            Rem Codigo Postal
            Campo9 = "    " + WPostal
        
            WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5 + Campo6 + Campo7 + Campo8 + Campo9
        
            Print #2, WImpre
            
            ZLugarEntra = ZLugarEntra + 1
            ZEntra(ZLugarEntra) = WCliente
                    
        
        End If

    Next Cicla
    
    Close #1
    Close #2
    Close #3
    
    Call Cancela_Click
        
End Sub

Private Sub Cancela_Click()
    Desde.SetFocus
    PrgProcesoPerceIbTucuman.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""


    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
End Sub

Private Sub Eval()

    Es = WCuit

    x = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               x = x + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               x = x + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               x = x + Y: MinusOk = 0

        End If

    Next

    WCuit = x

End Sub


Private Sub MOratoria_Click()


    WDrive = Drive.Drive
    WDir = Dir1.Path
    
    XNombre = WDir + "\" + "Datos.Txt"
    Open XNombre For Output As #1
    
    XNombre = WDir + "\" + "RetPer.Txt"
    Open XNombre For Output As #2
    
    XNombre = WDir + "\" + "NcFact.Txt"
    Open XNombre For Output As #3

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia

    OPEN_FILE_Moratoria
    
    Renglon = 0
    ZLugarEntra = 0
    Erase Vector
    Erase ZEntra
    
    With rstMoratoria
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            
            Do
            
                ZZFecha = !Fecha
                WAno = Right$(!Fecha, 4)
                WMes = Mid$(!Fecha, 4, 2)
                WDia = Left$(!Fecha, 2)
                ZZOrdFecha = WAno + WMes + WDia
            
                If ZZOrdFecha >= WDesde And ZZOrdFecha <= WHasta Then
                
                    Renglon = Renglon + 1
                    
                    ZZRazon = !Razon
                    ZZCuit = !Cuit
                    ZZNeto = !gravado
                    ZZTipo = !Tipo
                    ZZNumero = !Numero
                    ZZImpuesto = !Impuesto
                    ZZPadron = !Padron
                    ZZAlicuota = !Alicuota
                    ZZCoeficiente = !Coeficiente
                    
                    ZZPorce = ZZAlicuota * ZZCoeficiente
                    
                    Vector(Renglon, 1) = ZZOrdFecha
                    Vector(Renglon, 2) = ZZRazon
                    Vector(Renglon, 3) = ZZTipo
                    Vector(Renglon, 4) = Str$(ZZNumero)
                    Vector(Renglon, 5) = Str$(Abs(ZZNeto))
                    Vector(Renglon, 6) = Str$(Abs(ZZImpuesto))
                    Vector(Renglon, 7) = ZZCuit
                    Vector(Renglon, 8) = Str$(ZZPadron)
                    Vector(Renglon, 9) = Str$(ZZPorce)
                    
                    Rem Vector(Renglon, 6) = Str$(rstCtaCte!ImpoIb)
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With


    For Cicla = 1 To Renglon
    
        WFecha = Vector(Cicla, 1)
        WRazon = Vector(Cicla, 2)
        WTipo = Vector(Cicla, 3)
        WNumero = Vector(Cicla, 4)
        WNeto = Val(Vector(Cicla, 5))
        WImpoIb = Val(Vector(Cicla, 6))
        WCuit = Vector(Cicla, 7)
        WNroIbTucu = Vector(Cicla, 8)
        WPorceIb = Val(Vector(Cicla, 9))
        
        Call Redondeo(WPorceIb)
        XPorceIb = Str$(WPorceIb)
        Call Ceros(XPorceIb, 6)
        
        Call Redondeo(WImpoIb)
        XImpoIb = Str$(WImpoIb)
        Call Ceros(XImpoIb, 15)
        
        Call Redondeo(WNeto)
        XNeto = Str$(WNeto)
        Call Ceros(XNeto, 15)
            
        Rem WCuit = ""
        Rem WNroIbTucu = ""
        WNombre = ""
        WDomicilio = ""
        WPuerta = "00000"
        WLocalidad = ""
        WProvincia = "0"
        WPostal = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cuit = " + "'" + WCuit + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Rem WNroIbTucu = IIf(IsNull(rstClientes!NroIbTucu), "", rstClientes!NroIbTucu)
            WNombre = rstCliente!Razon
            WDomicilio = rstCliente!Direccion
            WLocalidad = rstCliente!Localidad
            WPuerta = "00000"
            WProvincia = rstCliente!Provincia
            WPostal = rstCliente!Postal
            rstCliente.Close
        End If
        
        
        Rem If Val(WNroIbTucu) <> 0 Then
        Rem     WNroIbTucu = Left$(WNroIbTucu, 3) + "-" + Mid$(WNroIbTucu, 4, 6) + "-" + Mid$(WNroIbTucu, 10, 1)
        Rem         Else
        Rem     WNroIbTucu = "99999999999"
        Rem End If
        If Val(WNroIbTucu) = 0 Then
            WNroIbTucu = "99999999999"
        End If
        
        Call Ceros(WNumero, 8)
        
        Call Eval
        Call Ceros(WCuit, 11)
        Call Ceros(WPostal, 4)
        
        
        Call Ceros(WNroIbTucu, 11)
        
        Rem fecha
        Campo1 = WFecha
        
        Rem tipo de documento
        Campo2 = "80"
        
        Rem documento
        Campo3 = WCuit
        
        Rem tipo de comprobante
        Select Case WTipo
            Case "FC"
                Campo4 = "01"
            Case "ND"
                Campo4 = "02"
            Case Else
                Campo4 = "03"
        End Select
        
        Rem Letra de comprobante
        Campo5 = "A"
        
        Rem Punto de Venta
        Campo6 = "0001"
        
        Rem Numero del Comprobante
        Campo7 = WNumero
        
        Rem Numero del Comprobante
        Campo8 = XNeto
        
        Rem alicutoa
        Campo9 = XPorceIb
        
        Rem importe de la retencion
        Campo10 = XImpoIb
        
        Rem nor de ingresos brutos
        Campo11 = Left$(WNroIbTucu + Space$(11), 11)
        
        WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5 + Campo6 + Campo7 + Campo8 + Campo9 + Campo10 + Campo11
        
        Print #1, WImpre
        
        WEntra = "S"
        
        For CicloII = 1 To ZLugarEntra
            If ZEntra(CicloII) = WCuit Then
                WEntra = "N"
                Exit For
            End If
        Next CicloII
        
        If WEntra = "S" Then
        
            Rem tipo de documento
            Campo1 = "80"
        
            Rem documento
            Campo2 = WCuit
        
            Rem razon Social
            Campo3 = Left$(WNombre + Space$(40), 40)
        
            Rem domicilio
            Campo4 = Left$(WDomicilio + Space$(40), 40)
        
            Rem altura
            Campo5 = WPuerta
        
            Rem Localidad
            Campo6 = Left$(WLocalidad + Space$(15), 15)
        
            Rem provincia
            WImpreProvincia = Provincia(Val(WProvincia))
            Campo7 = Left$(WImpreProvincia + Space$(15), 15)
            
            Rem nor de ingresos brutos
            Campo8 = Left$(WNroIbTucu + Space$(11), 11)
            
            Rem Codigo Postal
            Campo9 = "    " + WPostal
        
            WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5 + Campo6 + Campo7 + Campo8 + Campo9
        
            Print #2, WImpre
            
            ZLugarEntra = ZLugarEntra + 1
            ZEntra(ZLugarEntra) = WCuit
                    
        
        End If

    Next Cicla
    
    Close #1
    Close #2
    Close #3
    
    Call Cancela_Click
        



End Sub
