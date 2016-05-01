VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgGrabaCd 
   AutoRedraw      =   -1  'True
   Caption         =   "Grabacion electronica de Datos"
   ClientHeight    =   2970
   ClientLeft      =   3315
   ClientTop       =   2175
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   2970
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      Begin VB.TextBox Cai 
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
         Left            =   1680
         MaxLength       =   14
         TabIndex        =   10
         Text            =   " "
         Top             =   1320
         Width           =   2055
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   3360
         TabIndex        =   5
         Top             =   360
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
         Left            =   3360
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin MSMask.MaskEdBox VtoCai 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label4 
         Caption         =   "Cai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Vto. Cai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgGrabaCd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZCampo1 As String
Dim ZCampo2 As String
Dim ZCampo3 As String
Dim ZCampo4 As String
Dim ZCampo5 As String
Dim ZCampo6 As String
Dim ZCampo7 As String
Dim ZCampo8 As String
Dim ZCampo9 As String
Dim ZCampo10 As String
Dim ZCampo11 As String
Dim ZCampo12 As String
Dim ZCampo13 As String
Dim ZCampo14 As String
Dim ZCampo15 As String
Dim ZCampo16 As String
Dim ZCampo17 As String
Dim ZCampo18 As String
Dim ZCampo19 As String
Dim ZCampo20 As String
Dim ZCampo21 As String
Dim ZCampo22 As String
Dim ZCampo23 As String
Dim ZCampo24 As String
Dim ZCampo25 As String
Dim ZCampo26 As String
Dim ZCampo27 As String
Dim ZCampo28 As String
Dim ZCampo29 As String
Dim ZCampo30 As String

Dim WCuit As String
Dim WWNeto As Double
Dim WWIva1 As Double
Dim WWIva2 As Double
Dim WWImpoIb As Double
Dim WWParidad As Double
Dim WWIva21 As Double
Dim WWIva27 As Double
Dim WWIva5 As Double
Dim WWIb As Double
Dim WWExento As Double

Dim rstCtacte As Recordset
Dim spCtacte As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Acepta_Click()

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    WAno = Right$(VtoCai.Text, 4)
    WMes = Mid$(VtoCai.Text, 4, 2)
    WDia = Left$(VtoCai.Text, 2)
    ZVtoCai = WAno + WMes + WDia
    
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    
    Open "c:\rg1361\VENTAS_" + WAno + WMes + ".txt" For Output As #1
    
    ZRegistros = 0
    
    ZSuma8 = 0
    ZSuma9 = 0
    ZSuma10 = 0
    ZSuma12 = 0
    ZSuma13 = 0
    ZSuma14 = 0
    ZSuma15 = 0
    ZSuma16 = 0
    ZSuma17 = 0
    ZSuma18 = 0
    
    ZSql = ""
    ZSql = ZSql + "Select Ctacte.Tipo, Ctacte.Numero, Ctacte.Cliente, Ctacte.Fecha, Ctacte.OrdFecha, Ctacte.Neto, Ctacte.Iva1, Ctacte.Iva2, Ctacte.ImpoIb, CtaCte.Paridad, Cliente.Cuit as [ClienteCuit], Cliente.Razon as [ClienteRazon], Cliente.Provincia as [ClienteProvincia], Cliente.Iva as [ClienteIva]"
    ZSql = ZSql + " FROM Ctacte, Cliente"
    ZSql = ZSql + " Where Ctacte.Cliente = Cliente.Cliente"
    ZSql = ZSql + " and Ctacte.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Ctacte.OrdFecha <= " + "'" + WHasta + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
        With rstCtacte
            .MoveFirst
            Do
            
                If Val(rstCtacte!Tipo) >= 1 And Val(rstCtacte!Tipo) <= 5 Then
                
                    WTipo = rstCtacte!Tipo
                    WNumero = rstCtacte!Numero
                    WCliente = rstCtacte!Cliente
                    WFecha = rstCtacte!Fecha
                    WOrdFecha = rstCtacte!OrdFecha
                    WWNeto = Abs(rstCtacte!Neto)
                    WWIva1 = Abs(rstCtacte!Iva1)
                    WWIva2 = Abs(rstCtacte!Iva2)
                    WWImpoIb = Abs(rstCtacte!ImpoIb)
                    WWParidad = rstCtacte!Paridad
                    Call Redondeo(WWNeto)
                    Call Redondeo(WWIva1)
                    Call Redondeo(WWIva2)
                    Call Redondeo(WWImpoIb)
                    Call Redondeo(WWParidad)
                    WNeto = Int(WWNeto * 100)
                    WIva1 = Int(WWIva1 * 100)
                    WIva2 = Int(WWIva2 * 100)
                    WImpoIb = Int(WWImpoIb * 100)
                    WParidad = Int(WWParidad * 100) * 10000
                    WTotal = WNeto + WIva1 + WIva2 + WImpoIb
                    WCuit = rstCtacte!ClienteCuit
                    Call Eval
                    WRazon = rstCtacte!ClienteRazon
                    WProvincia = rstCtacte!ClienteProvincia
                    WIva = rstCtacte!ClienteIva
                    
                    Rem TIPO DE REGISTRO
                    ZCampo1 = "1"
                    
                    Rem FECHA DEL COMPREOBANTE AAAAMMDD
                    ZCampo2 = WOrdFecha
                    
                    Rem TIPO DE COMPROBANTE
                    Rem 01-FAC  02-ND  03-NC  19-EXPO
                    If WNumero > 800000 Then
                        ZCampo3 = "19"
                            Else
                        If Val(WTipo) = 1 Or Val(WTipo) = 3 Then
                            ZCampo3 = "01"
                                Else
                            If Val(WTipo) = 4 Then
                                ZCampo3 = "02"
                                    Else
                                ZCampo3 = "03"
                            End If
                        End If
                    End If
                    
                    Rem CONTROLADOR FISCAL
                    ZCampo4 = Space$(1)
                    
                    Rem PUNTO DE VENTA
                    ZCampo5 = "0001"
                    
                    Rem DESDE NUMERO DEL COMPROBANTE
                    If WNumero > 800000 Then
                        ZCampo6 = Str$(Val(WNumero) - 800000)
                            Else
                        ZCampo6 = WNumero
                    End If
                    
                    Rem HASTA NUMERO DEL COMPROBANTE
                    ZCampo7 = ZCampo6
                    
                    Rem CODIGO DE DOCUMENTO DEL CLIENTE
                    ZCampo8 = "80"
                    
                    Rem NUMERO DE DOCUEMNTO DEL CLIENTE
                    ZCampo9 = WCuit
                    
                    
                    Rem RAZON SOCIAL
                    ZCampo10 = WRazon + Space$(30)
                    ZCampo10 = Left$(ZCampo10, 30)
                    
                    Rem IMPORTE TOTAL DE LA OPERACION
                    ZCampo11 = Str$(WTotal)
                    
                    Rem OPERACIONES EXENTAS Y OPERARACIONES GRABADAS
                    If WIva1 <> 0 Then
                        ZCampo12 = "0"
                        ZCampo13 = Str$(WNeto)
                        ZCampo14 = "2100"
                        ZCampo15 = Str$(WIva1)
                        ZCampo16 = "0"
                        ZCampo17 = "0"
                            Else
                        If Val(WNumero) > 800000 Or WProvincia = 23 Then
                            ZCampo12 = "0"
                            ZCampo13 = "0"
                            ZCampo14 = "0"
                            ZCampo15 = "0"
                            ZCampo16 = "0"
                            ZCampo17 = Str$(WNeto)
                                Else
                            ZCampo12 = Str$(WNeto)
                            ZCampo13 = "0"
                            ZCampo14 = "0"
                            ZCampo15 = "0"
                            ZCampo16 = "0"
                            ZCampo17 = "0"
                        End If
                    End If
                    
                    Rem IMPUESTOS NACIONALES
                    ZCampo18 = "0"
                    
                    Rem percepcion de ingresos brutos
                    ZCampo19 = Str$(WImpoIb)
                    
                    Rem IMPUESTOS MUNICIPALES
                    ZCampo20 = "0"
                    
                    Rem IMPUESTOS INTERNOS
                    ZCampo21 = "0"
                            
                    Rem CATEGORIA DE IVA
                    Select Case WIva
                        Case 2
                            ZCampo22 = "02"
                        Case 3
                            ZCampo22 = "05"
                        Case 4
                            ZCampo22 = "04"
                        Case 5
                            ZCampo22 = "06"
                        Case Else
                            ZCampo22 = "01"
                    End Select
                    
                    Rem PARIDAD Y TIPO DE MONEDA
                    If WParidad <> 0 Then
                        ZCampo23 = "DOL"
                        ZCampo24 = Str$(WParidad)
                            Else
                        ZCampo23 = "PES"
                        ZCampo24 = Str$(WParidad)
                    End If
                    
                    Rem CANTIDAD DE ALICUOTTAS DE IVA
                    ZCampo25 = "1"
                    
                    Rem CODIGO DE OPERACION
                    If WIva1 <> 0 Then
                        ZCampo26 = Space$(1)
                            Else
                        If Val(WNumero) > 800000 Then
                            ZCampo26 = "X"
                                Else
                            If WTotal <> 0 Then
                                ZCampo26 = "E"
                                    Else
                                ZCampo26 = Space$(1)
                            End If
                        End If
                    End If
                    
                    Rem CAI
                    ZCampo27 = Cai.Text
                    
                    Rem FECHA DE VENCIMIENTO DEL CAI
                    ZCampo28 = ZVtoCai
                    
                    If WTotal <> 0 Then
                        ZCampo29 = "00000000"
                            Else
                        ZCampo29 = WOrdFecha
                    End If
                    
                    Rem VARIOS
                    ZCampo30 = Space$(75)
                    
                    Call Ceros(ZCampo6, 20)
                    Call Ceros(ZCampo7, 20)
                    Call Ceros(ZCampo9, 11)
                    Call Ceros(ZCampo11, 15)
                    Call Ceros(ZCampo12, 15)
                    Call Ceros(ZCampo13, 15)
                    Call Ceros(ZCampo14, 4)
                    Call Ceros(ZCampo15, 15)
                    Call Ceros(ZCampo16, 15)
                    Call Ceros(ZCampo17, 15)
                    Call Ceros(ZCampo18, 15)
                    Call Ceros(ZCampo19, 15)
                    Call Ceros(ZCampo20, 15)
                    Call Ceros(ZCampo21, 15)
                    Call Ceros(ZCampo24, 10)
                    Call Ceros(ZCampo27, 14)
                    
                    GrabaCampo = ZCampo1 + ZCampo2 + ZCampo3 + ZCampo4 + ZCampo5 + ZCampo6 + ZCampo7 + ZCampo8 + ZCampo9 + ZCampo10 + ZCampo11 + ZCampo12 + ZCampo13 + ZCampo14 + ZCampo15 + ZCampo16 + ZCampo17 + ZCampo18 + ZCampo19 + ZCampo20 + ZCampo21 + ZCampo22 + ZCampo23 + ZCampo24 + ZCampo25 + ZCampo26 + ZCampo27 + ZCampo28 + ZCampo29 + ZCampo30
                    Print #1, GrabaCampo
                    
                    ZRegistros = ZRegistros + 1
                    
                    If Val(ZCampo3) <> 3 Then
                    
                        ZSuma8 = ZSuma8 + Val(ZCampo11)
                        ZSuma9 = ZSuma9 + Val(ZCampo12)
                        ZSuma10 = ZSuma10 + Val(ZCampo13)
                        ZSuma12 = ZSuma12 + Val(ZCampo15)
                        ZSuma13 = ZSuma13 + Val(ZCampo16)
                        ZSuma14 = ZSuma14 + Val(ZCampo17)
                        ZSuma15 = ZSuma15 + Val(ZCampo18)
                        ZSuma16 = ZSuma16 + Val(ZCampo19)
                        ZSuma17 = ZSuma17 + Val(ZCampo20)
                        ZSuma18 = ZSuma18 + Val(ZCampo21)
                        
                            Else
                    
                        ZSuma8 = ZSuma8 - Val(ZCampo11)
                        ZSuma9 = ZSuma9 - Val(ZCampo12)
                        ZSuma10 = ZSuma10 - Val(ZCampo13)
                        ZSuma12 = ZSuma12 - Val(ZCampo15)
                        ZSuma13 = ZSuma13 - Val(ZCampo16)
                        ZSuma14 = ZSuma14 - Val(ZCampo17)
                        ZSuma15 = ZSuma15 - Val(ZCampo18)
                        ZSuma16 = ZSuma16 - Val(ZCampo19)
                        ZSuma17 = ZSuma17 - Val(ZCampo20)
                        ZSuma18 = ZSuma18 - Val(ZCampo21)
                    
                    End If
                    Rem If ZRegistros = 3 Then Exit Do
                                              
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        
        rstCtacte.Close
        
    End If
    
    Rem Graba registro Tipo 2
                    
    Rem TIPO DE REGISTRO
    ZCampo1 = "2"
                    
    Rem FECHA
    ZCampo2 = Left$(WHasta, 6)
                    
    Rem RELELNO
    ZCampo3 = Space$(29)
                    
    Rem CANTIDAD DE REGISTROS NRO 1
    ZCampo4 = Str$(ZRegistros)
                    
    Rem RELLENO
    ZCampo5 = Space$(10)
                    
    Rem CUIT
    ZCampo6 = "30549165083"
                    
    Rem RELLENO
    ZCampo7 = Space$(30)
                    
    Rem
    ZCampo8 = Str$(ZSuma8)
                    
    Rem
    ZCampo9 = Str$(ZSuma9)
                    
    Rem
    ZCampo10 = Str$(ZSuma10)
                    
    Rem RELLENO
    ZCampo11 = Space$(4)
                    
    Rem
    ZCampo12 = Str$(ZSuma12)
                    
    Rem
    ZCampo13 = Str$(ZSuma13)
                    
    Rem
    ZCampo14 = Str$(ZSuma14)
                    
    Rem
    ZCampo15 = Str$(ZSuma15)
                    
    Rem
    ZCampo16 = Str$(ZSuma16)
                    
    Rem
    ZCampo17 = Str$(ZSuma17)
                    
    Rem
    ZCampo18 = Str$(ZSuma18)
                    
    Rem RELLENO
    ZCampo19 = Space$(122)
    
    Call Ceros(ZCampo4, 12)
    Call Ceros(ZCampo8, 15)
    Call Ceros(ZCampo9, 15)
    Call Ceros(ZCampo10, 15)
    Call Ceros(ZCampo12, 15)
    Call Ceros(ZCampo13, 15)
    Call Ceros(ZCampo14, 15)
    Call Ceros(ZCampo15, 15)
    Call Ceros(ZCampo16, 15)
    Call Ceros(ZCampo17, 15)
    Call Ceros(ZCampo18, 15)

    GrabaCampo = ZCampo1 + ZCampo2 + ZCampo3 + ZCampo4 + ZCampo5 + ZCampo6 + ZCampo7 + ZCampo8 + ZCampo9 + ZCampo10 + ZCampo11 + ZCampo12 + ZCampo13 + ZCampo14 + ZCampo15 + ZCampo16 + ZCampo17 + ZCampo18 + ZCampo19
    Print #1, GrabaCampo
    
    Close #1
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Open "c:\rg1361\OTRAS_PERCEP_" + WAno + WMes + ".txt" For Output As #1
    
    ZRegistros = 0
    
    ZSuma8 = 0
    ZSuma9 = 0
    ZSuma10 = 0
    ZSuma12 = 0
    ZSuma13 = 0
    ZSuma14 = 0
    ZSuma15 = 0
    ZSuma16 = 0
    ZSuma17 = 0
    ZSuma18 = 0
    
    ZSql = ""
    ZSql = ZSql + "Select Ctacte.Tipo, Ctacte.Numero, Ctacte.Cliente, Ctacte.Fecha, Ctacte.OrdFecha, Ctacte.Neto, Ctacte.Iva1, Ctacte.Iva2, Ctacte.ImpoIb, CtaCte.Paridad, Cliente.Cuit as [ClienteCuit], Cliente.Razon as [ClienteRazon], Cliente.Provincia as [ClienteProvincia], Cliente.Iva as [ClienteIva]"
    ZSql = ZSql + " FROM Ctacte, Cliente"
    ZSql = ZSql + " Where Ctacte.Cliente = Cliente.Cliente"
    ZSql = ZSql + " and Ctacte.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Ctacte.OrdFecha <= " + "'" + WHasta + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
        With rstCtacte
            .MoveFirst
            Do
            
                If Val(rstCtacte!Tipo) >= 1 And Val(rstCtacte!Tipo) <= 5 Then
                
                If rstCtacte!ImpoIb <> 0 Then
                
                    WTipo = rstCtacte!Tipo
                    WNumero = rstCtacte!Numero
                    WCliente = rstCtacte!Cliente
                    WFecha = rstCtacte!Fecha
                    WOrdFecha = rstCtacte!OrdFecha
                    WWNeto = Abs(rstCtacte!Neto)
                    WWIva1 = Abs(rstCtacte!Iva1)
                    WWIva2 = Abs(rstCtacte!Iva2)
                    WWImpoIb = Abs(rstCtacte!ImpoIb)
                    WWParidad = rstCtacte!Paridad
                    Call Redondeo(WWNeto)
                    Call Redondeo(WWIva1)
                    Call Redondeo(WWIva2)
                    Call Redondeo(WWImpoIb)
                    Call Redondeo(WWParidad)
                    WNeto = Int(WWNeto * 100)
                    WIva1 = Int(WWIva1 * 100)
                    WIva2 = Int(WWIva2 * 100)
                    WImpoIb = Int(WWImpoIb * 100)
                    WParidad = Int(WWParidad * 100) * 10000
                    WTotal = WNeto + WIva1 + WIva2 + WImpoIb
                    WCuit = rstCtacte!ClienteCuit
                    Call Eval
                    WRazon = rstCtacte!ClienteRazon
                    WProvincia = rstCtacte!ClienteProvincia
                    WIva = rstCtacte!ClienteIva
                    
                    Rem FECHA DEL COMPREOBANTE AAAAMMDD
                    ZCampo1 = WOrdFecha
                    
                    Rem TIPO DE COMPROBANTE
                    Rem 01-FAC  02-ND  03-NC  19-EXPO
                    If WNumero > 800000 Then
                        ZCampo2 = "19"
                            Else
                        If Val(WTipo) = 1 Or Val(WTipo) = 3 Then
                            ZCampo2 = "01"
                                Else
                            If Val(WTipo) = 4 Then
                                ZCampo2 = "02"
                                    Else
                                ZCampo2 = "03"
                            End If
                        End If
                    End If
                    
                    Rem PUNTO DE VENTA
                    ZCampo3 = "0001"
                    
                    Rem DESDE NUMERO DEL COMPROBANTE
                    If WNumero > 800000 Then
                        ZCampo4 = Str$(Val(WNumero) - 800000)
                            Else
                        ZCampo4 = WNumero
                    End If
                    
                    Rem
                    ZCampo5 = "01"
                    
                    Rem percepcion de ingresos brutos
                    ZCampo6 = Str$(WImpoIb)
                    
                    Rem
                    ZCampo7 = Space(40)
                    
                    Rem IMPUESTOS INTERNOS
                    ZCampo8 = "0"
                            
                    Call Ceros(ZCampo4, 8)
                    Call Ceros(ZCampo6, 15)
                    Call Ceros(ZCampo8, 15)
                    
                    GrabaCampo = ZCampo1 + ZCampo2 + ZCampo3 + ZCampo4 + ZCampo5 + ZCampo6 + ZCampo7 + ZCampo8
                    Print #1, GrabaCampo
                    
                End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        
        rstCtacte.Close
        
    End If
    
    Close #1
    
    
    
    
    
    
    
    
    
    
    
    
    Open "c:\rg1361\COMPRAS_" + WAno + WMes + ".txt" For Output As #1
    
    ZRegistros = 0
    
    ZSuma17 = 0
    ZSuma18 = 0
    ZSuma19 = 0
    ZSuma21 = 0
    ZSuma22 = 0
    ZSuma23 = 0
    ZSuma24 = 0
    ZSuma25 = 0
    ZSuma26 = 0
    
    ZSql = ""
    ZSql = ZSql + "Select IvaComp.Periodo, IvaComp.Proveedor, IvaComp.Tipo, IvaComp.Letra, IvaComp.Punto, IvaComp.Numero, IvaComp.Fecha, IvaComp.Vencimiento, IvaComp.Neto, IvaComp.Iva21, IvaComp.Iva5, IvaComp.Iva27, IvaComp.Ib, IvaComp.Exento, IvaComp.OrdFecha, IvaComp.Pago, IvaComp.Paridad, IvaComp.Cai, IvaComp.VtoCai, "
    ZSql = ZSql + "Proveedor.Cuit as [ProveedorCuit], Proveedor.Nombre as [ProveedorNombre], Proveedor.Iva as [ProveedorIva], Proveedor.Cai as [ProveedorCai], Proveedor.VtoCai as [ProveedorVtoCai]"
    ZSql = ZSql + " FROM Ivacomp, Proveedor"
    ZSql = ZSql + " Where IvaComp.Proveedor = Proveedor.Proveedor"
    ZSql = ZSql + " and IvaComp.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and IvaComp.OrdFecha <= " + "'" + WHasta + "'"
    
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        With rstIvaComp
            .MoveFirst
            Do
                WFecha = Right$(rstIvaComp!Periodo, 4) + Mid$(rstIvaComp!Periodo, 4, 2) + Left$(rstIvaComp!Periodo, 2)
                WCuit = rstIvaComp!ProveedorCuit
                Call Eval
                Call Ceros(WCuit, 11)
                
                WCai = IIf(IsNull(rstIvaComp!Cai), "", rstIvaComp!Cai)
                WVtoCai = IIf(IsNull(rstIvaComp!VtoCai), "", rstIvaComp!VtoCai)
                    
                WProveedorCai = IIf(IsNull(rstIvaComp!ProveedorCai), "", rstIvaComp!ProveedorCai)
                WProveedorVtoCai = IIf(IsNull(rstIvaComp!ProveedorVtoCai), "", rstIvaComp!ProveedorVtoCai)
                
                If WDesde <= WFecha And WFecha <= WHasta Then
                
                Rem If Left$(WCuit, 1) = "2" Or Left$(WCuit, 1) = "3" Then
                
                Rem If Val(WCai) <> 0 Or Val(WProveedorCai) <> 0 Then
                
                    WProveedor = rstIvaComp!Proveedor
                    WTipo = rstIvaComp!Tipo
                    WLetra = rstIvaComp!Letra
                    WPunto = rstIvaComp!Punto
                    WNumero = rstIvaComp!Numero
                    WFecha = rstIvaComp!Fecha
                    Wvencimiento = rstIvaComp!Vencimiento
                    WPeriodo = rstIvaComp!Periodo
                    WNombre = rstIvaComp!ProveedorNombre
                    
                    WCai = IIf(IsNull(rstIvaComp!Cai), "", rstIvaComp!Cai)
                    WVtoCai = IIf(IsNull(rstIvaComp!VtoCai), "", rstIvaComp!VtoCai)
                    
                    WProveedorCai = IIf(IsNull(rstIvaComp!ProveedorCai), "", rstIvaComp!ProveedorCai)
                    WProveedorVtoCai = IIf(IsNull(rstIvaComp!ProveedorVtoCai), "", rstIvaComp!ProveedorVtoCai)
                    
                    WAno = Right$(WVtoCai, 4)
                    WMes = Mid$(WVtoCai, 4, 2)
                    WDia = Left$(WVtoCai, 2)
                    WOrdVtoCai = WAno + WMes + WDia
                    
                    WAno = Right$(WProveedorVtoCai, 4)
                    WMes = Mid$(WProveedorVtoCai, 4, 2)
                    WDia = Left$(WProveedorVtoCai, 2)
                    WOrdProveedorVtoCai = WAno + WMes + WDia
                    
                    WWNeto = Abs(rstIvaComp!Neto)
                    WWIva21 = Abs(rstIvaComp!Iva21)
                    WWIva5 = Abs(rstIvaComp!Iva5)
                    WWIva27 = Abs(rstIvaComp!Iva27)
                    WWIb = Abs(rstIvaComp!Ib)
                    WWExento = Abs(rstIvaComp!Exento)
                    
                    Call Redondeo(WWNeto)
                    Call Redondeo(WWIva21)
                    Call Redondeo(WWIva5)
                    Call Redondeo(WWIva27)
                    Call Redondeo(WWIb)
                    Call Redondeo(WWExento)
                    
                    WNeto = WWNeto * 100
                    WIva21 = WWIva21 * 100
                    WIva5 = WWIva5 * 100
                    WIva27 = WWIva27 * 100
                    WIb = WWIb * 100
                    WExento = WWExento * 100
                    
                    WNeto = Int(WNeto)
                    WIva21 = Int(WIva21)
                    WIva5 = Int(WIva5)
                    WIva27 = Int(WIva27)
                    WIb = Int(WIb)
                    WExento = Int(WExento)
                    
                    
                    
                    WOrdFecha = rstIvaComp!OrdFecha
                    WIva = rstIvaComp!ProveedorIva
                    WPago = rstIvaComp!Pago
                    
                    WWParidad = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
                    WParidad = Int(WWParidad * 100) * 10000
                    
                    WTotal = WNeto + WIva21 + WIva5 + WIva27 + WIb + WExento
                
                    Rem TIPO DE REGISTRO
                    ZCampo1 = "1"
                    
                    Rem FECHA DEL COMPREOBANTE AAAAMMDD
                    ZCampo2 = WOrdFecha
                    
                    Rem TIPO DE COMPROBANTE
                    Rem 01-FAC  02-ND  03-NC
                    If Val(WTipo) = 1 Then
                        ZCampo3 = "01"
                            Else
                        If Val(WTipo) = 2 Then
                            ZCampo3 = "02"
                                Else
                            ZCampo3 = "03"
                        End If
                    End If
                    
                    Rem CONTROLADOR FISCAL
                    ZCampo4 = Space$(1)
                    
                    Rem PUNTO DE VENTA
                    ZCampo5 = WPunto
                    Call Ceros(ZCampo5, 4)
                    
                    ZCampo5 = ZCampo5 + "000000000000"
                    
                    Rem DESDE NUMERO DEL COMPROBANTE
                    ZCampo6 = WNumero
                    Call Ceros(ZCampo6, 8)
                    
                    Rem numero de comprobantes
                    ZCampo7 = ""
                    
                    Rem Año del Documento
                    ZCampo8 = WOrdFecha
                    
                    Rem Codigo de Aduana
                    ZCampo9 = "000"
                    
                    Rem Codigo de destinacion
                    ZCampo10 = "    "
                    
                    Rem Numero de Despacho
                    ZCampo11 = "000000"
                    
                    Rem Digiti verificador
                    ZCampo12 = " "
                    
                    Rem fecha de despacho a plaza
                    ZCampo13 = ""
                    
                    Rem tipo de documento
                    ZCampo14 = "80"
                    
                    Rem tipo de documento
                    ZCampo15 = WCuit
                    
                    Rem RAZON SOCIAL
                    ZCampo16 = WNombre + Space$(30)
                    ZCampo16 = Left$(ZCampo16, 30)
                    
                    Rem IMPORTE TOTAL DE LA OPERACION
                    ZCampo17 = Str$(WTotal)
                    Call Ceros(ZCampo17, 15)
                    
                    Rem exen to
                    ZCampo18 = Str$(WExento)
                    Call Ceros(ZCampo18, 15)
                    
                    Rem neto
                    ZCampo19 = Str$(WNeto)
                    Call Ceros(ZCampo19, 15)
                    
                    Rem alicuotra del oiva
                    ZCampo20 = "2100"
                    Call Ceros(ZCampo20, 4)
                    
                    Rem importe del iva
                    ZCampo21 = Str$(WIva21 + WIva27)
                    Call Ceros(ZCampo21, 15)
                    
                    Rem importes excentos
                    ZCampo22 = "0"
                    Call Ceros(ZCampo22, 15)
                    
                    Rem percepcion de impuestos nacioneles
                    ZCampo23 = Str$(WIva5)
                    Call Ceros(ZCampo23, 15)
                    
                    Rem percepcion de ingresos brutos
                    ZCampo24 = Str$(WIb)
                    Call Ceros(ZCampo24, 15)
                    
                    Rem IMPUESTOS MUNICIPALES
                    ZCampo25 = "0"
                    Call Ceros(ZCampo25, 15)
                    
                    Rem IMPUESTOS INTERNOS
                    ZCampo26 = "0"
                    Call Ceros(ZCampo26, 30)
                            
                    Rem CATEGORIA DE IVA
                    Select Case Val(WIva)
                        Case 1
                            ZCampo27 = "02"
                        Case 2
                            ZCampo27 = "05"
                        Case 3
                            ZCampo27 = "01"
                        Case 4
                            ZCampo27 = "04"
                        Case 5
                            ZCampo27 = "03"
                        Case 6
                            ZCampo27 = "06"
                        Case 7
                            ZCampo27 = "07"
                        Case Else
                            ZCampo27 = "01"
                    End Select
                    
                    Rem PARIDAD Y TIPO DE MONEDA
                    If WPago = 2 Then
                        ZCampo28 = "DOL"
                        ZCampo29 = Str$(WParidad)
                            Else
                        ZCampo28 = "PES"
                        ZCampo29 = Str$(WParidad)
                    End If
                    Call Ceros(ZCampo29, 10)
                    
                    Rem CANTIDAD DE ALICUOTTAS DE IVA
                    ZCampo30 = "1"
                    
                    Rem CODIGO DE OPERACION
                    If WNeto = 0 Then
                        ZCampo31 = "E"
                            Else
                        ZCampo31 = Space$(1)
                    End If
                    
                    Rem CAI
                    Rem FECHA DE VENCIMIENTO DEL CAI
                    If Val(WCai) <> 0 Then
                        ZCampo32 = WCai
                        ZCampo33 = WOrdVtoCai
                            Else
                        If Val(WProveedorCai) <> 0 Then
                            ZCampo32 = WProveedorCai
                            ZCampo33 = WOrdProveedorVtoCai
                                Else
                            ZCampo32 = "00000000000000"
                            ZCampo33 = "00000000"
                        End If
                    End If
                    
                    Rem VARIOS
                    ZCampo34 = Space$(75)
                    
                    GrabaCampo = ZCampo1 + ZCampo2 + ZCampo3 + ZCampo4 + ZCampo5 + ZCampo6 + ZCampo7 + ZCampo8 + ZCampo9 + ZCampo10 + ZCampo11 + ZCampo12 + ZCampo13 + ZCampo14 + ZCampo15 + ZCampo16 + ZCampo17 + ZCampo18 + ZCampo19 + ZCampo20 + ZCampo21 + ZCampo22 + ZCampo23 + ZCampo24 + ZCampo25 + ZCampo26 + ZCampo27 + ZCampo28 + ZCampo29 + ZCampo30 + ZCampo31 + ZCampo32 + ZCampo33 + ZCampo34
                    Print #1, GrabaCampo
                    
                    ZRegistros = ZRegistros + 1
                    
                    If Val(ZCampo3) = 3 Then
                    
                        ZSuma17 = ZSuma17 - Val(ZCampo17)
                        ZSuma18 = ZSuma18 - Val(ZCampo18)
                        ZSuma19 = ZSuma19 - Val(ZCampo19)
                    
                        ZSuma21 = ZSuma21 - Val(ZCampo21)
                        ZSuma22 = ZSuma22 - Val(ZCampo22)
                        ZSuma23 = ZSuma23 - Val(ZCampo23)
                        ZSuma24 = ZSuma24 - Val(ZCampo24)
                        ZSuma25 = ZSuma25 - Val(ZCampo25)
                        ZSuma26 = ZSuma26 - Val(ZCampo26)
                        
                            Else
                            
                        ZSuma17 = ZSuma17 + Val(ZCampo17)
                        ZSuma18 = ZSuma18 + Val(ZCampo18)
                        ZSuma19 = ZSuma19 + Val(ZCampo19)
                    
                        ZSuma21 = ZSuma21 + Val(ZCampo21)
                        ZSuma22 = ZSuma22 + Val(ZCampo22)
                        ZSuma23 = ZSuma23 + Val(ZCampo23)
                        ZSuma24 = ZSuma24 + Val(ZCampo24)
                        ZSuma25 = ZSuma25 + Val(ZCampo25)
                        ZSuma26 = ZSuma26 + Val(ZCampo26)
                        
                    End If
                    
                    Rem If ZRegistros = 10 Then Exit Do
                    
                Rem End If
                
                Rem End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            
            Loop
        End With
        
        rstIvaComp.Close
        
    End If
    
    
    Rem Graba registro Tipo 2
                    
    Rem TIPO DE REGISTRO
    ZCampo1 = "2"
                    
    Rem FECHA
    ZCampo2 = Left$(WHasta, 6)
                    
    Rem RELELNO
    ZCampo3 = Space$(10)
                    
    Rem CANTIDAD DE REGISTROS NRO 1
    ZCampo4 = Str$(ZRegistros)
                    
    Rem RELLENO
    ZCampo5 = Space$(31)
                    
    Rem CUIT
    ZCampo6 = "30549165083"
                    
    Rem RELLENO
    ZCampo7 = Space$(30)
                    
    Rem
    ZCampo8 = Str$(ZSuma17)
                    
    Rem
    ZCampo9 = Str$(ZSuma18)
                    
    Rem
    ZCampo10 = Str$(ZSuma19)
                    
    Rem RELLENO
    ZCampo11 = Space$(4)
                    
    Rem
    ZCampo12 = Str$(ZSuma21)
                    
    Rem
    ZCampo13 = Str$(ZSuma22)
                    
    Rem
    ZCampo14 = Str$(ZSuma23)
                    
    Rem
    ZCampo15 = Str$(ZSuma24)
                    
    Rem
    ZCampo16 = Str$(ZSuma25)
                    
    Rem
    ZCampo17 = Str$(ZSuma26)
                    
    Rem RELLENO
    ZCampo18 = Space$(114)
    
    Call Ceros(ZCampo4, 12)
    Call Ceros(ZCampo8, 15)
    Call Ceros(ZCampo9, 15)
    Call Ceros(ZCampo10, 15)
    Call Ceros(ZCampo12, 15)
    Call Ceros(ZCampo13, 15)
    Call Ceros(ZCampo14, 15)
    Call Ceros(ZCampo15, 15)
    Call Ceros(ZCampo16, 15)
    Call Ceros(ZCampo17, 30)

    GrabaCampo = ZCampo1 + ZCampo2 + ZCampo3 + ZCampo4 + ZCampo5 + ZCampo6 + ZCampo7 + ZCampo8 + ZCampo9 + ZCampo10 + ZCampo11 + ZCampo12 + ZCampo13 + ZCampo14 + ZCampo15 + ZCampo16 + ZCampo17 + ZCampo18
    Print #1, GrabaCampo
    
    Close #1
    
    
    
    
    
    
    
    
    
    
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Desde.SetFocus
    PrgIvaven.Hide
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
            Cai.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub Cai_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VtoCai.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub VtoCai_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(VtoCai.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            VtoCai.SetFocus
        End If
    End If
End Sub

Sub Form_Load()



    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Cai.Text = "25016153843724"
    VtoCai.Text = "14/04/2007"
End Sub

Private Sub Eval()

    Es = WCuit

    x = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If (Y = "-" Or Y = ".") And MinusOk = 1 Then
               x = x + Y: MinusOk = 0

        Rem ElseIf Y = "." And DecOk = 1 Then
        Rem        x = x + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               x = x + Y: MinusOk = 0

        End If

    Next

    WCuit = x

End Sub


