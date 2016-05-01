VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCitinuevo 
   AutoRedraw      =   -1  'True
   Caption         =   "Exportacion de Datos al CITI "
   ClientHeight    =   4590
   ClientLeft      =   3345
   ClientTop       =   2250
   ClientWidth     =   5475
   LinkTopic       =   "Form2"
   ScaleHeight     =   4590
   ScaleWidth      =   5475
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   1560
         TabIndex        =   11
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Nombre 
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
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   240
         Width           =   975
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
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre"
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
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
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
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1335
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
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgCitinuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstIvaComp As Recordset
Dim spIvaComp As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String
Dim XVectorII(8000, 20) As String
Dim XVector(8000, 20) As String
Dim WTipo As String
Dim WPunto As String
Dim WNumero As String
Dim WCuit As String
Dim WNombre As String

Dim ZZZZNeto21 As Double
Dim ZZZZNeto27 As Double
Dim ZZZZNeto105 As Double
Dim ZZZZNeto As Double

Dim AAa As Double


Private Sub Acepta_Click()

    WDrive = Drive.Drive
    WDir = Dir1.Path
    

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    
    Rem GoTo da
    
    
    XNOmbre = WDir + "\" + "REGINFO_CV_CABECERA" + ".txt"
    Open XNOmbre For Output As #1
    
    
    Select Case Val(Wempresa)
        Case 1
            ZNombre = "surfa"
            WImpo1 = "30549165083"
        Case Else
            ZNombre = "pellital"
            WImpo1 = "30610524598"
    End Select
    
    WImpo2 = WAno + WMes
    WImpo3 = "00"
    WImpo4 = "N"
    WImpo5 = "N"
    WImpo6 = "2"
    WImpo7 = "000000000000000"
    WImpo8 = "000000000000000"
    WImpo9 = "000000000000000"
    WImpo10 = "000000000000000"
    WImpo11 = "000000000000000"
    WImpo12 = "000000000000000"
    
    WImpre = WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8 + WImpo9 + WImpo10 + WImpo11 + WImpo12
    
    Print #1, WImpre
    
    Close #1
    
    
    
    Erase XVector
    Erase XVectorII
    Renglon = 0
    RenglonII = 0
    
    
    
    XNOmbre = WDir + "\" + "REGINFO_CV_COMPRAS_CBTE" + ".txt"
    Open XNOmbre For Output As #1
    
    
    XNOmbre = WDir + "\" + "REGINFO_CV_COMPRAS_CBTEnRO" + ".txt"
    Open XNOmbre For Output As #2
    
    Rem XParam = "'" + WDesde + "','" _
    rem              + WHasta + "'"
    Rem spIvaComp = "ListaIvacompDesdeHasta " + XParam
    spIvaComp = "ListaIvacomp"
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
            
        With rstIvaComp
                
            .MoveFirst
            Do
            
                WFecha = Right$(!Periodo, 4) + Mid$(!Periodo, 4, 2) + Left$(!Periodo, 2)
                If WDesde <= WFecha And WFecha <= WHasta Then
                    If !Letra = "A" Or !Letra = "B" Or !Letra = "C" Or !Letra = "M" Then
                        If Val(!Tipo) = 1 Or Val(!Tipo) = 2 Or Val(!Tipo) = 3 Then
                        
                            Renglon = Renglon + 1
                            XVectorII(Renglon, 1) = !Letra
                            XVectorII(Renglon, 2) = !Tipo
                            If Val(!Punto) > 0 Then
                                XVectorII(Renglon, 3) = !Punto
                                    Else
                                XVectorII(Renglon, 3) = "1"
                            End If
                            XVectorII(Renglon, 4) = !Numero
                            ZZVeri = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                            If ZZVeri < "20130101" Then
                                XVectorII(Renglon, 5) = WDesde
                                    Else
                                XVectorII(Renglon, 5) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                            End If
                            XVectorII(Renglon, 6) = !Proveedor
                            XVectorII(Renglon, 7) = Str$(!Neto)
                            XVectorII(Renglon, 8) = Str$(!Exento)
                            XVectorII(Renglon, 9) = Str$(!Iva21)
                            XVectorII(Renglon, 10) = Str$(!Iva5)
                            XVectorII(Renglon, 11) = Str$(!Iva27)
                            XVectorII(Renglon, 12) = Str$(!Ib)
                            XVectorII(Renglon, 13) = IIf(IsNull(!Despacho), "", !Despacho)
                            XVectorII(Renglon, 14) = !NroInterno
                            XVectorII(Renglon, 15) = !Fecha
                            ZZIva105 = IIf(IsNull(!Iva105), "0", !Iva105)
                            XVectorII(Renglon, 16) = Str$(ZZIva105)
        
                            Rem If Val(!NroInterno) = 137286 Then Stop
                            
                            ZZIva21 = !Iva21
                            ZZIva27 = !Iva27
                            ZZIva105 = IIf(IsNull(!Iva105), "0", !Iva105)
                            ZZIva = ZZIva21 + ZZIva27 + ZZIva105
                            If !Neto = 0 And ZZIva <> 0 And Trim(XVectorII(Renglon, 13)) = "" Then
                                ZZZZNeto21 = ZZIva21 / 21 * 100
                                ZZZZNeto27 = ZZIva27 / 27 * 100
                                ZZZZNeto105 = ZZIva105 / 10.5 * 100
                                ZZZZNeto = ZZZZNeto21 + ZZZZNeto27 + ZZZZNeto105
                                Call Redondeo(ZZZZNeto)
                                XVectorII(Renglon, 7) = Str$(ZZZZNeto)
                            End If
                            
                        End If
                    End If
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstIvaComp.Close
    End If
    
    
    RenglonII = Renglon
    Renglon = 0
    
    For Ciclo = 1 To RenglonII
    
        ZZGraba = "S"
        
        Rem If Val(XVectorII(Ciclo, 14)) = 137850 Then Stop

        For A = 1 To 50
            
            Auxi = XVectorII(Ciclo, 14)
            Call Ceros(Auxi, 8)
                
            Auxi1 = Str$(A)
            Call Ceros(Auxi1, 2)
                
            ZZClave = Auxi + Auxi1
                
            ZSql = "Select *"
            ZSql = ZSql + " FROM IvaCompAdicional"
            ZSql = ZSql + " Where IvaCompAdicional.Clave = " + "'" + ZZClave + "'"
            spIvaCompAdicional = ZSql
            Set rstIvaCompAdicional = db.OpenRecordset(spIvaCompAdicional, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaCompAdicional.RecordCount > 0 Then
            
                ZZTipoFac = rstIvaCompAdicional!Tipo
                
                Select Case ZZTipoFac
                    Case "NC", "C"
                        ZZTipo = "03"
                        ZZImpre = "NC"
                    Case "ND", "D"
                        ZZTipo = "02"
                        ZZImpre = "NF"
                    Case Else
                        ZZTipo = "01"
                        ZZImpre = "FC"
                End Select

                
                ZZLetra = rstIvaCompAdicional!Letra
                ZZPunto = Trim(rstIvaCompAdicional!Punto)
                ZZNumero = Trim(rstIvaCompAdicional!Numero)
                ZZFecha = rstIvaCompAdicional!Fecha
                ZZOrdFecha = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                ZZvencimiento = rstIvaCompAdicional!Fecha
                ZZPeriodo = rstIvaCompAdicional!Fecha
                ZZNeto = rstIvaCompAdicional!Neto
                ZZIva21 = rstIvaCompAdicional!Iva21
                ZZIva5 = rstIvaCompAdicional!perceib
                ZZIva27 = rstIvaCompAdicional!Iva27
                ZZib = rstIvaCompAdicional!perceiva
                ZZIva105 = rstIvaCompAdicional!Iva105
                ZZExento = rstIvaCompAdicional!Exento
                
                ZZNombre = rstIvaCompAdicional!Razon
                ZZCuit = rstIvaCompAdicional!Cuit
                
                    
                    
                Renglon = Renglon + 1
                
                XVector(Renglon, 1) = ZZLetra
                XVector(Renglon, 2) = ZZTipo
                If Val(ZZPunto) > 0 Then
                    XVector(Renglon, 3) = ZZPunto
                        Else
                    XVector(Renglon, 3) = "1"
                End If
                XVector(Renglon, 4) = ZZNumero
                XVector(Renglon, 5) = ZZOrdFecha
                XVector(Renglon, 6) = ""
                XVector(Renglon, 7) = Str$(ZZNeto)
                XVector(Renglon, 8) = Str$(ZZExento)
                XVector(Renglon, 9) = Str$(ZZIva21)
                XVector(Renglon, 10) = Str$(ZZIva5)
                XVector(Renglon, 11) = Str$(ZZIva27)
                XVector(Renglon, 12) = Str$(ZZib)
                XVector(Renglon, 13) = ""
                XVector(Renglon, 14) = XVectorII(Ciclo, 14)
                XVector(Renglon, 15) = ZZFecha
                XVector(Renglon, 16) = Str$(ZZIva105)
                XVector(Renglon, 17) = ZZCuit
                XVector(Renglon, 18) = ZZNombre
    
                rstIvaCompAdicional.Close
                    
                ZZGraba = "N"
                    
                    Else
                    
                Exit For
                    
            End If
                
        Next A

        If ZZGraba = "S" Then
        
            Renglon = Renglon + 1
            For CicloII = 1 To 16
                XVector(Renglon, CicloII) = XVectorII(Ciclo, CicloII)
            Next CicloII
        
        End If
    
    Next Ciclo
    
    
    ZZSuma = 0
    
    For Ciclo = 1 To Renglon
    
        WLetra = XVector(Ciclo, 1)
        WTipo = XVector(Ciclo, 2)
        WPunto = XVector(Ciclo, 3)
        WNumero = XVector(Ciclo, 4)
        WFecha = XVector(Ciclo, 5)
        WProveedor = XVector(Ciclo, 6)
        WNroInterno = XVector(Ciclo, 14)
        WRazon = XVector(Ciclo, 18)
        
        
        Rem If Val(WNroInterno) = 137996 Then Stop
        Rem If Val(WNroInterno) = 138092 Then Stop
        
        ZZNeto = Val(XVector(Ciclo, 7))
        ZZNeto = Int(ZZNeto * 100)
        WNeto = ZZNeto
        
        ZZExento = Val(XVector(Ciclo, 8))
        ZZExento = Int(ZZExento * 100)
        WExento = ZZExento
        
        ZZIva21 = Val(XVector(Ciclo, 9))
        ZZIva21 = Int(ZZIva21 * 100)
        WIva21 = ZZIva21
        
        ZZIva5 = Val(XVector(Ciclo, 10))
        ZZIva5 = Int(ZZIva5 * 100)
        WIva5 = ZZIva5
        
        ZZIva27 = Val(XVector(Ciclo, 11))
        ZZIva27 = Int(ZZIva27 * 100)
        WIva27 = ZZIva27
        
        ZZIva105 = Val(XVector(Ciclo, 16))
        ZZIva105 = Int(ZZIva105 * 100)
        WIva105 = ZZIva105
        
        zzzzib = Val(XVector(Ciclo, 12))
        zzzzib = Int(zzzzib * 100)
        WIb = zzzzib
        
        
        WResto = 0
        
        WDespacho = Trim(XVector(Ciclo, 13))
        If Trim(WDespacho) <> "" Then
            ZZLargo = Len(WDespacho)
            For ZZCiclo = 1 To ZZLargo
                If Mid$(WDespacho, ZZCiclo, 1) = " " Then
                    WDespacho = Left$(WDespacho, ZZCiclo - 1) + "" + Mid$(WDespacho, ZZCiclo + 1, 50)
                End If
            Next ZZCiclo
        End If
        WDespacho = Left$(WDespacho + Space$(16), 16)
        If Trim(WDespacho) <> "" Then
            WDespacho = Left$(Trim(WDespacho) + "0000000000000000", 16)
        End If
        
        If WLetra = "B" Or WLetra = "C" Then
            WExento = 0
        End If
        
        WNroInterno = XVector(Ciclo, 14)
        WFechaII = XVector(Ciclo, 15)
        
        Select Case WProveedor
            Case "10065511620", "10070956507", "10065786411"
                WIva = WIva21 + WIva27 + WIva105
                WIva27 = WIva
                WIva21 = 0
                WIva105 = 0
            Case "10053718600", "10050001091", "10099924210", "10050000845"
                WIva = WIva21 + WIva27 + WIva105
                WIva105 = WIva
                WIva21 = 0
                WIva27 = 0
        End Select
        
        WIva = WIva21 + WIva27 + WIva105
        If WIva = 0 Then
            WNeto = WNeto + WExento
            WExento = 0
        End If
        
        WTotal = WNeto + WExento + WIva21 + WIva5 + WIva27 + WIva105 + WIb
        If Trim(WDespacho) <> "" And WNeto = 0 Then
            WImpo = Int(WIva21 / 21 * 100)
            WTotal = WTotal + WImpo
        End If
        
        If WIva = 0 Then
            WCodigoExento = "N"
            Rem z   zona de exportacion
            Rem x   exportaciones al enterior
            Rem e   operaciones exentas
            Rem C   Operaciones de canje
                Else
            WCodigoExento = " "
        End If
            
        
        
        WAlicuota = 0
        If WIva21 <> 0 Then
            WAlicuota = WAlicuota + 1
        End If
        If WIva27 <> 0 Then
            WAlicuota = WAlicuota + 1
        End If
        If WIva105 <> 0 Then
            WAlicuota = WAlicuota + 1
        End If
        If WIva = 0 And WNeto <> 0 Then
            WAlicuota = WAlicuota + 1
        End If
        
        If WLetra = "B" Or WLetra = "C" Then
            WAlicuota = 0
        End If
        
        If Trim(WProveedor) <> "" Then
            spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                WNombre = RstProveedor!Nombre
                WNombre = WNombre + Space$(30)
                WNombre = Left$(WNombre, 30)
                WCuit = RstProveedor!Cuit
                Call Eval
                RstProveedor.Close
            End If
                Else
            WNombre = WRazon
            WNombre = WNombre + Space$(30)
            WNombre = Left$(WNombre, 30)
            WCuit = XVector(Ciclo, 17)
        End If
                
        Call Ceros(WTipo, 2)
        Call Ceros(WPunto, 5)
        Call Ceros(WNumero, 20)
        Call Ceros(WCuit, 20)
        
        Rem fecha
        WImpo1 = WFecha
        
        Rem tipo de comprobante
        Select Case WLetra
            Case "A"
                Select Case Val(WTipo)
                    Case 1
                        WImpo2 = "001"
                    Case 2
                        WImpo2 = "002"
                    Case 3
                        WImpo2 = "003"
                    Case Else
                        WImpo2 = "000"
                End Select
            Case "B"
                Select Case Val(WTipo)
                    Case 1
                        WImpo2 = "006"
                    Case 2
                        WImpo2 = "007"
                    Case 3
                        WImpo2 = "008"
                    Case Else
                        WImpo2 = "000"
                End Select
            Case "C"
                Select Case Val(WTipo)
                    Case 1
                        WImpo2 = "011"
                    Case 2
                        WImpo2 = "012"
                    Case 3
                        WImpo2 = "013"
                    Case Else
                        WImpo2 = "000"
                End Select
            Case "M"
                Select Case Val(WTipo)
                    Case 1
                        WImpo2 = "051"
                    Case 2
                        WImpo2 = "052"
                    Case 3
                        WImpo2 = "053"
                    Case Else
                        WImpo2 = "000"
                End Select
            Case Else
                WImpo2 = "000"
        End Select
        
        If Trim(WDespacho) <> "" Then
            WImpo2 = "066"
            WPunto = "0"
            WNumero = "0"
            Call Ceros(WPunto, 5)
            Call Ceros(WNumero, 20)
        End If
        
        Rem punto
        WImpo3 = WPunto
        
        Rem Numero
        WImpo4 = WNumero
        
        Rem despacho de importacion
        WImpo5 = WDespacho
        
        Rem tipo de doc
        WImpo6 = "80"
        
        Rem numero de doc
        WImpo7 = WCuit
        
        Rem razon social
        WImpo8 = WNombre
        
        Rem total
        If WTotal >= 0 Then
            Auxi1 = Str$(WTotal)
            Call Ceros(Auxi1, 15)
            WImpo9 = Auxi1
                Else
            Auxi1 = Str$(Abs(WTotal))
            Call Ceros(Auxi1, 14)
            WImpo9 = "0" + Auxi1
        End If
        
        Rem resto del neto
        If WResto >= 0 Then
            Auxi1 = Str$(WResto)
            Call Ceros(Auxi1, 15)
            WImpo11 = Auxi1
                Else
            Auxi1 = Str$(Abs(WResto))
            Call Ceros(Auxi1, 14)
            WImpo11 = "0" + Auxi1
        End If
        
        Rem exento
        If WExento >= 0 Then
            Auxi1 = Str$(WExento)
            Call Ceros(Auxi1, 15)
            WImpo10 = Auxi1
                Else
            Auxi1 = Str$(Abs(WExento))
            Call Ceros(Auxi1, 14)
            WImpo10 = "0" + Auxi1
        End If
            
        Rem percepcion de iva
        If WIva5 >= 0 Then
            Auxi1 = Str$(WIva5)
            Call Ceros(Auxi1, 15)
            WImpo12 = Auxi1
                Else
            Auxi1 = Str$(Abs(WIva5))
            Call Ceros(Auxi1, 14)
            WImpo12 = "0" + Auxi1
        End If
                
        
        Rem otros impuuestos nacionales
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo13 = Auxi1
        
        Rem ingresos brutos
        If WIb >= 0 Then
            Auxi1 = Str$(WIb)
            Call Ceros(Auxi1, 15)
            WImpo14 = Auxi1
                Else
            Auxi1 = Str$(Abs(WIb))
            Call Ceros(Auxi1, 14)
            WImpo14 = "0" + Auxi1
        End If
        
        Rem otros impuuestos municipales
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo15 = Auxi1
        
        Rem otros impuuestos internos
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo16 = Auxi1
        
        Rem codigo de moneda
        WImpo17 = "PES"
        
        Rem PARIDAD
        Rem ZCAmbio = "0"
        Rem spCambios = "ConsultaCambio " + "'" + WFechaII + "'"
        Rem Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstCambios.RecordCount > 0 Then
        Rem     ZCAmbio = rstCambios!Cambio
        Rem     rstCambios.Close
        Rem             Else
        Rem     ZCAmbio = "1"
        Rem End If
        Rem Auxi1 = Str$(Int(ZCAmbio * 1000000))
        
        ZCAmbio = "1"
        Auxi1 = Str$(Int(ZCAmbio * 1000000))
        Call Ceros(Auxi1, 10)
        WImpo18 = Auxi1
        
        Rem CANTIDAD DE ALICUOTAS de iva
        If WLetra = "A" Or WLetra = "M" Then
            If WAlicuota = 0 Then
                WAlicuota = "1"
            End If
        End If
        Auxi1 = Str$(WAlicuota)
        Call Ceros(Auxi1, 1)
        WImpo19 = Auxi1
        
        Rem codigo de operacion
        WImpo20 = WCodigoExento
        
        Rem iva
        If WIva >= 0 Then
            Auxi1 = Str$(WIva)
            Call Ceros(Auxi1, 15)
            WImpo21 = Auxi1
                Else
            Auxi1 = Str$(Abs(WIva))
            Call Ceros(Auxi1, 14)
            WImpo21 = "0" + Auxi1
        End If
        
        Rem otros triburtos
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo22 = Auxi1
        
        Rem cuit del emisor ????
        Auxi1 = "0"
        Call Ceros(Auxi1, 11)
        WImpo23 = Auxi1
        
        Rem nombre del emisor ????
        WImpo24 = Space$(30)
        
        Rem iva co,ision ????
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo25 = Auxi1
        
        Rem If Val(WEmpresa) = 1 Then
        Rem     WCuitII = "30549165083"
        Rem     WNombreII = Left$("SURFACTAN S.A." + Space$(25), 25)
        Rem         Else
        Rem     WCuitII = ""
        Rem     WNombreII = ""
        Rem End If
        Rem WCuitII = "00000000000"
        Rem WNombreII = Space$(25)
        
        If Val(WCuit) <> 0 Then
            ZZSuma = ZZSuma + 1
            
            WImpre = WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8 + WImpo9 + WImpo10 + WImpo11 + WImpo12 + WImpo13 + WImpo14 + WImpo15 + WImpo16 + WImpo17 + WImpo18 + WImpo19 + WImpo20 + WImpo21 + WImpo22 + WImpo23 + WImpo24 + WImpo25
            Rem WImpre = Str$(Ciclo) + " " + WNroInterno + " " + WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8 + WImpo9 + WImpo10 + WImpo11 + WImpo12 + WImpo13 + WImpo14 + WImpo15 + WImpo16 + WImpo17 + WImpo18 + WImpo19 + WImpo20 + WImpo21 + WImpo22 + WImpo23 + WImpo24 + WImpo25
            Print #1, WImpre
            Print #2, Str$(ZZSuma) + " " + WNroInterno + " " + WImpre
        End If
    
        
    Next Ciclo
    
    Close #1
    Close #2
    
    
    
    
    
    XNOmbre = WDir + "\" + "REGINFO_CV_COMPRAS_ALICUOTAS" + ".txt"
    Open XNOmbre For Output As #1
    
    XNOmbre = WDir + "\" + "REGINFO_CV_COMPRAS_ALICUOTASNro" + ".txt"
    Open XNOmbre For Output As #2
    
    
    ZZSuma = 0
    For Ciclo = 1 To Renglon
    
        WLetra = XVector(Ciclo, 1)
        
        If WLetra = "A" Or WLetra = "M" Then
        
            WTipo = XVector(Ciclo, 2)
            WPunto = XVector(Ciclo, 3)
            WNumero = XVector(Ciclo, 4)
            WFecha = XVector(Ciclo, 5)
            WProveedor = XVector(Ciclo, 6)
            WNroInterno = XVector(Ciclo, 14)
            
            Rem If Val(WNroInterno) = 137996 Then Stop
            Rem If Val(WNroInterno) = 138092 Then Stop
            
            ZZNeto = Abs(Val(XVector(Ciclo, 7)))
            ZZNeto = Int(ZZNeto * 100)
            WNeto = ZZNeto
            
            ZZExento = Abs(Val(XVector(Ciclo, 8)))
            ZZExento = Int(ZZExento * 100)
            WExento = ZZExento
            
            ZZIva21 = Abs(Val(XVector(Ciclo, 9)))
            ZZIva21 = Int(ZZIva21 * 100)
            WIva21 = ZZIva21
            
            ZZIva5 = Abs(Val(XVector(Ciclo, 10)))
            ZZIva5 = Int(ZZIva5 * 100)
            WIva5 = ZZIva5
            
            ZZIva27 = Abs(Val(XVector(Ciclo, 11)))
            ZZIva27 = Int(ZZIva27 * 100)
            WIva27 = ZZIva27
            
            ZZIva105 = Abs(Val(XVector(Ciclo, 16)))
            ZZIva105 = Int(ZZIva105 * 100)
            WIva105 = ZZIva105
            
            zzzzib = Abs(Val(XVector(Ciclo, 12)))
            zzzzib = Int(zzzzib * 100)
            WIb = zzzzib
        
            WDespacho = Trim(XVector(Ciclo, 13))
            If Trim(WDespacho) <> "" Then
                ZZLargo = Len(WDespacho)
                For ZZCiclo = 1 To ZZLargo
                    If Mid$(WDespacho, ZZCiclo, 1) = " " Then
                        WDespacho = Left$(WDespacho, ZZCiclo - 1) + "" + Mid$(WDespacho, ZZCiclo + 1, 50)
                    End If
                Next ZZCiclo
            End If
            WDespacho = Left$(WDespacho + Space$(16), 16)
            If Trim(WDespacho) <> "" Then
                WDespacho = Left$(Trim(WDespacho) + "0000000000000000", 16)
            End If
        
            If WLetra = "B" Or WLetra = "C" Then
                WResto = WExento
                WExento = 0
            End If
            
            WNroInterno = XVector(Ciclo, 14)
            
            Select Case WProveedor
                Case "10065511620", "10070956507", "10065786411"
                    WIva = WIva21 + WIva27 + WIva105
                    WIva27 = WIva
                    WIva21 = 0
                    WIva105 = 0
                Case "10053718600", "10050001091", "10099924210", "10050000845"
                    WIva = WIva21 + WIva27 + WIva105
                    WIva105 = WIva
                    WIva21 = 0
                    WIva27 = 0
            End Select
            
            WIva = WIva21 + WIva27 + WIva105
            If WIva = 0 Then
                WNeto = WNeto + WExento
                WExento = 0
            End If
            
            WTotal = WNeto + WExento + WIva21 + WIva5 + WIva27 + WIva105 + WIb
            
            
            If WIva = 0 Then
                WCodigoExento = "N"
                Rem z   zona de exportacion
                Rem x   exportaciones al enterior
                Rem e   operaciones exentas
                Rem C   Operaciones de canje
                    Else
                WCodigoExento = " "
            End If
                
            WAlicuota = 0
            If WIva21 <> 0 Then
                WAlicuota = WAlicuota + 1
            End If
            If WIva27 <> 0 Then
                WAlicuota = WAlicuota + 1
            End If
            If WIva105 <> 0 Then
                WAlicuota = WAlicuota + 1
            End If
            If WIva = 0 And WNeto <> 0 Then
                WAlicuota = WAlicuota + 1
            End If
            If WLetra = "B" Or WLetra = "C" Then
                WAlicuota = 0
            End If
            
            If Trim(WProveedor) <> "" Then
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WNombre = RstProveedor!Nombre
                    WNombre = WNombre + Space$(30)
                    WNombre = Left$(WNombre, 30)
                    WCuit = RstProveedor!Cuit
                    Call Eval
                    RstProveedor.Close
                End If
                    Else
                WNombre = Space$(30)
                WCuit = XVector(Ciclo, 17)
            End If
            
            Call Ceros(WTipo, 2)
            Call Ceros(WPunto, 5)
            Call Ceros(WNumero, 20)
            Call Ceros(WCuit, 20)
            
            If Val(WCuit) <> 0 And Trim(WDespacho) = "" Then
        
                Rem tipo de comprobante
                Select Case WLetra
                    Case "A"
                        Select Case Val(WTipo)
                            Case 1
                                WImpo1 = "001"
                            Case 2
                                WImpo1 = "002"
                            Case 3
                                WImpo1 = "003"
                            Case Else
                                WImpo1 = "000"
                        End Select
                    Case "M"
                        Select Case Val(WTipo)
                            Case 1
                                WImpo1 = "051"
                            Case 2
                                WImpo1 = "052"
                            Case 3
                                WImpo1 = "053"
                            Case Else
                                WImpo1 = "000"
                        End Select
                    Case Else
                        WImpo1 = "000"
                End Select
        
                If Trim(WDespacho) <> "" Then
                    WImpo1 = "066"
                    WPunto = "0"
                    WNumero = "0"
                    Call Ceros(WPunto, 5)
                    Call Ceros(WNumero, 20)
                End If
                
                Rem punto
                WImpo2 = WPunto
                
                Rem Numero
                WImpo3 = WNumero
                
                Rem tipo de doc
                WImpo4 = "80"
                
                Rem numero de doc
                WImpo5 = WCuit
                
                If WIva21 <> 0 Then
                
                    If WAlicuota = 1 Then
                    
                        Rem neto
                        If WNeto >= 0 Then
                            Auxi1 = Str$(WNeto)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WNeto))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                            Else
                            
                        WImpo = Int(WIva21 / 21 * 100)
                        If WImpo >= 0 Then
                            Auxi1 = Str$(WImpo)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WImpo))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                    End If
                    
                    Rem Iva 21
                    WImpo7 = "0005"
                    
                    Rem impo iva
                    If WIva21 >= 0 Then
                        Auxi1 = Str$(WIva21)
                        Call Ceros(Auxi1, 15)
                        WImpo8 = Auxi1
                            Else
                        Auxi1 = Str$(Abs(WIva21))
                        Call Ceros(Auxi1, 14)
                        WImpo8 = "0" + Auxi1
                    End If
                    
        
                    WImpre = WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8
                    Print #1, WImpre
                    
                    ZZSuma = ZZSuma + 1
                    Print #2, Str$(ZZSuma) + " " + WNroInterno + " " + WImpre
                End If
                
                If WIva105 <> 0 Then
                
                    If WAlicuota = 1 Then
                    
                        Rem neto
                        If WNeto >= 0 Then
                            Auxi1 = Str$(WNeto)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WNeto))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                            Else
                            
                        WImpo = Int(WIva105 / 10.5 * 100)
                        If WImpo >= 0 Then
                            Auxi1 = Str$(WImpo)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WImpo))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                    End If
                    
                    Rem Iva 10.5
                    WImpo7 = "0004"
                    
                    Rem impo iva
                    Auxi1 = Str$(WIva105)
                    Call Ceros(Auxi1, 15)
                    WImpo8 = Auxi1
        
                    WImpre = WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8
                    Print #1, WImpre
                    
                    ZZSuma = ZZSuma + 1
                    Print #2, Str$(ZZSuma) + " " + WNroInterno + " " + WImpre
                    
                End If
                
                If WIva27 <> 0 Then
                
                    If WAlicuota = 1 Then
                    
                        Rem neto
                        If WNeto >= 0 Then
                            Auxi1 = Str$(WNeto)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WNeto))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                            Else
                            
                        WImpo = Int(WIva27 / 27 * 100)
                        If WImpo >= 0 Then
                            Auxi1 = Str$(WImpo)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WImpo))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                    End If
                    
                    
                    Rem Iva 27
                    WImpo7 = "0006"
                    
                    Rem impo iva
                    Auxi1 = Str$(WIva27)
                    Call Ceros(Auxi1, 15)
                    WImpo8 = Auxi1
        
                    WImpre = WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8
                    Print #1, WImpre
                    
                    ZZSuma = ZZSuma + 1
                    Print #2, Str$(ZZSuma) + " " + WNroInterno + " " + WImpre
                    
                End If
                            
                If WIva = 0 And WNeto <> 0 Then
                
                    Rem neto
                    If WNeto >= 0 Then
                        Auxi1 = Str$(WNeto)
                        Call Ceros(Auxi1, 15)
                        WImpo6 = Auxi1
                            Else
                        Auxi1 = Str$(Abs(WNeto))
                        Call Ceros(Auxi1, 14)
                        WImpo6 = "0" + Auxi1
                    End If
                    
                    Rem Iva 10.5
                    WImpo7 = "0003"
                    
                    Rem impo iva
                    Auxi1 = "0"
                    Call Ceros(Auxi1, 15)
                    WImpo8 = Auxi1
                                
                                
                                
                    WImpre = WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8
                    Print #1, WImpre
                    
                    ZZSuma = ZZSuma + 1
                    Print #2, Str$(ZZSuma) + " " + WNroInterno + " " + WImpre
                    
                    
                End If
            
            End If
            
        End If
        
    Next Ciclo
    
    Close #1
    
    
    
    
    
    XNOmbre = WDir + "\" + "REGINFO_CV_COMPRAS_IMPORTACIONES" + ".txt"
    Open XNOmbre For Output As #1
    
    For Ciclo = 1 To Renglon
    
        WDespacho = XVector(Ciclo, 13)
        
        If Trim(WDespacho) <> "" Then
        
            WLetra = XVector(Ciclo, 1)
            WTipo = XVector(Ciclo, 2)
            WPunto = XVector(Ciclo, 3)
            WNumero = XVector(Ciclo, 4)
            WFecha = XVector(Ciclo, 5)
            WProveedor = XVector(Ciclo, 6)
            WNeto = Int(Val(XVector(Ciclo, 7)) * 100)
            WExento = Int(Val(XVector(Ciclo, 8)) * 100)
            WIva21 = Int(Val(XVector(Ciclo, 9)) * 100)
            WIva5 = Int(Val(XVector(Ciclo, 10)) * 100)
            WIva27 = Int(Val(XVector(Ciclo, 11)) * 100)
            WIva105 = 0
            WIb = Int(Val(XVector(Ciclo, 12)) * 100)
            WDespacho = Trim(XVector(Ciclo, 13))
            If Trim(WDespacho) <> "" Then
                ZZLargo = Len(WDespacho)
                For ZZCiclo = 1 To ZZLargo
                    If Mid$(WDespacho, ZZCiclo, 1) = " " Then
                        WDespacho = Left$(WDespacho, ZZCiclo - 1) + "" + Mid$(WDespacho, ZZCiclo + 1, 50)
                    End If
                Next ZZCiclo
            End If
            WDespacho = Left$(WDespacho + Space$(16), 16)
            If Trim(WDespacho) <> "" Then
                WDespacho = Left$(Trim(WDespacho) + "0000000000000000", 16)
            End If
            
            
            WNroInterno = XVector(Ciclo, 14)
            
            Select Case WProveedor
                Case "10065511620", "10070956507", "10065786411"
                    WIva = WIva21 + WIva27 + WIva105
                    WIva27 = WIva
                    WIva21 = 0
                    WIva105 = 0
                Case "10053718600", "10050001091", "10099924210", "10050000845"
                    WIva = WIva21 + WIva27 + WIva105
                    WIva105 = WIva
                    WIva21 = 0
                    WIva27 = 0
            End Select
            
            WIva = WIva21 + WIva27 + WIva105
            If WIva = 0 Then
                WNeto = WNeto + WExento
                WExento = 0
            End If
            
            WTotal = WNeto + WExento + WIva21 + WIva5 + WIva27 + WIva105 + WIb
            
            
            If WIva = 0 Then
                WCodigoExento = "N"
                Rem z   zona de exportacion
                Rem x   exportaciones al enterior
                Rem e   operaciones exentas
                Rem C   Operaciones de canje
                    Else
                WCodigoExento = " "
            End If
                
            
            
            WAlicuota = 1
            
            spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                WNombre = RstProveedor!Nombre
                WNombre = WNombre + Space$(30)
                WNombre = Left$(WNombre, 30)
                WCuit = RstProveedor!Cuit
                Call Eval
                RstProveedor.Close
            End If
            
            Call Ceros(WTipo, 2)
            Call Ceros(WPunto, 5)
            Call Ceros(WNumero, 20)
            Call Ceros(WCuit, 20)
            
            WImpo1 = WDespacho
            
            Rem neto
            Rem Auxi1 = Str$(WNeto)
            If WNeto <> 0 Then
                WImpo = WNeto
                    Else
                WImpo = Int(WIva21 / 21 * 100)
            End If
            If WImpo >= 0 Then
                Auxi1 = Str$(WImpo)
                Call Ceros(Auxi1, 15)
                WImpo2 = Auxi1
                    Else
                Auxi1 = Str$(Abs(WImpo))
                Call Ceros(Auxi1, 14)
                WImpo2 = "0" + Auxi1
            End If
            
            
            Rem Iva 21
            WImpo3 = "0005"
            
            Rem impo iva
            If WIva21 >= 0 Then
                Auxi1 = Str$(WIva21)
                Call Ceros(Auxi1, 15)
                WImpo4 = Auxi1
                    Else
                Auxi1 = Str$(WIva21)
                Call Ceros(Auxi1, 14)
                WImpo4 = "0" + Auxi1
            End If

            WImpre = WImpo1 + WImpo2 + WImpo3 + WImpo4
            Print #1, WImpre
            
        End If
        
    Next Ciclo
    
    Close #1
    
    
    
    
da:
    
    
    
    
    
    
    
    Erase XVector
    Renglon = 0
    
    
    
    XNOmbre = WDir + "\" + "REGINFO_CV_VENTAS_CBTE" + ".txt"
    Open XNOmbre For Output As #1
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Ctacte"
    ZSql = ZSql + " Where Ctacte.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Ctacte.OrdFecha <= " + "'" + WHasta + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
        With rstCtacte
                
            .MoveFirst
            Do
            
                If Val(!Tipo) = 1 Or Val(!Tipo) = 2 Or Val(!Tipo) = 3 Or Val(!Tipo) = 4 Or Val(!Tipo) = 5 Then
                
                    Renglon = Renglon + 1
                    XVector(Renglon, 1) = "A"
                    XVector(Renglon, 2) = !Tipo
                    XVector(Renglon, 3) = "0006"
                    If !Numero < 200000 Then
                        XVector(Renglon, 4) = !Numero - 100000
                                Else
                        If !Numero < 300000 Then
                            XVector(Renglon, 1) = "B"
                            XVector(Renglon, 4) = !Numero - 200000
                                    Else
                            If !Numero < 810000 Then
                                XVector(Renglon, 4) = !Numero - 800000
                                        Else
                                XVector(Renglon, 4) = !Numero - 810000
                            End If
                        End If
                    End If
                    XVector(Renglon, 5) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                    XVector(Renglon, 6) = !Cliente
                    XVector(Renglon, 7) = Str$(!Neto)
                    XVector(Renglon, 8) = Str$(!Iva1)
                    XVector(Renglon, 9) = Str$(!Iva2)
                    XVector(Renglon, 10) = Str$(!ImpoIbTucu)
                    XVector(Renglon, 11) = Str$(!ImpoIbCiudad)
                    XVector(Renglon, 12) = Str$(!impoib)
                    Select Case Val(Mid$(!Fecha, 4, 2))
                        Case 1
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "02" + "01"
                        Case 2
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "03" + "01"
                        Case 3
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "04" + "01"
                        Case 4
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "05" + "01"
                        Case 5
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "06" + "01"
                        Case 6
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "07" + "01"
                        Case 7
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "08" + "01"
                        Case 8
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "09" + "01"
                        Case 9
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "10" + "01"
                        Case 10
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "11" + "01"
                        Case 11
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "12" + "01"
                        Case 12
                            XVector(Renglon, 13) = Right$(!Fecha, 4) + "12" + "31"
                        Case Else
                    End Select
                    
                    XVector(Renglon, 14) = !Numero
                    XVector(Renglon, 15) = Right$(!Vencimiento, 4) + Mid$(!Vencimiento, 4, 2) + Left$(!Vencimiento, 2)

                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstCtacte.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        WLetra = XVector(Ciclo, 1)
        WTipo = XVector(Ciclo, 2)
        WPunto = XVector(Ciclo, 3)
        WNumero = XVector(Ciclo, 4)
        WFecha = XVector(Ciclo, 5)
        WCliente = XVector(Ciclo, 6)
        WNumeroII = XVector(Ciclo, 14)
        WVto = XVector(Ciclo, 15)
        
        
        
        
        ZZNeto = Val(XVector(Ciclo, 7))
        ZZNeto = Int(ZZNeto * 100)
        WNeto = ZZNeto
        
        ZZIva1 = Val(XVector(Ciclo, 8))
        ZZIva1 = Int(ZZIva1 * 100)
        WIva1 = ZZIva1
        
        ZZIva2 = Val(XVector(Ciclo, 9))
        ZZIva2 = Int(ZZIva2 * 100)
        WIva2 = ZZIva2
        
        ZZIbTucu = Val(XVector(Ciclo, 10))
        ZZIbTucu = Int(ZZIbTucu * 100)
        WIbTucu = ZZIbTucu
        
        ZZIbCiudad = Val(XVector(Ciclo, 11))
        ZZIbCiudad = Int(ZZIbCiudado * 100)
        WIbCiudad = ZZIbCiudad
        
        zzzzib = Val(XVector(Ciclo, 12))
        zzzzib = Int(zzzzib * 100)
        WIb = zzzzib
        
        WExento = 0
        
        Wvencimiento = XVector(Ciclo, 13)
        
        WTotal = WNeto + WIva1 + WIva2 + WIbTucu + WIbCiudad + WIb
        WIva = WIva1 + WIva2
        
        Rem If WIva = 0 Then
        Rem     WExento = WNeto
        Rem rem End If
        
        
        If WIva = 0 Then
            WCodigoExento = "N"
            Rem z   zona de exportacion
            Rem x   exportaciones al enterior
            Rem e   operaciones exentas
            Rem C   Operaciones de canje
                Else
            WCodigoExento = " "
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WNombre = rstCliente!Razon
            WNombre = WNombre + Space$(30)
            WNombre = Left$(WNombre, 30)
            WCuit = rstCliente!Cuit
            Call Eval
            rstCliente.Close
        End If
        
        Call Ceros(WTipo, 2)
        Call Ceros(WPunto, 5)
        Call Ceros(WNumero, 20)
        Call Ceros(WCuit, 20)
        
        Rem fecha
        WImpo1 = WFecha
        
        Rem tipo de comprobante
        Select Case WLetra
            Case "A"
                Select Case Val(WTipo)
                    Case 1, 3
                        WImpo2 = "001"
                    Case 4
                        WImpo2 = "002"
                    Case 2, 5
                        WImpo2 = "003"
                    Case Else
                        WImpo2 = "000"
                End Select
            Case "B"
                Select Case Val(WTipo)
                    Case 1, 3
                        WImpo2 = "006"
                    Case 4
                        WImpo2 = "007"
                    Case 2, 5
                        WImpo2 = "008"
                    Case Else
                        WImpo2 = "000"
                End Select
            Case "C"
                Select Case Val(WTipo)
                    Case 1
                        WImpo2 = "011"
                    Case 2
                        WImpo2 = "012"
                    Case 3
                        WImpo2 = "013"
                    Case Else
                        WImpo2 = "000"
                End Select
            Case "M"
                Select Case Val(WTipo)
                    Case 1
                        WImpo2 = "051"
                    Case 2
                        WImpo2 = "052"
                    Case 3
                        WImpo2 = "053"
                    Case Else
                        WImpo2 = "000"
                End Select
            Case Else
                WImpo2 = "000"
        End Select
        
        If Val(WNumeroII) > 800000 Then
            WImpo2 = "019"
            WPunto = "00003"
        End If
        
        
        If Val(WNumeroII) > 810000 Then
            WImpo2 = "019"
            WPunto = "00003"
        End If
        
        
        Rem punto
        WImpo3 = WPunto
        
        Rem Numero desde
        WImpo4 = WNumero
        
        Rem Numero hasta
        WImpo5 = WNumero
        
        Rem tipo de doc
        WImpo6 = "80"
        
        Rem numero de doc
        WImpo7 = WCuit
        
        Rem razon social
        WImpo8 = WNombre
        
        Rem total
        If WTotal >= 0 Then
            Auxi1 = Str$(WTotal)
            Call Ceros(Auxi1, 15)
            WImpo9 = Auxi1
                Else
            Auxi1 = Str$(Abs(WTotal))
            Call Ceros(Auxi1, 14)
            WImpo9 = "0" + Auxi1
        End If
        
        Rem resto del neto
        Rem ZZSumaResto = Str$(WIb + WIbTucu + WIbCiudad)
        ZZSumaResto = 0
        If ZZSumaResto >= 0 Then
            Auxi1 = Str$(ZZSumaResto)
            Call Ceros(Auxi1, 15)
            WImpo10 = Auxi1
                Else
            Auxi1 = Str$(Abs(ZZSumaResto))
            Call Ceros(Auxi1, 14)
            WImpo10 = "0" + Auxi1
        End If
        
        Rem percepcion a jo categorizados
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo11 = Auxi1
        
        Rem importes exentos
        If WExento >= 0 Then
            Auxi1 = Str$(WExento)
            Call Ceros(Auxi1, 15)
            WImpo12 = Auxi1
                Else
            Auxi1 = Str$(Abs(WExento))
            Call Ceros(Auxi1, 14)
            WImpo12 = "0" + Auxi1
        End If
        
        Rem percepsion p pago a cuenta de impuestos nacionales
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo13 = Auxi1
        
        Rem percepciones  i.b.
        WTotalIb = WIbTucu + WIbCiudad + WIb
        If WTotalIb >= 0 Then
            Auxi1 = Str$(WTotalIb)
            Call Ceros(Auxi1, 15)
            WImpo14 = Auxi1
                Else
            Auxi1 = Str$(Abs(WTotalIb))
            Call Ceros(Auxi1, 14)
            WImpo14 = "0" + Auxi1
        End If
        
        Rem percepsion de impuestos municipales
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo15 = Auxi1
        
        Rem percepsion de impuestos internmos
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo16 = Auxi1
        
        Rem codigo de moneda
        WImpo17 = "PES"
        
        Rem PARIDAD
        ZCAmbio = "1"
        Auxi1 = Str$(Int(ZCAmbio * 1000000))
        Call Ceros(Auxi1, 10)
        WImpo18 = Auxi1
        
        Rem cantidad de alicuotas
        WImpo19 = "1"
        
        Rem codigo de operacion
        WImpo20 = WCodigoExento
        
        Rem otros tributos
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo21 = Auxi1
        
        Rem fecha
        WImpo22 = Wvencimiento
        If Val(WImpo2) = 19 Then
            WImpo22 = "00000000"
        End If
        
        WImpre = WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8 + WImpo9 + WImpo10 + WImpo11 + WImpo12 + WImpo13 + WImpo14 + WImpo15 + WImpo16 + WImpo17 + WImpo18 + WImpo19 + WImpo20 + WImpo21 + WImpo22
    
        Print #1, WImpre
        
    Next Ciclo
    
    Close #1
    
    
    
    
    
    XNOmbre = WDir + "\" + "REGINFO_CV_VENTAS_ALICUOTAS" + ".txt"
    Open XNOmbre For Output As #1
    
    For Ciclo = 1 To Renglon
    
        WLetra = XVector(Ciclo, 1)
        WTipo = XVector(Ciclo, 2)
        WPunto = XVector(Ciclo, 3)
        WNumero = XVector(Ciclo, 4)
        WFecha = XVector(Ciclo, 5)
        WCliente = XVector(Ciclo, 6)
        WNumeroII = XVector(Ciclo, 14)
        
        ZZNeto = Val(XVector(Ciclo, 7))
        ZZNeto = Int(ZZNeto * 100)
        WNeto = ZZNeto
        
        ZZIva1 = Val(XVector(Ciclo, 8))
        ZZIva1 = Int(ZZIva1 * 100)
        WIva1 = ZZIva1
        
        ZZIva2 = Val(XVector(Ciclo, 9))
        ZZIva2 = Int(ZZIva2 * 100)
        WIva2 = ZZIva2
        
        ZZIbTucu = Val(XVector(Ciclo, 10))
        ZZIbTucu = Int(ZZIbTucu * 100)
        WIbTucu = ZZIbTucu
        
        ZZIbCiudad = Val(XVector(Ciclo, 11))
        ZZIbCiudad = Int(ZZIbCiudado * 100)
        WIbCiudad = ZZIbCiudad
        
        zzzzib = Val(XVector(Ciclo, 12))
        zzzzib = Int(zzzzib * 100)
        WIb = zzzzib
        
        WTotal = WNeto + WIva1 + WIva2 + WIbTucu + WIbCiudad + WIb
        WIva = WIva1 + WIva2
        
        
        
        
        Rem If WIva = 0 Then
        Rem     WExento = WNeto
        Rem End If
        If WIva = 0 Then
            WCodigoExento = "N"
            Rem z   zona de exportacion
            Rem x   exportaciones al enterior
            Rem e   operaciones exentas
            Rem C   Operaciones de canje
                Else
            WCodigoExento = " "
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WNombre = rstCliente!Razon
            WNombre = WNombre + Space$(30)
            WNombre = Left$(WNombre, 30)
            WCuit = rstCliente!Cuit
            Call Eval
            rstCliente.Close
        End If
        
        Call Ceros(WTipo, 2)
        Call Ceros(WPunto, 5)
        Call Ceros(WNumero, 20)
        Call Ceros(WCuit, 20)
            
            
            
        Rem tipo de comprobante
        Select Case WLetra
            Case "A"
                Select Case Val(WTipo)
                    Case 1, 3
                        WImpo1 = "001"
                    Case 4
                        WImpo1 = "002"
                    Case 2, 5
                        WImpo1 = "003"
                    Case Else
                        WImpo1 = "000"
                End Select
            Case "B"
                Select Case Val(WTipo)
                    Case 1, 3
                        WImpo1 = "006"
                    Case 4
                        WImpo1 = "007"
                    Case 2, 5
                        WImpo1 = "008"
                    Case Else
                        WImpo1 = "000"
                End Select
            Case "C"
                Select Case Val(WTipo)
                    Case 1
                        WImpo1 = "011"
                    Case 2
                        WImpo1 = "012"
                    Case 3
                        WImpo1 = "013"
                    Case Else
                        WImpo1 = "000"
                End Select
            Case "M"
                Select Case Val(WTipo)
                    Case 1
                        WImpo1 = "051"
                    Case 2
                        WImpo1 = "052"
                    Case 3
                        WImpo1 = "053"
                    Case Else
                        WImpo1 = "000"
                End Select
            Case Else
                WImpo2 = "000"
        End Select
        
        If Val(WNumeroII) > 800000 Then
            WImpo1 = "019"
            WPunto = "00003"
        End If
        
        If Val(WNumeroII) > 810000 Then
            WImpo1 = "019"
            WPunto = "00003"
        End If
        
        Rem punto
        WImpo2 = WPunto
        
        Rem Numero
        WImpo3 = WNumero
        
        
        If Val(WNumeroII) < 800000 Then
        
            Rem neto
            If WNeto >= 0 Then
                Auxi1 = Str$(WNeto)
                Call Ceros(Auxi1, 15)
                WImpo4 = Auxi1
                    Else
                Auxi1 = Str$(Abs(WNeto))
                Call Ceros(Auxi1, 14)
                WImpo4 = "0" + Auxi1
            End If
            
            If WIva <> 0 Then
                Rem Iva 21
                WImpo5 = "0005"
                    Else
                Rem Iva 9
                WImpo5 = "0003"
            End If
            
            Rem impo iva
            If WIva >= 0 Then
                Auxi1 = Str$(WIva)
                Call Ceros(Auxi1, 15)
                WImpo6 = Auxi1
                    Else
                Auxi1 = Str$(Abs(WIva))
                Call Ceros(Auxi1, 14)
                WImpo6 = "0" + Auxi1
            End If
            
            WImpre = WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6
            Print #1, WImpre
            
            
                Else
        
            Rem neto
            If WNeto >= 0 Then
                Auxi1 = Str$(WNeto)
                Call Ceros(Auxi1, 15)
                WImpo4 = Auxi1
                    Else
                Auxi1 = Str$(Abs(WNeto))
                Call Ceros(Auxi1, 14)
                WImpo4 = "0" + Auxi1
            End If
            If WIva <> 0 Then
                Rem Iva 21
                WImpo5 = "0005"
                    Else
                Rem Iva 9
                WImpo5 = "0003"
            End If
            Rem impo iva
            If WIva >= 0 Then
                Auxi1 = Str$(WIva)
                Call Ceros(Auxi1, 15)
                WImpo6 = Auxi1
                    Else
                Auxi1 = Str$(Abs(WIva))
                Call Ceros(Auxi1, 14)
                WImpo6 = "0" + Auxi1
            End If
            
            WImpre = WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6
            Print #1, WImpre
            
        End If
        
    Next Ciclo
    
    Close #1
    
    
    
    
    
    
    
    
    
    Call Cancela_Click
    
End Sub

Private Sub Cancela_Click()
    Desde.SetFocus
    PrgCitinuevo.Hide
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

Private Sub Drive_Change()
    Dir1.Path = Drive.Drive
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
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Frame2.Visible = True
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


