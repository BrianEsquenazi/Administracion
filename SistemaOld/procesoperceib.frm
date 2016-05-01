VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProcesoPerceIb 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Traspaso de Percepcion de Ingresos Brutos"
   ClientHeight    =   6630
   ClientLeft      =   3060
   ClientTop       =   1425
   ClientWidth     =   6930
   LinkTopic       =   "Form2"
   ScaleHeight     =   6630
   ScaleWidth      =   6930
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   4815
      Begin VB.ComboBox TipoList 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Text            =   " "
         Top             =   5040
         Width           =   2415
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
         TabIndex        =   7
         Top             =   4200
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
         Left            =   2520
         TabIndex        =   6
         Top             =   4200
         Width           =   1215
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   1920
         TabIndex        =   4
         Top             =   1920
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
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1080
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
         Left            =   1920
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1920
         TabIndex        =   5
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
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
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgProcesoPerceIb"
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
Dim WImpoIb As Double
Dim XNeto As String
Dim XImpoIb As String
Dim WCuit As String

Private Sub Command1_Click()

    Auxi = "2222 222222"
    
    For ZZCiclo = 1 To 11
    
        If Mid$(Auxi, ZZCiclo, 1) = " " Then
            Auxi = Left$(Auxi, ZZCiclo - 1) + "0" + Mid$(Auxi, ZZCiclo + 1, 11)
        End If
    
    Next ZZCiclo
        
        
        
        

End Sub

Private Sub Drive_Change()
    Dir1.Path = Drive.Drive
End Sub

Private Sub Acepta_Click()

    WDrive = Drive.Drive
    WDir = Dir1.Path
    
    If Val(WEmpresa) = 1 Then
        XNombre = WDir + "\AR-30549165083-" + Nombre.Text + "-7-LOTE1.txt"
            Else
        XNombre = WDir + "\AR-30610524598-" + Nombre.Text + "-7-LOTE1.txt"
    End If
    
    Open XNombre For Output As #1

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    spCtaCte = "ModificaCtacteImporteIva0"
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    If TipoList.ListIndex = 0 Then
        
        
        Rem Procesa las cobranzas
        
        Renglon = 0
        Erase Vector
        
        XParam = "'" + WDesde + "','" _
                     + WHasta + "'"
        spRecibo = "ListaRecibosDifeI" + XParam
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
                        Vector(Renglon, 5) = rstRecibo!Cliente
                        Vector(Renglon, 6) = ""
                        Vector(Renglon, 7) = ""
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
                WNeto = IIf(IsNull(rstCtaCte!Neto), "0", rstCtaCte!Neto)
                WImpoIb = IIf(IsNull(rstCtaCte!impoib), "0", rstCtaCte!impoib)
                If WImpoIb = 0 Then
                    Vector(Cicla, 1) = ""
                    Vector(Cicla, 2) = ""
                    Vector(Cicla, 3) = ""
                    Vector(Cicla, 4) = ""
                    Vector(Cicla, 5) = ""
                    Vector(Cicla, 6) = ""
                    Vector(Cicla, 7) = ""
                        Else
                    Vector(Cicla, 6) = Str$(WNeto)
                    Vector(Cicla, 7) = Str$(WImpoIb)
                End If
                rstCtaCte.Close
                    Else
                Vector(Cicla, 1) = ""
                Vector(Cicla, 2) = ""
                Vector(Cicla, 3) = ""
                Vector(Cicla, 4) = ""
                Vector(Cicla, 5) = ""
                Vector(Cicla, 6) = ""
                Vector(Cicla, 7) = ""
            End If
            
        Next Cicla
        
        For Cicla = 1 To Renglon
        
            WClave = Vector(Cicla, 1)
            If WClave <> "" Then
            
                WTipo = Vector(Cicla, 3)
                WNumero = Vector(Cicla, 4)
                WRecibo = Val(Left$(WClave, 6))
                WSale = "N"
            
                XParam = "'" + WTipo + "','" _
                             + WNumero + "'"
                spRecibo = "ListaRecibosFactura " + XParam
                Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
                If rstRecibo.RecordCount > 0 Then
                    With rstRecibo
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                If Val(rstRecibo!Recibo) < Val(WRecibo) Then
                                    WSale = "S"
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstRecibo.Close
                End If
                
                If WSale = "S" Then
                    Vector(Cicla, 1) = ""
                    Vector(Cicla, 2) = ""
                    Vector(Cicla, 3) = ""
                    Vector(Cicla, 4) = ""
                    Vector(Cicla, 5) = ""
                    Vector(Cicla, 6) = ""
                    Vector(Cicla, 7) = ""
                End If
                
            End If
            
        Next Cicla
        
        
        For Cicla = 1 To Renglon
        
            WClave = Vector(Cicla, 1)
            
            If WClave <> "" Then
            
                WRecibo = "00" + Left$(Vector(Cicla, 1), 6)
            
                WClave = Vector(Cicla, 1)
                WFecha = Vector(Cicla, 2)
                WTipo = Vector(Cicla, 3)
                WNumero = Vector(Cicla, 4)
                WCliente = Vector(Cicla, 5)
                WNeto = Vector(Cicla, 6)
                WImpoIb = Val(Vector(Cicla, 7))
                Call Redondeo(WImpoIb)
                If WImpoIb > 0 Then
                    XNeto = WNeto
                    XImpoIb = Str$(WImpoIb)
                    Call Ceros(XNeto, 12)
                    Call Ceros(XImpoIb, 11)
                    
                    XNeto = Alinea("#########.##", XNeto)
                    XImpoIb = Alinea("########.##", XImpoIb)
                    
                    Auxi = XImpoIb
                    For ZZCiclo = 1 To 11
                        If Mid$(Auxi, ZZCiclo, 1) = " " Then
                            Auxi = Left$(Auxi, ZZCiclo - 1) + "0" + Mid$(Auxi, ZZCiclo + 1, 11)
                        End If
                    Next ZZCiclo
                    XImpoIb = Auxi
                    
                    Auxi = XNeto
                    For ZZCiclo = 1 To 12
                        If Mid$(Auxi, ZZCiclo, 1) = " " Then
                            Auxi = Left$(Auxi, ZZCiclo - 1) + "0" + Mid$(Auxi, ZZCiclo + 1, 11)
                        End If
                    Next ZZCiclo
                    XNeto = Auxi
                    
                    TipoFac = "F"
                        Else
                    XNeto = Str$(Abs(Val(WNeto)))
                    XImpoIb = Str$(Abs(WImpoIb))
                    Call Ceros(XNeto, 11)
                    Call Ceros(XImpoIb, 10)
                    
                    XNeto = Alinea("########.##", XNeto)
                    XImpoIb = Alinea("#######.##", XImpoIb)
                    
                    Auxi = XImpoIb
                    For ZZCiclo = 1 To 10
                        If Mid$(Auxi, ZZCiclo, 1) = " " Then
                            Auxi = Left$(Auxi, ZZCiclo - 1) + "0" + Mid$(Auxi, ZZCiclo + 1, 11)
                        End If
                    Next ZZCiclo
                    XImpoIb = Auxi
                    
                    Auxi = XNeto
                    For ZZCiclo = 1 To 11
                        If Mid$(Auxi, ZZCiclo, 1) = " " Then
                            Auxi = Left$(Auxi, ZZCiclo - 1) + "0" + Mid$(Auxi, ZZCiclo + 1, 11)
                        End If
                    Next ZZCiclo
                    XNeto = Auxi
                    
                    XNeto = "-" + XNeto
                    XImpoIb = "-" + XImpoIb
                    Rem XNeto = Str$(Abs(Val(WNeto)))
                    Rem XImpoIb = Str$(Abs(WImpoIb))
                    TipoFac = "C"
                End If
                XRecibo = "00" + Left$(Vector(Cicla, 1), 6)
                
                spClientes = "ConsultaClientes " + "'" + WCliente + "'"
                Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
                If rstClientes.RecordCount > 0 Then
                    WCuit = Left$(rstClientes!Cuit, 13)
                    Rem Call Eval
                    rstClientes.Close
                End If
                
                Call Ceros(WNumero, 8)
                Rem Call Ceros(WCuit, 11)
                Rem Call Ceros(XNeto, 12)
                Rem Call Ceros(XImpoIb, 11)
            
                WImpre = WCuit + WFecha + TipoFac + "A0001" + WNumero + XNeto + XImpoIb + WFecha + "A"
            
                Print #1, WImpre
                    
            End If
            
        Next Cicla
        
        Close #1
    
            Else
    
    

        
        Rem Procesa las ventas
        
        Renglon = 0
        Erase Vector
        
        XParam = "'" + WDesde + "','" _
                     + WHasta + "'"
        spCtaCte = "ListaCtaCteDesdeHastaFecha" + XParam
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            With rstCtaCte
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstCtaCte!impoib <> 0 Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstCtaCte!Clave
                            Vector(Renglon, 2) = rstCtaCte!Fecha
                            Vector(Renglon, 3) = rstCtaCte!Tipo
                            Vector(Renglon, 4) = rstCtaCte!Numero
                            Vector(Renglon, 5) = rstCtaCte!Cliente
                            Vector(Renglon, 6) = ""
                            Vector(Renglon, 7) = ""
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCtaCte.Close
        End If
         
        For Cicla = 1 To Renglon
        
            WClave = Vector(Cicla, 1)
            WFecha = Vector(Cicla, 2)
            WTipo = Vector(Cicla, 3)
            WNumero = Vector(Cicla, 4)
            
            ClaveCtacte = WTipo + WNumero + "01"
            spCtaCte = "ConsultaCtacte " + "'" + WClave + "'"
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                WNeto = IIf(IsNull(rstCtaCte!Neto), "0", rstCtaCte!Neto)
                WImpoIb = IIf(IsNull(rstCtaCte!impoib), "0", rstCtaCte!impoib)
                If WImpoIb = 0 Then
                    Vector(Cicla, 1) = ""
                    Vector(Cicla, 2) = ""
                    Vector(Cicla, 3) = ""
                    Vector(Cicla, 4) = ""
                    Vector(Cicla, 5) = ""
                    Vector(Cicla, 6) = ""
                    Vector(Cicla, 7) = ""
                        Else
                    Vector(Cicla, 6) = Str$(WNeto)
                    Vector(Cicla, 7) = Str$(WImpoIb)
                End If
                rstCtaCte.Close
                    Else
                Vector(Cicla, 1) = ""
                Vector(Cicla, 2) = ""
                Vector(Cicla, 3) = ""
                Vector(Cicla, 4) = ""
                Vector(Cicla, 5) = ""
                Vector(Cicla, 6) = ""
                Vector(Cicla, 7) = ""
            End If
            
        Next Cicla
        
        For Cicla = 1 To Renglon
        
            WClave = Vector(Cicla, 1)
            
            If WClave <> "" Then
            
                WClave = Vector(Cicla, 1)
                WFecha = Vector(Cicla, 2)
                WTipo = Vector(Cicla, 3)
                WNumero = Vector(Cicla, 4)
                WCliente = Vector(Cicla, 5)
                WNeto = Vector(Cicla, 6)
                WImpoIb = Val(Vector(Cicla, 7))
                Call Redondeo(WImpoIb)
                
                If WImpoIb > 0 Then
                    XNeto = WNeto
                    XImpoIb = Str$(WImpoIb)
                    Call Ceros(XNeto, 12)
                    Call Ceros(XImpoIb, 11)
                    
                    XNeto = Alinea("#########.##", XNeto)
                    XImpoIb = Alinea("########.##", XImpoIb)
                    
                    Auxi = XImpoIb
                    For ZZCiclo = 1 To 11
                        If Mid$(Auxi, ZZCiclo, 1) = " " Then
                            Auxi = Left$(Auxi, ZZCiclo - 1) + "0" + Mid$(Auxi, ZZCiclo + 1, 11)
                        End If
                    Next ZZCiclo
                    XImpoIb = Auxi
                    
                    Auxi = XNeto
                    For ZZCiclo = 1 To 12
                        If Mid$(Auxi, ZZCiclo, 1) = " " Then
                            Auxi = Left$(Auxi, ZZCiclo - 1) + "0" + Mid$(Auxi, ZZCiclo + 1, 11)
                        End If
                    Next ZZCiclo
                    XNeto = Auxi
                    
                    TipoFac = "F"
                        Else
                    XNeto = Str$(Abs(Val(WNeto)))
                    XImpoIb = Str$(Abs(WImpoIb))
                    Call Ceros(XNeto, 11)
                    Call Ceros(XImpoIb, 10)
                    
                    XNeto = Alinea("########.##", XNeto)
                    XImpoIb = Alinea("#######.##", XImpoIb)
                    
                    Auxi = XImpoIb
                    For ZZCiclo = 1 To 10
                        If Mid$(Auxi, ZZCiclo, 1) = " " Then
                            Auxi = Left$(Auxi, ZZCiclo - 1) + "0" + Mid$(Auxi, ZZCiclo + 1, 11)
                        End If
                    Next ZZCiclo
                    XImpoIb = Auxi
                    
                    Auxi = XNeto
                    For ZZCiclo = 1 To 11
                        If Mid$(Auxi, ZZCiclo, 1) = " " Then
                            Auxi = Left$(Auxi, ZZCiclo - 1) + "0" + Mid$(Auxi, ZZCiclo + 1, 11)
                        End If
                    Next ZZCiclo
                    XNeto = Auxi
                    
                    XNeto = "-" + XNeto
                    XImpoIb = "-" + XImpoIb
                    Rem XNeto = Str$(Abs(Val(WNeto)))
                    Rem XImpoIb = Str$(Abs(WImpoIb))
                    TipoFac = "C"
                End If
                XRecibo = Numero
                
                spClientes = "ConsultaClientes " + "'" + WCliente + "'"
                Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
                If rstClientes.RecordCount > 0 Then
                    WCuit = Left$(rstClientes!Cuit, 13)
                    Rem Call Eval
                    rstClientes.Close
                End If
                
                Call Ceros(WNumero, 8)
                Rem Call Ceros(WCuit, 11)
                Rem Call Ceros(XNeto, 12)
                Rem Call Ceros(XImpoIb, 11)
            
                Rem WImpre = WCuit + WFecha + TipoFac + "A0001" + WNumero + XNeto + XImpoIb + WFecha + "A"
                WImpre = WCuit + WFecha + TipoFac + "A0001" + WNumero + XNeto + XImpoIb + WFecha + "A"
            
                Print #1, WImpre
                    
            End If
            
        Next Cicla
        
        Close #1
    
    
    End If
    
    
    
    
    Call Cancela_Click
        
End Sub

Private Sub Cancela_Click()
    Desde.SetFocus
    PrgProcesoPerceIb.Hide
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
            Nombre.SetFocus
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

    TipoList.AddItem "Cobranzas"
    TipoList.AddItem "Facturacion"

    TipoList.ListIndex = 0

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Nombre.Text = ""
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


