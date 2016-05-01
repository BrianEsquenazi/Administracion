VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProcesoReteGanan 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Traspaso de Retenciones de Ganancias de Terceros"
   ClientHeight    =   5580
   ClientLeft      =   3060
   ClientTop       =   1455
   ClientWidth     =   7185
   LinkTopic       =   "Form2"
   ScaleHeight     =   5580
   ScaleWidth      =   7185
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   4815
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
         TabIndex        =   9
         Top             =   1440
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   1920
         TabIndex        =   7
         Top             =   1920
         Width           =   2055
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   1920
         TabIndex        =   6
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
      Begin MSMask.MaskEdBox Desdefecha 
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
         TabIndex        =   3
         Top             =   4200
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
         TabIndex        =   2
         Top             =   4200
         Width           =   1215
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
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
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
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
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
         TabIndex        =   5
         Top             =   720
         Width           =   1695
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
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "PrgProcesoReteGanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPagos As Recordset
Dim spPagos As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String
Dim Vector(10000, 10) As String
Private LugarVector As String
Private WOrden As String
Private WCuit As String
Private WImporte As String
Private WRetencion As String
Private WBase As String
Private WCertificadoGan As String
Private WImpre4 As String
Private WImpre8 As String
Private WImpre11 As String
Private WImpre16 As String


Private Sub Drive_Change()
    Dir1.Path = Drive.Drive
End Sub

Private Sub Acepta_Click()
    
    WDrive = Drive.Drive
    WDir = Dir1.Path
    
    XNombre = WDir + "\" + Nombre.Text + ".txt"
    
    Open XNombre For Output As #1

    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia
    
    Erase Vector
    LugarVector = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *, Proveedor.Cuit as [ProveedorCuit], Proveedor.Tipo as [ProveedorTipo]"
    ZSql = ZSql + " FROM Pagos, Proveedor"
    ZSql = ZSql + " Where Pagos.Proveedor = Proveedor.Proveedor"
    ZSql = ZSql + " and Pagos.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Pagos.FechaOrd <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Pagos.Retencion <> 0"
    ZSql = ZSql + " and Pagos.Renglon = 1"
    
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
            
        With rstPagos
            .MoveFirst
            Do
            
                    WFecha = !Fecha
                    WOrden = !Orden
                    WImporte = Str$(!Importe)
                    WBase = Str$(!Importe)
                    WRetencion = Str$(!Retencion)
                    WProveedor = !Proveedor
                    WCuit = !ProveedorCuit
                    WCertificadoGan = IIf(IsNull(!certificadoGan), "0", !certificadoGan)
                    Rem      WCertificadoGan = !certificadoGan
                    WTipoprv = !ProveedorTipo
                    Call Eval
                    
                    Rem WDesProveedor = !Nombre
                    Rem WPrvDireccion = !Direccion
                    Rem WTipoprv = !Ganancia
                    Rem WTipoiva = !Iva
                    Rem WTipoReteiva = !Reteiva
                    Rem WExepcion = !PorceReteIva
                    WImpre16 = WCertificadoGan
                    Call Ceros(WImpre16, 16)
    
                    Call Ceros(WOrden, 16)
                    Call Ceros(WCuit, 11)
                    Call Ceros(WImporte, 16)
                    Call Ceros(WRetencion, 14)
                    Call Ceros(WBase, 14)
        
                    WImpre1 = "06"
                    WImpre2 = WFecha
                    WImpre3 = WImpre16
                    WImpre4 = WImporte
                    WImpre5 = "217"
                    
                    Select Case Val(WTipoprv)
                        Case 1
                            WImpre6 = "116"
                        Case 2
                            WImpre6 = "27 "
                        Case 4
                            WImpre6 = "124"
                        Case 5
                            WImpre6 = "094"
                        Case 6
                            WImpre6 = "95 "
                        Case Else
                            WImpre6 = "78 "
                    End Select
                    
                    WImpre7 = "1"
                    WImpre8 = WBase
                    WImpre9 = WFecha
                    WImpre10 = "01 "
                    WImpre11 = WRetencion
                    WImpre12 = "000000"
                    WImpre13 = Space$(10)
                    WImpre14 = "80"
                    WImpre15 = Left$(WCuit + Space$(20), 20)
                    WImpre16 = WCertificadoGan
                    
                    WImpre17 = Space$(30)
                    WImpre18 = "0"
                    WImpre19 = Space$(11)
                    WImpre20 = Space$(11)
                    
                    Call Ceros(WImpre4, 16)
                    Call Ceros(WImpre8, 14)
                    Call Ceros(WImpre11, 14)
                    Call Ceros(WImpre16, 14)
                    
                    WImpre = WImpre1 + WImpre2 + WImpre3 + WImpre4 + WImpre5 + WImpre6 + WImpre7 + WImpre8 + WImpre9 + WImpre10 + WImpre11 + WImpre12 + WImpre13 + WImpre14 + WImpre15 + WImpre16 + WImpre17 + WImpre18 + WImpre19 + WImpre20
        
                    Print #1, WImpre
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
        
        rstPagos.Close
        
    End If
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Desdefecha.SetFocus
    PrgProcesoReteGanan.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Nombre.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefecha.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    Desdefecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
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

