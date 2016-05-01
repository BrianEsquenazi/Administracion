VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProcesoReteIb 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Traspaso de Retenciones de Ingresos Brutos"
   ClientHeight    =   5580
   ClientLeft      =   3060
   ClientTop       =   1455
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   ScaleHeight     =   5580
   ScaleWidth      =   6240
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   600
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
         MaxLength       =   50
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
Attribute VB_Name = "PrgProcesoReteIb"
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

Private Sub Drive_Change()
    Dir1.Path = Drive.Drive
End Sub

Private Sub Acepta_Click()
    
    WDrive = Drive.Drive
    WDir = Dir1.Path
    
    If Val(WEmpresa) = 1 Then
        XNombre = WDir + "\AR-30549165083-" + Nombre.Text + "-6-LOTE1.txt"
            Else
        XNombre = WDir + "\AR-30610524598-" + Nombre.Text + "-6-LOTE1.txt"
    End If
    
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
            
    spPagos = "ListaPagos"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
            
        With rstPagos
            .MoveFirst
            Do
            
                If WDesde <= !FechaOrd And !FechaOrd <= WHasta Then
                    If !RetOtra <> 0 Then
                        If !Renglon = 1 Then
                            LugarVector = LugarVector + 1
                            Vector(LugarVector, 1) = !Proveedor
                            Vector(LugarVector, 2) = !Fecha
                            Vector(LugarVector, 3) = !Orden
                            Vector(LugarVector, 4) = Str$(!RetOtra)
                        End If
                    End If
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        
        rstPagos.Close
        
    End If
    
    For Ciclo = 1 To LugarVector
    
        WProveedor = Vector(Ciclo, 1)
        WFecha = Left$(Vector(Ciclo, 2), 2) + "/" + Mid$(Vector(Ciclo, 2), 4, 2) + "/" + Right$(Vector(Ciclo, 2), 4)
        WOrden = Vector(Ciclo, 3)
        WImporte = Vector(Ciclo, 4)
    
        spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            WCuit = Trim(RstProveedor!Cuit)
            Rem Call Eval
            RstProveedor.Close
        End If
    
        Call Ceros(WOrden, 8)
        Rem Call Ceros(WCuit, 11)
        Call Ceros(WImporte, 11)
        WSucursal = "0001"
        
        WImporte = Alinea("########.##", WImporte)
        Auxi = WImporte
        For ZZCiclo = 1 To 11
            If Mid$(Auxi, ZZCiclo, 1) = " " Then
                Auxi = Left$(Auxi, ZZCiclo - 1) + "0" + Mid$(Auxi, ZZCiclo + 1, 11)
            End If
        Next ZZCiclo
        WImporte = Auxi
        
        WImpre = WCuit + WFecha + WSucursal + WOrden + WImporte + "A"
        
        Print #1, WImpre
        
    Next Ciclo
    
    Close #1
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Desdefecha.SetFocus
    PrgProcesoReteIb.Hide
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

