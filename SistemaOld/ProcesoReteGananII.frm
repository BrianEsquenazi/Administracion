VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProcesoReteGananII 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Traspaso de Retenciones de Ganancias"
   ClientHeight    =   5580
   ClientLeft      =   3060
   ClientTop       =   1455
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   ScaleHeight     =   5580
   ScaleWidth      =   6540
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   840
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
Attribute VB_Name = "PrgProcesoReteGananII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstRecibos As Recordset
Dim spRecibos As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Dim Vector(10000, 10) As String
Private LugarVector As String
Private WOrden As String
Private WCuit As String
Private WImporte As String
Private WComproGanan As String

Private Sub Command1_Click()

    Open "a:david.txt" For Input As #9
    Open "a:davidok.txt" For Output As #1
    
    Do
        Line Input #9, WLinea
            
        WCuit = Mid$(WLinea, 1, 11)
        WComproGanan = Mid$(WLinea, 12, 7)
        WFecha = Mid$(WLinea, 19, 2) + "/" + Mid$(WLinea, 21, 2) + "/" + Mid$(WLinea, 23, 4)
        WRegimen = "078"
        WImporte = Mid$(WLinea, 30, 13)
        
        Call Ceros(WComproGanan, 12)
        Call Ceros(WImporte, 15)
        
        WImpre1 = WCuit
        WImpre2 = WFecha
        WImpre3 = "078"
        WImpre4 = WImporte
        WImpre5 = WComproGanan
        
        WImpre = WImpre1 + WImpre2 + WImpre3 + WImpre4 + WImpre5
        
        Print #1, WImpre
        
            
    Loop Until EOF(9)
    
    Close #1
    Close #9

End Sub

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
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Recibos.FechaOrd <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Recibos.RetGanancias <> 0"
    ZSql = ZSql + " and Recibos.Renglon = 1"
    
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
            
        With rstRecibos
            .MoveFirst
            Do
            
                LugarVector = LugarVector + 1
                
                Vector(LugarVector, 1) = !Cliente
                Vector(LugarVector, 2) = !Fecha
                WComproGanan = IIf(IsNull(!ComproGanan), "", !ComproGanan)
                Vector(LugarVector, 3) = WComproGanan
                Vector(LugarVector, 4) = Str$(!Retganancias)
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
        
        rstRecibos.Close
        
    End If
    
    For Ciclo = 1 To LugarVector
    
        WCliente = Vector(Ciclo, 1)
        WFecha = Left$(Vector(Ciclo, 2), 2) + "/" + Mid$(Vector(Ciclo, 2), 4, 2) + "/" + Right$(Vector(Ciclo, 2), 4)
        WComproGanan = Vector(Ciclo, 3)
        Rem WComproGanan = Str$(Ciclo)
        WImporte = Vector(Ciclo, 4)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WCuit = Left$(rstCliente!Cuit, 13)
            Call Eval
            rstCliente.Close
        End If
    
        Call Ceros(WComproGanan, 30)
        Call Ceros(WImporte, 15)
        
        WImpre1 = "1" + WCuit
        WImpre2 = Left$(WFecha, 2) + "/" + Mid$(WFecha, 4, 2) + "/" + Mid$(WFecha, 7, 4)
        WImpre3 = "078"
        WImpre4 = WImporte
        WImpre5 = WComproGanan
        
        WImpre = WImpre1 + WImpre2 + WImpre3 + WImpre4 + WImpre5
        
        Print #1, WImpre
        
    Next Ciclo
    
    









    Erase Vector
    LugarVector = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Imputac"
    ZSql = ZSql + " Where Imputac.Cuenta = '144'"
    ZSql = ZSql + " and Imputac.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Imputac.FechaOrd <= " + "'" + WHasta + "'"
    spImpu = ZSql
    Set RstImpu = db.OpenRecordset(spImpu, dbOpenSnapshot, dbSQLPassThrough)
    If RstImpu.RecordCount > 0 Then
        With RstImpu
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    LugarVector = LugarVector + 1
                    Vector(LugarVector, 1) = !Proveedor
                    Vector(LugarVector, 2) = !Fecha
                    Vector(LugarVector, 3) = !NroComp
                    Vector(LugarVector, 4) = Str$(!Debito)
                        
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        RstImpu.Close
    End If
    
    
    For Ciclo = 1 To LugarVector
    
        WProveedor = Vector(Ciclo, 1)
        WFecha = Left$(Vector(Ciclo, 2), 2) + "/" + Mid$(Vector(Ciclo, 2), 4, 2) + "/" + Right$(Vector(Ciclo, 2), 4)
        WComproGanan = Vector(Ciclo, 3)
        WImporte = Vector(Ciclo, 4)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
        spProveedor = ZSql
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            WCuit = Left$(RstProveedor!Cuit, 13)
            Call Eval
            RstProveedor.Close
        End If
    
        Call Ceros(WComproGanan, 30)
        Call Ceros(WImporte, 15)
        
        WImpre1 = "2" + WCuit
        WImpre2 = Left$(WFecha, 2) + "/" + Mid$(WFecha, 4, 2) + "/" + Mid$(WFecha, 7, 4)
        WImpre3 = "531"
        WImpre4 = WImporte
        WImpre5 = WComproGanan
        
        WImpre = WImpre1 + WImpre2 + WImpre3 + WImpre4 + WImpre5
        
        Print #1, WImpre
        
    Next Ciclo
    
    Close #1
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Desdefecha.SetFocus
    PrgProcesoReteGananII.Hide
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
    MinusOk = 0                'a minus sign is okay only once, and only
                               'if it preceeds the first numeric character
    DecOk = 0                  'only the first decimal point is okay

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

