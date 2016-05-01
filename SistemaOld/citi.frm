VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCiti 
   AutoRedraw      =   -1  'True
   Caption         =   "Exportacion de Datos al CITI COMPRAS"
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
Attribute VB_Name = "PrgCiti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstIvaComp As Recordset
Dim spIvaComp As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String
Dim XVector(10000, 6) As String
Dim WTipo As String
Dim WPunto As String
Dim WNumero As String
Dim WCuit As String
Dim WNombre As String
Dim WIva As String

Private Sub Acepta_Click()

    WDrive = Drive.Drive
    WDir = Dir1.Path
    
    XNombre = WDir + "\" + Nombre.Text + ".txt"
    
    Open XNombre For Output As #1

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    Erase XVector
    Renglon = 0
    
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
                    If Val(!Tipo) = 1 Or Val(!Tipo) = 2 Then
                        Iva = !Iva21 + !Iva27
                        If Iva <> 0 Then
                            Renglon = Renglon + 1
                            XVector(Renglon, 1) = !Tipo
                            XVector(Renglon, 2) = !Punto
                            XVector(Renglon, 3) = !Numero
                            XVector(Renglon, 4) = Left$(!Fecha, 2) + Mid$(!Fecha, 4, 2) + Right$(!Fecha, 4)
                            XVector(Renglon, 5) = Str$(Int(Iva * 100))
                            XVector(Renglon, 6) = !Proveedor
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
    
    For Ciclo = 1 To Renglon
    
        WTipo = XVector(Ciclo, 1)
        WPunto = XVector(Ciclo, 2)
        WNumero = XVector(Ciclo, 3)
        WFecha = XVector(Ciclo, 4)
        WIva = XVector(Ciclo, 5)
        WProveedor = XVector(Ciclo, 6)
        
        spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            WNombre = RstProveedor!Nombre
            WNombre = WNombre + Space$(25)
            WNombre = Left$(WNombre, 25)
            WCuit = RstProveedor!Cuit
            Call Eval
            RstProveedor.Close
        End If
        
        Call Ceros(WTipo, 2)
        Call Ceros(WPunto, 4)
        Call Ceros(WNumero, 20)
        Call Ceros(WIva, 12)
        Call Ceros(WCuit, 11)
        
        If Val(WEmpresa) = 1 Then
            WCuitII = "30549165083"
            WNombreII = Left$("SURFACTAN S.A." + Space$(25), 25)
                Else
            WCuitII = ""
            WNombreII = ""
        End If
        WCuitII = "00000000000"
        WNombreII = Space$(25)
        
        
        
        WImpre = WTipo + WPunto + WNumero + WFecha + WCuit + WNombre + WIva + WCuitII + WNombreII + "000000000000"
    
        Print #1, WImpre
        
    Next Ciclo
    
    Close #1
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Desde.SetFocus
    PrgCiti.Hide
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


