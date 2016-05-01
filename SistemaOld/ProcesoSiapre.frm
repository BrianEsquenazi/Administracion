VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProcesoSiapre 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Traspaso de Percepcioes SIAPRE (Aduana)"
   ClientHeight    =   5775
   ClientLeft      =   3060
   ClientTop       =   1425
   ClientWidth     =   6930
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   6930
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   4815
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
Attribute VB_Name = "PrgProcesoSiapre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstImputac As Recordset
Dim spImputac As String
Dim XParam As String
Dim Vector(10000, 10) As String
Dim WClave As String
Dim WFecha As String
Dim WTipo As String
Dim WNumero As String
Dim WImpoIva As Double
Dim XNeto As String
Dim XImpoIva As String
Dim WCuit As String

Private Sub Drive_Change()
    Dir1.Path = Drive.Drive
End Sub

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
    
    Renglon = 0
    Erase Vector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *, IvaComp.Despacho as [WDespacho], Proveedor.Cuit as [WCuit]"
    ZSql = ZSql + " FROM Imputac, IvaComp, Proveedor"
    ZSql = ZSql + " Where Imputac.NroInterno = IvaComp.NroInterno"
    ZSql = ZSql + " and Imputac.Proveedor = Proveedor.Proveedor"
    ZSql = ZSql + " and Imputac.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Imputac.FechaOrd <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Imputac.Debito <> 0"
    
    spImputac = ZSql
    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
    If rstImputac.RecordCount > 0 Then
            
        With rstImputac
            .MoveFirst
            Do
            
                If rstImputac!Debito <> 0 Then
                    
                    If !Proveedor = "10069345023" Or !Proveedor = "10014123562" Or !Proveedor = "10022098824" Then
                            
                        WCodigo = ""
                        Select Case Val(rstImputac!Cuenta)
                            Case 168
                                WCodigo = "924"
                            Case Else
                        End Select
                        
                        If Val(WCodigo) <> 0 Then
                        
                            WNumero = IIf(IsNull(rstImputac!WDespacho), "", rstImputac!WDespacho)
                            WNumero = Left$(WNumero + Space$(20), 20)
                            
                            WImpoIva = rstImputac!Debito
                            Call Redondeo(WImpoIva)
                            
                            WCuit = Left$(rstImputac!WCuit, 13)
                
                            XImpoIva = Str$(WImpoIva)
                            XImpoIva = Pusing("########.##", XImpoIva)
                            Call Ceros(XImpoIva, 10)
                            Mid$(XImpoIva, 8, 1) = ","
                        
                            Campo1 = WCodigo
                            Campo2 = WCuit
                            Campo3 = rstImputac!Fecha
                            Campo4 = WNumero
                            Campo5 = XImpoIva
                        
                            If WImpoIva > 0 Then
                                WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5
                                Print #1, WImpre
                            End If
                            
                        End If
                        
                    End If
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                End If
            Loop
        End With
        rstImputac.Close
    End If
    
    Close #1
    
    Call Cancela_click
        
End Sub

Private Sub Cancela_click()
    PrgProcesoSiapre.Hide
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


