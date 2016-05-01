VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProcesoGananciaAduana 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Traspaso de Percepcion de Ganancias Aduana"
   ClientHeight    =   5775
   ClientLeft      =   3060
   ClientTop       =   1425
   ClientWidth     =   7290
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   7290
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   1200
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   4200
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
Attribute VB_Name = "PrgProcesoGananciaAduana"
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
Dim Vector(10000, 15) As String
Dim WClave As String
Dim WFecha As String
Dim WTipo As String
Dim WNumero As String

Dim WNeto As Double
Dim WIva1 As Double
Dim WIva2 As Double
Dim WImpoIb As Double
Dim WImpoIbTucu As Double
Dim WImpoIbCiudad As Double
Dim WTotal As Double

Dim WIbTucu As Integer

Dim WPorceIb As Double
Dim WPorceIbCiudad As Double

Dim XNeto As String
Dim XImpoIb As String
Dim XPorceIb As String
Dim XImpoIbCiudad As String
Dim XPorceIbCiudad As String

Dim WCuit As String
Dim ZEntra(1000) As String
Dim WNroIbTucu As String
Dim WNombre As String
Dim WDomicilio As String
Dim WPuerta As String
Dim WLocalidad As String
Dim WProvincia As String
Dim WPostal As String
Dim Provincia(100) As String
Dim XTotal As String
Dim WOtros As Double
Dim XOtros As String
Dim WIva As Double
Dim XIva As String
Dim ZZAlicuota As Double

Dim WNroIbCiudad As String
Dim WNroIbCiudadII As String

Dim WRegimen As String
Dim WImporte As String


Private Sub Drive_Change()
    Dir1.Path = Drive.Drive
End Sub

Private Sub Acepta_Click()

    WDrive = Drive.Drive
    WDir = Dir1.Path
    
    XNombre = WDir + "\" + "aduana.Txt"
    Open XNombre For Output As #1
    
    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    
    
    Rem Procesa las ventas
    OPEN_FILE_Aduana
    
    da = ""
    With rstAduana
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
            
            
                WCuit = !Cuit
                WFecha = !Fecha
                WRegimen = !regimen
                Call Ceros(WRegimen, 3)
                WImporte = !Importe
                Call Ceros(WImporte, 15)
                WNroComprobante = Left$(!NroComprobante + Space$(30), 30)
                
            
        
                Rem Tipo de Operacion
                Campo1 = "2"
        
                Rem Cuit
                Campo2 = WCuit
        
                Rem fecha
                Campo3 = WFecha
        
                Rem Codigo de Regimen
                Campo4 = WRegimen
                
            
                Rem Importe
                Campo5 = WImporte
        
                Rem Numero del Comprobante
                Campo6 = WNroComprobante
        
                WImpre = Campo1 + Campo2 + Campo3 + Campo4 + Campo5 + Campo6 + Campo7 + Campo8 + Campo9 + Campo10 + Campo11 + Campo12 + Campo13 + Campo14 + Campo15 + Campo16 + Campo17 + Campo18 + Campo19 + Campo20 + Campo21
                
                
            
                Print #1, WImpre
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    Close #1
    
    Call Cancela_Click
        
End Sub

Private Sub Cancela_Click()
    PrgProcesoGananciaAduana.Hide
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


