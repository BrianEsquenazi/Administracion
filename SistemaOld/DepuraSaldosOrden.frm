VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDepuraSaldoOrden 
   AutoRedraw      =   -1  'True
   Caption         =   "Depuracion de Saldos de Ordenes de Compra"
   ClientHeight    =   2625
   ClientLeft      =   2025
   ClientTop       =   1050
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   2625
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   6
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
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
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
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
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1560
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
         Left            =   840
         TabIndex        =   4
         Top             =   1560
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1575
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgDepuraSaldoOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WOrden As String
Private WClave As String

Dim rstOrden As Recordset
Dim spOrden As String

Dim XParam As String

Dim Vector(10000, 4) As String
Dim Empe(12, 10) As String

Private WDescripcion As String
Private WSaldo As Double
Private XSaldo As Double

Private Sub Acepta_Click()

    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)

    Erase Vector
    Renglon = 0
    
    spOrden = "ListaOrdenTotal"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    If !FechaOrd >= WDesde And !FechaOrd <= WHasta Then
                        If !Saldo <> 0 Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstOrden!Clave
                            Vector(Renglon, 2) = rstOrden!Orden
                            Vector(Renglon, 3) = rstOrden!Articulo
                            Vector(Renglon, 4) = Str$(rstOrden!Cantidad)
                        End If
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        WClave = Vector(Ciclo, 1)
        WOrden = Vector(Ciclo, 2)
        WArticulo = Vector(Ciclo, 3)
        WCantidad = Val(Vector(Ciclo, 4))
        WResta = 0
        
        ZSql = ""
        ZSql = ZSql & "UPDATE Orden SET "
        ZSql = ZSql & "Cantidad = Recibida"
        ZSql = ZSql & " Where Clave = " + "'" + WClave + "'"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    PrgDepuraSaldoOrden.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Cancela_click()
    PrgDepuraSaldoOrden.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


