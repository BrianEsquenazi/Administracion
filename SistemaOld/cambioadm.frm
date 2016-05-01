VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCambioAdm 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Cambios"
   ClientHeight    =   4560
   ClientLeft      =   2505
   ClientTop       =   780
   ClientWidth     =   7245
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   7245
   Begin VB.Frame Clave 
      Caption         =   "  Ingreso de Clave de Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1920
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox WClave 
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
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
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
         TabIndex        =   17
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C000&
         Caption         =   "Ingrese su Password"
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
         Left            =   960
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCambios.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cambios"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "cambioadm.frx":0000
      Left            =   480
      List            =   "cambioadm.frx":0007
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
      Width           =   3255
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
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
         Left            =   1680
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
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
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
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
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
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
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Cambio 
      Alignment       =   1  'Right Justify
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
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox Opcion 
      Height          =   1230
      Left            =   840
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Paridad del Dolar"
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fecha"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "PrgCambioAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCambioAdm As Recordset
Dim spCambioAdm As String
Dim XParam As String
Private WGraba As String
Private WProceso As String

Sub Verifica_datos()
    If Val(Cambio.Text) = 0 Then
        Cambio.Text = "0"
    End If
End Sub

Sub Format_datos()
    Cambio.Text = Pusing("###,###.####", Cambio.Text)
End Sub

Sub Imprime_Datos()

    WFecha = Fecha.Text
    spCambioAdm = "ConsultaCambioAdm " + "'" + Fecha.Text + "'"
    Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambioAdm.RecordCount > 0 Then
        Fecha.Text = rstCambioAdm!Fecha
        Cambio.Text = rstCambioAdm!Cambio
        Call Format_datos
    End If

End Sub

Private Sub cmdAdd_Click()
    If Fecha.Text <> "" Then
    
        WProceso = 0
        If WGraba <> "S" Then
        
            Call Ingresa_clave
            
                Else
                
            WGraba = ""
    
    
    
            spCambioAdm = "ConsultaCambioAdm " + "'" + Fecha.Text + "'"
            Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambioAdm.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
        
       
        
        
        
        
            
           
             WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        
            Call Verifica_datos
            If WPasa = "N" Then
                XParam = "'" + Fecha.Text + "','" + Cambio.Text + "','" + WOrdFecha + "'"
                Set rstCambioAdm = db.OpenRecordset("AltaCambioAdm " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + Fecha.Text + "','" + Cambio.Text + "','" + WOrdFecha + "'"
                Set rstCambioAdm = db.OpenRecordset("ModificaCambioAdm " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            End If
    
            Call CmdLimpiar_Click
            Fecha.SetFocus
            
      
       
       
            
        End If
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Fecha.Text <> "" Then
    
        WProceso = 1
        If WGraba <> "S" Then
        
            Call Ingresa_clave
            
                Else
    
            spCambioAdm = "ConsultaCambioAdm " + "'" + Fecha.Text + "'"
            Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambioAdm.RecordCount > 0 Then
                T$ = "Borrar Registro"
                m$ = "Desea Borrar el Registro "
                Respuesta% = MsgBox(m$, 32 + 4, T$)
                If Respuesta% = 6 Then
                    spCambioAdm = "BorrarCambioAdm " + "'" + Fecha.Text + "'"
                    Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenDynaset, dbSQLPassThrough)
                    Call CmdLimpiar_Click
                    Fecha.SetFocus
                End If
            End If
            
        End If
        
    End If
End Sub

Private Sub CmdLimpiar_Click()
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cambio.Text = ""
    Fecha.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Fecha.SetFocus
    PrgCambioAdm.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Cambio_Keypress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    '      Cuenta.SetFocus
    'End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Fecha.Text <> "" Then
            WFecha = Fecha.Text
            spCambioAdm = "ConsultaCambioAdm " + "'" + Fecha.Text + "'"
            Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambioAdm.RecordCount > 0 Then
                Fecha.Text = rstCambioAdm!Fecha
                Call Imprime_Datos
                    Else
                WFecha = Fecha.Text
                CmdLimpiar_Click
                Fecha.Text = WFecha
            End If
        End If
        Cambio.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     'Opcion.Clear
     '
     'Opcion.AddItem "Fecha"
     'Opcion.AddItem "Cuentas Contables"

     'Opcion.Visible = True
     
'End Sub

'' Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    'XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
            spCambioAdm = "ListaCambioAdm"
            Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstCambioAdm
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCambioAdm!Fecha + " " + Str$(rstCambioAdm!Cambio)
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCambioAdm!Fecha
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCambioAdm.Close
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WFecha = WIndice.List(Indice)
            spCambioAdm = "ConsultaCambioAdm " + "'" + WFecha + "'"
            Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambioAdm.RecordCount > 0 Then
                Fecha.Text = rstCambioAdm!Fecha
                Call Imprime_Datos
                        Else
                CmdLimpiar_Click
                Fecha.Text = WFecha
            End If
        
            Fecha.SetFocus
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cambio.Text = ""
End Sub


Private Sub Primer_Click()

    On Error GoTo WError
    
    spCambioAdm = "ListaCambioAdm"
    Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstCambioAdm
        .MoveFirst
        Fecha.Text = rstCambioAdm!Fecha
    End With
    
    rstCambioAdm.Close
    Call Imprime_Datos
    Fecha.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cambio", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Fecha.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spCambioAdm = "ListaCambioAdm"
    Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstCambioAdm
        .MoveLast
        Fecha.Text = rstCambioAdm!Fecha
    End With
    
    rstCambioAdm.Close
    Call Imprime_Datos
    Fecha.SetFocus
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Cambio", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Fecha.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spCambioAdm = "AnteriorCambioAdm " + "'" + Fecha.Text + "'"
    Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstCambioAdm
        .MoveLast
        Fecha.Text = rstCambioAdm!Fecha
    End With
    
    rstCambioAdm.Close
    Call Imprime_Datos
    Fecha.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cambio", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Fecha.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spCambioAdm = "PosteriorCambioAdm " + "'" + Fecha.Text + "'"
    Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstCambioAdm
        .MoveFirst
        Fecha.Text = rstCambioAdm!Fecha
    End With
    
    rstCambioAdm.Close
    Call Imprime_Datos
    Fecha.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Cambio", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Fecha.SetFocus
    
End Sub

Sub Ingresa_clave()

    WClave.Text = ""
    Clave.Visible = True
    WClave.SetFocus
    
End Sub

Private Sub CancelaGraba_Click()

    Clave.Visible = False
    Fecha.SetFocus

End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
        If WClave.Text = "SERGIO" Then
            WGraba = "S"
            Clave.Visible = False
            If WProceso = 0 Then
                Call cmdAdd_Click
                    Else
                Call cmdDelete_Click
            End If
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Archivo de Cambios")
            WClave.SetFocus
        End If
    End If
End Sub

