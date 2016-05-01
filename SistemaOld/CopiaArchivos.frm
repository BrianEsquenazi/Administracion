VERSION 5.00
Begin VB.Form PrgCopiaArchivos 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Archivos Adjuntados a la OC"
   ClientHeight    =   9180
   ClientLeft      =   3810
   ClientTop       =   330
   ClientWidth     =   7245
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   9180
   ScaleWidth      =   7245
   Visible         =   0   'False
   Begin VB.Frame IngresaArchivo 
      BackColor       =   &H00FFC0C0&
      Height          =   8775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton CierraPanta 
         Caption         =   "Cierra Pantalla"
         Height          =   615
         Left            =   2520
         TabIndex        =   5
         Top             =   8040
         Width           =   1695
      End
      Begin VB.TextBox Archivo 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Text            =   " "
         Top             =   7560
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.FileListBox File1 
         Height          =   3405
         Left            =   360
         TabIndex        =   3
         Top             =   3840
         Width           =   6015
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   6015
      End
      Begin VB.DirListBox Dir1 
         Height          =   3015
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   6015
      End
      Begin VB.Image AceptaFoto 
         Height          =   480
         Left            =   600
         MouseIcon       =   "CopiaArchivos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "CopiaArchivos.frx":030A
         ToolTipText     =   "Confirma el Proceso"
         Top             =   8160
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "PrgCopiaArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZBusca(10000) As String
Dim ZZLugarBusca As Integer

Private Sub cmdAgregar_Click()
    IngresaArchivo.Visible = True
End Sub

Private Sub AceptaFoto_Click()
    Call File1_dBLClick
End Sub

Private Sub CierraPanta_Click()
    Unload Me
    PrgOrdenArchivos.Show
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path  ' Establece la ruta del archivo.
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
End Sub

Private Sub File1_dBLClick()

    Rem On Error GoTo WError
    
    ZZNombre = Trim(File1.filename)
    ZZCaracteres = Len(ZZNombre)
    
    For ZZCiclo = ZZCaracteres To 1 Step -1
        If Mid$(ZZNombre, ZZCiclo, 1) = "." Then
            ZZExtension = Trim(Mid$(ZZNombre, ZZCiclo + 1, 10))
            Exit For
        End If
    Next ZZCiclo
    
    WNombre = Trim(File1.filename)
    Auxi = UCase(WNombre)
    For ZZCiclo = 1 To 100
        If Mid$(Auxi, ZZCiclo, 1) = " " Then
            Auxi = Left$(Auxi, ZZCiclo - 1) + "_" + Mid$(Auxi, ZZCiclo + 1, 100)
        End If
    Next ZZCiclo
    WNombre = Auxi
    
    
    Rem ZZExtension = Right$(UCase(File1.filename), 3)
    
    Select Case Trim(UCase(ZZExtension))
        Case "JPG", "BMP", "PDF", "DOC", "XLS", "XLSX", "DOCX"
            WDrive = Drive1.Drive
            WDir = Dir1.Path
            WLon = Len(WDir)
            If Right$(WDir, 1) = "\" Then
                WDir = Mid(WDir, 1, WLon - 1)
            End If
            XNombre = WDir + "\"
            
            ZZOrigen = XNombre + File1.filename
            ZZDestino = "W:\orden\" + Trim(Str$(Val(WPasaOrden))) + "\" + WNombre
            ZZDestinoII = "W:\orden\" + Trim(Str$(Val(WPasaOrden))) + "\"
        
            Rem Open "W:\COPIA.bat" For Output As #1
            Rem Print #1, "XCOPY " + ZZOrigen + " " + ZZDestinoII
            Rem Close #1
            Rem Shell "W:\COPIA.bat"
        
            FileCopy ZZOrigen, ZZDestino
            
            m$ = "Copia Realizada " + Chr$(13) + _
                 ZZOrigen + Chr$(13) + _
                 ZZDestino
            
            G% = MsgBox(m$, 0, "Mercaderia en Consignacion")
        
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Rem MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
    
End Sub

Private Sub Form_Load()

    Drive1.Drive = "C:"
    Auxi = "C:\"
    
    ZZEstado = Dir(Auxi)
    If ZZEstado <> "" Then
        Dir1.Path = Auxi
    End If
    
End Sub

