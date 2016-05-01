VERSION 5.00
Begin VB.Form PrgOrdenArchivos 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Archivos Adjuntados a la OC"
   ClientHeight    =   9075
   ClientLeft      =   3810
   ClientTop       =   330
   ClientWidth     =   8085
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   9075
   ScaleWidth      =   8085
   Visible         =   0   'False
   Begin VB.Frame IngresaArchivo 
      BackColor       =   &H00C0FFFF&
      Height          =   8775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton FinIngresaObserva 
         Caption         =   "Fin de Ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   6
         Top             =   7920
         Width           =   1815
      End
      Begin VB.CommandButton CopiaArchivo 
         Caption         =   "Agrega Archivos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   5
         Top             =   7920
         Width           =   1575
      End
      Begin VB.TextBox Archivo 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Text            =   " "
         Top             =   7440
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.FileListBox File1 
         Height          =   3210
         Left            =   360
         TabIndex        =   3
         Top             =   4080
         Width           =   6975
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   6975
      End
      Begin VB.DirListBox Dir1 
         Height          =   3240
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   6975
      End
      Begin VB.Image AceptaFoto 
         Height          =   480
         Left            =   480
         MouseIcon       =   "ordenarchivos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ordenarchivos.frx":030A
         ToolTipText     =   "Confirma el Proceso"
         Top             =   8040
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "PrgOrdenArchivos"
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

Private Sub Command1_Click()

End Sub

Private Sub CopiaArchivo_Click()
    PrgCopiaArchivos.Show
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path  ' Establece la ruta del archivo.
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
End Sub

Private Sub File1_dBLClick()

    On Error GoTo WError
    
    WDrive = Drive1.Drive
    WDir = Dir1.Path
    WLon = Len(WDir)
    If Right$(WDir, 1) = "\" Then
        WDir = Mid(WDir, 1, WLon - 1)
    End If
    XNombre = WDir + "\"
    
    WPasoUnifica = XNombre + File1.filename
    Archivo.Text = UCase(XNombre + File1.filename)
    
    
    ZZNombre = Trim(File1.filename)
    ZZCaracteres = Len(ZZNombre)
    For ZZCiclo = ZZCaracteres To 1 Step -1
        If Mid$(ZZNombre, ZZCiclo, 1) = "." Then
            ZZExtension = Trim(Mid$(ZZNombre, ZZCiclo + 1, 10))
            Exit For
        End If
    Next ZZCiclo
    
    
    ZZExtension = UCase(ZZExtension)
    
    
    Rem If ZZExtension = "JPG" Or ZZExtension = "BMP" Then
    Rem     WWWPasaFoto = Archivo.Text
    Rem     PrgOrdenArchivosFoto.Show
    Rem End If
    
    If ZZExtension = "PDF" Or ZZExtension = "DOC" Or ZZExtension = "XLS" Or ZZExtension = "DOCX" Or ZZExtension = "XLSX" Or ZZExtension = "JPG" Or ZZExtension = "BMP" Then
    
        Erase ZZBusca
        ZLugarbusca = 0
       
        ' Muestra los nombres en C:\ que representan directorios.
        Select Case ZZExtension
            Case "PDF"
                ZZCodigoExe = "AcroRd32.exe"
                ZZPasaExe = ""
                Erase ZZBusca
                ZZLugarBusca = 1
                ZZBusca(ZZLugarBusca) = "c:\Archivos de programa\Adobe\"
                
            Case "DOC"
                ZZCodigoExe = "WINWORD.exe"
                ZZPasaExe = ""
                Erase ZZBusca
                ZZLugarBusca = 1
                ZZBusca(ZZLugarBusca) = "c:\Archivos de programa\Microsoft Office\"
                
            Case "XLS"
                ZZCodigoExe = "EXCEL.exe"
                ZZPasaExe = ""
                Erase ZZBusca
                ZZLugarBusca = 1
                ZZBusca(ZZLugarBusca) = "c:\Archivos de programa\Microsoft Office\"
                
            Case "DOCX"
                ZZCodigoExe = "WINWORD.exe"
                ZZPasaExe = ""
                Erase ZZBusca
                ZZLugarBusca = 1
                ZZBusca(ZZLugarBusca) = "c:\Archivos de programa\Microsoft Office\"
                
            Case "XLSX"
                ZZCodigoExe = "EXCEL.exe"
                ZZPasaExe = ""
                Erase ZZBusca
                ZZLugarBusca = 1
                ZZBusca(ZZLugarBusca) = "c:\Archivos de programa\Microsoft Office\"
                
            Case "JPG", "BMP"
                ZZCodigoExe = "OIS.exe"
                ZZPasaExe = ""
                Erase ZZBusca
                ZZLugarBusca = 1
                ZZBusca(ZZLugarBusca) = "c:\Archivos de programa\Microsoft Office\"
                
        End Select
        
        CicloBusca = 1
        ZZSalida = "N"
        
        Do
        
            MiRuta = ZZBusca(CicloBusca)
            MiNombre = Dir(MiRuta, vbDirectory) ' Recupera la primera entrada.
            Do While MiNombre <> "" ' Inicia el bucle.
                    
                If MiNombre <> "." And MiNombre <> ".." Then
            
                    If (GetAttr(MiRuta & MiNombre) And vbDirectory) = vbDirectory Then
                        
                        ZZLugarBusca = ZZLugarBusca + 1
                        ZZBusca(ZZLugarBusca) = MiRuta & MiNombre + "\"
                        
                            Else
                            
                        WEspacios = Len(ZZCodigoExe)
                        DA = Len(MiNombre) - WEspacios
                        If UCase(Trim(ZZCodigoExe)) = UCase(Trim(MiNombre)) Then
                            ZZPasaExe = MiRuta & MiNombre
                            ZZSalida = "S"
                            Exit Do
                        End If
                        
                    End If
                
                End If
                MiNombre = Trim(UCase(Dir))  ' Obtiene siguiente entrada.
                
            Loop
    
            If CicloBusca = ZZLugarBusca Or ZZSalida = "S" Then
                Exit Do
                    Else
                CicloBusca = CicloBusca + 1
            End If
    
        Loop
                     
        ZZRuta = Archivo.Text
        ZZEstado = Dir(ZZRuta)
        If ZZEstado <> "" Then
            RetVal = Shell(ZZPasaExe + " " + ZZRuta + " ", 3)
        End If

    End If
    
    Exit Sub
    
WError:
    Rem MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
    
End Sub

Private Sub FinIngresaObserva_Click()
    PrgOrdenArchivos.Hide
End Sub

Private Sub Form_Activate()

    On Error GoTo WError
    
    Dir1.Path = "W:\orden"
    File1.Path = Dir1.Path
    Dir1.Path = "W:\orden\" + Trim(Str$(Val(WPasaOrden)))
    File1.Path = Dir1.Path
    
    Exit Sub
    
WError:
    Rem MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
    

End Sub

Private Sub Form_Load()

    On Error GoTo WError
    
    Drive1.Drive = "W"
    Auxi = "W:\orden\" + Trim(Str$(Val(WPasaOrden)))
    
    ZZEstado = Dir(Auxi)
    If ZZEstado = "" Then
        MkDir Auxi
    End If
    
    Dir1.Path = "W:\orden\" + Trim(Str$(Val(WPasaOrden)))
    
    Exit Sub
    
WError:
    Rem MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
    
End Sub
