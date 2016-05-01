VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form PrgConsultaPartidaDy 
   Caption         =   "Consulta de Partidas"
   ClientHeight    =   6735
   ClientLeft      =   3195
   ClientTop       =   930
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   4515
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgConsultaPartidaDy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XLote(100, 7) As String
Dim ZZVector(10000, 2) As String
Dim ZZLugar As Integer


Private Sub Limpia_Vector()

    WVector1.Clear
    WVector1.Font.Bold = True
    
    WVector1.FixedCols = 1
    WVector1.Cols = 3
    WVector1.FixedRows = 1
    WVector1.Rows = ZZLugar + 1
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Articulo"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WVector1.Text = "Partida"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub


Private Sub Form_Activate()
    
    ZZLugar = 0
    Erase ZZVector
    
    ZZCorte = ""
    ZZCorteII = ""
    ZZPasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select Laudo.PartiOri, Laudo.Articulo"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo.PartiOri > " + "'" + "" + "'"
    ZSql = ZSql + " Order by Laudo.PartiOri"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then

        With rstLaudo

            .MoveFirst
    
            If .NoMatch = False Then
            Do
    
                If .EOF = True Then
                    Exit Do
                End If
        
                If ZZPasa = 0 Then
                    ZZPasa = 1
                    ZZCorte = rstLaudo!Articulo
                    ZZCorteII = rstLaudo!partiori
                End If
                
                If ZZCorte <> rstLaudo!Articulo Or ZZCorteII <> rstLaudo!partiori Then
                
                    ZZLugar = ZZLugar + 1
                    
                    ZZVector(ZZLugar, 1) = ZZCorte
                    ZZVector(ZZLugar, 2) = ZZCorteII
                    
                    ZZCorte = rstLaudo!Articulo
                    ZZCorteII = rstLaudo!partiori
                    
                End If
    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
        
            Loop
            End If
    
        End With
    End If

    If ZZPasa <> 0 Then
    
        ZZLugar = ZZLugar + 1
        
        ZZVector(ZZLugar, 1) = ZZCorte
        ZZVector(ZZLugar, 2) = ZZCorteII
        
    End If

    Call Limpia_Vector
    
    For Ciclo = 1 To ZZLugar
    
        WVector1.TextMatrix(Ciclo, 1) = ZZVector(Ciclo, 1)
        WVector1.TextMatrix(Ciclo, 2) = ZZVector(Ciclo, 2)
    
    Next Ciclo
    
End Sub

