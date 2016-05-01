VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ok 
   AutoRedraw      =   -1  'True
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8205
   ScaleWidth      =   11880
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
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   7080
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   6120
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
      Index           =   2
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   6000
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   6855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "ok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String
Dim XParam As String
Dim ZZLugar As Integer

Dim ZZTipo(1000, 2) As String
Dim ZZAyudaI(1000) As String
Dim ZZAyudaII(1000) As String
Dim ZZAyudaIII(1000) As String
Dim ZZAyudaIV(1000) As String
Dim ZZAyudaV(1000) As String
Dim ZZAyudaVI(1000) As String


Private Sub Cancela_click()
    PrgAgendaPlanifica.Hide
    Unload Me
    PrgAgendaTotal.Show
End Sub

Sub Form_Load()

    Call Limpia_Ayuda
    
    Erase ZZAyudaI
    Erase ZZAyudaII
    Erase ZZAyudaIII
    Erase ZZAyudaIV
    Erase ZZAyudaV
    Erase ZZAyudaVI
    
    ZZAyudaI(1) = "Iniciada"
    ZZAyudaI(2) = "Investig."
    ZZAyudaI(3) = "Implemen."
    ZZAyudaI(4) = "Impl.a Veri"
    ZZAyudaI(5) = "Impl.Verifi"
    ZZAyudaI(6) = "Cerrada"
    
    ZZAyudaII(1) = "Auditoria"
    ZZAyudaII(2) = "Reclamo"
    ZZAyudaII(3) = "I.No Conf"
    ZZAyudaII(4) = "Proc/Sist"
    ZZAyudaII(5) = "Otro"
    
    Erase ZZTipo
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoSac"
    ZSql = ZSql + " Order by TipoSac.Codigo"
    spTipoSac = ZSql
    Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoSac.RecordCount > 0 Then
        With rstTipoSac
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Lugar = Lugar + 1
                    
                    ZZTipo(Lugar, 1) = rstTipoSac!Codigo
                    ZZTipo(Lugar, 2) = rstTipoSac!Descripcion
                    
                    ZZAyudaIII(rstTipoSac!Codigo) = Trim(rstTipoSac!Descripcion)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTipoSac.Close
    End If
    
   
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CentroSac"
    ZSql = ZSql + " Order by CentroSac.Codigo"
    spCentroSac = ZSql
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        With rstCentroSac
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZAyudaIV(rstCentroSac!Codigo) = Trim(rstCentroSac!Descripcion)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCentroSac.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ResponsableSac"
    ZSql = ZSql + " Order by ResponsableSac.Codigo"
    spResponsableSac = ZSql
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        With rstResponsableSac
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZAyudaV(rstResponsableSac!Codigo) = Trim(rstResponsableSac!Descripcion)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstResponsableSac.Close
    End If
    
    


    ZLugar = 0

    For Ciclo = 1 To 100
    
        If Val(ZZPasaDatos(Ciclo, 1)) <> 0 Then
        
            ZTipo = Left$(ZZPasaDatos(Ciclo, 1), 2)
            ZAno = Mid$(ZZPasaDatos(Ciclo, 1), 3, 4)
            ZNumero = Right$(ZZPasaDatos(Ciclo, 1), 6)
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + ZTipo + "'"
            ZSql = ZSql + " and CargaSac.Ano = " + "'" + ZAno + "'"
            ZSql = ZSql + " and CargaSac.Numero = " + "'" + ZNumero + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
            
                ZLugar = ZLugar + 1
                
                Pantalla.TextMatrix(ZLugar, 1) = ZZAyudaIII(rstCargaSac!Tipo)
                Pantalla.TextMatrix(ZLugar, 2) = rstCargaSac!Ano
                Pantalla.TextMatrix(ZLugar, 3) = rstCargaSac!Numero
                Pantalla.TextMatrix(ZLugar, 4) = rstCargaSac!Fecha
                Pantalla.TextMatrix(ZLugar, 5) = ZZAyudaI(rstCargaSac!Estado)
                Pantalla.TextMatrix(ZLugar, 6) = rstCargaSac!Titulo
                Pantalla.TextMatrix(ZLugar, 7) = rstCargaSac!Referencia
                Pantalla.TextMatrix(ZLugar, 8) = ZZAyudaIV(rstCargaSac!Centro)
                Pantalla.TextMatrix(ZLugar, 9) = ZZAyudaII(rstCargaSac!Origen)
                Pantalla.TextMatrix(ZLugar, 10) = ZZAyudaV(rstCargaSac!ResponsableEmisor)
                Pantalla.TextMatrix(ZLugar, 11) = ZZAyudaV(rstCargaSac!ResponsableDestino)
                Pantalla.TextMatrix(ZLugar, 12) = rstCargaSac!Clave
            
                rstCargaSac.Close
            End If
        End If
    Next Ciclo
End Sub
Private Sub Limpia_Ayuda()

    Pantalla.Clear
    Pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    Select Case ZZLugar
        Case 1, 2, 4
            Pantalla.FixedCols = 1
            Pantalla.Cols = 3
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
            
            Pantalla.ColWidth(0) = 200
            Pantalla.Row = 0
            
            For Ciclo = 1 To Pantalla.Cols - 1
                Pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        Pantalla.Text = "Codigo"
                        Pantalla.ColWidth(Ciclo) = 1000
                        Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        Pantalla.Text = "Nombre"
                        Pantalla.ColWidth(Ciclo) = 6000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
            
            Rem DESPILEGA LOS TITULOS
            
            WTitulo(1).Visible = False
            WTitulo(2).Visible = False
            
        Case Else
            Pantalla.FixedCols = 1
            Pantalla.Cols = 13
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
            
            Pantalla.ColWidth(0) = 200
            Pantalla.Row = 0
            
            For Ciclo = 1 To Pantalla.Cols - 1
                Pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        Pantalla.Text = "Tipo"
                        Pantalla.ColWidth(Ciclo) = 1200
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 2
                        Pantalla.Text = "Año"
                        Pantalla.ColWidth(Ciclo) = 600
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 3
                        Pantalla.Text = "Nro"
                        Pantalla.ColWidth(Ciclo) = 700
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 4
                        Pantalla.Text = "Fecha"
                        Pantalla.ColWidth(Ciclo) = 1200
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 5
                        Pantalla.Text = "Estado"
                        Pantalla.ColWidth(Ciclo) = 1000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 6
                        Pantalla.Text = "Titulo"
                        Pantalla.ColWidth(Ciclo) = 3000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 7
                        Pantalla.Text = "Referencia"
                        Pantalla.ColWidth(Ciclo) = 3000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 8
                        Pantalla.Text = "Centro"
                        Pantalla.ColWidth(Ciclo) = 1400
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 9
                        Pantalla.Text = "Origen"
                        Pantalla.ColWidth(Ciclo) = 1000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 10
                        Pantalla.Text = "Emisor"
                        Pantalla.ColWidth(Ciclo) = 800
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 11
                        Pantalla.Text = "Respon."
                        Pantalla.ColWidth(Ciclo) = 800
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 12
                        Pantalla.Text = ""
                        Pantalla.ColWidth(Ciclo) = 10
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
            
            Rem DESPILEGA LOS TITULOS
            
            WTitulo(1).Visible = False
            WTitulo(2).Visible = False
            
    End Select
    
    Pantalla.Row = 0
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        Rem WTitulo(Ciclo).Text = Pantalla.Text
        Rem WTitulo(Ciclo).Left = Pantalla.CellLeft + Pantalla.Left
        Rem WTitulo(Ciclo).Top = Pantalla.CellTop + Pantalla.Top
        Rem WTitulo(Ciclo).Width = Pantalla.CellWidth
        Rem WTitulo(Ciclo).Height = Pantalla.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To Pantalla.Cols - 1
        WAncho = WAncho + Pantalla.ColWidth(Ciclo)
    Next Ciclo
    Rem Pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = Pantalla.Font.Name
    Font.Size = Pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Pantalla.AllowUserResizing = flexResizeBoth
    
    Pantalla.Col = 1
    Pantalla.Row = 1
    
End Sub


Private Sub Pantalla_Click()
    WPasaNumero = Pantalla.TextMatrix(Pantalla.Row, 12)
    PrgAgendaConsultaSacauto.Show
End Sub
