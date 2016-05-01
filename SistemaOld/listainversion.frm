VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaInversion 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Proyectos de Inversion"
   ClientHeight    =   2385
   ClientLeft      =   2010
   ClientTop       =   735
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   2385
   ScaleWidth      =   8085
   Begin VB.ListBox WIndice 
      Height          =   255
      ItemData        =   "listainversion.frx":0000
      Left            =   7080
      List            =   "listainversion.frx":0002
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox Tipo 
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
         Left            =   2040
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Ano 
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
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Left            =   3000
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
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
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   1080
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
         Left            =   4440
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo Listado"
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
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo"
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WListaCursoLegajo.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgListaInversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sppruefarma As String
Dim ZVector(5000, 2) As String

Private Sub Acepta_Click()


    Erase ZVector
    ZLugar = 0
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM asigna"
   Rem ZSql = ZSql + " Where asigna.producto = " + "'" + Producto + "'"
  Rem  ZSql = ZSql + " Order by prueterfarma.partida"
    spAsigna = ZSql
    Set rstAsigna = db.OpenRecordset(spAsigna, dbOpenSnapshot, dbSQLPassThrough)
    If rstAsigna.RecordCount > 0 Then
        With rstAsigna
            .MoveFirst
            Do
                If .EOF = False Then
                    ZLugar = ZLugar + 1
                 Rem   ZVector(ZLugar, 1) = rstAsigna!partida
                 Rem   ZVector(ZLugar, 2) = Str$(rstAsigna!Valor)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstAsigna.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZProyecto = ZVector(Ciclo, 1)
        ZGasto = 0
        ZImporte = Val(ZVector(Ciclo, 2))
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Avance"
        ZSql = ZSql + " Where Avance.Proyecto = " + "'" + ZProyecto + "'"
        spAvance = ZSql
        Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
        If rstAvance.RecordCount > 0 Then
            With rstAvance
                .MoveFirst
                Do
                    If .EOF = False Then
                        ZPeriodo = 0
                        If rstAvance!ordfecha >= "20070701" And rstAvance!ordfecha <= "20080631" Then
                            ZPeriodo = 1
                        End If
                        If rstAvance!ordfecha >= "20080701" And rstAvance!ordfecha <= "20090631" Then
                            ZPeriodo = 2
                        End If
                        If rstAvance!ordfecha >= "20090701" And rstAvance!ordfecha <= "20100631" Then
                            ZPeriodo = 3
                        End If
                        If rstAvance!ordfecha >= "20100701" And rstAvance!ordfecha <= "20110631" Then
                            ZPeriodo = 4
                        End If
                        If rstAvance!ordfecha >= "20110701" And rstAvance!ordfecha <= "20120631" Then
                            ZPeriodo = 5
                        End If
                        If rstAvance!ordfecha >= "20120701" And rstAvance!ordfecha <= "20130631" Then
                            ZPeriodo = 6
                        End If
                        If rstAvance!ordfecha >= "20130701" And rstAvance!ordfecha <= "20140631" Then
                            ZPeriodo = 7
                        End If
                        If rstAvance!ordfecha >= "20140701" And rstAvance!ordfecha <= "20150631" Then
                            ZPeriodo = 8
                        End If
                        If rstAvance!ordfecha >= "20150701" And rstAvance!ordfecha <= "20160631" Then
                            ZPeriodo = 9
                        End If
                        If rstAvance!ordfecha >= "20160701" And rstAvance!ordfecha <= "20070631" Then
                            ZPeriodo = 10
                        End If
                        If ZPeriodo = Val(Ano.Text) Then
                            ZGasto = ZGasto + rstAvance!Importe
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstAvance.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Proyecto SET "
        ZSql = ZSql + " Proyecto.Gasto = " + "'" + Str$(ZGasto) + "',"
        ZSql = ZSql + " Proyecto.Importe = " + "'" + Str$(ZImporte) + "',"
        ZSql = ZSql + " Proyecto.Marca = " + "'" + "X" + "'"
        ZSql = ZSql + " Where Proyecto.Codigo = " + "'" + ZProyecto + "'"
        spProyecto = ZSql
        Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo


    Listado.WindowTitle = "Listado de Proyectos de Inversion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Select Case tipo.ListIndex
         Case 1
            Uno = "{Proyecto.Planta} in 1 to 1"
            Dos = " and {Proyecto.Marca} = " + Chr$(34) + "X" + Chr$(34)
            
            Listado.SQLQuery = "SELECT Proyecto.Codigo, Proyecto.Sector, Proyecto.Descripcion, Proyecto.Centro, Proyecto.Presupuesto, Proyecto.Estado, Proyecto.Prioridad, Proyecto.FechaInicio, Proyecto.FechaFinal, Proyecto.Gasto, Proyecto.Ano, Proyecto.Planta, Proyecto.FechaAprobado, Proyecto.Solicitante, Proyecto.Marca, Proyecto.Importe, " _
                    + "SectorInve.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Proyecto Proyecto, " _
                    + DSQ + ".dbo.SectorInve SectorInve " _
                    + "Where " _
                    + "Proyecto.Sector = SectorInve.Codigo AND " _
                    + "Proyecto.Planta >= 1 AND " _
                    + "Proyecto.Planta <= 1 AND " _
                    + "Proyecto.Marca = '" + "X" + "'"
                    
            Listado.ReportFileName = "ListaInversion.rpt"
            
         Case 2
            Uno = "{Proyecto.Planta} in 1 to 1"
            Dos = " and {Proyecto.Marca} = " + Chr$(34) + "X" + Chr$(34)
            
            Listado.SQLQuery = "SELECT Proyecto.Codigo, Proyecto.Sector, Proyecto.Descripcion, Proyecto.Centro, Proyecto.Presupuesto, Proyecto.Estado, Proyecto.Prioridad, Proyecto.FechaInicio, Proyecto.FechaFinal, Proyecto.Gasto, Proyecto.Ano, Proyecto.Planta, Proyecto.FechaAprobado, Proyecto.Solicitante, Proyecto.Marca, Proyecto.Importe, " _
                    + "SectorInve.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Proyecto Proyecto, " _
                    + DSQ + ".dbo.SectorInve SectorInve " _
                    + "Where " _
                    + "Proyecto.Sector = SectorInve.Codigo AND " _
                    + "Proyecto.Planta >= 2 AND " _
                    + "Proyecto.Planta <= 2 AND " _
                    + "Proyecto.Marca = '" + "X" + "'"
                    
            Listado.ReportFileName = "ListaInversion.rpt"
            
         Case 3
            Uno = "{Proyecto.Planta} in 1 to 1"
            Dos = " and {Proyecto.Marca} = " + Chr$(34) + "X" + Chr$(34)
            
            Listado.SQLQuery = "SELECT Proyecto.Codigo, Proyecto.Sector, Proyecto.Descripcion, Proyecto.Centro, Proyecto.Presupuesto, Proyecto.Estado, Proyecto.Prioridad, Proyecto.FechaInicio, Proyecto.FechaFinal, Proyecto.Gasto, Proyecto.Ano, Proyecto.Planta, Proyecto.FechaAprobado, Proyecto.Solicitante, Proyecto.Marca, Proyecto.Importe, " _
                    + "SectorInve.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Proyecto Proyecto, " _
                    + DSQ + ".dbo.SectorInve SectorInve " _
                    + "Where " _
                    + "Proyecto.Sector = SectorInve.Codigo AND " _
                    + "Proyecto.Planta >= 3 AND " _
                    + "Proyecto.Planta <= 3 AND " _
                    + "Proyecto.Marca = '" + "X" + "'"
                    
            Listado.ReportFileName = "ListaInversion.rpt"
            
         Case 4
            Uno = "{Proyecto.Planta} in 1 to 1"
            Dos = " and {Proyecto.Marca} = " + Chr$(34) + "X" + Chr$(34)
            
            Listado.SQLQuery = "SELECT Proyecto.Codigo, Proyecto.Sector, Proyecto.Descripcion, Proyecto.Centro, Proyecto.Presupuesto, Proyecto.Estado, Proyecto.Prioridad, Proyecto.FechaInicio, Proyecto.FechaFinal, Proyecto.Gasto, Proyecto.Ano, Proyecto.Planta, Proyecto.FechaAprobado, Proyecto.Solicitante, Proyecto.Marca, Proyecto.Importe, " _
                    + "SectorInve.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Proyecto Proyecto, " _
                    + DSQ + ".dbo.SectorInve SectorInve " _
                    + "Where " _
                    + "Proyecto.Sector = SectorInve.Codigo AND " _
                    + "Proyecto.Planta >= 5 AND " _
                    + "Proyecto.Planta <= 5 AND " _
                    + "Proyecto.Marca = '" + "X" + "'"
                    
            Listado.ReportFileName = "ListaInversion.rpt"
            
         Case Else
            Uno = ""
            Dos = "{Proyecto.Marca} = " + Chr$(34) + "X" + Chr$(34)
            
            Listado.SQLQuery = "SELECT Proyecto.Codigo, Proyecto.SectorProyecto.Descripcion, Proyecto.Centro, Proyecto.Presupuesto, Proyecto.Estado, Proyecto.Prioridad, Proyecto.FechaInicio, Proyecto.FechaFinal, Proyecto.Gasto, Proyecto.Ano, Proyecto.Planta, Proyecto.FechaAprobado, Proyecto.Solicitante, Proyecto.Marca, Proyecto.Importe, " _
                    + "SectorInve.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Proyecto Proyecto, " _
                    + DSQ + ".dbo.SectorInve SectorInve " _
                    + "Where " _
                    + "Proyecto.Sector = SectorInve.Codigo AND " _
                    + "Proyecto.Marca = '" + "X" + "'"
                    
            Listado.ReportFileName = "ListaInversionConsol.rpt"
            
    End Select
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Connect = Connect()
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgListaInversion.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    tipo.Clear
    
    tipo.AddItem "Unificado"
    tipo.AddItem "Planta I"
    tipo.AddItem "Planta II"
    tipo.AddItem "Planta III"
    tipo.AddItem ""
    tipo.AddItem "Planta V"
    
    tipo.ListIndex = 0
    Panta.Value = True
    Impresora.Value = False
End Sub

Private Sub Ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

