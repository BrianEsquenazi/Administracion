VERSION 5.00
Begin VB.Form proyectocostos 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
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
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
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
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "proyectocostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZVector(5000) As String

Private Sub Command1_Click()


Rem   ZSql = ""
Rem    ZSql = ZSql + "UPDATE Proyecto SET "
Rem    ZSql = ZSql + " Proyecto.Gasto = 0"
Rem    spProyecto = ZSql
Rem    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    
    
  Rem  Erase ZVector
  Rem  ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proyecto"
    ZSql = ZSql + " Where Proyecto.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " Order by Proyecto.Codigo"
    spProyecto = ZSql
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        With rstProyecto
            .MoveFirst
            Do
                If .EOF = False Then
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar) = rstProyecto!Codigo
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProyecto.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZProyecto = ZVector(Ciclo)
        ZGasto = 0
        
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
                        
                        
                        Tipo = rstAvance!Tipo
                           If Tipo = 1 Then
                              
                               ZGasto2 = ZGasto2 + rstAvance!Importe
                                 
                                  Else
                                       If Tipo = 2 Then
                                             
                                             ZGasto3 = ZGasto3 + rstAvance!Importe
                                             
                                             Else
                                              If Tipo = 3 Then
                                                ZGasto3 = ZGasto3 + rstAvance!Importe
                                          
                                          End If
                                      End If
                           End If
                        
                          .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstAvance.Close
        End If
        
    Rem    ZSql = ""
    Rem    ZSql = ZSql + "UPDATE Proyectocosto SET "
    Rem    ZSql = ZSql + " Proyecto.costo1 = " + "'" + Str$(ZGast1) + "'"
    Rem    ZS    Rem ql = ZSql + " Proyecto.costo1 = " + "'" + Str$(ZGast1)
    Rem Zql = ZSql + " Where Proyecto.Codigo = " + "'" + ZProyecto + "'"
    Rem    spProyecto = ZSql
    Rem    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo


    Listado.WindowTitle = "Listado de Proyectos de Inversion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Select Case Tipo.ListIndex
         Case 1
            Uno = "{Proyecto.Ano} in " + Ano.Text + " to " + Ano.Text
            Dos = " and {Proyecto.Planta} in 1 to 1"
            
            Listado.SQLQuery = "SELECT Proyecto.Codigo, Proyecto.Sector, Proyecto.Descripcion, Proyecto.Centro, Proyecto.Presupuesto, Proyecto.Estado, Proyecto.Prioridad, Proyecto.FechaInicio, Proyecto.FechaFinal, Proyecto.Gasto, Proyecto.Ano, Proyecto.Planta, Proyecto.FechaAprobado, Proyecto.Solicitante, " _
                    + "SectorInve.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Proyecto Proyecto, " _
                    + DSQ + ".dbo.SectorInve SectorInve " _
                    + "Where " _
                    + "Proyecto.Sector = SectorInve.Codigo AND " _
                    + "Proyecto.Ano >= " + Ano.Text + " AND " _
                    + "Proyecto.Ano <= " + Ano.Text + " AND " _
                    + "Proyecto.Planta >= 1 AND " _
                    + "Proyecto.Planta <= 1"
                    
            Listado.ReportFileName = "ListaInversion.rpt"
            
         Case 2
            Uno = "{Proyecto.Ano} in " + Ano.Text + " to " + Ano.Text
            Dos = " and {Proyecto.Planta} in 2 to 2"
            
            Listado.SQLQuery = "SELECT Proyecto.Codigo, Proyecto.Sector, Proyecto.Descripcion, Proyecto.Centro, Proyecto.Presupuesto, Proyecto.Estado, Proyecto.Prioridad, Proyecto.FechaInicio, Proyecto.FechaFinal, Proyecto.Gasto, Proyecto.Ano, Proyecto.Planta, Proyecto.FechaAprobado, Proyecto.Solicitante, " _
                    + "SectorInve.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Proyecto Proyecto, " _
                    + DSQ + ".dbo.SectorInve SectorInve " _
                    + "Where " _
                    + "Proyecto.Sector = SectorInve.Codigo AND " _
                    + "Proyecto.Ano >= " + Ano.Text + " AND " _
                    + "Proyecto.Ano <= " + Ano.Text + " AND " _
                    + "Proyecto.Planta >= 2 AND " _
                    + "Proyecto.Planta <= 2"
                    
            Listado.ReportFileName = "ListaInversion.rpt"
            
         Case 3
            Uno = "{Proyecto.Ano} in " + Ano.Text + " to " + Ano.Text
            Dos = " and {Proyecto.Planta} in 3 to 3"
            
            Listado.SQLQuery = "SELECT Proyecto.Codigo, Proyecto.Sector, Proyecto.Descripcion, Proyecto.Centro, Proyecto.Presupuesto, Proyecto.Estado, Proyecto.Prioridad, Proyecto.FechaInicio, Proyecto.FechaFinal, Proyecto.Gasto, Proyecto.Ano, Proyecto.Planta, Proyecto.FechaAprobado, Proyecto.Solicitante, " _
                    + "SectorInve.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Proyecto Proyecto, " _
                    + DSQ + ".dbo.SectorInve SectorInve " _
                    + "Where " _
                    + "Proyecto.Sector = SectorInve.Codigo AND " _
                    + "Proyecto.Ano >= " + Ano.Text + " AND " _
                    + "Proyecto.Ano <= " + Ano.Text + " AND " _
                    + "Proyecto.Planta >= 3 AND " _
                    + "Proyecto.Planta <= 3"
                    
            Listado.ReportFileName = "ListaInversion.rpt"
            
         Case 4
            Uno = "{Proyecto.Ano} in " + Ano.Text + " to " + Ano.Text
            Dos = " and {Proyecto.Planta} in 5 to 5"
            
            Listado.SQLQuery = "SELECT Proyecto.Codigo, Proyecto.Sector, Proyecto.Descripcion, Proyecto.Centro, Proyecto.Presupuesto, Proyecto.Estado, Proyecto.Prioridad, Proyecto.FechaInicio, Proyecto.FechaFinal, Proyecto.Gasto, Proyecto.Ano, Proyecto.Planta, Proyecto.FechaAprobado, Proyecto.Solicitante, " _
                    + "SectorInve.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Proyecto Proyecto, " _
                    + DSQ + ".dbo.SectorInve SectorInve " _
                    + "Where " _
                    + "Proyecto.Sector = SectorInve.Codigo AND " _
                    + "Proyecto.Ano >= " + Ano.Text + " AND " _
                    + "Proyecto.Ano <= " + Ano.Text + " AND " _
                    + "Proyecto.Planta >= 5 AND " _
                    + "Proyecto.Planta <= 5"
                    
            Listado.ReportFileName = "ListaInversion.rpt"
            
         Case Else
            Uno = "{Proyecto.Ano} in " + Ano.Text + " to " + Ano.Text
            Dos = ""
            
            Listado.SQLQuery = "SELECT Proyecto.Codigo, Proyecto.SectorProyecto.Descripcion, Proyecto.Centro, Proyecto.Presupuesto, Proyecto.Estado, Proyecto.Prioridad, Proyecto.FechaInicio, Proyecto.FechaFinal, Proyecto.Gasto, Proyecto.Ano, Proyecto.Planta, Proyecto.FechaAprobado, Proyecto.Solicitante,  " _
                    + "SectorInve.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Proyecto Proyecto, " _
                    + DSQ + ".dbo.SectorInve SectorInve " _
                    + "Where " _
                    + "Proyecto.Sector = SectorInve.Codigo AND " _
                    + "Proyecto.Ano >= " + Ano.Text + " AND " _
                    + "Proyecto.Ano <= " + Ano.Text
                    
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
   Rem Tipo.Clear
    
   Rem Tipo.AddItem "Unificado"
   Rem Tipo.AddItem "Planta I"
  Rem  Tipo.AddItem "Planta II"
  Rem  Tipo.AddItem "Planta III"
  Rem  Tipo.AddItem ""
  Rem  Tipo.AddItem "Planta V"
    
  Rem  Tipo.ListIndex = 0
  Rem  Panta.Value = True
  Rem  Impresora.Value = False
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
