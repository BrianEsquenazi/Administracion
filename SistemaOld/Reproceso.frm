VERSION 5.00
Begin VB.Form PrgReproceso 
   AutoRedraw      =   -1  'True
   Caption         =   "Reproceso de Grabacion de Cursos"
   ClientHeight    =   2925
   ClientLeft      =   2010
   ClientTop       =   735
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   2925
   ScaleWidth      =   8085
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5655
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
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   0
         Top             =   720
         Width           =   855
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
         Left            =   3000
         TabIndex        =   3
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
         Left            =   1200
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Año"
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
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "PrgReproceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstLegajo As Recordset
Dim spLegajo As String
Dim rstCurso As Recordset
Dim spCurso As String
Dim rstCursadas As Recordset
Dim spCursadas As String

Dim ZVector(10000, 5) As String

Private Sub Acepta_Click()

    ZSql = ""
    ZSql = ZSql + "UPDATE Cronograma SET "
    ZSql = ZSql + " Realizado = 0"
    ZSql = ZSql + " Where Ano = " + "'" + Ano.Text + "'"
    spCronograma = ZSql
    Set rstCronograma = db.OpenRecordset(spCronograma, dbOpenSnapshot, dbSQLPassThrough)

    WRenglon = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cursadas"
    ZSql = ZSql + " Order by Cursadas.Clave"
    rsCursadas = ZSql
    Set rstCursadas = db.OpenRecordset(rsCursadas, dbOpenSnapshot, dbSQLPassThrough)
    If rstCursadas.RecordCount > 0 Then
        With rstCursadas
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZFecha = Mid$(rstCursadas!Fecha, 7, 4)
                    
                    If Val(ZFecha) = Val(Ano.Text) Then
                
                        WRenglon = WRenglon + 1
                        
                        ZVector(WRenglon, 1) = rstCursadas!Curso
                        ZVector(WRenglon, 2) = rstCursadas!Legajo
                        ZVector(WRenglon, 3) = Str$(rstCursadas!Horas)
                        ZVector(WRenglon, 4) = rstCursadas!Fecha
                        ZVector(WRenglon, 5) = rstCursadas!Clave
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCursadas.Close
    End If
    
    For Ciclo = 1 To WRenglon
    
        ZCurso = ZVector(Ciclo, 1)
        ZLegajo = ZVector(Ciclo, 2)
        ZHoras = ZVector(Ciclo, 3)
        ZFecha = ZVector(Ciclo, 4)
        ZClave = ZVector(Ciclo, 5)
        
        ZAno = Mid$(ZFecha, 7, 4)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cronograma"
        ZSql = ZSql + " Where Ano = " + "'" + ZAno + "'"
        ZSql = ZSql + " and Legajo = " + "'" + ZLegajo + "'"
        ZSql = ZSql + " and Curso = " + "'" + ZCurso + "'"
        rsCursadas = ZSql
        Set rstCursadas = db.OpenRecordset(rsCursadas, dbOpenSnapshot, dbSQLPassThrough)
        If rstCursadas.RecordCount > 0 Then
            rstCursadas.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Cronograma SET "
            ZSql = ZSql + " Realizado = Realizado + " + "'" + ZHoras + "'"
            ZSql = ZSql + " Where Ano = " + "'" + ZAno + "'"
            ZSql = ZSql + " and Legajo = " + "'" + ZLegajo + "'"
            ZSql = ZSql + " and Curso = " + "'" + ZCurso + "'"
            spCronograma = ZSql
            Set rstCronograma = db.OpenRecordset(spCronograma, dbOpenSnapshot, dbSQLPassThrough)
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Cursadas SET "
            ZSql = ZSql + " TipoCursada = 0"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spCursadas = ZSql
            Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Cursadas SET "
            ZSql = ZSql + " TipoCursada = 1"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spCursadas = ZSql
            Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
        
    m$ = "El proceso a finalizado con exito"
    A% = MsgBox(m$, 0, "Reproceso de Cursos Realizados")
        
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgReproceso.Hide
    Unload Me
    Menu.Show
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

