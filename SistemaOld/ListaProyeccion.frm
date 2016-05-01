VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaProyeccion 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado  de  Proyeccion de Fabricacion"
   ClientHeight    =   4080
   ClientLeft      =   2175
   ClientTop       =   945
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4080
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   5895
      Begin MSMask.MaskEdBox Fecha 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1320
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
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   2640
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
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   2640
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
         Left            =   4200
         TabIndex        =   5
         Top             =   1440
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
         Height          =   495
         Left            =   4200
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
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
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
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
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListaProyeccion.rpt"
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
Attribute VB_Name = "PrgListaProyeccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCargaProyeccion As Recordset
Dim spCargaProyeccion As String
Dim rstProyeccionFabrica As Recordset
Dim spProyeccionFabrica As String

Dim XParam As String

Dim WFecha As String
Dim WVencimiento As String
Dim WDias1 As Integer
Dim Impre(100, 2) As String
Dim Dia(100) As String
Dim ZVector(1000, 100) As String

Private Sub Acepta_Click()

    Rem On Error GoTo WError
    Rem DateDiff("d", Now, LaFecha)
    
    ZSql = "DELETE ProyeccionFabrica"
    spProyeccionFabrica = ZSql
    Set rstProyeccionFabrica = db.OpenRecordset(spProyeccionFabrica, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase ZVector
    LugarVector = 0
    
    ZClave = 0
        
    WDesdeFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    ZFecha = Fecha.Text
    
    Impre(1, 1) = Left$(ZFecha, 2)
    Impre(1, 2) = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
    
    For Ciclo = 1 To 30
    
        WFecha = ZFecha
        WDias1 = 2
        
        Call Calcula_vencimiento(WFecha, WDias1, WVencimiento)
        
        Impre(Ciclo + 1, 1) = Left$(WVencimiento, 2)
        Impre(Ciclo + 1, 2) = Right$(WVencimiento, 4) + Mid$(WVencimiento, 4, 2) + Left$(WVencimiento, 2)
        
        ZFecha = WVencimiento
        WHastaFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        
    Next Ciclo
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaProyeccion"
    ZSql = ZSql + " Where CargaProyeccion.Articulo >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and CargaProyeccion.Articulo <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " and CargaProyeccion.Saldo > 0"
    spCargaProyeccion = ZSql
    Set rstCargaProyeccion = db.OpenRecordset(spCargaProyeccion, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaProyeccion.RecordCount > 0 Then
        
        With rstCargaProyeccion
    
            .MoveFirst
            
            Do
            
                ZClave = ZClave + 1
                
                zFechaI = rstCargaProyeccion!FechaI
                zFechaII = rstCargaProyeccion!FechaII
                ZOrdFechaI = rstCargaProyeccion!OrdFechaI
                ZOrdFechaII = rstCargaProyeccion!OrdFechaII
                ZDetalle = Left$(Trim(rstCargaProyeccion!Detalle), 50)
                ZSaldo = Str$(rstCargaProyeccion!Saldo)
                
                ZArticulo = rstCargaProyeccion!Articulo
                ZCorte = "1"
                
                For Ciclo = 1 To 31
                    If Impre(Ciclo, 2) >= ZOrdFechaI And Impre(Ciclo, 2) <= ZOrdFechaII Then
                        Dia(Ciclo) = "X"
                            Else
                        Dia(Ciclo) = ""
                    End If
                Next Ciclo
                
                ZDia1 = Dia(1)
                ZImpre1 = Impre(1, 1)
                ZDia2 = Dia(2)
                ZImpre2 = Impre(2, 1)
                ZDia3 = Dia(3)
                ZImpre3 = Impre(3, 1)
                ZDia4 = Dia(4)
                ZImpre4 = Impre(4, 1)
                ZDia5 = Dia(5)
                ZImpre5 = Impre(5, 1)
                ZDia6 = Dia(6)
                ZImpre6 = Impre(6, 1)
                ZDia7 = Dia(7)
                ZImpre7 = Impre(7, 1)
                ZDia8 = Dia(8)
                ZImpre8 = Impre(8, 1)
                ZDia9 = Dia(9)
                ZImpre9 = Impre(9, 1)
                ZDia10 = Dia(10)
                ZImpre10 = Impre(10, 1)
                ZDia11 = Dia(11)
                ZImpre11 = Impre(11, 1)
                ZDia12 = Dia(12)
                ZImpre12 = Impre(12, 1)
                ZDia13 = Dia(13)
                ZImpre13 = Impre(13, 1)
                ZDia14 = Dia(14)
                ZImpre14 = Impre(14, 1)
                ZDia15 = Dia(15)
                ZImpre15 = Impre(15, 1)
                ZDia16 = Dia(16)
                ZImpre16 = Impre(16, 1)
                ZDia17 = Dia(17)
                ZImpre17 = Impre(17, 1)
                ZDia18 = Dia(18)
                ZImpre18 = Impre(18, 1)
                ZDia19 = Dia(19)
                ZImpre19 = Impre(19, 1)
                ZDia20 = Dia(20)
                ZImpre20 = Impre(20, 1)
                ZDia21 = Dia(21)
                ZImpre21 = Impre(21, 1)
                ZDia22 = Dia(22)
                ZImpre22 = Impre(22, 1)
                ZDia23 = Dia(23)
                ZImpre23 = Impre(23, 1)
                ZDia24 = Dia(24)
                ZImpre24 = Impre(24, 1)
                ZDia25 = Dia(25)
                ZImpre25 = Impre(25, 1)
                ZDia26 = Dia(26)
                ZImpre26 = Impre(26, 1)
                ZDia27 = Dia(27)
                ZImpre27 = Impre(27, 1)
                ZDia28 = Dia(28)
                ZImpre28 = Impre(28, 1)
                ZDia29 = Dia(29)
                ZImpre29 = Impre(29, 1)
                ZDia30 = Dia(30)
                ZImpre30 = Impre(30, 1)
                ZDia31 = Dia(31)
                ZImpre31 = Impre(31, 1)
                
                ZVector(ZClave, 1) = Str$(ZClave)
                ZVector(ZClave, 2) = zFechaI
                ZVector(ZClave, 3) = zFechaII
                ZVector(ZClave, 4) = ZOrdFechaI
                ZVector(ZClave, 5) = ZOrdFechaII
                ZVector(ZClave, 6) = ZArticulo
                ZVector(ZClave, 7) = ZDetalle
                ZVector(ZClave, 8) = ZCorte
                ZVector(ZClave, 9) = ZDia1
                ZVector(ZClave, 10) = ZImpre1
                ZVector(ZClave, 11) = ZDia2
                ZVector(ZClave, 12) = ZImpre2
                ZVector(ZClave, 13) = ZDia3
                ZVector(ZClave, 14) = ZImpre3
                ZVector(ZClave, 15) = ZDia4
                ZVector(ZClave, 16) = ZImpre4
                ZVector(ZClave, 17) = ZDia5
                ZVector(ZClave, 18) = ZImpre5
                ZVector(ZClave, 19) = ZDia6
                ZVector(ZClave, 20) = ZImpre6
                ZVector(ZClave, 21) = ZDia7
                ZVector(ZClave, 22) = ZImpre7
                ZVector(ZClave, 23) = ZDia8
                ZVector(ZClave, 24) = ZImpre8
                ZVector(ZClave, 25) = ZDia9
                ZVector(ZClave, 26) = ZImpre9
                ZVector(ZClave, 27) = ZDia10
                ZVector(ZClave, 28) = ZImpre10
                ZVector(ZClave, 29) = ZDia11
                ZVector(ZClave, 30) = ZImpre11
                ZVector(ZClave, 31) = ZDia12
                ZVector(ZClave, 32) = ZImpre12
                ZVector(ZClave, 33) = ZDia13
                ZVector(ZClave, 34) = ZImpre13
                ZVector(ZClave, 35) = ZDia14
                ZVector(ZClave, 36) = ZImpre14
                ZVector(ZClave, 37) = ZDia15
                ZVector(ZClave, 38) = ZImpre15
                ZVector(ZClave, 39) = ZDia16
                ZVector(ZClave, 40) = ZImpre16
                ZVector(ZClave, 41) = ZDia17
                ZVector(ZClave, 42) = ZImpre17
                ZVector(ZClave, 43) = ZDia18
                ZVector(ZClave, 44) = ZImpre18
                ZVector(ZClave, 45) = ZDia19
                ZVector(ZClave, 46) = ZImpre19
                ZVector(ZClave, 47) = ZDia20
                ZVector(ZClave, 48) = ZImpre20
                ZVector(ZClave, 49) = ZDia21
                ZVector(ZClave, 50) = ZImpre21
                ZVector(ZClave, 51) = ZDia22
                ZVector(ZClave, 52) = ZImpre22
                ZVector(ZClave, 53) = ZDia23
                ZVector(ZClave, 54) = ZImpre23
                ZVector(ZClave, 55) = ZDia24
                ZVector(ZClave, 56) = ZImpre24
                ZVector(ZClave, 57) = ZDia25
                ZVector(ZClave, 58) = ZImpre25
                ZVector(ZClave, 59) = ZDia26
                ZVector(ZClave, 60) = ZImpre26
                ZVector(ZClave, 61) = ZDia27
                ZVector(ZClave, 62) = ZImpre27
                ZVector(ZClave, 63) = ZDia28
                ZVector(ZClave, 64) = ZImpre28
                ZVector(ZClave, 65) = ZDia29
                ZVector(ZClave, 66) = ZImpre29
                ZVector(ZClave, 67) = ZDia30
                ZVector(ZClave, 68) = ZImpre30
                ZVector(ZClave, 69) = ZDia31
                ZVector(ZClave, 70) = ZImpre31
                ZVector(ZClave, 71) = ZSaldo
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
        rstCargaProyeccion.Close
    End If
    
    For Ciclo = 1 To ZClave
    
        ZClave = Val(ZVector(Ciclo, 1))
        zFechaI = ZVector(Ciclo, 2)
        zFechaII = ZVector(Ciclo, 3)
        ZOrdFechaI = ZVector(Ciclo, 4)
        ZOrdFechaII = ZVector(Ciclo, 5)
        ZArticulo = ZVector(Ciclo, 6)
        ZDetalle = ZVector(Ciclo, 7)
        ZCorte = ZVector(Ciclo, 8)
        ZDia1 = ZVector(Ciclo, 9)
        ZImpre1 = ZVector(Ciclo, 10)
        ZDia2 = ZVector(Ciclo, 11)
        ZImpre2 = ZVector(Ciclo, 12)
        ZDia3 = ZVector(Ciclo, 13)
        ZImpre3 = ZVector(Ciclo, 14)
        ZDia4 = ZVector(Ciclo, 15)
        ZImpre4 = ZVector(Ciclo, 16)
        ZDia5 = ZVector(Ciclo, 17)
        ZImpre5 = ZVector(Ciclo, 18)
        ZDia6 = ZVector(Ciclo, 19)
        ZImpre6 = ZVector(Ciclo, 20)
        ZDia7 = ZVector(Ciclo, 21)
        ZImpre7 = ZVector(Ciclo, 22)
        ZDia8 = ZVector(Ciclo, 23)
        ZImpre8 = ZVector(Ciclo, 24)
        ZDia9 = ZVector(Ciclo, 25)
        ZImpre9 = ZVector(Ciclo, 26)
        ZDia10 = ZVector(Ciclo, 27)
        ZImpre10 = ZVector(Ciclo, 28)
        ZDia11 = ZVector(Ciclo, 29)
        ZImpre11 = ZVector(Ciclo, 30)
        ZDia12 = ZVector(Ciclo, 31)
        ZImpre12 = ZVector(Ciclo, 32)
        ZDia13 = ZVector(Ciclo, 33)
        ZImpre13 = ZVector(Ciclo, 34)
        ZDia14 = ZVector(Ciclo, 35)
        ZImpre14 = ZVector(Ciclo, 36)
        ZDia15 = ZVector(Ciclo, 37)
        ZImpre15 = ZVector(Ciclo, 38)
        ZDia16 = ZVector(Ciclo, 39)
        ZImpre16 = ZVector(Ciclo, 40)
        ZDia17 = ZVector(Ciclo, 41)
        ZImpre17 = ZVector(Ciclo, 42)
        ZDia18 = ZVector(Ciclo, 43)
        ZImpre18 = ZVector(Ciclo, 44)
        ZDia19 = ZVector(Ciclo, 45)
        ZImpre19 = ZVector(Ciclo, 46)
        ZDia20 = ZVector(Ciclo, 47)
        ZImpre20 = ZVector(Ciclo, 48)
        ZDia21 = ZVector(Ciclo, 49)
        ZImpre21 = ZVector(Ciclo, 50)
        ZDia22 = ZVector(Ciclo, 51)
        ZImpre22 = ZVector(Ciclo, 52)
        ZDia23 = ZVector(Ciclo, 53)
        ZImpre23 = ZVector(Ciclo, 54)
        ZDia24 = ZVector(Ciclo, 55)
        ZImpre24 = ZVector(Ciclo, 56)
        ZDia25 = ZVector(Ciclo, 57)
        ZImpre25 = ZVector(Ciclo, 58)
        ZDia26 = ZVector(Ciclo, 59)
        ZImpre26 = ZVector(Ciclo, 60)
        ZDia27 = ZVector(Ciclo, 61)
        ZImpre27 = ZVector(Ciclo, 62)
        ZDia28 = ZVector(Ciclo, 63)
        ZImpre28 = ZVector(Ciclo, 64)
        ZDia29 = ZVector(Ciclo, 65)
        ZImpre29 = ZVector(Ciclo, 66)
        ZDia30 = ZVector(Ciclo, 67)
        ZImpre30 = ZVector(Ciclo, 68)
        ZDia31 = ZVector(Ciclo, 69)
        ZImpre31 = ZVector(Ciclo, 70)
        ZSaldo = ZVector(Ciclo, 71)
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ProyeccionFabrica ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "FechaI ,"
        ZSql = ZSql + "FechaII ,"
        ZSql = ZSql + "OrdFechaI ,"
        ZSql = ZSql + "OrdFechaII ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Detalle ,"
        ZSql = ZSql + "Corte ,"
        ZSql = ZSql + "Dia1 ,"
        ZSql = ZSql + "Impre1 ,"
        ZSql = ZSql + "Dia2 ,"
        ZSql = ZSql + "Impre2 ,"
        ZSql = ZSql + "Dia3 ,"
        ZSql = ZSql + "Impre3 ,"
        ZSql = ZSql + "Dia4 ,"
        ZSql = ZSql + "Impre4 ,"
        ZSql = ZSql + "Dia5 ,"
        ZSql = ZSql + "Impre5 ,"
        ZSql = ZSql + "Dia6 ,"
        ZSql = ZSql + "Impre6 ,"
        ZSql = ZSql + "Dia7 ,"
        ZSql = ZSql + "Impre7 ,"
        ZSql = ZSql + "Dia8 ,"
        ZSql = ZSql + "Impre8 ,"
        ZSql = ZSql + "Dia9 ,"
        ZSql = ZSql + "Impre9 ,"
        ZSql = ZSql + "Dia10 ,"
        ZSql = ZSql + "Impre10 ,"
        ZSql = ZSql + "Dia11 ,"
        ZSql = ZSql + "Impre11 ,"
        ZSql = ZSql + "Dia12 ,"
        ZSql = ZSql + "Impre12 ,"
        ZSql = ZSql + "Dia13 ,"
        ZSql = ZSql + "Impre13 ,"
        ZSql = ZSql + "Dia14 ,"
        ZSql = ZSql + "Impre14 ,"
        ZSql = ZSql + "Dia15 ,"
        ZSql = ZSql + "Impre15 ,"
        ZSql = ZSql + "Dia16 ,"
        ZSql = ZSql + "Impre16 ,"
        ZSql = ZSql + "Dia17 ,"
        ZSql = ZSql + "Impre17 ,"
        ZSql = ZSql + "Dia18 ,"
        ZSql = ZSql + "Impre18 ,"
        ZSql = ZSql + "Dia19 ,"
        ZSql = ZSql + "Impre19 ,"
        ZSql = ZSql + "Dia20 ,"
        ZSql = ZSql + "Impre20 ,"
        ZSql = ZSql + "Dia21 ,"
        ZSql = ZSql + "Impre21 ,"
        ZSql = ZSql + "Dia22 ,"
        ZSql = ZSql + "Impre22 ,"
        ZSql = ZSql + "Dia23 ,"
        ZSql = ZSql + "Impre23 ,"
        ZSql = ZSql + "Dia24 ,"
        ZSql = ZSql + "Impre24 ,"
        ZSql = ZSql + "Dia25 ,"
        ZSql = ZSql + "Impre25 ,"
        ZSql = ZSql + "Dia26 ,"
        ZSql = ZSql + "Impre26 ,"
        ZSql = ZSql + "Dia27 ,"
        ZSql = ZSql + "Impre27 ,"
        ZSql = ZSql + "Dia28 ,"
        ZSql = ZSql + "Impre28 ,"
        ZSql = ZSql + "Dia29 ,"
        ZSql = ZSql + "Impre29 ,"
        ZSql = ZSql + "Dia30 ,"
        ZSql = ZSql + "Impre30 ,"
        ZSql = ZSql + "Dia31 ,"
        ZSql = ZSql + "Impre31 ,"
        ZSql = ZSql + "Saldo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Str$(ZClave) + "',"
        ZSql = ZSql + "'" + zFechaI + "',"
        ZSql = ZSql + "'" + zFechaII + "',"
        ZSql = ZSql + "'" + ZOrdFechaI + "',"
        ZSql = ZSql + "'" + ZOrdFechaII + "',"
        ZSql = ZSql + "'" + ZArticulo + "',"
        ZSql = ZSql + "'" + ZDetalle + "',"
        ZSql = ZSql + "'" + ZCorte + "',"
        ZSql = ZSql + "'" + ZDia1 + "',"
        ZSql = ZSql + "'" + ZImpre1 + "',"
        ZSql = ZSql + "'" + ZDia2 + "',"
        ZSql = ZSql + "'" + ZImpre2 + "',"
        ZSql = ZSql + "'" + ZDia3 + "',"
        ZSql = ZSql + "'" + ZImpre3 + "',"
        ZSql = ZSql + "'" + ZDia4 + "',"
        ZSql = ZSql + "'" + ZImpre4 + "',"
        ZSql = ZSql + "'" + ZDia5 + "',"
        ZSql = ZSql + "'" + ZImpre5 + "',"
        ZSql = ZSql + "'" + ZDia6 + "',"
        ZSql = ZSql + "'" + ZImpre6 + "',"
        ZSql = ZSql + "'" + ZDia7 + "',"
        ZSql = ZSql + "'" + ZImpre7 + "',"
        ZSql = ZSql + "'" + ZDia8 + "',"
        ZSql = ZSql + "'" + ZImpre8 + "',"
        ZSql = ZSql + "'" + ZDia9 + "',"
        ZSql = ZSql + "'" + ZImpre9 + "',"
        ZSql = ZSql + "'" + ZDia10 + "',"
        ZSql = ZSql + "'" + ZImpre10 + "',"
        ZSql = ZSql + "'" + ZDia11 + "',"
        ZSql = ZSql + "'" + ZImpre11 + "',"
        ZSql = ZSql + "'" + ZDia12 + "',"
        ZSql = ZSql + "'" + ZImpre12 + "',"
        ZSql = ZSql + "'" + ZDia13 + "',"
        ZSql = ZSql + "'" + ZImpre13 + "',"
        ZSql = ZSql + "'" + ZDia14 + "',"
        ZSql = ZSql + "'" + ZImpre14 + "',"
        ZSql = ZSql + "'" + ZDia15 + "',"
        ZSql = ZSql + "'" + ZImpre15 + "',"
        ZSql = ZSql + "'" + ZDia16 + "',"
        ZSql = ZSql + "'" + ZImpre16 + "',"
        ZSql = ZSql + "'" + ZDia17 + "',"
        ZSql = ZSql + "'" + ZImpre17 + "',"
        ZSql = ZSql + "'" + ZDia18 + "',"
        ZSql = ZSql + "'" + ZImpre18 + "',"
        ZSql = ZSql + "'" + ZDia19 + "',"
        ZSql = ZSql + "'" + ZImpre19 + "',"
        ZSql = ZSql + "'" + ZDia20 + "',"
        ZSql = ZSql + "'" + ZImpre20 + "',"
        ZSql = ZSql + "'" + ZDia21 + "',"
        ZSql = ZSql + "'" + ZImpre21 + "',"
        ZSql = ZSql + "'" + ZDia22 + "',"
        ZSql = ZSql + "'" + ZImpre22 + "',"
        ZSql = ZSql + "'" + ZDia23 + "',"
        ZSql = ZSql + "'" + ZImpre23 + "',"
        ZSql = ZSql + "'" + ZDia24 + "',"
        ZSql = ZSql + "'" + ZImpre24 + "',"
        ZSql = ZSql + "'" + ZDia25 + "',"
        ZSql = ZSql + "'" + ZImpre25 + "',"
        ZSql = ZSql + "'" + ZDia26 + "',"
        ZSql = ZSql + "'" + ZImpre26 + "',"
        ZSql = ZSql + "'" + ZDia27 + "',"
        ZSql = ZSql + "'" + ZImpre27 + "',"
        ZSql = ZSql + "'" + ZDia28 + "',"
        ZSql = ZSql + "'" + ZImpre28 + "',"
        ZSql = ZSql + "'" + ZDia29 + "',"
        ZSql = ZSql + "'" + ZImpre29 + "',"
        ZSql = ZSql + "'" + ZDia30 + "',"
        ZSql = ZSql + "'" + ZImpre30 + "',"
        ZSql = ZSql + "'" + ZDia31 + "',"
        ZSql = ZSql + "'" + ZImpre31 + "',"
        ZSql = ZSql + "'" + ZSaldo + "')"
                
        spProyeccionFabrica = ZSql
        Set rstProyeccionFabrica = db.OpenRecordset(spProyeccionFabrica, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    WTitulo = "entre al " + Fecha.Text
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Proyeccion de Fabricacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + Whasta + Chr$(34)
    Rem Dos = " and {Estadistica.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Rem Listado.GroupSelectionFormula = Uno + Dos
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ProyeccionFabrica.Clave, ProyeccionFabrica.OrdFechaI, ProyeccionFabrica.Articulo, ProyeccionFabrica.Dia1, ProyeccionFabrica.Dia2, ProyeccionFabrica.Dia3, ProyeccionFabrica.Dia4, ProyeccionFabrica.Dia5, ProyeccionFabrica.Dia6, ProyeccionFabrica.Dia7, ProyeccionFabrica.Dia8, ProyeccionFabrica.Dia9, ProyeccionFabrica.Dia10, ProyeccionFabrica.Dia11, ProyeccionFabrica.Dia12, ProyeccionFabrica.Dia13, ProyeccionFabrica.Dia14, ProyeccionFabrica.Dia15, ProyeccionFabrica.Dia16, ProyeccionFabrica.Dia17, ProyeccionFabrica.Dia18, ProyeccionFabrica.Dia19, ProyeccionFabrica.Dia20, ProyeccionFabrica.Dia21, ProyeccionFabrica.Dia22, ProyeccionFabrica.Dia23, ProyeccionFabrica.Dia24, ProyeccionFabrica.Dia25, ProyeccionFabrica.Dia26, ProyeccionFabrica.Dia27, ProyeccionFabrica.Dia28, " _
                    + "ProyeccionFabrica.Dia29, ProyeccionFabrica.Dia30, ProyeccionFabrica.Dia31, ProyeccionFabrica.Detalle, ProyeccionFabrica.Corte, ProyeccionFabrica.Impre1, " _
                    + "ProyeccionFabrica.Impre2, ProyeccionFabrica.Impre3, ProyeccionFabrica.Impre4, ProyeccionFabrica.Impre5, " _
                    + "ProyeccionFabrica.Impre6, ProyeccionFabrica.Impre7, ProyeccionFabrica.Impre8, ProyeccionFabrica.Impre9, ProyeccionFabrica.Impre10, ProyeccionFabrica.Impre11, ProyeccionFabrica.Impre12, ProyeccionFabrica.Impre13, ProyeccionFabrica.Impre14, ProyeccionFabrica.Impre15, ProyeccionFabrica.Impre16, ProyeccionFabrica.Impre17, ProyeccionFabrica.Impre18, ProyeccionFabrica.Impre19, ProyeccionFabrica.Impre20, ProyeccionFabrica.Impre21, ProyeccionFabrica.Impre22, ProyeccionFabrica.Impre23, ProyeccionFabrica.Impre24, ProyeccionFabrica.Impre25, ProyeccionFabrica.Impre26, ProyeccionFabrica.Impre27, ProyeccionFabrica.Impre28, ProyeccionFabrica.Impre29, ProyeccionFabrica.Impre30, ProyeccionFabrica.Impre31, ProyeccionFabrica.Saldo" _
                    + "From " _
                    + DSQ + ".dbo.ProyeccionFabrica ProyeccionFabrica " _
                    + "Where " _
                    + "ProyeccionFabrica.Clave >= 0 AND " _
                    + "ProyeccionFabrica.Clave <= 999999 "
                      
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Rem Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.ReportFileName = "ProyecionFabrica.rpt"
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaProyeccion.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
            Else
        Desde.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Auxiliar
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Fecha.SetFocus
            Else
        Hasta.SetFocus
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Sub Form_Load()
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    Fecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub







