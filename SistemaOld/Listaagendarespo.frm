VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaAgendaRespo 
   Caption         =   "Listado de Agendas de Responsables"
   ClientHeight    =   3210
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3210
   ScaleWidth      =   8145
   Begin Crystal.CrystalReport Listado 
      Left            =   7800
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.ComboBox TipoI 
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
         TabIndex        =   2
         Top             =   1200
         Width           =   2295
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
         Left            =   2880
         TabIndex        =   8
         Top             =   1920
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
         Left            =   1440
         TabIndex        =   7
         Top             =   1920
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
         MaskColor       =   &H00000000&
         TabIndex        =   6
         Top             =   600
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
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Top             =   600
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
         Left            =   2040
         TabIndex        =   0
         Top             =   240
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
         Left            =   360
         TabIndex        =   11
         Top             =   600
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
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Emisor"
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
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7560
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaAgendaRespo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String
Dim rstAgendaRespo As Recordset
Dim spAgendaRespo As String
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String

Dim XParam As String
Dim ZZLugar As Integer
Dim ZTipo As String
Dim ZAno As String
Dim ZNumero As String
Dim ZFecha As String
Dim ZVto As String
Dim SumaDia As Integer
Dim ZLugar As Integer

Dim ZZTipo(1000, 2) As String
Dim ZZAyudaI(1000) As String
Dim ZZAyudaII(1000) As String
Dim ZZAyudaIII(1000) As String
Dim ZZAyudaIV(1000) As String
Dim ZZAyudaV(1000) As String
Dim ZZAyudaVI(1000) As String

Dim WWVto As String
Dim WWTipo As String
Dim WWAno As String
Dim WWNumero As String
Dim WWFecha As String
Dim WWEstado As String
Dim WWTitulo As String
Dim WWReferencia As String
Dim WWCentro As String
Dim WWOrigen As String
Dim WWEmisor As String
Dim WWResponsable As String
Dim WWEmisorII As String
Dim WWDescripcionII As String
Dim WWObservacionesII As String
Dim WWTipoAgenda As String

Dim ZZResponsable As String

Dim ZResponsable(1000) As String
Dim ZPlanifica(10000, 10) As String
Dim ZSac(1000, 10) As String
Dim ZImple(10, 3) As String

Private Sub Acepta_Click()

    ZSql = "DELETE AgendaRespo"
    spAgendaRespo = ZSql
    Set rstAgendaRespo = db.OpenRecordset(spAgendaRespo, dbOpenSnapshot, dbSQLPassThrough)

    If Desde.Text = "  /  /    " Then
        Desde.Text = "01/01/2000"
    End If
    If Hasta.Text = "  /  /    " Then
        Hasta.Text = "01/01/2999"
    End If
    
    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    ZZResponsable = ZResponsable(TipoI.ListIndex + 1)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.ResponsableDestino = " + "'" + ZZResponsable + "'"
    ZSql = ZSql + " AND CargaSac.Estado < 3"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
        With rstCargaSac
        .MoveFirst
        Do
            If .EOF = False Then
            
                ZLugarII = ZLugarII + 1
                
                ZSac(ZLugarII, 1) = rstCargaSac!Tipo
                ZSac(ZLugarII, 2) = rstCargaSac!Ano
                ZSac(ZLugarII, 3) = rstCargaSac!Numero
                ZSac(ZLugarII, 4) = rstCargaSac!Fecha
                
                .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaSac.Close
    End If
    
    For Ciclo = 1 To ZLugarII
    
        ZTipo = ZSac(Ciclo, 1)
        ZAno = ZSac(Ciclo, 2)
        ZNumero = ZSac(Ciclo, 3)
        ZFecha = ZSac(Ciclo, 4)
        
        Call Ceros(ZTipo, 2)
        Call Ceros(ZAno, 4)
        Call Ceros(ZNumero, 6)
        
        
        ZEntra = "S"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSacII"
        ZSql = ZSql + " Where CargaSacII.Tipo = " + "'" + ZTipo + "'"
        ZSql = ZSql + " and CargaSacII.Ano = " + "'" + ZAno + "'"
        ZSql = ZSql + " and CargaSacII.Numero = " + "'" + ZNumero + "'"
        spCargaSacII = ZSql
        Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSacII.RecordCount > 0 Then
        
            ZZAccion11 = Trim(rstCargaSacII!Accion11)
            ZZAccion12 = Trim(rstCargaSacII!Accion12)
            ZZAccion21 = Trim(rstCargaSacII!Accion21)
            ZZAccion22 = Trim(rstCargaSacII!Accion22)
            ZZAccion31 = Trim(rstCargaSacII!Accion31)
            ZZAccion32 = Trim(rstCargaSacII!Accion32)
            ZZAccion41 = Trim(rstCargaSacII!Accion41)
            ZZAccion42 = Trim(rstCargaSacII!Accion42)
            ZZAccion51 = Trim(rstCargaSacII!Accion51)
            ZZAccion52 = Trim(rstCargaSacII!Accion52)
            ZZAccion61 = Trim(rstCargaSacII!Accion61)
            ZZAccion62 = Trim(rstCargaSacII!Accion62)
            
            ZZResponsable1 = rstCargaSacII!Responsable1
            ZZResponsable2 = rstCargaSacII!Responsable2
            ZZResponsable3 = rstCargaSacII!Responsable3
            ZZResponsable4 = rstCargaSacII!Responsable4
            ZZResponsable5 = rstCargaSacII!Responsable5
            ZZResponsable6 = rstCargaSacII!Responsable6
            
            If ZZAccion11 <> "" Or ZZAccion12 <> "" Then
                ZEntra = "N"
            End If
            If ZZAccion21 <> "" Or ZZAccion22 <> "" Then
                ZEntra = "N"
            End If
            If ZZAccion31 <> "" Or ZZAccion32 <> "" Then
                ZEntra = "N"
            End If
            If ZZAccion41 <> "" Or ZZAccion42 <> "" Then
                ZEntra = "N"
            End If
            If ZZAccion51 <> "" Or ZZAccion52 <> "" Then
                ZEntra = "N"
            End If
            If ZZAccion61 <> "" Or ZZAccion62 <> "" Then
                ZEntra = "N"
            End If
            
            If ZZResponsable1 <> 0 Then
                ZEntra = "N"
            End If
            If ZZResponsable2 <> 0 Then
                ZEntra = "N"
            End If
            If ZZResponsable3 <> 0 Then
                ZEntra = "N"
            End If
            If ZZResponsable4 <> 0 Then
                ZEntra = "N"
            End If
            If ZZResponsable5 <> 0 Then
                ZEntra = "N"
            End If
            If ZZResponsable6 <> 0 Then
                ZEntra = "N"
            End If
            
            rstCargaSacII.Close
            
        End If
        
        If ZEntra = "S" Then
        
            SumaDia = 31
            Call Calcula_vencimiento(ZFecha, SumaDia, ZVto)
            ZPasa = 0
            ZFechaII = ZVto
            WFechaOrdII = Right$(ZFechaII, 4) + Mid$(ZFechaII, 4, 2) + Left$(ZFechaII, 2)
            
            If WFechaOrdII <= WHasta Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CargaSac"
                ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + ZTipo + "'"
                ZSql = ZSql + " and CargaSac.Ano = " + "'" + ZAno + "'"
                ZSql = ZSql + " and CargaSac.Numero = " + "'" + ZNumero + "'"
                spCargaSac = ZSql
                Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstCargaSac.RecordCount > 0 Then
                
                    WWVto = ZVto
                    WWTipo = Left$(Trim(ZZAyudaIII(rstCargaSac!Tipo)), 30)
                    WWAno = Str$(rstCargaSac!Ano)
                    WWNumero = Str$(rstCargaSac!Numero)
                    WWFecha = rstCargaSac!Fecha
                    WWEstado = ZZAyudaI(rstCargaSac!Estado)
                    WWTitulo = Left$(Trim(rstCargaSac!Titulo), 50)
                    WWReferencia = Left$(Trim(rstCargaSac!Referencia), 50)
                    WWCentro = ZZAyudaIV(rstCargaSac!Centro)
                    WWOrigen = ZZAyudaII(rstCargaSac!Origen)
                    WWEmisor = ZZAyudaV(rstCargaSac!ResponsableEmisor)
                    WWResponsable = ZZAyudaV(rstCargaSac!ResponsableDestino)
                    WWTipoAgenda = "1"
                    WWEmisorII = ""
                    WWDescripcionII = ""
                    WWObservacionesII = ""
                
                    rstCargaSac.Close
                    
                    Call Graba_Datos
                    
                End If
            End If
            
        End If
        
    Next Ciclo
    
    
    
    
    
    ZLugarII = 0
    Erase ZSac
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacII"
    ZSql = ZSql + " Where CargaSacII.Responsable1 = " + "'" + Str$(ZZOperadorResponsable) + "'"
    ZSql = ZSql + " or CargaSacII.Responsable2 = " + "'" + Str$(ZZOperadorResponsable) + "'"
    ZSql = ZSql + " or CargaSacII.Responsable3 = " + "'" + Str$(ZZOperadorResponsable) + "'"
    ZSql = ZSql + " or CargaSacII.Responsable4 = " + "'" + Str$(ZZOperadorResponsable) + "'"
    ZSql = ZSql + " or CargaSacII.Responsable5 = " + "'" + Str$(ZZOperadorResponsable) + "'"
    ZSql = ZSql + " or CargaSacII.Responsable6 = " + "'" + Str$(ZZOperadorResponsable) + "'"
    spCargaSacII = ZSql
    Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacII.RecordCount > 0 Then
        With rstCargaSacII
        .MoveFirst
        Do
            If .EOF = False Then
            
                ZLugarII = ZLugarII + 1
                ZSac(ZLugarII, 1) = rstCargaSacII!Tipo
                ZSac(ZLugarII, 2) = rstCargaSacII!Ano
                ZSac(ZLugarII, 3) = rstCargaSacII!Numero
                
                .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaSacII.Close
    End If
    
    For Ciclo = 1 To ZLugarII
    
        ZTipo = ZSac(Ciclo, 1)
        ZAno = ZSac(Ciclo, 2)
        ZNumero = ZSac(Ciclo, 3)
        
        Call Ceros(ZTipo, 2)
        Call Ceros(ZAno, 4)
        Call Ceros(ZNumero, 6)
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSac"
        ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + ZTipo + "'"
        ZSql = ZSql + " and CargaSac.Ano = " + "'" + ZAno + "'"
        ZSql = ZSql + " and CargaSac.Numero = " + "'" + ZNumero + "'"
        spCargaSac = ZSql
        Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSac.RecordCount > 0 Then
        
            rstCargaSac.Close
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSacII"
            ZSql = ZSql + " Where CargaSacII.Tipo = " + "'" + ZTipo + "'"
            ZSql = ZSql + " and CargaSacII.Ano = " + "'" + ZAno + "'"
            ZSql = ZSql + " and CargaSacII.Numero = " + "'" + ZNumero + "'"
            spCargaSacII = ZSql
            Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSacII.RecordCount > 0 Then
                ZImple(1, 1) = rstCargaSacII!Responsable1
                ZImple(2, 1) = rstCargaSacII!Responsable2
                ZImple(3, 1) = rstCargaSacII!Responsable3
                ZImple(4, 1) = rstCargaSacII!Responsable4
                ZImple(5, 1) = rstCargaSacII!Responsable5
                ZImple(6, 1) = rstCargaSacII!Responsable6
                ZImple(1, 2) = rstCargaSacII!Plazo1
                ZImple(2, 2) = rstCargaSacII!Plazo2
                ZImple(3, 2) = rstCargaSacII!Plazo3
                ZImple(4, 2) = rstCargaSacII!Plazo4
                ZImple(5, 2) = rstCargaSacII!Plazo5
                ZImple(6, 2) = rstCargaSacII!Plazo6
                rstCargaSacII.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSacIII"
            ZSql = ZSql + " Where CargaSacIII.Tipo = " + "'" + ZTipo + "'"
            ZSql = ZSql + " and CargaSacIII.Ano = " + "'" + ZAno + "'"
            ZSql = ZSql + " and CargaSacIII.Numero = " + "'" + ZNumero + "'"
            spCargaSacIII = ZSql
            Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSacIII.RecordCount > 0 Then
                ZImple(1, 3) = rstCargaSacIII!Estado1
                ZImple(2, 3) = rstCargaSacIII!Estado2
                ZImple(3, 3) = rstCargaSacIII!Estado3
                ZImple(4, 3) = rstCargaSacIII!Estado4
                ZImple(5, 3) = rstCargaSacIII!Estado5
                ZImple(6, 3) = rstCargaSacIII!Estado6
                rstCargaSacIII.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + ZTipo + "'"
            ZSql = ZSql + " and CargaSac.Ano = " + "'" + ZAno + "'"
            ZSql = ZSql + " and CargaSac.Numero = " + "'" + ZNumero + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
                ZFecha = rstCargaSac!Fecha
                ZEstado = rstCargaSac!Estado
                rstCargaSac.Close
            End If
            
            If ZEstado <= 3 Then
            
                For CicloRes = 1 To 6
                
                    ZEntra = "N"
                    
                    ZResponsable1 = ZImple(CicloRes, 1)
                    ZEstado1 = ZImple(CicloRes, 3)
                    ZPlazo1 = ZImple(CicloRes, 2)
                    If Trim(ZPlazo1) = "" Or ZPlazo1 = "  /  /    " Then
                        SumaDia = 31
                        Call Calcula_vencimiento(ZFecha, SumaDia, ZVto)
                        ZPlazo1 = ZVto
                    End If
                    
                    If Val(ZResponsable1) = ZZOperadorResponsable And Val(ZEstado1) = 0 Then
                        ZEntra = "S"
                        ZPlazo = ZPlazo1
                    End If
            
                    If ZEntra = "S" Then
                    
                        ZPasa = 0
                        ZFechaII = ZPlazo
                        WFechaOrdII = Right$(ZFechaII, 4) + Mid$(ZFechaII, 4, 2) + Left$(ZFechaII, 2)
            
                        If WFechaOrdII <= WHasta Then
                            
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM CargaSac"
                            ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + ZTipo + "'"
                            ZSql = ZSql + " and CargaSac.Ano = " + "'" + ZAno + "'"
                            ZSql = ZSql + " and CargaSac.Numero = " + "'" + ZNumero + "'"
                            spCargaSac = ZSql
                            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
                            If rstCargaSac.RecordCount > 0 Then
                            
                                WWVto = ZPlazo
                                WWTipo = Left$(Trim(ZZAyudaIII(rstCargaSac!Tipo)), 30)
                                WWAno = rstCargaSac!Ano
                                WWNumero = rstCargaSac!Numero
                                WWFecha = rstCargaSac!Fecha
                                WWEstado = ZZAyudaI(rstCargaSac!Estado)
                                WWTitulo = Left$(Trim(rstCargaSac!Titulo), 50)
                                WWReferencia = Left$(Trim(rstCargaSac!Referencia), 50)
                                WWCentro = ZZAyudaIV(rstCargaSac!Centro)
                                WWOrigen = ZZAyudaII(rstCargaSac!Origen)
                                WWEmisor = ZZAyudaV(rstCargaSac!ResponsableEmisor)
                                WWResponsable = ZZAyudaV(rstCargaSac!ResponsableDestino)
                                WWTipoAgenda = "1"
                                WWEmisorII = ""
                                WWDescripcionII = ""
                                WWObservacionesII = ""
                            
                                rstCargaSac.Close
                                
                                Call Graba_Datos
                                
                            End If
                            
                        End If
                        
                    End If
            
                Next CicloRes
                
            End If
            
        End If
        
    Next Ciclo
    
    
    
    
    
    
    
    
    Erase ZPlanifica
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Planifica"
    ZSql = ZSql + " Where Planifica.ResponsableII = " + "'" + Str$(ZZOperadorResponsable) + "'"
    ZSql = ZSql + " and Planifica.Responsable <> " + "'" + Str$(ZZOperadorResponsable) + "'"
    ZSql = ZSql + " and Planifica.Estado = " + "'" + "1" + "'"
    spPlanifica = ZSql
    Set rstPlanifica = db.OpenRecordset(spPlanifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstPlanifica.RecordCount > 0 Then
        With rstPlanifica
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZPasa = 0
                    ZFechaII = rstPlanifica!Vencimiento
                    WFechaOrdII = Right$(ZFechaII, 4) + Mid$(ZFechaII, 4, 2) + Left$(ZFechaII, 2)
                    
                    If WFechaOrdII <= WHasta Then
                        ZLugar = ZLugar + 1
                        ZPlanifica(ZLugar, 1) = ZZAyudaV(rstPlanifica!Responsable)
                        ZPlanifica(ZLugar, 2) = rstPlanifica!Descripcion
                        ZPlanifica(ZLugar, 3) = rstPlanifica!Observaciones
                        ZPlanifica(ZLugar, 4) = rstPlanifica!Vencimiento
                        ZPlanifica(ZLugar, 5) = rstPlanifica!Fecha
                        ZPlanifica(ZLugar, 6) = rstPlanifica!Estado
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPlanifica.Close
    End If
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Planifica"
    ZSql = ZSql + " Where Planifica.ResponsableII = " + "'" + Str$(ZZOperadorResponsable) + "'"
    ZSql = ZSql + " and Planifica.Responsable = " + "'" + Str$(ZZOperadorResponsable) + "'"
    ZSql = ZSql + " and Planifica.Estado = " + "'" + "1" + "'"
    spPlanifica = ZSql
    Set rstPlanifica = db.OpenRecordset(spPlanifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstPlanifica.RecordCount > 0 Then
        With rstPlanifica
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZPasa = 0
                    ZFechaII = rstPlanifica!Vencimiento
                    WFechaOrdII = Right$(ZFechaII, 4) + Mid$(ZFechaII, 4, 2) + Left$(ZFechaII, 2)
                    
                    If WFechaOrdII <= WHasta Then
                        ZLugar = ZLugar + 1
                        ZPlanifica(ZLugar, 1) = ZZAyudaV(rstPlanifica!Responsable)
                        ZPlanifica(ZLugar, 2) = rstPlanifica!Descripcion
                        ZPlanifica(ZLugar, 3) = rstPlanifica!Observaciones
                        ZPlanifica(ZLugar, 4) = rstPlanifica!Vencimiento
                        ZPlanifica(ZLugar, 5) = rstPlanifica!Fecha
                        ZPlanifica(ZLugar, 6) = rstPlanifica!Estado
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPlanifica.Close
    End If
    

    For Ciclo = 1 To ZLugar

        WWVto = ZPlanifica(Ciclo, 4)
        WWTipo = "Asignacion Tarea"
        WWAno = ""
        WWNumero = ""
        WWFecha = ZPlanifica(Ciclo, 5)
        If Val(ZPlanifica(Ciclo, 6)) = 1 Then
            WWEstado = "Pendiente"
                Else
            WWEstado = "Finalizado"
        End If
        WWTitulo = ""
        WWReferencia = ""
        WWCentro = ""
        WWOrigen = ""
        WWEmisor = ""
        WWResponsable = ""
        WWEmisorII = ZPlanifica(Ciclo, 1)
        WWDescripcionII = ZPlanifica(Ciclo, 2)
        WWObservacionesII = ZPlanifica(Ciclo, 3)
        WWTipoAgenda = "2"
                        
        Call Graba_Datos

    Next Ciclo



    Rem On Error GoTo WError
    
    Listado.WindowTitle = "Listado de Tareas Asignadas a Responsables"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT AgendaRespo.Persona, AgendaRespo.DesPersona, AgendaRespo.OrdVto, AgendaRespo.Vto, AgendaRespo.Tipo, AgendaRespo.Ano, AgendaRespo.Numero, AgendaRespo.Fecha, AgendaRespo.Estado, AgendaRespo.Titulo, AgendaRespo.Referencia, AgendaRespo.Centro, AgendaRespo.Origen, AgendaRespo.Emisor, AgendaRespo.EmisorII, AgendaRespo.DescripcionII, AgendaRespo.ObservacionesII, AgendaRespo.TipoAgenda " _
            + "From " _
            + DSQ + ".dbo.AgendaRespo AgendaRespo"
    
    Rem Uno = "{Planifica.ResponsableII} in " + ZDesdeII + " to " + ZHastaII
    Rem Dos = " and {Planifica.OrdVencimiento} in " + Chr$(34) + DesdeFecha + Chr$(34) + " to " + Chr$(34) + HastaFecha + Chr$(34)
    Rem Tres = " and {Planifica.Estado} in " + ZDesdeIII + " to " + ZHastaIII
    Rem Cuatro = " and {Planifica.Responsable} in " + ZDesdeI + " to " + ZHastaI
    
    Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
    Rem Listado.SelectionFormula = Uno + Dos + Tres + Cuatro
    
    Listado.Connect = Connect()
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.ReportFileName = "ListaAgendaRespo.Rpt"
    Listado.Action = 1
    
    Exit Sub

WError:
    Resume Next
    
End Sub

Private Sub Graba_Datos()

    WWPersona = ZZResponsable
    WWDesPersona = TipoI.Text
    
    WWOrdVto = Right$(WWVto, 4) + Mid$(WWVto, 4, 2) + Left$(WWVto, 2)
    
    ZSql = ""
    ZSql = ZSql + "INSERT INTO AgendaRespo ("
    ZSql = ZSql + "Persona ,"
    ZSql = ZSql + "DesPersona ,"
    ZSql = ZSql + "TipoAgenda ,"
    ZSql = ZSql + "OrdVto ,"
    ZSql = ZSql + "Vto ,"
    ZSql = ZSql + "Tipo ,"
    ZSql = ZSql + "Ano ,"
    ZSql = ZSql + "Numero ,"
    ZSql = ZSql + "Fecha ,"
    ZSql = ZSql + "Estado ,"
    ZSql = ZSql + "Titulo ,"
    ZSql = ZSql + "Referencia ,"
    ZSql = ZSql + "Centro ,"
    ZSql = ZSql + "Origen ,"
    ZSql = ZSql + "Emisor ,"
    ZSql = ZSql + "Responsable ,"
    ZSql = ZSql + "EmisorII ,"
    ZSql = ZSql + "DescripcionII ,"
    ZSql = ZSql + "ObservacionesII )"
    ZSql = ZSql + "Values ("
    ZSql = ZSql + "'" + WWPersona + "',"
    ZSql = ZSql + "'" + WWDesPersona + "',"
    ZSql = ZSql + "'" + WWTipoAgenda + "',"
    ZSql = ZSql + "'" + WWOrdVto + "',"
    ZSql = ZSql + "'" + WWVto + "',"
    ZSql = ZSql + "'" + WWTipo + "',"
    ZSql = ZSql + "'" + WWAno + "',"
    ZSql = ZSql + "'" + WWNumero + "',"
    ZSql = ZSql + "'" + WWFecha + "',"
    ZSql = ZSql + "'" + WWEstado + "',"
    ZSql = ZSql + "'" + WWTitulo + "',"
    ZSql = ZSql + "'" + WWReferencia + "',"
    ZSql = ZSql + "'" + WWCentro + "',"
    ZSql = ZSql + "'" + WWOrigen + "',"
    ZSql = ZSql + "'" + WWEmisor + "',"
    ZSql = ZSql + "'" + WWResponsable + "',"
    ZSql = ZSql + "'" + WWEmisorII + "',"
    ZSql = ZSql + "'" + WWDescripcionII + "',"
    ZSql = ZSql + "'" + WWObservacionesII + "')"
    
     spAgendaRespo = ZSql
     Set rstAgenadRespo = db.OpenRecordset(spAgendaRespo, dbOpenSnapshot, dbSQLPassThrough)
    
End Sub


Private Sub Cancela_click()
    PrgListaTareaAsignada.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    
    Erase ZZAyudaI
    Erase ZZAyudaII
    Erase ZZAyudaIII
    Erase ZZAyudaIV
    Erase ZZAyudaV
    
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
    
    TipoI.Clear
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ResponsableSac"
    spResponsableSac = ZSql
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        With rstResponsableSac
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZLugar = ZLugar + 1
                    ZResponsable(ZLugar) = rstResponsableSac!Codigo
                    
                    If Val(ZZOperadorResponsable) = rstResponsableSac!Codigo Then
                        ZPuntero = ZLugar
                    End If
                
                    TipoI.AddItem Trim(rstResponsableSac!Descripcion)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstResponsableSac.Close
    End If
    
    TipoI.ListIndex = ZPuntero - 1
    
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    
    Panta.Value = True
    Impresora.Value = False
    
    Frame2.Visible = True
    
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

