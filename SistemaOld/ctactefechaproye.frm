VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCtaCtefechaProye 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Proyeccion de deuda a Fecha"
   ClientHeight    =   2775
   ClientLeft      =   1605
   ClientTop       =   585
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   ScaleHeight     =   2775
   ScaleWidth      =   9015
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   7215
      Begin VB.Frame Frame4 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   2775
         Begin VB.OptionButton Pesos 
            Caption         =   "Pesos"
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
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Dolares 
            Caption         =   "Dolares"
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
            Left            =   1560
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
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
         Left            =   4440
         TabIndex        =   7
         Top             =   240
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
         Left            =   4440
         TabIndex        =   6
         Top             =   600
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
         TabIndex        =   5
         Top             =   1200
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
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox FechaII 
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
      Begin VB.Label Label1 
         Caption         =   "Hasta Periodo"
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
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Periodo"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7920
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ctacte.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Clientes"
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
      Left            =   7680
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7680
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCtaCtefechaProye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WNume As String
Private WPasa As String
Private WTitulo As String
Private Importe3 As Double
Private Acumula As Double
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstRecibos As Recordset
Dim spRecibos As String
Dim rstImpCtaCteProy As Recordset
Dim spImpCtaCteProy As String


Dim XParam As String
Dim WRecibo(20000, 10) As String
Dim WTrabajo(10000) As String

Dim ZZImporte As Double
Dim ZZDatos(100, 15) As String



Private Sub Acepta_Click()

    Dim ZZPeriodo(100, 2) As String
        
    OPEN_FILE_ProyectaDias
    
    DaII = 0
    With rstProyectaDias
        .Index = "Clave"
        .Seek ">=", DaII
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With

    WAno = Val(Right$(Fecha.Text, 4))
    WMes = Val(Mid$(Fecha.Text, 4, 2))
    WDia = Val(Left$(Fecha.Text, 2))

    WAnoII = Val(Right$(FechaII.Text, 4))
    WMesII = Val(Mid$(FechaII.Text, 4, 2))
    WDiaII = Val(Left$(FechaII.Text, 2))
    
    ZZLugar = 0
    
    Do
    
        ZZLugar = ZZLugar + 1
        
        ZZPeriodo(ZZLugar, 1) = Str$(WAno)
        ZZPeriodo(ZZLugar, 2) = Str$(WMes)
        
        If WAno = WAnoII And WMes = WMesII Then
            Exit Do
        End If
        
        WMes = WMes + 1
        If WMes > 12 Then
            WAno = WAno + 1
            WMes = 1
        End If
        
    Loop
    
    Erase ZZDatos
        
    For ZZZCiclo = 1 To ZZLugar
    
        Auxi1 = ZZPeriodo(ZZZCiclo, 1)
        Auxi2 = ZZPeriodo(ZZZCiclo, 2)
        Call Ceros(Auxi1, 4)
        Call Ceros(Auxi2, 2)
        WFecha = Auxi1 + Auxi2 + "31"
        Select Case Val(Auxi2)
            Case 1, 3, 5, 7, 8, 10, 12
                WFechaII = "31" + "/" + Auxi2 + "/" + Auxi1
            Case 2
                WFechaII = "28" + "/" + Auxi2 + "/" + Auxi1
            Case Else
                WFechaII = "30" + "/" + Auxi2 + "/" + Auxi1
        End Select
        
        ZZDatos(ZZZCiclo, 1) = Auxi2 + "/" + Auxi1
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CtaCte SET "
        ZSql = ZSql + "ProyectaSaldo = " + "'" + "0" + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spImpCtaCteProy, dbOpenSnapshot, dbSQLPassThrough)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CtaCte SET "
        ZSql = ZSql + "ProyectaSaldo = Saldo"
        ZSql = ZSql + " Where Tipo < " + "'" + "06" + "'"
        ZSql = ZSql + " and OrdFecha <= " + "'" + WFecha + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        Erase WRecibo
        Renglon = 0
        
        
        ZSql = ""
        ZSql = ZSql + "Select Recibos.Importe1, Recibos.Fechaord, Recibos.Clave, Recibos.Tipo1, Recibos.Numero1, Recibos.Importe1, Recibos.Cliente, Recibos.Recibo"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Importe1 <> 0 "
        ZSql = ZSql + " and Recibos.FechaOrd > " + "'" + WFecha + "'"
        ZSql = ZSql + " Order by Recibos.Clave"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
                
            With rstRecibos
                .MoveFirst
                Do
                            
                    If WFecha < !FechaOrd Then
                    
                        If !Importe1 <> 0 Then
                        
                            Renglon = Renglon + 1
                    
                            WRecibo(Renglon, 1) = !Tipo1
                            WRecibo(Renglon, 2) = !Numero1
                            WRecibo(Renglon, 3) = Str$(!Importe1)
                            WRecibo(Renglon, 4) = !Clave
                            WRecibo(Renglon, 5) = !Recibo
                            WRecibo(Renglon, 6) = !Cliente
                                
                        End If
                    
                    End If
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End With
            
            rstRecibos.Close
            
        End If
        
        
        For Ciclo = 1 To Renglon
        
            WTipo = WRecibo(Ciclo, 1)
            WNumero = WRecibo(Ciclo, 2)
            WImporte = Val(WRecibo(Ciclo, 3))
            XClave = WRecibo(Ciclo, 4)
            XRecibo = WRecibo(Ciclo, 5)
            WCliente = WRecibo(Ciclo, 6)
                        
            Call Ceros(WTipo, 2)
            Call Ceros(WNumero, 8)
                        
            WClave = WTipo + WNumero + "01"
            
            WProv = 0
            WParidad = 0
            
            spClientes = "ConsultaClientes " + "'" + WCliente + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                WProv = rstClientes!Provincia
                rstClientes.Close
            End If
            
            If WProv = 24 Then
                Auxi1 = XRecibo
                Call Ceros(Auxi1, 8)
                ClaveCtacte = "06" + Auxi1 + "01"
                spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtacte.RecordCount > 0 Then
                    WParidad = Str$(rstCtacte!Paridad)
                    rstCtacte.Close
                        Else
                    ClaveCtacte = "07" + Auxi1 + "01"
                    spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCtacte.RecordCount > 0 Then
                        WParidad = Str$(rstCtacte!Paridad)
                        rstCtacte.Close
                    End If
                End If
            End If
            
            If WProv = 24 And WParidad <> 0 Then
               ZZImporte = (WImporte / Val(WParidad))
                    Else
               ZZImporte = WImporte
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CtaCte SET "
            ZSql = ZSql + "ProyectaSaldo = ProyectaSaldo + " + "'" + Str$(WImporte) + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
            spCtacte = ZSql
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    
        Next Ciclo
        
        
        WAno = Right$(Fecha.Text, 4)
        WMes = Mid$(Fecha.Text, 4, 2)
        WDia = Left$(Fecha.Text, 2)
        XClave = WAno + WMes + WDia
        WParidad = 0
    
        spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
            With rstCambios
                .MoveLast
                WParidad = Str$(rstCambios!Cambio)
            End With
            rstCambios.Close
        End If
        
        
        Erase WTrabajo
        ZZLugarII = 0
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.ProyectaSaldo <> 0 "
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            
            With rstCtacte
                    .MoveFirst
                    Do
                    
                        ZZLugarII = ZZLugarII + 1
                        WTrabajo(ZZLugarII) = !Clave
                        
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
            End With
            rstCtacte.Close
        
        End If
        
        
        For ZZCiclo = 1 To ZZLugarII
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Clave = " + WTrabajo(ZZCiclo)
            spCtacte = ZSql
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtacte.RecordCount > 0 Then
            
                ZZCliente = rstCtacte!Cliente
                ZZImporte = rstCtacte!ProyectaSaldo
                ZZFecha = rstCtacte!Fecha
                ZZTotal = rstCtacte!Total
                ZZTotalUs = rstCtacte!Totalus
                
                rstCtacte.Close
                                
                If Trim(ZZCliente) <> "" Then
                    WRazon = ""
                    spCliente = "ConsultaCliente " + ZZCliente
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WRazon = rstCliente!Razon
                        WProv = rstCliente!Provincia
                        rstCliente.Close
                    End If
                        Else
                    WProv = 1
                End If
                        
                If Pesos.Value = True Then
                    If WProv = 24 And WParidad <> 0 Then
                        ZZImporte = ZZImporte * Val(WParidad)
                    End If
                End If
                
                If Dolares.Value = True Then
                    If ZZTotalUs <> 0 Then
                        Pari = ZZTotal / ZZTotalUs
                        ZZImporte3 = ZZImporte3 / Pari
                    End If
                End If
                
                XFecha = ZZFecha
                ZZDias = DateDiff("d", XFecha, WFechaII)
                
                If ZZDias <= 30 Then
                    ZZDatos(ZZZCiclo, 2) = Str$(Val(ZZDatos(ZZZCiclo, 2)) + ZZImporte)
                        Else
                    If ZZDias <= 60 Then
                        ZZDatos(ZZZCiclo, 3) = Str$(Val(ZZDatos(ZZZCiclo, 3)) + ZZImporte)
                            Else
                        If ZZDias <= 90 Then
                            ZZDatos(ZZZCiclo, 4) = Str$(Val(ZZDatos(ZZZCiclo, 4)) + ZZImporte)
                                Else
                            If ZZDias <= 120 Then
                                ZZDatos(ZZZCiclo, 5) = Str$(Val(ZZDatos(ZZZCiclo, 5)) + ZZImporte)
                                    Else
                                If ZZDias <= 150 Then
                                    ZZDatos(ZZZCiclo, 6) = Str$(Val(ZZDatos(ZZZCiclo, 6)) + ZZImporte)
                                        Else
                                    If ZZDias <= 180 Then
                                        ZZDatos(ZZZCiclo, 7) = Str$(Val(ZZDatos(ZZZCiclo, 7)) + ZZImporte)
                                            Else
                                        ZZDatos(ZZZCiclo, 8) = Str$(Val(ZZDatos(ZZZCiclo, 8)) + ZZImporte)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
                
        Next ZZCiclo
        
        With rstProyectaDias

            .Index = "Clave"
                            
            .AddNew
    
            !Clave = ZZZCiclo
            !Fecha = ZZDatos(ZZZCiclo, 1)
            !Impo1 = Val(ZZDatos(ZZZCiclo, 2))
            !Impo2 = Val(ZZDatos(ZZZCiclo, 3))
            !Impo3 = Val(ZZDatos(ZZZCiclo, 4))
            !Impo4 = Val(ZZDatos(ZZZCiclo, 5))
            !Impo5 = Val(ZZDatos(ZZZCiclo, 6))
            !Impo6 = Val(ZZDatos(ZZZCiclo, 7))
            !Impo7 = Val(ZZDatos(ZZZCiclo, 8))
            !Empresa = 1

            .Update
    
        End With

        
        
    Next ZZZCiclo
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    WTitulo = ""
    
    If Pesos.Value = True Then
        WTitulo = WTitulo + "Pesos"
    End If
    If Dolares.Value = True Then
        WTitulo = WTitulo + "Dolares"
    End If
    
    WTitulo = WTitulo + " al " + Fecha.Text
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Proyeccion de dias a Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.ReportFileName = "ProyectaDias.rpt"
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Rem Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub




Private Sub AceptaAnterior_Click()

    Dim ZZPeriodo(100, 2) As String
        
    OPEN_FILE_ProyectaDias
    
    DaII = 0
    With rstProyectaDias
        .Index = "Clave"
        .Seek ">=", DaII
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With

    WAno = Val(Right$(Fecha.Text, 4))
    WMes = Val(Mid$(Fecha.Text, 4, 2))
    WDia = Val(Left$(Fecha.Text, 2))

    WAnoII = Val(Right$(FechaII.Text, 4))
    WMesII = Val(Mid$(FechaII.Text, 4, 2))
    WDiaII = Val(Left$(FechaII.Text, 2))
    
    ZZLugar = 0
    
    Do
    
        ZZLugar = ZZLugar + 1
        
        ZZPeriodo(ZZLugar, 1) = Str$(WAno)
        ZZPeriodo(ZZLugar, 2) = Str$(WMes)
        
        If WAno = WAnoII And WMes = WMesII Then
            Exit Do
        End If
        
        WMes = WMes + 1
        If WMes > 12 Then
            WAno = WAno + 1
            WMes = 1
        End If
        
    Loop
    
    Erase ZZDatos
        
    For ZZZCiclo = 1 To ZZLugar
    
        Auxi1 = ZZPeriodo(ZZZCiclo, 1)
        Auxi2 = ZZPeriodo(ZZZCiclo, 2)
        Call Ceros(Auxi1, 4)
        Call Ceros(Auxi2, 2)
        WFecha = Auxi1 + Auxi2 + "31"
        Select Case Val(Auxi2)
            Case 1, 3, 5, 7, 8, 10, 12
                WFechaII = "31" + "/" + Auxi2 + "/" + Auxi1
            Case 2
                WFechaII = "28" + "/" + Auxi2 + "/" + Auxi1
            Case Else
                WFechaII = "30" + "/" + Auxi2 + "/" + Auxi1
        End Select
        
        ZZDatos(ZZZCiclo, 1) = Auxi2 + "/" + Auxi1
        
        ZSql = ""
        ZSql = ZSql + "DELETE ImpCtaCteProy"
        spImpCtaCteProy = Sql1 + Sql2
        Set rstImpCtaCteProy = db.OpenRecordset(spImpCtaCteProy, dbOpenSnapshot, dbSQLPassThrough)
        
        
        ZSql = ""
        ZSql = ZSql + "Select CtaCte.Tipo, CtaCte.Impre, CtaCte.Numero, CtaCte.Renglon, CtaCte.Cliente, CtaCte.Fecha, CtaCte.Total, CtaCte.TotalUs, CtaCte.Saldo, CtaCte.SaldoUs, CtaCte.OrdFecha, CtaCte.Paridad, CtaCte.Clave"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.Tipo < " + "'" + "06" + "'"
        ZSql = ZSql + " and CtaCte.OrdFecha <= " + "'" + WFecha + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
    
            With rstCtacte
                    .MoveFirst
                    Do
                    
                        If Val(!Tipo) < 6 Then
                        
                            If !OrdFecha <= WFecha Then
                        
                                WTipo = !Tipo
                                WImpre = !Impre
                                WNumero = !Numero
                                WRenglon = !Renglon
                                WCliente = !Cliente
                                XFecha = !Fecha
                                WEstado = ""
                                Wvencimiento = ""
                                WVencimiento1 = ""
                                WTotal = !Total
                                WTotalUs = !Totalus
                                WSaldo = !Saldo
                                WSaldoUs = !Saldous
                                WNeto = 0
                                WIva1 = 0
                                WWIva2 = 0
                                WOrdFecha = !OrdFecha
                                WOrdVencimiento = ""
                                WOrdVencimiento1 = ""
                                WPedido = ""
                                WRemito = ""
                                WOrden = ""
                                WParidad = !Paridad
                                WProvincia = ""
                                WVendedor = 0
                                WRubro = 0
                                WCcomprobante = ""
                                WAceptada = ""
                                WCosto = 0
                                WImporte1 = 0
                                WImporte2 = 0
                                WImporte3 = 0
                                WImporte4 = 0
                                WImporte5 = 0
                                WImporte6 = 0
                                WImporte7 = 0
                                WClave = !Clave
                        
                                With rstImpCtaCteProy
                
                                    .Index = "Clave"
                                                    
                                    .AddNew
                            
                                    !Tipo = WTipo
                                    !Impre = WImpre
                                    !Numero = WNumero
                                    !Renglon = WRenglon
                                    !Cliente = WCliente
                                    !Fecha = XFecha
                                    !Estado = WEstado
                                    !Vencimiento = Wvencimiento
                                    !Vencimiento1 = WVencimiento1
                                    !Total = WTotal
                                    !Totalus = WTotalUs
                                    !Saldo = WSaldo
                                    !Saldous = WSaldoUs
                                    !Neto = WNeto
                                    !Iva1 = WIva1
                                    !Iva2 = WIva2
                                    !OrdFecha = WOrdFecha
                                    !OrdVencimiento = WOrdVencimiento
                                    !OrdVencimiento1 = WOrdVencimiento1
                                    !Pedido = WPedido
                                    !Remito = WRemito
                                    !Orden = WOrden
                                    !Paridad = WParidad
                                    !Provincia = WProvincia
                                    !vendedor = WVendedor
                                    !Rubro = WRubro
                                    !Comprobante = WComprobante
                                    !Aceptada = WAceptada
                                    !Costo = WCosto
                                    !Importe1 = 0
                                    !Importe2 = 0
                                    !Importe3 = 0
                                    !Importe4 = 0
                                    !Importe5 = 0
                                    !Importe6 = 0
                                    !Importe7 = 0
                                    !Clave = WClave
                                    WNume = Str$(!Numero)
                                    Call Ceros(WNume, 8)
                                    !ClaveImpre = !Cliente + !OrdFecha + !Tipo + WNume
        
                                    If !Total > 0 Then
                                        !Importe1 = !Total
                                        !Importe2 = 0
                                            Else
                                        !Importe1 = 0
                                        !Importe2 = !Total
                                    End If
                                    !Importe3 = !Saldo
        
                                    .Update
                            
                                End With
                        
                            End If
                        
                        End If
                        
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
            End With
            rstCtacte.Close
        
        End If
        
        
        Erase WRecibo
        Renglon = 0
        
        
        ZSql = ""
        ZSql = ZSql + "Select Recibos.Importe1, Recibos.Fechaord, Recibos.Clave, Recibos.Tipo1, Recibos.Numero1, Recibos.Importe1, Recibos.Cliente, Recibos.Recibo"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Importe1 <> 0 "
        ZSql = ZSql + " and Recibos.FechaOrd > " + "'" + WFecha + "'"
        ZSql = ZSql + " Order by Recibos.Clave"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
                
            With rstRecibos
                .MoveFirst
                Do
                            
                    If WFecha < !FechaOrd Then
                    
                        If !Importe1 <> 0 Then
                        
                            Renglon = Renglon + 1
                    
                            WRecibo(Renglon, 1) = !Tipo1
                            WRecibo(Renglon, 2) = !Numero1
                            WRecibo(Renglon, 3) = Str$(!Importe1)
                            WRecibo(Renglon, 4) = !Clave
                            WRecibo(Renglon, 5) = !Recibo
                            WRecibo(Renglon, 6) = !Cliente
                                
                        End If
                    
                    End If
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End With
            
            rstRecibos.Close
            
        End If
        
        
        For Ciclo = 1 To Renglon
        
            WTipo = WRecibo(Ciclo, 1)
            WNumero = WRecibo(Ciclo, 2)
            WImporte = Val(WRecibo(Ciclo, 3))
            XClave = WRecibo(Ciclo, 4)
            XRecibo = WRecibo(Ciclo, 5)
            WCliente = WRecibo(Ciclo, 6)
                        
            Call Ceros(WTipo, 2)
            Call Ceros(WNumero, 8)
                        
            WClave = WTipo + WNumero + "01"
            
            WProv = 0
            WParidad = 0
            
            spClientes = "ConsultaClientes " + "'" + WCliente + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                WProv = rstClientes!Provincia
                rstClientes.Close
            End If
            
            If WProv = 24 Then
                Auxi1 = XRecibo
                Call Ceros(Auxi1, 8)
                ClaveCtacte = "06" + Auxi1 + "01"
                spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtacte.RecordCount > 0 Then
                    WParidad = Str$(rstCtacte!Paridad)
                    rstCtacte.Close
                        Else
                    ClaveCtacte = "07" + Auxi1 + "01"
                    spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCtacte.RecordCount > 0 Then
                        WParidad = Str$(rstCtacte!Paridad)
                        rstCtacte.Close
                    End If
                End If
            End If
            
            With rstImpCtaCteProy
                .Index = "Clave"
                .Seek "=", WClave
                If .NoMatch = False Then
                    .Edit
                    If WProv = 24 And WParidad <> 0 Then
                        !Importe3 = !Importe3 + (WImporte / Val(WParidad))
                            Else
                        !Importe3 = !Importe3 + WImporte
                    End If
                    .Update
                End If
            End With
        
        Next Ciclo
        
        
        WAno = Right$(Fecha.Text, 4)
        WMes = Mid$(Fecha.Text, 4, 2)
        WDia = Left$(Fecha.Text, 2)
        XClave = WAno + WMes + WDia
        WParidad = 0
    
        spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
            With rstCambios
                .MoveLast
                WParidad = Str$(rstCambios!Cambio)
            End With
            rstCambios.Close
        End If
        
        With rstImpCtaCteProy
                .Index = "ClaveImpre"
                .MoveFirst
                Do
                
                    ZZImporte = !Importe3
                    Call Redondeo(ZZImporte)
                    If ZZImporte = 0 Then
                    
                        .Delete
                            
                            Else
                            
                        .Edit
                        
                        If Trim(!Cliente) <> "" Then
                            WRazon = ""
                            spCliente = "ConsultaCliente " + !Cliente
                            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                            If rstCliente.RecordCount > 0 Then
                                WRazon = rstCliente!Razon
                                WProv = rstCliente!Provincia
                                rstCliente.Close
                            End If
                                Else
                            WProv = 1
                        End If
                        
                        If Pesos.Value = True Then
                            If WProv = 24 And WParidad <> 0 Then
                                !Importe3 = !Importe3 * Val(WParidad)
                            End If
                        End If
                        
                        If Dolares.Value = True Then
                            If rstImpCtaCteProy!Totalus <> 0 Then
                                Pari = rstImpCtaCteProy!Total / rstImpCtaCteProy!Totalus
                                !Importe1 = !Importe1 / Pari
                                !Importe2 = !Importe2 / Pari
                                !Importe3 = !Importe3 / Pari
                            End If
                        End If
                        
                        !Razon = WRazon
                        
                        XFecha = !Fecha
                        ZZDias = DateDiff("d", XFecha, WFechaII)
                        !Importe4 = ZZDias
                        
                        If ZZDias <= 30 Then
                            ZZDatos(ZZZCiclo, 2) = Str$(Val(ZZDatos(ZZZCiclo, 2)) + !Importe3)
                                Else
                            If ZZDias <= 60 Then
                                ZZDatos(ZZZCiclo, 3) = Str$(Val(ZZDatos(ZZZCiclo, 3)) + !Importe3)
                                    Else
                                If ZZDias <= 90 Then
                                    ZZDatos(ZZZCiclo, 4) = Str$(Val(ZZDatos(ZZZCiclo, 4)) + !Importe3)
                                        Else
                                    If ZZDias <= 120 Then
                                        ZZDatos(ZZZCiclo, 5) = Str$(Val(ZZDatos(ZZZCiclo, 5)) + !Importe3)
                                            Else
                                        If ZZDias <= 150 Then
                                            ZZDatos(ZZZCiclo, 6) = Str$(Val(ZZDatos(ZZZCiclo, 6)) + !Importe3)
                                                Else
                                            If ZZDias <= 180 Then
                                                ZZDatos(ZZZCiclo, 7) = Str$(Val(ZZDatos(ZZZCiclo, 7)) + !Importe3)
                                                    Else
                                                ZZDatos(ZZZCiclo, 8) = Str$(Val(ZZDatos(ZZZCiclo, 8)) + !Importe3)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        .Update
                           
                    End If
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
        End With
        
        With rstProyectaDias

            .Index = "Clave"
                            
            .AddNew
    
            !Clave = ZZZCiclo
            !Fecha = ZZDatos(ZZZCiclo, 1)
            !Impo1 = Val(ZZDatos(ZZZCiclo, 2))
            !Impo2 = Val(ZZDatos(ZZZCiclo, 3))
            !Impo3 = Val(ZZDatos(ZZZCiclo, 4))
            !Impo4 = Val(ZZDatos(ZZZCiclo, 5))
            !Impo5 = Val(ZZDatos(ZZZCiclo, 6))
            !Impo6 = Val(ZZDatos(ZZZCiclo, 7))
            !Impo7 = Val(ZZDatos(ZZZCiclo, 8))
            !Empresa = 1

            .Update
    
        End With

        
        
    Next ZZZCiclo
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    WTitulo = ""
    
    If Pesos.Value = True Then
        WTitulo = WTitulo + "Pesos"
    End If
    If Dolares.Value = True Then
        WTitulo = WTitulo + "Dolares"
    End If
    
    WTitulo = WTitulo + " al " + Fecha.Text
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Proyeccion de dias a Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.ReportFileName = "ProyectaDias.rpt"
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Rem Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgCtaCtefechaProye.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCliente!Cliente + "     " + rstCliente!Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCliente!Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
        End With
    End If
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_ImpCtacteProy
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
       
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    
    spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Desde.Text = rstCliente!Cliente
        Hasta.Text = rstCliente!Cliente
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
    End If
End Sub

Sub Form_Load()
    Fecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Pesos.Value = True
    Dolares.Value = False
    Frame2.Visible = True
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
            
                    DA = Len(rstCliente!Razon) - WEspacios
                
                    For aa = 1 To DA
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                            Auxi = rstCliente!Cliente
                            IngresaItem = Auxi + "    " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next aa
                    .MoveNext
                    
                        Else
                        
                    Exit Do
                
                End If
            Loop
        End With
        rstCliente.Close
    End If
    End If

End Sub



