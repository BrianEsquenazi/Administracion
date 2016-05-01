VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListIbVen 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Percepciones de Ingresos Brutos"
   ClientHeight    =   2730
   ClientLeft      =   3315
   ClientTop       =   2175
   ClientWidth     =   5685
   LinkTopic       =   "Form2"
   ScaleHeight     =   2730
   ScaleWidth      =   5685
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   4575
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
         Left            =   1680
         TabIndex        =   10
         Top             =   1200
         Width           =   2655
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   240
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   1680
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   1680
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
         Left            =   3240
         TabIndex        =   5
         Top             =   240
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
         Left            =   3240
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1095
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5280
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wListIbVen.rpt"
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
Attribute VB_Name = "PrgListIbVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstRecibo As Recordset
Dim spRecibo As String
Dim XParam As String
Dim Vector(10000, 4) As String
Dim WClave As String
Dim WFecha As String
Dim WTipo As String
Dim WNumero As String
Dim WImpoIb As Double

Private Sub Acepta_Click()

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    Select Case Tipo.ListIndex
        Case 0
            ZSql = ""
            ZSql = ZSql + "UPDATE Ctacte SET "
            ZSql = ZSql + " ImpoIbTucu = 0"
            ZSql = ZSql + " Where ImpoIbTucu IS NULL"
            spCtacte = ZSql
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            
            spCtacte = "ModificaCtacteImporteIva0"
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
            Rem Procesa las cobranzas
    
            Renglon = 0
            Erase Vector
    
            XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
            spRecibo = "ListaRecibosDifeI" + XParam
            Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibo.RecordCount > 0 Then
                With rstRecibo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstRecibo!Clave
                            Vector(Renglon, 2) = rstRecibo!Fecha
                            Vector(Renglon, 3) = rstRecibo!Tipo1
                            Vector(Renglon, 4) = rstRecibo!Numero1
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibo.Close
            End If
     
            For Cicla = 1 To Renglon
    
                WClave = Vector(Cicla, 1)
                WFecha = Vector(Cicla, 2)
                WTipo = Vector(Cicla, 3)
                WNumero = Vector(Cicla, 4)
        
                ClaveCtacte = WTipo + WNumero + "01"
                spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtacte.RecordCount > 0 Then
                    WImpoIb = IIf(IsNull(rstCtacte!ImpoIb), "0", rstCtacte!ImpoIb)
                    If WImpoIb = 0 Then
                        Vector(Cicla, 1) = ""
                        Vector(Cicla, 2) = ""
                        Vector(Cicla, 3) = ""
                        Vector(Cicla, 4) = ""
                    End If
                    rstCtacte.Close
                        Else
                    Vector(Cicla, 1) = ""
                    Vector(Cicla, 2) = ""
                    Vector(Cicla, 3) = ""
                    Vector(Cicla, 4) = ""
                End If
        
            Next Cicla
    
            For Cicla = 1 To Renglon
    
                WClave = Vector(Cicla, 1)
                If WClave <> "" Then
        
                    WTipo = Vector(Cicla, 3)
                    WNumero = Vector(Cicla, 4)
                    WRecibo = Val(Left$(WClave, 6))
                    WSale = "N"
        
                    XParam = "'" + WTipo + "','" _
                                 + WNumero + "'"
                    spRecibo = "ListaRecibosFactura " + XParam
                    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstRecibo.RecordCount > 0 Then
                        With rstRecibo
                            .MoveFirst
                            Do
                                If .EOF = False Then
                                    If Val(rstRecibo!Recibo) < WRecibo Then
                                        WSale = "S"
                                    End If
                                    .MoveNext
                                        Else
                                    Exit Do
                                End If
                            Loop
                        End With
                        rstRecibo.Close
                    End If
            
                    If WSale = "S" Then
                        Vector(Cicla, 1) = ""
                        Vector(Cicla, 2) = ""
                        Vector(Cicla, 3) = ""
                        Vector(Cicla, 4) = ""
                    End If
            
                End If
        
            Next Cicla
    
            For Cicla = 1 To Renglon
                
                WClave = Vector(Cicla, 1)
                If WClave <> "" Then
        
                    WTipo = Vector(Cicla, 3)
                    WNumero = Vector(Cicla, 4)
                    
                    ClaveCtacte = WTipo + WNumero + "01"
                    XParam = "'" + ClaveCtacte + "','" _
                                 + WClave + "'"
                    spCtacte = "ModificaCtacteIb " + XParam
                    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
                End If
        
            Next Cicla
        
            With rstEmpresa
                .Index = "Empresa"
                .Seek "=", Val(WEmpresa)
                If .NoMatch = False Then
                    WAuxiliar = !Nombre
                End If
            End With
    
            WTitulo = "del " + Desde.Text + " al " + Hasta.Text
    
            With rstAuxiliar
                .Index = "Clave"
                .Seek "=", 1
                If .NoMatch = False Then
                    .Edit
                    !Nombre = WAuxiliar
                    !Varios = Left$(WTitulo, 50)
                    .Update
                End If
            End With
    
            Listado.WindowTitle = "Listado de Percepcion de Ingresos Brutos (Buenos Aires)"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
    
            Rem Listado.GroupSelectionFormula = "{CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            If Impresora.Value = True Then
                Listado.Destination = 1
                    Else
                Listado.Destination = 0
            End If
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT CtaCte.Tipo, CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.Importe4, CtaCte.Importe8, " _
                    + "Cliente.Razon, Cliente.Cuit, " _
                    + "Recibos.Recibo " _
                    + "From " _
                    + DSQ + ".dbo.CtaCte CtaCte, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Recibos Recibos " _
                    + "Where " _
                    + "CtaCte.Cliente = Cliente.Cliente AND " _
                    + "CtaCte.ClaveRecibo = Recibos.Clave AND " _
                    + "CtaCte.Tipo >= '01' AND " _
                    + "CtaCte.Tipo <= '05' AND " _
                    + "CtaCte.OrdFecha >= '00000000' AND " _
                    + "CtaCte.OrdFecha <= '99999999' AND " _
                    + "CtaCte.Importe8 <> 0"
                    
                    
            Uno = "{CtaCte.Tipo} in " + Chr$(34) + "01" + Chr$(34) + " to " + Chr$(34) + "05" + Chr$(34) + " and "
            Dos = "{CtaCte.Importe8} <> 0 and "
            Tres = "{CtaCte.OrdFecha} in " + Chr$(34) + "00000000" + Chr$(34) + " to " + Chr$(34) + "99999999" + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos + Tres
            Listado.SelectionFormula = Uno + Dos + Tres
                    
    
            Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
            Listado.Connect = Connect()
            Listado.ReportFileName = "WListIbVen.rpt"
    
            Listado.Action = 1
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "UPDATE Ctacte SET "
            ZSql = ZSql + " ImpoIbTucu = 0"
            ZSql = ZSql + " Where ImpoIbTucu IS NULL"
            spCtacte = ZSql
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
            With rstEmpresa
                .Index = "Empresa"
                .Seek "=", Val(WEmpresa)
                If .NoMatch = False Then
                    WAuxiliar = !Nombre
                End If
            End With
    
            WTitulo = "del " + Desde.Text + " al " + Hasta.Text
    
            With rstAuxiliar
                .Index = "Clave"
                .Seek "=", 1
                If .NoMatch = False Then
                    .Edit
                    !Nombre = WAuxiliar
                    !Varios = Left$(WTitulo, 50)
                    .Update
                End If
            End With
    
            Listado.WindowTitle = "Listado de Percepcion de Ingresos Brutos (Tucuman)"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
    
            If Impresora.Value = True Then
                Listado.Destination = 1
                    Else
                Listado.Destination = 0
            End If
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT CtaCte.Tipo, CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.Neto, CtaCte.Importe8, CtaCte.ImpoIbTucu, " _
                    + "Cliente.Razon, Cliente.Cuit " _
                    + "From " _
                    + DSQ + ".dbo.CtaCte CtaCte, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "CtaCte.Cliente = Cliente.Cliente AND " _
                    + "CtaCte.Tipo >= '01' AND " _
                    + "CtaCte.Tipo <= '05' AND " _
                    + "CtaCte.OrdFecha >= '" + WDesde + "' AND " _
                    + "CtaCte.OrdFecha <= '" + WHasta + "' AND " _
                    + "CtaCte.ImpoIbTucu <> 0"
                    
            Uno = "{CtaCte.Tipo} in " + Chr$(34) + "01" + Chr$(34) + " to " + Chr$(34) + "05" + Chr$(34) + " and "
            Dos = "{CtaCte.ImpoIbTucu} <> 0 and "
            Tres = "{CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos + Tres
            Listado.SelectionFormula = Uno + Dos + Tres
                    
            Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
            Listado.Connect = Connect()
            Listado.ReportFileName = "ListIbTucu.rpt"
    
            Listado.Action = 1
            
        Case 99
            ZSql = ""
            ZSql = ZSql + "UPDATE Ctacte SET "
            ZSql = ZSql + " ImpoIbCiudad = 0"
            ZSql = ZSql + " Where ImpoIbCiudad IS NULL"
            spCtacte = ZSql
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
            With rstEmpresa
                .Index = "Empresa"
                .Seek "=", Val(WEmpresa)
                If .NoMatch = False Then
                    WAuxiliar = !Nombre
                End If
            End With
    
            WTitulo = "del " + Desde.Text + " al " + Hasta.Text
    
            With rstAuxiliar
                .Index = "Clave"
                .Seek "=", 1
                If .NoMatch = False Then
                    .Edit
                    !Nombre = WAuxiliar
                    !Varios = Left$(WTitulo, 50)
                    .Update
                End If
            End With
    
            Listado.WindowTitle = "Listado de Percepcion de Ingresos Brutos (Ciudad)"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
    
            If Impresora.Value = True Then
                Listado.Destination = 1
                    Else
                Listado.Destination = 0
            End If
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT CtaCte.Tipo, CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.Neto, CtaCte.Importe8, CtaCte.ImpoIbCiudad, " _
                    + "Cliente.Razon, Cliente.Cuit " _
                    + "From " _
                    + DSQ + ".dbo.CtaCte CtaCte, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "CtaCte.Cliente = Cliente.Cliente AND " _
                    + "CtaCte.Tipo >= '01' AND " _
                    + "CtaCte.Tipo <= '05' AND " _
                    + "CtaCte.OrdFecha >= '" + WDesde + "' AND " _
                    + "CtaCte.OrdFecha <= '" + WHasta + "' AND " _
                    + "CtaCte.ImpoIbCiudad <> 0"
                    
            Uno = "{CtaCte.Tipo} in " + Chr$(34) + "01" + Chr$(34) + " to " + Chr$(34) + "05" + Chr$(34) + " and "
            Dos = "{CtaCte.ImpoIbCiudad} <> 0 and "
            Tres = "{CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos + Tres
            Listado.SelectionFormula = Uno + Dos + Tres
                    
            Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
            Listado.Connect = Connect()
            Listado.ReportFileName = "ListIbCiudad.rpt"
    
            Listado.Action = 1
            
        Case Else
            ZSql = ""
            ZSql = ZSql + "UPDATE Ctacte SET "
            ZSql = ZSql + " ImpoIbTucu = 0"
            ZSql = ZSql + " Where ImpoIbTucu IS NULL"
            spCtacte = ZSql
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Ctacte SET "
            ZSql = ZSql + " ImpoIbCiudad = 0"
            ZSql = ZSql + " Where ImpoIbCiudad IS NULL"
            spCtacte = ZSql
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            
            spCtacte = "ModificaCtacteImporteIva0"
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
            Rem Procesa las cobranzas
    
            Renglon = 0
            Erase Vector
    
            XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
            spRecibo = "ListaRecibosDifeI" + XParam
            Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibo.RecordCount > 0 Then
                With rstRecibo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstRecibo!Clave
                            Vector(Renglon, 2) = rstRecibo!Fecha
                            Vector(Renglon, 3) = rstRecibo!Tipo1
                            Vector(Renglon, 4) = rstRecibo!Numero1
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibo.Close
            End If
     
            For Cicla = 1 To Renglon
    
                WClave = Vector(Cicla, 1)
                WFecha = Vector(Cicla, 2)
                WTipo = Vector(Cicla, 3)
                WNumero = Vector(Cicla, 4)
        
                ClaveCtacte = WTipo + WNumero + "01"
                spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtacte.RecordCount > 0 Then
                    WImpoIb = IIf(IsNull(rstCtacte!ImpoIbCiudad), "0", rstCtacte!ImpoIbCiudad)
                    If WImpoIb = 0 Then
                        Vector(Cicla, 1) = ""
                        Vector(Cicla, 2) = ""
                        Vector(Cicla, 3) = ""
                        Vector(Cicla, 4) = ""
                    End If
                    rstCtacte.Close
                        Else
                    Vector(Cicla, 1) = ""
                    Vector(Cicla, 2) = ""
                    Vector(Cicla, 3) = ""
                    Vector(Cicla, 4) = ""
                End If
        
            Next Cicla
    
            For Cicla = 1 To Renglon
    
                WClave = Vector(Cicla, 1)
                If WClave <> "" Then
        
                    WTipo = Vector(Cicla, 3)
                    WNumero = Vector(Cicla, 4)
                    WRecibo = Val(Left$(WClave, 6))
                    WSale = "N"
        
                    XParam = "'" + WTipo + "','" _
                                 + WNumero + "'"
                    spRecibo = "ListaRecibosFactura " + XParam
                    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstRecibo.RecordCount > 0 Then
                        With rstRecibo
                            .MoveFirst
                            Do
                                If .EOF = False Then
                                    If Val(rstRecibo!Recibo) < WRecibo Then
                                        WSale = "S"
                                    End If
                                    .MoveNext
                                        Else
                                    Exit Do
                                End If
                            Loop
                        End With
                        rstRecibo.Close
                    End If
            
                    If WSale = "S" Then
                        Vector(Cicla, 1) = ""
                        Vector(Cicla, 2) = ""
                        Vector(Cicla, 3) = ""
                        Vector(Cicla, 4) = ""
                    End If
            
                End If
        
            Next Cicla
    
            For Cicla = 1 To Renglon
                
                WClave = Vector(Cicla, 1)
                If WClave <> "" Then
        
                    WTipo = Vector(Cicla, 3)
                    WNumero = Vector(Cicla, 4)
                    
                    ClaveCtacte = WTipo + WNumero + "01"
                    XParam = "'" + ClaveCtacte + "','" _
                                 + WClave + "'"
                    spCtacte = "ModificaCtacteIbCiudad " + XParam
                    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
                End If
        
            Next Cicla
        
            With rstEmpresa
                .Index = "Empresa"
                .Seek "=", Val(WEmpresa)
                If .NoMatch = False Then
                    WAuxiliar = !Nombre
                End If
            End With
    
            WTitulo = "del " + Desde.Text + " al " + Hasta.Text
    
            With rstAuxiliar
                .Index = "Clave"
                .Seek "=", 1
                If .NoMatch = False Then
                    .Edit
                    !Nombre = WAuxiliar
                    !Varios = Left$(WTitulo, 50)
                    .Update
                End If
            End With
    
            Listado.WindowTitle = "Listado de Percepcion de Ingresos Brutos (Ciudad)"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
    
            Rem Listado.GroupSelectionFormula = "{CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            If Impresora.Value = True Then
                Listado.Destination = 1
                    Else
                Listado.Destination = 0
            End If
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT CtaCte.Tipo, CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.Importe4, CtaCte.Importe8, " _
                    + "Cliente.Razon, Cliente.Cuit, " _
                    + "Recibos.Recibo " _
                    + "From " _
                    + DSQ + ".dbo.CtaCte CtaCte, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Recibos Recibos " _
                    + "Where " _
                    + "CtaCte.Cliente = Cliente.Cliente AND " _
                    + "CtaCte.ClaveRecibo = Recibos.Clave AND " _
                    + "CtaCte.Tipo >= '01' AND " _
                    + "CtaCte.Tipo <= '05' AND " _
                    + "CtaCte.OrdFecha >= '00000000' AND " _
                    + "CtaCte.OrdFecha <= '99999999' AND " _
                    + "CtaCte.Importe8 <> 0"
                    
                    
            Uno = "{CtaCte.Tipo} in " + Chr$(34) + "01" + Chr$(34) + " to " + Chr$(34) + "05" + Chr$(34) + " and "
            Dos = "{CtaCte.Importe8} <> 0 and "
            Tres = "{CtaCte.OrdFecha} in " + Chr$(34) + "00000000" + Chr$(34) + " to " + Chr$(34) + "99999999" + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos + Tres
            Listado.SelectionFormula = Uno + Dos + Tres
                    
    
            Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
            Listado.Connect = Connect()
            Listado.ReportFileName = "WListIbVENCiudad.rpt"
    
            Listado.Action = 1
            
            
            
        
    End Select
            
            
End Sub

Private Sub Cancela_click()
    Desde.SetFocus
    PrgListIbVen.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Buenos Aires"
    Tipo.AddItem "Tucuman"
    Tipo.AddItem "Ciudad"
    
    Tipo.ListIndex = 0

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

