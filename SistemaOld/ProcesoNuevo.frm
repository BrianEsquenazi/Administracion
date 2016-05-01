VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProcesoNueva 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Produccion"
   ClientHeight    =   7680
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7680
   ScaleWidth      =   11685
   Visible         =   0   'False
   Begin VB.TextBox ZTiempoI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   148
      Top             =   840
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   6720
   End
   Begin VB.TextBox EquipoProceso 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10440
      MaxLength       =   10
      TabIndex        =   89
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox HoraInicio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   86
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Teorico 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10440
      MaxLength       =   10
      TabIndex        =   84
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Operario 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   30
      Text            =   " "
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Hoja 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   28
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   9840
      TabIndex        =   27
      Top             =   120
      Width           =   1695
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   5280
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   9313
      _Version        =   327680
      Tabs            =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Procedimiento"
      TabPicture(0)   =   "ProcesoNuevo.frx":0000
      Tab(0).ControlCount=   6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "WTexto3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "WVector1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "WTitulo(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "WTexto2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "WCombo1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "WTexto1"
      Tab(0).Control(5).Enabled=   0   'False
      TabCaption(1)   =   "Controles de Calidad"
      TabPicture(1)   =   "ProcesoNuevo.frx":001C
      Tab(1).ControlCount=   5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WVector2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "WTexto32"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "WTexto12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "WTexto22"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "WCombo12"
      Tab(1).Control(4).Enabled=   0   'False
      TabCaption(2)   =   "Equipos y Metodos de Seguridad"
      TabPicture(2)   =   "ProcesoNuevo.frx":0038
      Tab(2).ControlCount=   10
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DesSeguridadIII"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "DesSeguridadII"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "Seguridad"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "DesSeguridadI"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "DesEquipoI"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "Equipo"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "DesEquipoII"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "DesEquipoIII"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "lblLabels(1)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "lblLabels(0)"
      Tab(2).Control(9).Enabled=   0   'False
      TabCaption(3)   =   "Carga de Lotes"
      TabPicture(3)   =   "ProcesoNuevo.frx":0054
      Tab(3).ControlCount=   43
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Descri1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Descri2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Descri3"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Descri4"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Descri5"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Descri6"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Descri7"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Descri8"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label16"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label12"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Mp8"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Pt8"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Mp7"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Pt7"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Mp6"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Pt6"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Mp5"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Pt5"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "Mp4"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "Pt4"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "Mp3"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "Pt3"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "Mp2"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "Pt2"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "Mp1"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "Pt1"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "Canti1"
      Tab(3).Control(26).Enabled=   -1  'True
      Tab(3).Control(27)=   "Tipo1"
      Tab(3).Control(27).Enabled=   -1  'True
      Tab(3).Control(28)=   "Canti2"
      Tab(3).Control(28).Enabled=   -1  'True
      Tab(3).Control(29)=   "Tipo2"
      Tab(3).Control(29).Enabled=   -1  'True
      Tab(3).Control(30)=   "Canti3"
      Tab(3).Control(30).Enabled=   -1  'True
      Tab(3).Control(31)=   "Tipo3"
      Tab(3).Control(31).Enabled=   -1  'True
      Tab(3).Control(32)=   "Canti4"
      Tab(3).Control(32).Enabled=   -1  'True
      Tab(3).Control(33)=   "Tipo4"
      Tab(3).Control(33).Enabled=   -1  'True
      Tab(3).Control(34)=   "Canti5"
      Tab(3).Control(34).Enabled=   -1  'True
      Tab(3).Control(35)=   "Tipo5"
      Tab(3).Control(35).Enabled=   -1  'True
      Tab(3).Control(36)=   "Canti6"
      Tab(3).Control(36).Enabled=   -1  'True
      Tab(3).Control(37)=   "Tipo6"
      Tab(3).Control(37).Enabled=   -1  'True
      Tab(3).Control(38)=   "Canti7"
      Tab(3).Control(38).Enabled=   -1  'True
      Tab(3).Control(39)=   "Tipo7"
      Tab(3).Control(39).Enabled=   -1  'True
      Tab(3).Control(40)=   "Canti8"
      Tab(3).Control(40).Enabled=   -1  'True
      Tab(3).Control(41)=   "Tipo8"
      Tab(3).Control(41).Enabled=   -1  'True
      Tab(3).Control(42)=   "CargaLote"
      Tab(3).Control(42).Enabled=   0   'False
      TabCaption(4)   =   "Observaciones"
      TabPicture(4)   =   "ProcesoNuevo.frx":0070
      Tab(4).ControlCount=   10
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ObservaI"
      Tab(4).Control(0).Enabled=   -1  'True
      Tab(4).Control(1)=   "ObservaII"
      Tab(4).Control(1).Enabled=   -1  'True
      Tab(4).Control(2)=   "ObservaIII"
      Tab(4).Control(2).Enabled=   -1  'True
      Tab(4).Control(3)=   "ObservaIV"
      Tab(4).Control(3).Enabled=   -1  'True
      Tab(4).Control(4)=   "ObservaV"
      Tab(4).Control(4).Enabled=   -1  'True
      Tab(4).Control(5)=   "ObservaVI"
      Tab(4).Control(5).Enabled=   -1  'True
      Tab(4).Control(6)=   "ObservaVII"
      Tab(4).Control(6).Enabled=   -1  'True
      Tab(4).Control(7)=   "ObservaVIII"
      Tab(4).Control(7).Enabled=   -1  'True
      Tab(4).Control(8)=   "ObservaIX"
      Tab(4).Control(8).Enabled=   -1  'True
      Tab(4).Control(9)=   "ObservaX"
      Tab(4).Control(9).Enabled=   -1  'True
      TabCaption(5)   =   "Controles de Proceso"
      TabPicture(5)   =   "ProcesoNuevo.frx":008C
      Tab(5).ControlCount=   49
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label3"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label4"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "lblLabels(2)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "lblLabels(3)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "lblLabels(4)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "lblLabels(5)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Label9"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "lblLabels(11)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "lblLabels(10)"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Label10"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "lblLabels(9)"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "lblLabels(8)"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "Label11"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "lblLabels(6)"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "lblLabels(7)"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "lblLabels(12)"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "lblLabels(13)"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "lblLabels(14)"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "lblLabels(15)"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "lblLabels(16)"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "lblLabels(17)"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "lblLabels(18)"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "lblLabels(19)"
      Tab(5).Control(22).Enabled=   0   'False
      Tab(5).Control(23)=   "ZTiempoII"
      Tab(5).Control(23).Enabled=   -1  'True
      Tab(5).Control(24)=   "Temperatura"
      Tab(5).Control(24).Enabled=   -1  'True
      Tab(5).Control(25)=   "ControlI"
      Tab(5).Control(25).Enabled=   -1  'True
      Tab(5).Control(26)=   "ControlII"
      Tab(5).Control(26).Enabled=   -1  'True
      Tab(5).Control(27)=   "DesdeI"
      Tab(5).Control(27).Enabled=   -1  'True
      Tab(5).Control(28)=   "HastaI"
      Tab(5).Control(28).Enabled=   -1  'True
      Tab(5).Control(29)=   "TiempoII"
      Tab(5).Control(29).Enabled=   -1  'True
      Tab(5).Control(30)=   "TiempoI"
      Tab(5).Control(30).Enabled=   -1  'True
      Tab(5).Control(31)=   "ControlV"
      Tab(5).Control(31).Enabled=   -1  'True
      Tab(5).Control(32)=   "DesdeV"
      Tab(5).Control(32).Enabled=   -1  'True
      Tab(5).Control(33)=   "HastaV"
      Tab(5).Control(33).Enabled=   -1  'True
      Tab(5).Control(34)=   "ControlIV"
      Tab(5).Control(34).Enabled=   -1  'True
      Tab(5).Control(35)=   "DesdeIV"
      Tab(5).Control(35).Enabled=   -1  'True
      Tab(5).Control(36)=   "HastaIV"
      Tab(5).Control(36).Enabled=   -1  'True
      Tab(5).Control(37)=   "ControlIII"
      Tab(5).Control(37).Enabled=   -1  'True
      Tab(5).Control(38)=   "DesdeIII"
      Tab(5).Control(38).Enabled=   -1  'True
      Tab(5).Control(39)=   "HastaIII"
      Tab(5).Control(39).Enabled=   -1  'True
      Tab(5).Control(40)=   "ValorV"
      Tab(5).Control(40).Enabled=   -1  'True
      Tab(5).Control(41)=   "ValorIV"
      Tab(5).Control(41).Enabled=   -1  'True
      Tab(5).Control(42)=   "ValorIII"
      Tab(5).Control(42).Enabled=   -1  'True
      Tab(5).Control(43)=   "WAlarma"
      Tab(5).Control(43).Enabled=   -1  'True
      Tab(5).Control(44)=   "WAlarmaI"
      Tab(5).Control(44).Enabled=   -1  'True
      Tab(5).Control(45)=   "WAlarmaII"
      Tab(5).Control(45).Enabled=   -1  'True
      Tab(5).Control(46)=   "ZTiempoIII"
      Tab(5).Control(46).Enabled=   -1  'True
      Tab(5).Control(47)=   "WAlarmaITempe"
      Tab(5).Control(47).Enabled=   -1  'True
      Tab(5).Control(48)=   "WAlarmaITiempo"
      Tab(5).Control(48).Enabled=   -1  'True
      Begin VB.TextBox WAlarmaITiempo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Height          =   285
         Left            =   -68160
         Locked          =   -1  'True
         TabIndex        =   152
         Top             =   4000
         Width           =   735
      End
      Begin VB.TextBox WAlarmaITempe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Height          =   285
         Left            =   -65520
         Locked          =   -1  'True
         TabIndex        =   151
         Top             =   4000
         Width           =   735
      End
      Begin VB.TextBox ZTiempoIII 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   -65520
         Locked          =   -1  'True
         TabIndex        =   150
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox WAlarmaII 
         Caption         =   "Alarma Tiempo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   -72840
         TabIndex        =   147
         Top             =   4440
         Width           =   3495
      End
      Begin VB.CheckBox WAlarmaI 
         Caption         =   "Alarma Temperatura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   -72840
         TabIndex        =   146
         Top             =   3960
         Width           =   3255
      End
      Begin VB.CheckBox WAlarma 
         Caption         =   "Alarma Rampa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   -72840
         TabIndex        =   142
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Frame CargaLote 
         Caption         =   "Ingreso de Partidas"
         Height          =   1815
         Left            =   -66600
         TabIndex        =   129
         Top             =   840
         Visible         =   0   'False
         Width           =   2655
         Begin VB.TextBox WLote1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   138
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox WLote2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   137
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox WLote3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   136
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox WCanti1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   135
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox WCanti2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   134
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox WCanti3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   133
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox WControl1 
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   132
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox WControl2 
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   131
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox WControl3 
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   130
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label dada 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Partida"
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   1200
            TabIndex        =   139
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox ValorIII 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64560
         TabIndex        =   125
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox ValorIV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64560
         TabIndex        =   124
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox ValorV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64560
         TabIndex        =   123
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox HastaIII 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66480
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox DesdeIII 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68400
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox ControlIII 
         Height          =   315
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox HastaIV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66480
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox DesdeIV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68400
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox ControlIV 
         Height          =   315
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox HastaV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66480
         Locked          =   -1  'True
         TabIndex        =   107
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox DesdeV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68400
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   2880
         Width           =   735
      End
      Begin VB.ComboBox ControlV 
         Height          =   315
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   2880
         Width           =   3255
      End
      Begin VB.TextBox TiempoI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66360
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TiempoII 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox HastaI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68160
         TabIndex        =   96
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox DesdeI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox ControlII 
         Height          =   315
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   1380
         Width           =   1695
      End
      Begin VB.ComboBox ControlI 
         Height          =   315
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Temperatura 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   -64680
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox ZTiempoII 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   -64680
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   1380
         Width           =   735
      End
      Begin VB.TextBox ObservaX 
         Height          =   285
         Left            =   -74640
         TabIndex        =   82
         Top             =   4080
         Width           =   6855
      End
      Begin VB.TextBox ObservaIX 
         Height          =   285
         Left            =   -74640
         TabIndex        =   81
         Top             =   3720
         Width           =   6855
      End
      Begin VB.TextBox ObservaVIII 
         Height          =   285
         Left            =   -74640
         TabIndex        =   80
         Top             =   3360
         Width           =   6855
      End
      Begin VB.TextBox ObservaVII 
         Height          =   285
         Left            =   -74640
         TabIndex        =   79
         Top             =   3000
         Width           =   6855
      End
      Begin VB.TextBox ObservaVI 
         Height          =   285
         Left            =   -74640
         TabIndex        =   78
         Top             =   2640
         Width           =   6855
      End
      Begin VB.TextBox ObservaV 
         Height          =   285
         Left            =   -74640
         TabIndex        =   77
         Top             =   2280
         Width           =   6855
      End
      Begin VB.TextBox ObservaIV 
         Height          =   285
         Left            =   -74640
         TabIndex        =   76
         Top             =   1920
         Width           =   6855
      End
      Begin VB.TextBox ObservaIII 
         Height          =   285
         Left            =   -74640
         TabIndex        =   75
         Top             =   1560
         Width           =   6855
      End
      Begin VB.TextBox ObservaII 
         Height          =   285
         Left            =   -74640
         TabIndex        =   74
         Top             =   1200
         Width           =   6855
      End
      Begin VB.TextBox ObservaI 
         Height          =   285
         Left            =   -74640
         TabIndex        =   73
         Top             =   840
         Width           =   6855
      End
      Begin VB.TextBox Tipo8 
         Height          =   285
         Left            =   -74760
         MaxLength       =   1
         TabIndex        =   70
         Text            =   " "
         Top             =   4320
         Width           =   495
      End
      Begin VB.TextBox Canti8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   68
         Text            =   " "
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox Tipo7 
         Height          =   285
         Left            =   -74760
         MaxLength       =   1
         TabIndex        =   65
         Text            =   " "
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox Canti7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   63
         Text            =   " "
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Tipo6 
         Height          =   285
         Left            =   -74760
         MaxLength       =   1
         TabIndex        =   60
         Text            =   " "
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Canti6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   58
         Text            =   " "
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Tipo5 
         Height          =   285
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   55
         Text            =   " "
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox Canti5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   53
         Text            =   " "
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Tipo4 
         Height          =   285
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   50
         Text            =   " "
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Canti4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   48
         Text            =   " "
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Tipo3 
         Height          =   285
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   45
         Text            =   " "
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Canti3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   43
         Text            =   " "
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Tipo2 
         Height          =   285
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   40
         Text            =   " "
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Canti2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   38
         Text            =   " "
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Tipo1 
         Height          =   285
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   35
         Text            =   " "
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Canti1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   33
         Text            =   " "
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox DesSeguridadIII 
         Height          =   285
         Left            =   -72720
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   25
         Top             =   2700
         Width           =   8775
      End
      Begin VB.TextBox DesSeguridadII 
         Height          =   285
         Left            =   -72720
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   24
         Top             =   2340
         Width           =   8775
      End
      Begin VB.TextBox Seguridad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   23
         Text            =   " "
         Top             =   1980
         Width           =   855
      End
      Begin VB.TextBox DesSeguridadI 
         Height          =   285
         Left            =   -72720
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   22
         Top             =   1980
         Width           =   8775
      End
      Begin VB.TextBox DesEquipoI 
         Height          =   285
         Left            =   -72720
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   20
         Top             =   800
         Width           =   8775
      End
      Begin VB.TextBox Equipo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   19
         Text            =   " "
         Top             =   800
         Width           =   855
      End
      Begin VB.TextBox DesEquipoII 
         Height          =   285
         Left            =   -72720
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   18
         Top             =   1150
         Width           =   8775
      End
      Begin VB.TextBox DesEquipoIII 
         Height          =   285
         Left            =   -72720
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   17
         Top             =   1500
         Width           =   8775
      End
      Begin VB.ComboBox WCombo12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71160
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   -72720
         TabIndex        =   13
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox WTexto12 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -73440
         TabIndex        =   12
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox WTexto1 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   1980
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Top             =   1980
         Width           =   375
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2520
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   7011
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16776960
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
      Begin MSMask.MaskEdBox WTexto32 
         Height          =   285
         Left            =   -72000
         TabIndex        =   15
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16776960
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
      Begin MSFlexGridLib.MSFlexGrid WVector2 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   16
         Top             =   900
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7011
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox Pt1 
         Height          =   285
         Left            =   -74280
         TabIndex        =   34
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox Mp1 
         Height          =   300
         Left            =   -72720
         TabIndex        =   36
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Pt2 
         Height          =   285
         Left            =   -74280
         TabIndex        =   39
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox Mp2 
         Height          =   300
         Left            =   -72720
         TabIndex        =   41
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Pt3 
         Height          =   285
         Left            =   -74280
         TabIndex        =   44
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox Mp3 
         Height          =   300
         Left            =   -72720
         TabIndex        =   46
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Pt4 
         Height          =   285
         Left            =   -74280
         TabIndex        =   49
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox Mp4 
         Height          =   300
         Left            =   -72720
         TabIndex        =   51
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Pt5 
         Height          =   285
         Left            =   -74280
         TabIndex        =   54
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox Mp5 
         Height          =   300
         Left            =   -72720
         TabIndex        =   56
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         Enabled         =   0   'False
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Pt6 
         Height          =   285
         Left            =   -74280
         TabIndex        =   59
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox Mp6 
         Height          =   300
         Left            =   -72720
         TabIndex        =   61
         Top             =   3600
         Width           =   1455
         _ExtentX        =   2566
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Pt7 
         Height          =   285
         Left            =   -74280
         TabIndex        =   64
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox Mp7 
         Height          =   300
         Left            =   -72720
         TabIndex        =   66
         Top             =   3960
         Width           =   1455
         _ExtentX        =   2566
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Pt8 
         Height          =   285
         Left            =   -74280
         TabIndex        =   69
         Top             =   4320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox Mp8 
         Height          =   300
         Left            =   -72720
         TabIndex        =   83
         Top             =   4320
         Width           =   1455
         _ExtentX        =   2566
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Label lblLabels 
         Caption         =   "Temperatura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   19
         Left            =   -67320
         TabIndex        =   154
         Top             =   4005
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tiempo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   18
         Left            =   -69240
         TabIndex        =   153
         Top             =   4000
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tiempo Rampa"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   17
         Left            =   -66480
         TabIndex        =   145
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Caption         =   "Temp."
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   16
         Left            =   -64680
         TabIndex        =   144
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tiempo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   15
         Left            =   -65520
         TabIndex        =   143
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MATERIALES"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -72480
         TabIndex        =   141
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Caption         =   "Valor"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   14
         Left            =   -65400
         TabIndex        =   128
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Valor"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   13
         Left            =   -65400
         TabIndex        =   127
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Valor"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   12
         Left            =   -65400
         TabIndex        =   126
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   -69240
         TabIndex        =   122
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Maximo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   -67320
         TabIndex        =   121
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Controla Presion"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   120
         Top             =   2205
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Maximo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   8
         Left            =   -67320
         TabIndex        =   119
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   9
         Left            =   -69240
         TabIndex        =   118
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Controla ??"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   117
         Top             =   2565
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Maximo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   10
         Left            =   -67320
         TabIndex        =   116
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   11
         Left            =   -69240
         TabIndex        =   115
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Controla ??"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   114
         Top             =   2925
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tiempo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   -70560
         TabIndex        =   104
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Durante"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   -67200
         TabIndex        =   103
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Maximo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   -69000
         TabIndex        =   102
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   -70560
         TabIndex        =   101
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Controla Rampa"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   100
         Top             =   1380
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Controla Temperatura"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   99
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ADICIONALES"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -72480
         TabIndex        =   72
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Descri8 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   -71280
         TabIndex        =   71
         Top             =   4320
         Width           =   3255
      End
      Begin VB.Label Descri7 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   -71280
         TabIndex        =   67
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   -71280
         TabIndex        =   62
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   -71280
         TabIndex        =   57
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   -71280
         TabIndex        =   52
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   -71280
         TabIndex        =   47
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label Descri2 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   -71280
         TabIndex        =   42
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   -71280
         TabIndex        =   37
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label lblLabels 
         Caption         =   "Seguridad"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   26
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Equipo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.TextBox Paso 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      MaxLength       =   4
      TabIndex        =   4
      Text            =   " "
      Top             =   120
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
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
   Begin MSMask.MaskEdBox FechaInicio 
      Height          =   285
      Left            =   2880
      TabIndex        =   87
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin VB.Label Label14 
      Caption         =   "Tiempo de la Etapa"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5520
      TabIndex        =   149
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Equipo"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8640
      TabIndex        =   90
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label23 
      Caption         =   "Fecha y Hora Inicio  Etapa"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   88
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "Cantidad"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8640
      TabIndex        =   85
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Operario"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   32
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label DesOperario 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   5520
      TabIndex        =   31
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Hoja de Produccion"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Etapa"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label DesTerminado 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   9840
      MouseIcon       =   "ProcesoNuevo.frx":00A8
      MousePointer    =   99  'Custom
      Picture         =   "ProcesoNuevo.frx":03B2
      ToolTipText     =   "Salida"
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7080
      MouseIcon       =   "ProcesoNuevo.frx":0BF4
      MousePointer    =   99  'Custom
      Picture         =   "ProcesoNuevo.frx":0EFE
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   8400
      MouseIcon       =   "ProcesoNuevo.frx":1740
      MousePointer    =   99  'Custom
      Picture         =   "ProcesoNuevo.frx":1A4A
      ToolTipText     =   "Consulta de Datos"
      Top             =   6840
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Producto"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgProcesoNueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOperarios As Recordset
Dim spOperarios As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnsayos As Recordset
Dim spEnsayos As String
Dim rstProduI As Recordset
Dim rsProduI As String
Dim rstProduII As Recordset
Dim rsProduII As String
Dim rstProduIII As Recordset
Dim rsProduIII As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Dim XPaso As String
Dim Renglon As Integer
Dim ZEnsayo As String
Dim ZValor As String
Dim ZItem As String
Dim ZDescri(100) As String
Dim ZTimer As Double

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Rem para el vector II

Dim WBorraII(1000, 20) As String
Dim WParametrosII(10, 20) As Double
Dim WFormatoII(20) As String
Dim WControlII As String
Dim ZLugar As Integer
Dim ZTipo As String
Dim ZArticulo As String
Dim ZTerminado As String
Dim ZCantidad As String
Dim XLote(10, 10) As String
Dim XSaldo1 As String
Dim XSaldo2 As String
Dim XSaldo3 As String
Dim WSaldo1 As Double
Dim WSaldo2 As Double
Dim WSaldo3 As Double
Dim WEstado As String


Rem para el vector III

Dim WBoraIII(1000, 20) As String
Dim WParametrosIII(10, 20) As Double
Dim WFormatoIII(20) As String
Dim WControlIII As String



Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"
     Opcion.AddItem "Ensayos"

     Opcion.Visible = True
     
End Sub



Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    Select Case XIndice
        Case 0, 2
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Order by Codigo"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTerminado!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTerminado.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Order by Codigo"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case 3
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 9, 10
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Order by Codigo"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                With rstEnsayos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEnsayos!Codigo) + " " + rstEnsayos!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstEnsayos!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnsayos.Close
            End If
    
            Call Conecta_Empresa
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose1_Click()

    PrgCargaNueva.Hide
    Unload Me
    If ZZOrigenProceso = 1 Then
        PrgConsultaHoja.Show
            Else
        PrgConsultaHojaTotal.Show
    End If
    
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector
    Call Limpia_VectorII
    
    Tablas.Tab = 0

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Paso.Text = ""
    
    Equipo.Text = ""
    DesEquipoI.Text = ""
    DesEquipoII.Text = ""
    DesEquipoIII.Text = ""
    
    Seguridad.Text = ""
    DesSeguridadI.Text = ""
    DesSeguridadII.Text = ""
    DesSeguridadIII.Text = ""
    
    DesdeI.Text = ""
    HastaI.Text = ""
    TiempoI.Text = ""
    TiempoII.Text = ""
    
    ControlI.ListIndex = 0
    ControlII.ListIndex = 0
    Tipo.ListIndex = 0
    
    Renglon = 0
    Graba.Enabled = True
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1
    
    ControlI.ListIndex = 0
    ControlII.ListIndex = 0
    
    Terminado.SetFocus

End Sub



Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Terminado.Text = WIndice.List(Indice)
            Call Terminado_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + WArticulo + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = Trim(rstArticulo!Codigo)
                WVector1.Col = 4
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstArticulo!Descripcion)
                End If
                WVector1.Col = 3
                rstArticulo.Close
                Call StartEdit
            End If
            Rem Ayuda.Visible = False
            
        Case 2
            Indice = Pantalla.ListIndex
            WPTerminado = WIndice.List(Indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Where Terminado.Codigo = " + "'" + WPTerminado + "'"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = Trim(rstTerminado!Codigo)
                WVector1.Col = 4
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstTerminado!Descripcion)
                End If
                WVector1.Col = 3
                rstTerminado.Close
                Call StartEdit
            End If
            Rem Ayuda.Visible = False
            
        Case 3
            Indice = Pantalla.ListIndex
            WEnsayo = WIndice.List(Indice)
            
            WTexto12.Visible = False
            WTexto22.Visible = False
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WEnsayo + "'"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                WVector2.Col = 1
                WVector2.Text = Trim(rstEnsayos!Codigo)
                WVector2.Col = 2
                WVector2.Text = Trim(rstEnsayos!Descripcion)
                WVector2.Col = 3
                rstEnsayos.Close
                Call StartEditII
            End If
            Rem Ayuda.Visible = False
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub


Private Sub Form_Activate()

    Call Limpia_Vector
    Call Limpia_VectorII
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Paso.Text = ""
    
    ControlI.Clear
    
    ControlI.AddItem "No Controla"
    ControlI.AddItem "Controla"
    
    ControlI.ListIndex = 0
    
    ControlII.Clear
    
    ControlII.AddItem "No Controla"
    ControlII.AddItem "Controla"
    
    ControlII.ListIndex = 0
    
    ControlIII.Clear
    
    ControlIII.AddItem "No Controla Presion"
    ControlIII.AddItem "Controla Presion"
    
    ControlIII.ListIndex = 0
    
    ControlIV.Clear
    
    ControlIV.AddItem "No Controla ??"
    ControlIV.AddItem "Controla ??"
    
    ControlIV.ListIndex = 0
    
    ControlV.Clear
    
    ControlV.AddItem "No Controla ??"
    ControlV.AddItem "Controla ??"
    
    ControlV.ListIndex = 0
    
    Tipo.Clear
    
    Tipo.AddItem "Fabricacion"
    Tipo.AddItem "Envasamiento"
    
    Tipo.ListIndex = 0
    
    Equipo.Text = ""
    DesEquipoI.Text = ""
    DesEquipoII.Text = ""
    DesEquipoIII.Text = ""
    
    Seguridad.Text = ""
    DesSeguridadI.Text = ""
    DesSeguridadII.Text = ""
    DesSeguridadIII.Text = ""
    
    DesdeI.Text = ""
    HastaI.Text = ""
    TiempoI.Text = ""
    TiempoII.Text = ""
    
    ControlI.ListIndex = 0
    ControlII.ListIndex = 0
    
    Tipo1.Text = ""
    Mp1.Text = "  -   -   "
    Pt1.Text = "  -     -   "
    Descri1.Caption = ""
    Canti1.Text = ""
    
    Tipo2.Text = ""
    Mp2.Text = "  -   -   "
    Pt2.Text = "  -     -   "
    Descri2.Caption = ""
    Canti2.Text = ""
    
    Tipo3.Text = ""
    Mp3.Text = "  -   -   "
    Pt3.Text = "  -     -   "
    Descri3.Caption = ""
    Canti3.Text = ""
    
    Tipo4.Text = ""
    Mp4.Text = "  -   -   "
    Pt4.Text = "  -     -   "
    Descri4.Caption = ""
    Canti4.Text = ""
    
    Tipo5.Text = ""
    Mp5.Text = "  -   -   "
    Pt5.Text = "  -     -   "
    Descri5.Caption = ""
    Canti5.Text = ""
    
    Tipo6.Text = ""
    Mp6.Text = "  -   -   "
    Pt6.Text = "  -     -   "
    Descri6.Caption = ""
    Canti6.Text = ""
    
    Tipo7.Text = ""
    Mp7.Text = "  -   -   "
    Pt7.Text = "  -     -   "
    Descri7.Caption = ""
    Canti7.Text = ""
    
    Tipo8.Text = ""
    Mp8.Text = "  -   -   "
    Pt8.Text = "  -     -   "
    Descri8.Caption = ""
    Canti8.Text = ""
    
    Renglon = 0
    
    Hoja.Text = ZHojaProceso
    Terminado.Text = ZTerminadoProceso
    Paso.Text = ZEtapaProceso
    Teorico.Text = ZCantidadProceso
    Operario.Text = ZOperarioProceso
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.Hoja = " + "'" + Hoja.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        FechaInicio.Text = rstHoja!FechaInicioEtapa
        HoraInicio.Text = rstHoja!HoraInicioEtapa
        Operario.Text = rstHoja!Operario
        EquipoProceso.Text = rstHoja!Equipo
        ZTimer = rstHoja!TimerInicioEtapa
        
        If rstHoja!alarma = "S" Then
            WAlarma.Value = 1
                Else
            WAlarma.Value = 0
        End If
        
        If rstHoja!AlarmaI = "S" Then
            WAlarmaI.Value = 1
                Else
            WAlarmaI.Value = 0
        End If
        
        If rstHoja!AlarmaII = "S" Then
            WAlarmaII.Value = 1
                Else
            WAlarmaII.Value = 0
        End If
        
        ZTiempoII.Text = IIf(IsNull(rstHoja!TiempoII), "0", rstHoja!TiempoII)
        WAlarmaITiempo.Text = IIf(IsNull(rstHoja!AlarmaITiempo), "", rstHoja!AlarmaITiempo)
        WAlarmaITempe.Text = IIf(IsNull(rstHoja!AlarmaITempe), "0", rstHoja!AlarmaITempe)
        
        rstHoja.Close
    End If
    
    ZTimeractual = Int(Timer)
    ZSegundos = ZTimeractual - ZTimer
    ZTiempoI.Text = Int(ZSegundos / 60)
    
    Sql1 = "Select *"
    Sql2 = " FROM Terminado"
    Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
    spTerminado = Sql1 + Sql2 + Sql3
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesTerminado.Caption = Trim(rstTerminado!Descripcion)
        rstTerminado.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM Operarios"
    Sql3 = " Where Operarios.Codigo = " + "'" + Operario.Text + "'"
    spOperarios = Sql1 + Sql2 + Sql3
    Set rstOperarios = db.OpenRecordset(spOperarios, dbOpenSnapshot, dbSQLPassThrough)
    If rstOperarios.RecordCount > 0 Then
        DesOperario.Caption = rstOperarios!Descripcion
        rstOperarios.Close
    End If
    
    Call Proceso_Click
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    Call Limpia_VectorII
    
    WRenglon = 0
    Erase ZDescri
    
    ZSql = " "
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ProduI"
    ZSql = ZSql + " Where ProduI.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " and ProduI.Etapa = " + "'" + Paso.Text + "'"
    ZSql = ZSql + " Order by ProduI.Clave"
    
    rsProduI = ZSql
    Set rstProduI = db.OpenRecordset(rsProduI, dbOpenSnapshot, dbSQLPassThrough)
    If rstProduI.RecordCount > 0 Then
        With rstProduI
            .MoveFirst
            Do
                If .EOF = False Then
                    WRenglon = WRenglon + 1
                    ZDescri(WRenglon) = Trim(rstProduI!instruccion)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProduI.Close
    End If
    
    
    
    
    For Ciclo = 1 To WRenglon
    
        WVector1.Row = Ciclo
        Renglon = Ciclo
                
        WVector1.Col = 1
        WVector1.Text = Trim(ZDescri(Ciclo))
                    
        Xlugar = 0
                    
        Do
            ZLugar = InStr(WVector1.Text, "{")
            If ZLugar = 0 Then Exit Do
                        
            If ZLugar > 0 Then
                        
                ZLugar1 = InStr(WVector1.Text, "}")
                If ZLugar1 = 0 Then Exit Do
                ZDife = ZLugar1 - ZLugar - 1
                            
                ZItem = Mid$(WVector1.Text, ZLugar + 1, ZDife)
                Call Ceros(ZItem, 2)
                            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Composicion"
                ZSql = ZSql + " Where Composicion.Terminado = " + "'" + Terminado.Text + "'"
                ZSql = ZSql + " and Composicion.Renglon = " + "'" + ZItem + "'"
                ZSql = ZSql + " Order by Composicion.Clave"
                spComposicion = ZSql
                Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                If rstComposicion.RecordCount > 0 Then
                    Xlugar = Xlugar + 1
                    Select Case Xlugar
                        Case 1
                            Tipo1.Text = rstComposicion!Tipo
                            Mp1.Text = rstComposicion!Articulo1
                            Pt1.Text = rstComposicion!Articulo2
                            Canti1.Text = Str$(rstComposicion!Cantidad * Val(Teorico.Text))
                            rstComposicion.Close
                                        
                            If Tipo1.Text = "T" Then
                                spTerminado = "ConsultaTerminado " + "'" + Pt1.Text + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
                                    Descri1.Caption = Trim(Left$(rstTerminado!Descripcion, 30))
                                    rstTerminado.Close
                                End If
                                    Else
                                spArticulo = "ConsultaArticulo " + "'" + Mp1.Text + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    Descri1.Caption = Trim(Left$(rstArticulo!Descripcion, 30))
                                    rstArticulo.Close
                                End If
                            End If
                                        
                            ZAuxi1 = Mid$(WVector1.Text, 1, ZLugar - 1)
                            ZAuxi2 = Mid$(WVector1.Text, ZLugar1 + 1, 100)
                                        
                            If Tipo1.Text = "M" Then
                                WVector1.Text = ZAuxi1 + Mp1.Text + ZAuxi2
                                    Else
                                WVector1.Text = ZAuxi1 + Pt1.Text + ZAuxi2
                            End If
                            
                        Case 2
                            Tipo2.Text = rstComposicion!Tipo
                            Mp2.Text = rstComposicion!Articulo1
                            Pt2.Text = rstComposicion!Articulo2
                            Canti2.Text = Str$(rstComposicion!Cantidad * Val(Teorico.Text))
                            rstComposicion.Close
                                        
                            If Tipo2.Text = "T" Then
                                spTerminado = "ConsultaTerminado " + "'" + Pt2.Text + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
                                    Descri2.Caption = Trim(Left$(rstTerminado!Descripcion, 30))
                                    rstTerminado.Close
                                End If
                                    Else
                                spArticulo = "ConsultaArticulo " + "'" + Mp2.Text + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    Descri2.Caption = Trim(Left$(rstArticulo!Descripcion, 30))
                                    rstArticulo.Close
                                End If
                            End If
                                        
                            ZAuxi1 = Mid$(WVector1.Text, 1, ZLugar - 1)
                            ZAuxi2 = Mid$(WVector1.Text, ZLugar1 + 1, 100)
                                        
                            If Tipo2.Text = "M" Then
                                WVector1.Text = ZAuxi1 + Mp2.Text + ZAuxi2
                                    Else
                                WVector1.Text = ZAuxi1 + Pt2.Text + ZAuxi2
                            End If
                            
                        Case 3
                            Tipo3.Text = rstComposicion!Tipo
                            Mp3.Text = rstComposicion!Articulo1
                            Pt3.Text = rstComposicion!Articulo2
                            Canti3.Text = Str$(rstComposicion!Cantidad * Val(Teorico.Text))
                            rstComposicion.Close
                                        
                            If Tipo3.Text = "T" Then
                                spTerminado = "ConsultaTerminado " + "'" + Pt3.Text + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
                                    Descri3.Caption = Trim(Left$(rstTerminado!Descripcion, 30))
                                    rstTerminado.Close
                                End If
                                    Else
                                spArticulo = "ConsultaArticulo " + "'" + Mp3.Text + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    Descri3.Caption = Trim(Left$(rstArticulo!Descripcion, 30))
                                    rstArticulo.Close
                                End If
                            End If
                                        
                            ZAuxi1 = Mid$(WVector1.Text, 1, ZLugar - 1)
                            ZAuxi2 = Mid$(WVector1.Text, ZLugar1 + 1, 100)
                                        
                            If Tipo3.Text = "M" Then
                                WVector1.Text = ZAuxi1 + Mp3.Text + ZAuxi2
                                    Else
                                WVector1.Text = ZAuxi1 + Pt3.Text + ZAuxi2
                            End If
                            
                        Case 4
                            Tipo4.Text = rstComposicion!Tipo
                            Mp4.Text = rstComposicion!Articulo1
                            Pt4.Text = rstComposicion!Articulo2
                            Canti4.Text = Str$(rstComposicion!Cantidad * Val(Teorico.Text))
                            rstComposicion.Close
                                        
                            If Tipo2.Text = "T" Then
                                spTerminado = "ConsultaTerminado " + "'" + Pt4.Text + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
                                    Descri4.Caption = Trim(Left$(rstTerminado!Descripcion, 30))
                                    rstTerminado.Close
                                End If
                                    Else
                                spArticulo = "ConsultaArticulo " + "'" + Mp4.Text + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    Descri4.Caption = Trim(Left$(rstArticulo!Descripcion, 30))
                                    rstArticulo.Close
                                End If
                            End If
                                        
                            ZAuxi1 = Mid$(WVector1.Text, 1, ZLugar - 1)
                            ZAuxi2 = Mid$(WVector1.Text, ZLugar1 + 1, 100)
                                        
                            If Tipo4.Text = "M" Then
                                WVector1.Text = ZAuxi1 + Mp4.Text + ZAuxi2
                                    Else
                                WVector1.Text = ZAuxi1 + Pt4.Text + ZAuxi2
                            End If
                            
                        Case 5
                            Tipo5.Text = rstComposicion!Tipo
                            Mp5.Text = rstComposicion!Articulo1
                            Pt5.Text = rstComposicion!Articulo2
                            Canti5.Text = Str$(rstComposicion!Cantidad * Val(Teorico.Text))
                            rstComposicion.Close
                                        
                            If Tipo5.Text = "T" Then
                                spTerminado = "ConsultaTerminado " + "'" + Pt5.Text + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
                                    Descri5.Caption = Trim(Left$(rstTerminado!Descripcion, 30))
                                    rstTerminado.Close
                                End If
                                    Else
                                spArticulo = "ConsultaArticulo " + "'" + Mp5.Text + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    Descri5.Caption = Trim(Left$(rstArticulo!Descripcion, 30))
                                    rstArticulo.Close
                                End If
                            End If
                                        
                            ZAuxi1 = Mid$(WVector1.Text, 1, ZLugar - 1)
                            ZAuxi2 = Mid$(WVector1.Text, ZLugar1 + 1, 100)
                                        
                            If Tipo5.Text = "M" Then
                                WVector1.Text = ZAuxi1 + Mp5.Text + ZAuxi2
                                    Else
                                WVector1.Text = ZAuxi1 + Pt5.Text + ZAuxi2
                            End If
                            
                        Case Else
                        
                    End Select
                End If
                            
            End If
        Loop
    Next Ciclo
    
    
    
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ProduII"
    ZSql = ZSql + " Where ProduII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " and ProduII.Etapa = " + "'" + Paso.Text + "'"
    ZSql = ZSql + " Order by ProduII.Clave"
    
    rsProduII = ZSql
    Set rstProduII = db.OpenRecordset(rsProduII, dbOpenSnapshot, dbSQLPassThrough)
    If rstProduII.RecordCount > 0 Then
        With rstProduII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    WVector2.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector2.Col = 1
                    WVector2.Text = Trim(Str$(rstProduII!Ensayo))
            
                    WVector2.Col = 2
                    WVector2.Text = ""
            
                    WVector2.Col = 3
                    WVector2.Text = Trim(rstProduII!Valor)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProduII.Close
    End If
          
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 9, 10
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    For Ciclo = 1 To WRenglon
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ensayos"
        ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WVector2.TextMatrix(Ciclo, 1) + "'"
        spEnsayos = ZSql
        Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayos.RecordCount > 0 Then
            WVector2.TextMatrix(Ciclo, 2) = Trim(rstEnsayos!Descripcion)
            rstEnsayos.Close
        End If
        
    Next Ciclo
    
    Call Conecta_Empresa
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ProduIII"
    ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " and ProduIII.Etapa = " + "'" + Paso.Text + "'"
    
    rsProduIII = ZSql
    Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstProduIII.RecordCount > 0 Then
    
        Equipo.Text = rstProduIII!Equipo
        Seguridad.Text = rstProduIII!Seguridad
        ControlI.ListIndex = rstProduIII!ControlI
        DesdeI.Text = rstProduIII!TemperaturaI
        HastaI.Text = rstProduIII!TemperaturaII
        TiempoI.Text = rstProduIII!Tiempo
        ControlII.ListIndex = rstProduIII!ControlII
        TiempoII.Text = rstProduIII!TiempoII
        DesEquipoI.Text = rstProduIII!DesEquipoI
        DesEquipoII.Text = rstProduIII!DesEquipoII
        DesEquipoIII.Text = rstProduIII!DesEquipoIII
        DesSeguridadI.Text = rstProduIII!DesSeguridadI
        DesSeguridadII.Text = rstProduIII!DesSeguridadII
        DesSeguridadIII.Text = rstProduIII!DesSeguridadIII
        Tipo.ListIndex = rstProduIII!Tipo
        ControlIII.ListIndex = rstProduIII!ControlIII
        DesdeIII.Text = rstProduIII!DesdeIII
        HastaIII.Text = rstProduIII!HastaIII
        ControlIV.ListIndex = rstProduIII!ControlIV
        DesdeIV.Text = rstProduIII!DesdeIV
        HastaIV.Text = rstProduIII!HastaIV
        ControlV.ListIndex = rstProduIII!ControlV
        DesdeV.Text = rstProduIII!DesdeV
        HastaV.Text = rstProduIII!HastaV
    
        rstProduIII.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM Terminado"
    Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
    spTerminado = Sql1 + Sql2 + Sql3
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesTerminado.Caption = Trim(rstTerminado!Descripcion)
        rstTerminado.Close
    End If
    
    Call Temperatura_Click
    
    Tablas.Tab = 0
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1
    
    Call StartEdit
    
    Graba.Enabled = True

End Sub

Private Sub Form_Load()
    Timer1.Interval = 30000
End Sub

Private Sub Graba_Click()

    T$ = "Grabacion de Datos"
    m$ = "Confirma la Finalizacion de la Etapa"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        Rem graba datos de la etapa
        
        Auxi1 = Hoja.Text
        Auxi2 = Paso.Text
        
        Call Ceros(Auxi1, 6)
        Call Ceros(Auxi2, 4)
        
        ZClave = Auxi1 + Auxi2
        ZFechaFinal = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZHoraFinal = Left$(Time$, 5)
        
        If WAlarma.Value = 1 Then
            ZAlarmaRampa = "S"
                Else
            ZAlarmaRampa = "N"
        End If
        
        If WAlarmaI.Value = 1 Then
            ZAlarmaTempe = "S"
                Else
            ZAlarmaTempe = "N"
        End If
        
        If WAlarmaII.Value = 1 Then
            ZAlarmaTiempo = "S"
                Else
            ZAlarmaTiempo = "N"
        End If
        
        ZTiempoII = ZTiempoII.Text
        
        
        
        
        
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO HojaEtapa ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Hoja ,"
        ZSql = ZSql + "Etapa ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Operario ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "FechaInicio ,"
        ZSql = ZSql + "HoraInicio ,"
        ZSql = ZSql + "FechaFinal ,"
        ZSql = ZSql + "HoraFinal ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "ResultadoI ,"
        ZSql = ZSql + "ResultadoII ,"
        ZSql = ZSql + "ResultadoIII ,"
        ZSql = ZSql + "ResultadoIV ,"
        ZSql = ZSql + "ResultadoV ,"
        ZSql = ZSql + "ResultadoVI ,"
        ZSql = ZSql + "ResultadoVII ,"
        ZSql = ZSql + "ResultadoVIII ,"
        ZSql = ZSql + "ResultadoIX ,"
        ZSql = ZSql + "ResultadoX ,"
        ZSql = ZSql + "ObservaI ,"
        ZSql = ZSql + "ObservaII ,"
        ZSql = ZSql + "ObservaIII ,"
        ZSql = ZSql + "ObservaIV ,"
        ZSql = ZSql + "ObservaV ,"
        ZSql = ZSql + "ObservaVI ,"
        ZSql = ZSql + "ObservaVII ,"
        ZSql = ZSql + "ObservaVIII ,"
        ZSql = ZSql + "ObservaIX ,"
        ZSql = ZSql + "ObservaX ,"
        ZSql = ZSql + "AlarmaTempe ,"
        ZSql = ZSql + "AlarmaRampa ,"
        ZSql = ZSql + "AlarmaTiempo ,"
        ZSql = ZSql + "AlarmaTempeTiempo ,"
        ZSql = ZSql + "AlarmaTempeTempe ,"
        ZSql = ZSql + "Temperatura ,"
        ZSql = ZSql + "Tiempo ,"
        ZSql = ZSql + "Rampa ,"
        ZSql = ZSql + "TiempoTempe ,"
        ZSql = ZSql + "ValorIII ,"
        ZSql = ZSql + "ValorIV ,"
        ZSql = ZSql + "ValorV )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZClave + "',"
        ZSql = ZSql + "'" + Hoja.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(Tipo.ListIndex) + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Operario.Text + "',"
        ZSql = ZSql + "'" + Teorico.Text + "',"
        ZSql = ZSql + "'" + FechaInicio.Text + "',"
        ZSql = ZSql + "'" + Trim(HoraInicio.Text) + "',"
        ZSql = ZSql + "'" + ZFechaFinal + "',"
        ZSql = ZSql + "'" + ZHoraFinal + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + WVector2.TextMatrix(1, 4) + "',"
        ZSql = ZSql + "'" + WVector2.TextMatrix(2, 4) + "',"
        ZSql = ZSql + "'" + WVector2.TextMatrix(3, 4) + "',"
        ZSql = ZSql + "'" + WVector2.TextMatrix(4, 4) + "',"
        ZSql = ZSql + "'" + WVector2.TextMatrix(5, 4) + "',"
        ZSql = ZSql + "'" + WVector2.TextMatrix(6, 4) + "',"
        ZSql = ZSql + "'" + WVector2.TextMatrix(7, 4) + "',"
        ZSql = ZSql + "'" + WVector2.TextMatrix(8, 4) + "',"
        ZSql = ZSql + "'" + WVector2.TextMatrix(9, 4) + "',"
        ZSql = ZSql + "'" + WVector2.TextMatrix(10, 4) + "',"
        ZSql = ZSql + "'" + ObservaI.Text + "',"
        ZSql = ZSql + "'" + ObservaII.Text + "',"
        ZSql = ZSql + "'" + ObservaIII.Text + "',"
        ZSql = ZSql + "'" + ObservaIV.Text + "',"
        ZSql = ZSql + "'" + ObservaV.Text + "',"
        ZSql = ZSql + "'" + ObservaVI.Text + "',"
        ZSql = ZSql + "'" + ObservaVII.Text + "',"
        ZSql = ZSql + "'" + ObservaVIII.Text + "',"
        ZSql = ZSql + "'" + ObservaIX.Text + "',"
        ZSql = ZSql + "'" + ObservaX.Text + "',"
        ZSql = ZSql + "'" + ZAlarmaTempe + "',"
        ZSql = ZSql + "'" + ZAlarmaRampa + "',"
        ZSql = ZSql + "'" + ZAlarmaTiempo + "',"
        ZSql = ZSql + "'" + WAlarmaITiempo.Text + "',"
        ZSql = ZSql + "'" + WAlarmaITempe.Text + "',"
        ZSql = ZSql + "'" + Temperatura.Text + "',"
        ZSql = ZSql + "'" + ZTiempoI.Text + "',"
        ZSql = ZSql + "'" + ZTiempoII.Text + "',"
        ZSql = ZSql + "'" + ZTiempoIII.Text + "',"
        ZSql = ZSql + "'" + ValorIII.Text + "',"
        ZSql = ZSql + "'" + ValorIV.Text + "',"
        ZSql = ZSql + "'" + ValorV.Text + "')"
        
        spHojaEtapa = ZSql
        Set rstHojaEtapa = db.OpenRecordset(spHojaEtapa, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
        
        
        
        
        XLote(1, 7) = Tipo1.Text
        XLote(1, 8) = Mp1.Text
        XLote(1, 9) = Pt1.Text
        XLote(1, 10) = Canti1.Text
        
        XLote(2, 7) = Tipo2.Text
        XLote(2, 8) = Mp2.Text
        XLote(2, 9) = Pt2.Text
        XLote(2, 10) = Canti2.Text
        
        XLote(3, 7) = Tipo3.Text
        XLote(3, 8) = Mp3.Text
        XLote(3, 9) = Pt3.Text
        XLote(3, 10) = Canti3.Text
        
        XLote(4, 7) = Tipo4.Text
        XLote(4, 8) = Mp4.Text
        XLote(4, 9) = Pt4.Text
        XLote(4, 10) = Canti4.Text
        
        XLote(5, 7) = Tipo5.Text
        XLote(5, 8) = Mp5.Text
        XLote(5, 9) = Pt5.Text
        XLote(5, 10) = Canti5.Text
        
        XLote(6, 7) = Tipo6.Text
        XLote(6, 8) = Mp6.Text
        XLote(6, 9) = Pt6.Text
        XLote(6, 10) = Canti6.Text
        
        XLote(7, 7) = Tipo7.Text
        XLote(7, 8) = Mp7.Text
        XLote(7, 9) = Pt7.Text
        XLote(7, 10) = Canti7.Text
        
        XLote(8, 7) = Tipo8.Text
        XLote(8, 8) = Mp8.Text
        XLote(8, 9) = Pt8.Text
        XLote(8, 10) = Canti8.Text
        
        For ZCiclo = 1 To 8
        
            If Val(XLote(ZCiclo, 10)) <> 0 Then
            
                Auxi1 = Hoja.Text
                Auxi2 = Paso.Text
                Auxi3 = Str$(ZCiclo)
        
                Call Ceros(Auxi1, 6)
                Call Ceros(Auxi2, 4)
                Call Ceros(Auxi3, 2)
        
                ZClave = Auxi1 + Auxi2 + Auxi3
            
        
                ZSql = ""
                ZSql = ZSql + "INSERT INTO HojaEtapaII ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Hoja ,"
                ZSql = ZSql + "Etapa ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Articulo ,"
                ZSql = ZSql + "Terminado ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Lote1 ,"
                ZSql = ZSql + "Cantidad1 ,"
                ZSql = ZSql + "Lote2 ,"
                ZSql = ZSql + "Cantidad2 ,"
                ZSql = ZSql + "Lote3 ,"
                ZSql = ZSql + "Cantidad3 )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZClave + "',"
                ZSql = ZSql + "'" + Hoja.Text + "',"
                ZSql = ZSql + "'" + Paso.Text + "',"
                ZSql = ZSql + "'" + Str$(ZCiclo) + "',"
                ZSql = ZSql + "'" + XLote(ZCiclo, 7) + "',"
                ZSql = ZSql + "'" + XLote(ZCiclo, 8) + "',"
                ZSql = ZSql + "'" + XLote(ZCiclo, 9) + "',"
                ZSql = ZSql + "'" + XLote(ZCiclo, 10) + "',"
                ZSql = ZSql + "'" + XLote(ZCiclo, 1) + "',"
                ZSql = ZSql + "'" + XLote(ZCiclo, 2) + "',"
                ZSql = ZSql + "'" + XLote(ZCiclo, 3) + "',"
                ZSql = ZSql + "'" + XLote(ZCiclo, 4) + "',"
                ZSql = ZSql + "'" + XLote(ZCiclo, 5) + "',"
                ZSql = ZSql + "'" + XLote(ZCiclo, 6) + "')"

                spHojaEtapaII = ZSql
                Set rstHojaEtapaII = db.OpenRecordset(spHojaEtapaII, dbOpenSnapshot, dbSQLPassThrough)
                    
            End If
        
        Next ZCiclo
    
        If Tipo.ListIndex = 0 Then
        
            ZEtapaProceso = Str$(Val((ZEtapaProceso) + 1))
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
            ZTimer = Int(Timer)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
            ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
            ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
            ZSql = ZSql + " Alarma = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
    
            ZTipo = 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM ProduIII"
            ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + Terminado.Text + "'"
            ZSql = ZSql + " and ProduIII.Etapa = " + "'" + Trim(ZEtapaProceso) + "'"
            rsProduIII = ZSql
            Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
            If rstProduIII.RecordCount > 0 Then
                ZTipo = rstProduIII!Tipo
                rstProduIII.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " TipoEtapa = " + "'" + ZTipo + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
            If ZTipo = 0 Then
                Call Form_Activate
                    Else
                Call cmdClose1_Click
            End If
            
                Else
            
            ZEtapaProceso = Str$(Val((ZEtapaProceso) + 1))
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " EstadoHoja = " + "'" + "2" + "',"
            ZSql = ZSql + " FechaFinal = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraFinal = " + "'" + ZHora + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
            Call cmdClose1_Click
            
        End If
        
    End If

End Sub

Private Sub Temperatura_Click()

    Select Case Trim(UCase(EquipoProceso.Text))
        Case "I", "1"
            OPEN_FILE_Temperatura0
            WFecha = "31/12/2100"
            WPasa = 0

            With rstTemperatura0
                .Index = "Fecha"
                .Seek "<=", WFecha
                Do
                    If .EOF = False Then
                        If WPasa = 0 Then
                            WPasa = 1
                            WFecha = !Fecha
                        End If
                        If !Fecha <> WFecha Then
                            Exit Do
                        End If
                        Temperatura.Text = !Valor
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
    
        Case "II", "2"
            OPEN_FILE_Temperatura1
            WFecha = "31/12/2100"
            WPasa = 0

            With rstTemperatura1
                .Index = "Fecha"
                .Seek "<=", WFecha
                Do
                    If .EOF = False Then
                        If WPasa = 0 Then
                            WPasa = 1
                            WFecha = !Fecha
                        End If
                        If !Fecha <> WFecha Then
                            Exit Do
                        End If
                        Temperatura.Text = !Valor
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case "III", "2"
            OPEN_FILE_Temperatura2
            WFecha = "31/12/2100"
            WPasa = 0

            With rstTemperatura2
                .Index = "Fecha"
                .Seek "<=", WFecha
                Do
                    If .EOF = False Then
                        If WPasa = 0 Then
                            WPasa = 1
                            WFecha = !Fecha
                        End If
                        If !Fecha <> WFecha Then
                            Exit Do
                        End If
                        Temperatura.Text = !Valor
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case "IV", "4"
            OPEN_FILE_Temperatura3
            WFecha = "31/12/2100"
            WPasa = 0

            With rstTemperatura3
                .Index = "Fecha"
                .Seek "<=", WFecha
                Do
                    If .EOF = False Then
                        If WPasa = 0 Then
                            WPasa = 1
                            WFecha = !Fecha
                        End If
                        If !Fecha <> WFecha Then
                            Exit Do
                        End If
                        Temperatura.Text = !Valor
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case "V", "5"
            OPEN_FILE_Temperatura4
            WFecha = "31/12/2100"
            WPasa = 0

            With rstTemperatura4
                .Index = "Fecha"
                .Seek "<=", WFecha
                Do
                    If .EOF = False Then
                        If WPasa = 0 Then
                            WPasa = 1
                            WFecha = !Fecha
                        End If
                        If !Fecha <> WFecha Then
                            Exit Do
                        End If
                        Temperatura.Text = !Valor
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case "VI", "6"
            OPEN_FILE_Temperatura5
            WFecha = "31/12/2100"
            WPasa = 0

            With rstTemperatura5
                .Index = "Fecha"
                .Seek "<=", WFecha
                Do
                    If .EOF = False Then
                        If WPasa = 0 Then
                            WPasa = 1
                            WFecha = !Fecha
                        End If
                        If !Fecha <> WFecha Then
                            Exit Do
                        End If
                        Temperatura.Text = !Valor
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case "VII", "7"
            OPEN_FILE_Temperatura6
            WFecha = "31/12/2100"
            WPasa = 0

            With rstTemperatura6
                .Index = "Fecha"
                .Seek "<=", WFecha
                Do
                    If .EOF = False Then
                        If WPasa = 0 Then
                            WPasa = 1
                            WFecha = !Fecha
                        End If
                        If !Fecha <> WFecha Then
                            Exit Do
                        End If
                        Temperatura.Text = !Valor
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case "VIII", "8"
            OPEN_FILE_Temperatura7
            WFecha = "31/12/2100"
            WPasa = 0

            With rstTemperatura7
                .Index = "Fecha"
                .Seek "<=", WFecha
                Do
                    If .EOF = False Then
                        If WPasa = 0 Then
                            WPasa = 1
                            WFecha = !Fecha
                        End If
                        If !Fecha <> WFecha Then
                            Exit Do
                        End If
                        Temperatura.Text = !Valor
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
    
        Case Else
    End Select
    
    ZTimeractual = Int(Timer)
    ZSegundos = ZTimeractual - ZTimer
    ZTiempoI.Text = Int(ZSegundos / 60)
    
    If ControlII.ListIndex = 1 And Val(ZTiempoI.Text) > Val(TiempoII.Text) Then
        If Val(Temperatura.Text) < Val(DesdeI.Text) Then
            If WAlarma.Value = 0 Then
                m$ = "Ha pasado el tiempo estimado para alacanzar la temperatura establecida en esta etapa"
                G% = MsgBox(m$, 16, "Carga de Procesos")
                WAlarma.Value = 1
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " Alarma = " + "'" + "S" + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    End If
    
    If Val(ZTiempoII.Text) = 0 Then
        If Val(Temperatura.Text) >= Val(DesdeI.Text) And Val(Temperatura.Text) < Val(HastaI.Text) Then
            ZTiempoII.Text = ZTiempoI.Text
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " TiempoII = " + "'" + ZTiempoII.Text + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        End If
    End If
    
    ZTiempoIII.Text = Str$(Val(ZTiempoI.Text) - Val(ZTiempoII.Text))
    
    If ControlI.ListIndex = 1 Then
        If Val(ZTiempoIII.Text) > Val(TiempoI.Text) Then
            If WAlarmaII.Value = 0 Then
                m$ = "Se ha excedido el tiempo establecido para la etapa"
                G% = MsgBox(m$, 16, "Carga de Procesos")
                WAlarmaII.Value = 1
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " AlarmaII = " + "'" + "S" + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    End If
    
    If ControlI.ListIndex = 1 And Val(ZTiempoII.Text) <> 0 Then
        If Val(Temperatura.Text) < Val(DesdeI.Text) Or Val(Temperatura.Text) > Val(HastaI.Text) Then
            If WAlarmaI.Value = 0 Then
                m$ = "La temperatura no se encuentra dentro del rango establecido"
                G% = MsgBox(m$, 16, "Carga de Procesos")
                WAlarmaI.Value = 1
                WAlarmaITiempo.Text = ZTiempoIII.Text
                WAlarmaITempe.Text = Temperatura.Text
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " AlarmaI = " + "'" + "S" + "',"
                ZSql = ZSql + " AlarmaITiempo = " + "'" + ZTiempoIII.Text + "',"
                ZSql = ZSql + " AlarmaITempe = " + "'" + Temperatura.Text + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    End If
    
End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM Terminado"
        Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
        spTerminado = Sql1 + Sql2 + Sql3
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = Trim(rstTerminado!Descripcion)
            rstTerminado.Close
            Paso.SetFocus
                Else
            Terminado.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Terminado.Text = ""
        DesTerminado.Caption = ""
    End If
End Sub

Private Sub Paso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Existe = "N"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ProduI"
        ZSql = ZSql + " Where ProduI.Terminado = " + "'" + Terminado.Text + "'"
        ZSql = ZSql + " and ProduI.Etapa = " + "'" + Paso.Text + "'"
        rsProduI = ZSql
        Set rstProduI = db.OpenRecordset(rsProduI, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduI.RecordCount > 0 Then
            rstProduI.Close
            Existe = "S"
        End If
        
        If Existe = "S" Then
            Call Proceso_Click
                Else
            Graba.Enabled = True
            WTerminado = Terminado.Text
            WPaso = Terminado.Text
            Terminado.Text = WTerminado
            Paso.Text = Paso
            Call Limpia_Vector
            Call Limpia_VectorII
            Tablas.Tab = 0
            WVector1.TopRow = 1
            WVector1.Col = 1
            WVector1.Row = 1
            WVector2.TopRow = 1
            WVector2.Col = 1
            WVector2.Row = 1
            Call StartEdit
        End If
        
    End If
    If KeyAscii = 27 Then
        Paso.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    Busqueda = Left$(Ayuda.Text, WEspacios)
    
    Select Case XIndice
        Case 0, 2
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Order by Codigo"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstTerminado!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstTerminado!Descripcion, aa, WEspacios) Then
                                    IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstTerminado!Codigo
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
                rstTerminado.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM MaterialAuxiliar"
            Sql3 = " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            Sql4 = " Order by Codigo"
            spMaterialAuxiliar = Sql1 + Sql2 + Sql3 + Sql4
            Set rstMaterialAuxiliar = db.OpenRecordset(spMaterialAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
            If rstMaterialAuxiliar.RecordCount > 0 Then
                With rstMaterialAuxiliar
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstMaterialAuxiliar.Close
            End If
            
        Case 3
            Rem XEmpresa = WEmpresa
            Rem Select Case Val(WEmpresa)
            Rem     Case 1, 3, 5, 6, 7, 9.10
            Rem         WEmpresa = "0003"
            Rem        txtOdbc = "Empresa03"
            Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Rem     Case Else
            Rem         WEmpresa = "0004"
            Rem         txtOdbc = "Empresa04"
            Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Rem End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Codigo"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                With rstEnsayos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnsayos.Close
            End If
            
            Rem Call Conecta_Empresa
            
        Case Else
    End Select
            
    End If

End Sub

Private Sub Terminado_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Productos Terminados"
    Opcion.AddItem "Material Auxiliar a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub Timer1_Timer()
    Call Temperatura_Click
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 1
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM articulo"
            Sql3 = " Where articulo.Codigo = " + "'" + WVector1.Text + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 4
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstArticulo!Descripcion)
                End If
                WVector1.Col = 3
                rstArticulo.Close
            End If
            
        Case 2
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Where Terminado.Codigo = " + "'" + WVector1.Text + "'"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WVector1.Col = 4
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstTerminado!Descripcion)
                End If
                WVector1.Col = 3
                rstTerminado.Close
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
    If WVector1.Col = 2 Then

    Opcion.Clear
    
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 2
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Intrucciones"
                WVector1.ColWidth(Ciclo) = 10000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 70
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
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
    
    Rem WVector1.Col = 1
    Rem WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub














Rem
Rem Controles de la WVector2
Rem

Private Sub GridEditTextII(ByVal KeyAscii As Integer)

    XColumna = WVector2.Col
    XTipoDato = WParametrosII(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto12.Left = WVector2.CellLeft + WVector2.Left
            WTexto12.Top = WVector2.CellTop + WVector2.Top
            WTexto12.Width = WVector2.CellWidth
            WTexto12.Height = WVector2.CellHeight
            WTexto12.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto12.Text = WVector2.Text
                    WTexto12.SelStart = Len(WTexto12.Text)
                Case Else
                    WTexto12.Text = Chr$(KeyAscii)
                    WTexto12.SelStart = 1
            End Select
            WTexto12.Visible = True
            WTexto12.SetFocus
        Case 1
            WTexto22.Left = WVector2.CellLeft + WVector2.Left
            WTexto22.Top = WVector2.CellTop + WVector2.Top
            WTexto22.Width = WVector2.CellWidth
            WTexto22.Height = WVector2.CellHeight
            WTexto22.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto22.Text = WVector2.Text
                    Rem WTexto22.SelStart = Len(WTexto22.Text)
                    WTexto22.SelStart = 0
                Case Else
                    WTexto22.Text = Chr$(KeyAscii)
                    WTexto22.SelStart = 1
            End Select
            WTexto22.Visible = True
            WTexto22.SetFocus
        Case 2
            WTexto32.Left = WVector2.CellLeft + WVector2.Left
            WTexto32.Top = WVector2.CellTop + WVector2.Top
            WTexto32.Width = WVector2.CellWidth
            WTexto32.Height = WVector2.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector2.Text) = 10 Then
                        WTexto32.Text = WVector2.Text
                            Else
                        WTexto32.Text = "  /  /    "
                    End If
                    WTexto32.SelStart = 0
                Case Else
                    WTexto32.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto32.SelStart = 1
            End Select
            WTexto32.Visible = True
            WTexto32.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditII()
    Pasa = 0
    If WCombo12.Visible Then
        Pasa = 0
        WVector2.Text = WCombo12.Text
        WCombo12.Visible = False
            Else
        If WTexto12.Visible Then
            Pasa = 1
            WVector2.Text = WTexto12.Text
            WTexto12.Visible = False
                Else
            If WTexto22.Visible Then
                Pasa = 1
                WVector2.Text = WTexto22.Text
                WTexto22.Visible = False
                    Else
                If WTexto32.Visible Then
                    Pasa = 1
                    WVector2.Text = WTexto32.Text
                    WTexto32.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoII(WVector2.Col) <> "" Then
            WVector2.Text = Pusing(WFormatoII(WVector2.Col), WVector2.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboII()
    ' Position the ComboBox over the cell.
    WCombo12.Left = WVector2.CellLeft + WVector2.Left
    WCombo12.Top = WVector2.CellTop + WVector2.Top
    WCombo12.Width = WVector2.CellWidth
    WCombo12.Visible = True
    WCombo12.SetFocus
End Sub

Private Sub WTexto12_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto12.Text = ""
            
        Rem F1
        Case 113
            WTexto12.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 123
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Col > 1 Then
                WVector2.Col = WVector2.Col - 1
            End If
            Call StartEditII

    End Select
End Sub

Private Sub WTexto22_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto22.Text = ""
            
        Rem F1
        Case 113
            WTexto22.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

Private Sub Wtexto32_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto32.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto32.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto12_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto22_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub Wtexto32_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo12_Click()
    WVector2.SetFocus
End Sub


Private Sub WVector2_Click()
    StartEditII
End Sub

Private Sub WVector2_LeaveCell()
    EndEditII
End Sub

Private Sub WVector2_GotFocus()
    EndEditII
End Sub

Private Sub WVector2_KeyPress(KeyAscii As Integer)
    XColumna = WVector2.Col
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
        Case Else
            If WParametrosII(2, XColumna) = 0 Then
                GridEditTextII KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditII()
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo12.Clear
            WCombo12.AddItem "Campo1"
            WCombo12.AddItem "Campo2"
            On Error Resume Next
            WCombo12.Text = WVector2.Text
            On Error GoTo 0
            GridEditComboII
        Case Else
            If WParametrosII(2, WVector2.Col) = 0 Then
                GridEditTextII Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVector2()
    Select Case WVector2.Col
        Case 4
            If WVector2.Row < WVector2.Rows - 1 Then
                WVector2.Row = WVector2.Row + 1
            End If
            WVector2.Col = 4
        Case Else
            If WVector2.Col < WVector2.Cols - 1 Then
                WVector2.Col = WVector2.Col + 1
            End If
    End Select
    WVector2.SetFocus
    GridEditTextII KeyAscii
End Sub

Private Sub Control_CampoII()
    XColumna = WVector2.Col
    XFila = WVector2.Row
    WControlII = "S"
    Select Case XColumna
        Case 1
            Rem XEmpresa = WEmpresa
            Rem Select Case Val(WEmpresa)
            Rem     Case 1, 3, 5, 6, 7, 9,10
            Rem         WEmpresa = "0003"
            Rem         txtOdbc = "Empresa03"
            Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Rem     Case Else
            Rem         WEmpresa = "0004"
            Rem         txtOdbc = "Empresa04"
            Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Rem End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WVector2.Text + "'"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                WVector2.Col = 2
                WVector2.Text = Trim(rstEnsayos!Descripcion)
                rstEnsayos.Close
                    Else
                WControlII = "N"
            End If
    
            Rem Call Conecta_Empresa
            
        Case 3, 6, 7
            Rem If Val(WVector2.Text) <> 0 Then
            Rem     ZCodigo = Val(WVector2.Text)
            Rem     Call Ceros(ZCodigo, 4)
            Rem
            Rem     Sql1 = "Select *"
            Rem     Sql2 = " FROM EquipoFabrica"
            Rem     Sql3 = " Where EquipoFabrica.Codigo = " + "'" + ZCodigo + "'"
            Rem     spEquipoFabrica = Sql1 + Sql2 + Sql3
            Rem     Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstEquipoFabrica.RecordCount > 0 Then
            Rem         rstEquipoFabrica.Close
            Rem     End If
            Rem End If
            
        Case Else
            WVector2.Col = XColumna
    End Select
End Sub

Private Sub WVector2_DblClick()

    If WVector2.Col = 0 Or WVector2.Col = 1 Then
    
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
    
    RenglonAuxiliar = WVector2.Row

    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WVector2.Text = ""
    Next Ciclo
    
    Erase WBorraII
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        Ensayo = WVector2.TextMatrix(iRow, 1)
            
        If Ensayo <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector2.Row = Ciclo
        WVector2.Col = 1
        WAuxi1 = WVector2.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector2.Cols - 1
                WVector2.Col = Ciclo1
                WBorraII(EntraVector, Ciclo1) = WVector2.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorII
    
    For Ciclo = 1 To EntraVector
        WVector2.Row = Ciclo
        For da = 0 To WVector2.Cols - 1
            WVector2.Col = da
            WVector2.Text = WBorraII(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub Limpia_VectorII()

    WVector2.Clear

    Rem ponga la WVector2 en negritas
    WVector2.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto12.FontName = WVector2.FontName
    WTexto12.FontSize = WVector2.FontSize
    WTexto12.Visible = False
    WTexto22.FontName = WVector2.FontName
    WTexto22.FontSize = WVector2.FontSize
    WTexto22.Visible = False
    WTexto32.FontName = WVector2.FontName
    WTexto32.FontSize = WVector2.FontSize
    WTexto32.Visible = False
    WCombo12.FontName = WVector2.FontName
    WCombo12.FontSize = WVector2.FontSize
    WCombo12.Visible = False

    ' Establesco loa Valores de la WVector2
    
    WVector2.FixedCols = 1
    WVector2.Cols = 5
    WVector2.FixedRows = 1
    WVector2.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector2.Text = "Articulo"
    
    Rem Longitud
    Rem WVector2.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosII(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosII(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosII(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosII(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Ensayos"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 4
                WParametrosII(2, Ciclo) = 1
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 2
                WVector2.Text = "Descripcion"
                WVector2.ColWidth(Ciclo) = 3000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 50
                WParametrosII(2, Ciclo) = 1
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 3
                WVector2.Text = "Valor"
                WVector2.ColWidth(Ciclo) = 3000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 50
                WParametrosII(2, Ciclo) = 1
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 4
                WVector2.Text = "Resultado"
                WVector2.ColWidth(Ciclo) = 3000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 50
                WParametrosII(2, Ciclo) = 1
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector2.Text
        Rem WTitulo(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        Rem WTitulo(Ciclo).Top = WVector2.CellTop + WVector2.Top
        Rem WTitulo(Ciclo).Width = WVector2.CellWidth
        Rem WTitulo(Ciclo).Height = WVector2.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector2
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 4
    WVector2.Row = 1
    
End Sub

Private Sub WVector2_Scroll()
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
End Sub
















Private Sub Equipo_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(Equipo.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM EquipoFabrica"
            Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Equipo.Text + "'"
            spEquipoFabrica = Sql1 + Sql2 + Sql3
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipoFabrica.RecordCount > 0 Then
                DesEquipoI.Text = Trim(rstEquipoFabrica!Descripcion)
                DesEquipoII.Text = Trim(rstEquipoFabrica!DescripcionII)
                DesEquipoIII.Text = Trim(rstEquipoFabrica!DescripcionIII)
                rstEquipoFabrica.Close
                Seguridad.SetFocus
                    Else
                Equipo.SetFocus
            End If
                Else
            DesEquipoI.SetFocus
        End If
    End If
    
    If KeyAscii = 27 Then
        Equipo.Text = ""
        DesEquipoI.Text = ""
        DesEquipoII.Text = ""
        DesEquipoIII.Text = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DesEquipoI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesEquipoII.SetFocus
    End If
    If KeyAscii = 27 Then
        DesEquipoI.Text = ""
    End If
End Sub

Private Sub DesEquipoII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesEquipoIII.SetFocus
    End If
    If KeyAscii = 27 Then
        DesEquipoII.Text = ""
    End If
End Sub

Private Sub DesEquipoIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Seguridad.SetFocus
    End If
    If KeyAscii = 27 Then
        DesEquipoIII.Text = ""
    End If
End Sub

Private Sub Seguridad_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(Seguridad.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM EquipoFabrica"
            Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Seguridad.Text + "'"
            spEquipoFabrica = Sql1 + Sql2 + Sql3
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipoFabrica.RecordCount > 0 Then
                DesSeguridadI.Text = Trim(rstEquipoFabrica!Descripcion)
                DesSeguridadII.Text = Trim(rstEquipoFabrica!DescripcionII)
                DesSeguridadIII.Text = Trim(rstEquipoFabrica!DescripcionIII)
                rstEquipoFabrica.Close
                ControlI.SetFocus
                    Else
                Seguridad.SetFocus
            End If
                Else
            DesSeguridadI.SetFocus
        End If
    End If
    
    If KeyAscii = 27 Then
        Seguridad.Text = ""
        DesSeguridadI.Text = ""
        DesSeguridadII.Text = ""
        DesSeguridadIII.Text = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DesSeguridadI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesSeguridadII.SetFocus
    End If
    If KeyAscii = 27 Then
        DesSeguridadI.Text = ""
    End If
End Sub

Private Sub DesSeguridadII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesSeguridadIII.SetFocus
    End If
    If KeyAscii = 27 Then
        DesSeguridadII.Text = ""
    End If
End Sub

Private Sub DesSeguridadIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ControlI.SetFocus
    End If
    If KeyAscii = 27 Then
        DesSeguridadIII.Text = ""
    End If
End Sub








Private Sub Tablas_Click(PreviousTab As Integer)
    
    Select Case Tablas.Tab
        Case 0
            Rem WVector1.Col = 1
            Rem WVector1.Row = 1
            Rem Call StartEdit
        Case 1
            Rem WVector2.Col = 4
            Rem WVector2.Row = 1
            Rem call StartEditII
        Case 2
            Equipo.SetFocus
        Case 4
            ObservaI.SetFocus
        Case Else
    End Select
End Sub

Private Sub ObservaI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaII.SetFocus
    End If
End Sub

Private Sub ObservaII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaIII.SetFocus
    End If
End Sub

Private Sub ObservaIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaIV.SetFocus
    End If
End Sub

Private Sub ObservaIV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaV.SetFocus
    End If
End Sub

Private Sub ObservaV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaVI.SetFocus
    End If
End Sub

Private Sub ObservaVI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaVII.SetFocus
    End If
End Sub

Private Sub ObservaVII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaVIII.SetFocus
    End If
End Sub

Private Sub ObservaVIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaIX.SetFocus
    End If
End Sub

Private Sub ObservaIX_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaX.SetFocus
    End If
End Sub

Private Sub ObservaX_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaI.SetFocus
    End If
End Sub

Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub



Private Sub Inicio_Carga()

    CargaLote.Visible = True
    If ZTipo = "M" Then
        CargaLote.Caption = "Ingreso de Lote"
        dada.Caption = "Lote"
            Else
        CargaLote.Caption = "Ingreso de Partida"
        dada.Caption = "Partida"
    End If
    
    WLote1.Text = ""
    WCanti1.Text = ""
    WLote2.Text = ""
    WCanti2.Text = ""
    WLote3.Text = ""
    WCanti3.Text = ""
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True
        
    If Val(XLote(ZLugar, 1)) <> 0 Then
        WLote1.Text = XLote(ZLugar, 1)
        WCanti1.Text = XLote(ZLugar, 2)
        WControl1.Locked = False
        WControl1.Text = ""
        WControl1.Locked = True
    End If
        
    If Val(XLote(ZLugar, 3)) <> 0 Then
        WLote2.Text = XLote(ZLugar, 3)
        WCanti2.Text = XLote(ZLugar, 4)
        WControl2.Locked = False
        WControl2.Text = ""
        WControl2.Locked = True
    End If
        
    If Val(XLote(ZLugar, 5)) <> 0 Then
        WLote3.Text = XLote(ZLugar, 5)
        WCanti3.Text = XLote(ZLugar, 6)
        WControl3.Locked = False
        WControl3.Text = ""
        WControl3.Locked = True
    End If
        
    WLote1.SetFocus

End Sub

Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If ZTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            ZArticulo = UCase(ZArticulo)
            spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                            + ZArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZArticulo + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
            
            If Val(WLote1.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    XLote(ZLugar, 1) = WLote1.Text
                    XLote(ZLugar, 2) = WCanti1.Text
                    XLote(ZLugar, 3) = WLote2.Text
                    XLote(ZLugar, 4) = WCanti2.Text
                    XLote(ZLugar, 5) = WLote3.Text
                    XLote(ZLugar, 6) = WCanti3.Text
                    Select Case ZLugar
                        Case 1
                            Canti2.SetFocus
                        Case 2
                            Canti3.SetFocus
                        Case 3
                            Canti4.SetFocus
                        Case 4
                            Canti5.SetFocus
                        Case 5
                            Tipo6.SetFocus
                        Case 6
                            Tipo7.SetFocus
                        Case 7
                            Tipo8.SetFocus
                        Case 8
                            Tipo9.SetFocus
                        Case 9
                            Tipo10.SetFocus
                        Case Else
                            Canti1.SetFocus
                    End Select
                    CargaLote.Visible = False
                    Exit Sub
                        Else
                    WLote1.SetFocus
                    Exit Sub
                End If
            End If
            
            If WEntra = "S" Then
                WCanti1.SetFocus
                    Else
                m$ = ZArticulo + " Articulo inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
            
                XParam = "'" + Str$(Int(Val(WLote1.Text))) + "','" _
                             + ZTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZTerminado + "','" _
                            + Str$(Int(Val(WLote1.Text))) + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote1.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    XLote(ZLugar, 1) = WLote1.Text
                    XLote(ZLugar, 2) = WCanti1.Text
                    XLote(ZLugar, 3) = WLote2.Text
                    XLote(ZLugar, 4) = WCanti2.Text
                    XLote(ZLugar, 5) = WLote3.Text
                    XLote(ZLugar, 6) = WCanti3.Text
                    Select Case ZLugar
                        Case 1
                            Canti2.SetFocus
                        Case 2
                            Canti3.SetFocus
                        Case 3
                            Canti4.SetFocus
                        Case 4
                            Canti5.SetFocus
                        Case 5
                            Tipo6.SetFocus
                        Case 6
                            Tipo7.SetFocus
                        Case 7
                            Tipo8.SetFocus
                        Case 8
                            Tipo9.SetFocus
                        Case 9
                            Tipo10.SetFocus
                        Case Else
                            Canti1.SetFocus
                    End Select
                    CargaLote.Visible = False
                    Exit Sub
                        Else
                    WLote1.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti1.SetFocus
                    Else
                m$ = ZTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WSaldo1 = 0
        If ZTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            ZArticulo = UCase(ZArticulo)
            spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                            + ZArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZArticulo + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = ZArticulo + " Articulo inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                        + ZTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    wdada = rstHoja!Hoja
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZTerminado + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = ZTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
        If WSaldo1 >= Val(WCanti1.Text) Or WControla > 0 Then
            WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
            WControl1.Locked = False
            WControl1.Text = "X"
            WControl1.Locked = True
            WLote2.SetFocus
                Else
            XSaldo1 = WSaldo1
            XSaldo1 = Pusing("###,###.##", XSaldo1)
            If ZTipo = "M" Then
                m$ = ZArticulo + " Cantidad Insuficiente Stock : " + XSaldo1
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    Else
                m$ = ZTerminado + " Cantidad Insuficiente Stock : " + XSaldo1
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            WLote1.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If ZTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            ZArticulo = UCase(ZArticulo)
            spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                            + ZArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZArticulo + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
            
            If Val(WLote2.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    XLote(ZLugar, 1) = WLote1.Text
                    XLote(ZLugar, 2) = WCanti1.Text
                    XLote(ZLugar, 3) = WLote2.Text
                    XLote(ZLugar, 4) = WCanti2.Text
                    XLote(ZLugar, 5) = WLote3.Text
                    XLote(ZLugar, 6) = WCanti3.Text
                    Select Case ZLugar
                        Case 1
                            Canti2.SetFocus
                        Case 2
                            Canti3.SetFocus
                        Case 3
                            Canti4.SetFocus
                        Case 4
                            Canti5.SetFocus
                        Case 5
                            Tipo6.SetFocus
                        Case 6
                            Tipo7.SetFocus
                        Case 7
                            Tipo8.SetFocus
                        Case 8
                            Tipo9.SetFocus
                        Case 9
                            Tipo10.SetFocus
                        Case Else
                            Canti1.SetFocus
                    End Select
                    CargaLote.Visible = False
                    Exit Sub
                        Else
                    WLote2.SetFocus
                    Exit Sub
                End If
            End If
            
            If WEntra = "S" Then
                WCanti2.SetFocus
                    Else
                m$ = ZArticulo + " Articulo inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
            
                XParam = "'" + Str$(Int(Val(WLote2.Text))) + "','" _
                             + ZTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZTerminado + "','" _
                            + Str$(Int(Val(WLote2.Text))) + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote2.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    XLote(ZLugar, 1) = WLote1.Text
                    XLote(ZLugar, 2) = WCanti1.Text
                    XLote(ZLugar, 3) = WLote2.Text
                    XLote(ZLugar, 4) = WCanti2.Text
                    XLote(ZLugar, 5) = WLote3.Text
                    XLote(ZLugar, 6) = WCanti3.Text
                    Select Case ZLugar
                        Case 1
                            Canti2.SetFocus
                        Case 2
                            Canti3.SetFocus
                        Case 3
                            Canti4.SetFocus
                        Case 4
                            Canti5.SetFocus
                        Case 5
                            Tipo6.SetFocus
                        Case 6
                            Tipo7.SetFocus
                        Case 7
                            Tipo8.SetFocus
                        Case 8
                            Tipo9.SetFocus
                        Case 9
                            Tipo10.SetFocus
                        Case Else
                            Canti1.SetFocus
                    End Select
                    CargaLote.Visible = False
                    Exit Sub
                        Else
                    WLote2.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti2.SetFocus
                    Else
                m$ = ZTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WSaldo2 = 0
        If ZTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            ZArticulo = UCase(ZArticulo)
            spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                            + ZArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZArticulo + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = ZArticulo + " Articulo inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                        + ZTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    wdada = rstHoja!Hoja
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZTerminado + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = ZTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
        If WSaldo2 >= Val(WCanti2.Text) Or WControla > 0 Then
            WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
            WControl2.Locked = False
            WControl2.Text = "X"
            WControl2.Locked = True
            WLote3.SetFocus
                Else
            XSaldo2 = WSaldo2
            XSaldo2 = Pusing("###,###.##", XSaldo2)
            If ZTipo = "M" Then
                m$ = ZArticulo + " Cantidad Insuficiente Stock : " + XSaldo2
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    Else
                m$ = ZTerminado + " Cantidad Insuficiente Stock : " + XSaldo2
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            WLote2.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If ZTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            ZArticulo = UCase(ZArticulo)
            spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                            + ZArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZArticulo + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
            
            If Val(WLote3.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    XLote(ZLugar, 1) = WLote1.Text
                    XLote(ZLugar, 2) = WCanti1.Text
                    XLote(ZLugar, 3) = WLote2.Text
                    XLote(ZLugar, 4) = WCanti2.Text
                    XLote(ZLugar, 5) = WLote3.Text
                    XLote(ZLugar, 6) = WCanti3.Text
                    Select Case ZLugar
                        Case 1
                            Canti2.SetFocus
                        Case 2
                            Canti3.SetFocus
                        Case 3
                            Canti4.SetFocus
                        Case 4
                            Canti5.SetFocus
                        Case 5
                            Tipo6.SetFocus
                        Case 6
                            Tipo7.SetFocus
                        Case 7
                            Tipo8.SetFocus
                        Case 8
                            Tipo9.SetFocus
                        Case 9
                            Tipo10.SetFocus
                        Case Else
                            Canti1.SetFocus
                    End Select
                    CargaLote.Visible = False
                    Exit Sub
                        Else
                    WLote3.SetFocus
                    Exit Sub
                End If
            End If
            
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = ZArticulo + " Articulo inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
            
                XParam = "'" + Str$(Int(Val(WLote3.Text))) + "','" _
                             + ZTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZTerminado + "','" _
                            + Str$(Int(Val(WLote3.Text))) + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote3.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    XLote(ZLugar, 1) = WLote1.Text
                    XLote(ZLugar, 2) = WCanti1.Text
                    XLote(ZLugar, 3) = WLote2.Text
                    XLote(ZLugar, 4) = WCanti2.Text
                    XLote(ZLugar, 5) = WLote3.Text
                    XLote(ZLugar, 6) = WCanti3.Text
                    Select Case ZLugar
                        Case 1
                            Canti2.SetFocus
                        Case 2
                            Canti3.SetFocus
                        Case 3
                            Canti4.SetFocus
                        Case 4
                            Canti5.SetFocus
                        Case 5
                            Tipo6.SetFocus
                        Case 6
                            Tipo7.SetFocus
                        Case 7
                            Tipo8.SetFocus
                        Case 8
                            Tipo9.SetFocus
                        Case 9
                            Tipo10.SetFocus
                        Case Else
                            Canti1.SetFocus
                    End Select
                    CargaLote.Visible = False
                    Exit Sub
                        Else
                    WLote3.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = ZTerminado + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WSaldo3 = 0
        If ZTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            ZArticulo = UCase(ZArticulo)
            spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                            + ZArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZArticulo + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = ZArticulo + " Articulo inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                        + ZTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    wdada = rstHoja!Hoja
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZTerminado + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = ZTerminado + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
        If WSaldo3 >= Val(WCanti3.Text) Or WControla > 0 Then
            WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
            WControl3.Locked = False
            WControl3.Text = "X"
            WControl3.Locked = True
            
            Call Verifica_Lote
            If WEstado = "S" Then
                XLote(ZLugar, 1) = WLote1.Text
                XLote(ZLugar, 2) = WCanti1.Text
                XLote(ZLugar, 3) = WLote2.Text
                XLote(ZLugar, 4) = WCanti2.Text
                XLote(ZLugar, 5) = WLote3.Text
                XLote(ZLugar, 6) = WCanti3.Text
                Select Case ZLugar
                    Case 1
                        Canti2.SetFocus
                    Case 2
                        Canti3.SetFocus
                    Case 3
                        Canti4.SetFocus
                    Case 4
                        Canti5.SetFocus
                    Case 5
                        Tipo6.SetFocus
                    Case 6
                        Tipo7.SetFocus
                    Case 7
                        Tipo8.SetFocus
                    Case 8
                        Tipo9.SetFocus
                    Case 9
                        Tipo10.SetFocus
                    Case Else
                        Canti1.SetFocus
                End Select
                CargaLote.Visible = False
                Exit Sub
                    Else
                WLote1.SetFocus
                Exit Sub
            End If
                    
                Else
                
            XSaldo3 = WSaldo3
            XSaldo3 = Pusing("###,###.##", XSaldo3)
            If ZTipo = "M" Then
                m$ = ZArticulo + " Cantidad Insuficiente Stock : " + XSaldo3
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    Else
                m$ = ZTerminado + " Cantidad Insuficiente Stock : " + XSaldo3
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            WLote3.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Verifica_Lote()

    WEstado = "N"
    Suma = 0
    
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True

    
    WSaldo1 = 0
    WSaldo2 = 0
    WSaldo3 = 0
    
    If Val(WLote1.Text) <> 0 Then
        If ZTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            ZArticulo = UCase(ZArticulo)
            spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
            
                XParam = "'" + WLote1.Text + "','" _
                            + ZArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZArticulo + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = ZArticulo + " Articulo inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                        + ZTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZTerminado + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = ZTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
        
        If WSaldo1 >= Val(WCanti1.Text) Then
            WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
            WControl1.Locked = False
            WControl1.Text = "X"
            WControl1.Locked = True
        End If
        
    End If
    
    If Val(WLote2.Text) <> 0 Then
        If ZTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            ZArticulo = UCase(ZArticulo)
            spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                            + ZArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZArticulo + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = ZArticulo + " Articulo inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                        + ZTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZTerminado + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = ZTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
            
        If WSaldo2 >= Val(WCanti2.Text) Then
            WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
            WControl2.Locked = False
            WControl2.Text = "X"
            WControl2.Locked = True
        End If
        
    End If
    
    
    If Val(WLote3.Text) <> 0 Then
        If ZTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            ZArticulo = UCase(ZArticulo)
            spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                            + ZArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZArticulo + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = ZArticulo + " Articulo inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                        + ZTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + ZTerminado + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = ZTerminado + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
        
        If WSaldo3 >= Val(WCanti3.Text) Then
            WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
            WControl3.Locked = False
            WControl3.Text = "X"
            WControl3.Locked = True
        End If
        
    End If
    
    If Val(WLote1.Text) <> 0 And WControl1.Text = "X" Then
        Suma = Suma + Val(WCanti1.Text)
    End If
    If Val(WLote2.Text) <> 0 And WControl2.Text = "X" Then
        Suma = Suma + Val(WCanti2.Text)
    End If
    If Val(WLote3.Text) <> 0 And WControl3.Text = "X" Then
        Suma = Suma + Val(WCanti3.Text)
    End If
    
    If Suma = Val(ZCantidad) Then
        WEstado = "S"
    End If
    
    If WControla <> 0 Then
        WEstado = "S"
    End If
    
End Sub


























Private Sub Canti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Canti1.Text = Pusing("###,###.##", Canti1.Text)
        Call Canti1_DblClick
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti1_DblClick()
    
    If Tipo1.Text = "M" Then
        ZTipo = Tipo1.Text
        ZArticulo = Mp1.Text
        ZCantidad = Canti1.Text
        If ZArticulo <> "  -   -   " And Val(ZCantidad) <> 0 Then
            ZLugar = 1
            Call Inicio_Carga
        End If
            Else
        If Tipo1.Text = "T" Then
            ZTipo = Tipo1.Text
            ZTerminado = Pt1.Text
            ZCantidad = Canti1.Text
            If ZTerminado <> "  -   -   " And Val(ZCantidad) <> 0 Then
                ZLugar = 1
                Call Inicio_Carga
            End If
        End If
    End If

End Sub

Private Sub Canti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Canti2.Text = Pusing("###,###.##", Canti2.Text)
        Call Canti2_DblClick
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti2_DblClick()
    
    If Tipo2.Text = "M" Then
        ZTipo = Tipo2.Text
        ZArticulo = Mp2.Text
        ZCantidad = Canti2.Text
        If ZArticulo <> "  -   -   " And Val(ZCantidad) <> 0 Then
            ZLugar = 2
            Call Inicio_Carga
        End If
            Else
        If Tipo2.Text = "T" Then
            ZTipo = Tipo2.Text
            ZTerminado = Pt2.Text
            ZCantidad = Canti2.Text
            If ZTerminado <> "  -   -   " And Val(ZCantidad) <> 0 Then
                ZLugar = 2
                Call Inicio_Carga
            End If
        End If
    End If

End Sub

Private Sub Canti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Canti3.Text = Pusing("###,###.##", Canti3.Text)
        Call Canti3_DblClick
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti3_DblClick()
    
    If Tipo3.Text = "M" Then
        ZTipo = Tipo3.Text
        ZArticulo = Mp3.Text
        ZCantidad = Canti3.Text
        If ZArticulo <> "  -   -   " And Val(ZCantidad) <> 0 Then
            ZLugar = 3
            Call Inicio_Carga
        End If
            Else
        If Tipo3.Text = "T" Then
            ZTipo = Tipo3.Text
            ZTerminado = Pt3.Text
            ZCantidad = Canti3.Text
            If ZTerminado <> "  -   -   " And Val(ZCantidad) <> 0 Then
                ZLugar = 3
                Call Inicio_Carga
            End If
        End If
    End If

End Sub

Private Sub Canti4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Canti4.Text = Pusing("###,###.##", Canti4.Text)
        Call Canti4_DblClick
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti4_DblClick()
    
    If Tipo4.Text = "M" Then
        ZTipo = Tipo4.Text
        ZArticulo = Mp4.Text
        ZCantidad = Canti4.Text
        If ZArticulo <> "  -   -   " And Val(ZCantidad) <> 0 Then
            ZLugar = 4
            Call Inicio_Carga
        End If
            Else
        If Tipo4.Text = "T" Then
            ZTipo = Tipo4.Text
            ZTerminado = Pt4.Text
            ZCantidad = Canti4.Text
            If ZTerminado <> "  -   -   " And Val(ZCantidad) <> 0 Then
                ZLugar = 4
                Call Inicio_Carga
            End If
        End If
    End If

End Sub

Private Sub Canti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Canti5.Text = Pusing("###,###.##", Canti5.Text)
        Call Canti5_DblClick
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti5_DblClick()
    
    If Tipo5.Text = "M" Then
        ZTipo = Tipo5.Text
        ZArticulo = Mp5.Text
        ZCantidad = Canti5.Text
        If ZArticulo <> "  -   -   " And Val(ZCantidad) <> 0 Then
            ZLugar = 5
            Call Inicio_Carga
        End If
            Else
        If Tipo5.Text = "T" Then
            ZTipo = Tipo5.Text
            ZTerminado = Pt5.Text
            ZCantidad = Canti5.Text
            If ZTerminado <> "  -   -   " And Val(ZCantidad) <> 0 Then
                ZLugar = 5
                Call Inicio_Carga
            End If
        End If
    End If

End Sub

Private Sub Canti6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Canti6.Text = Pusing("###,###.##", Canti6.Text)
        Call Canti6_DblClick
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti6_DblClick()
    
    If Tipo6.Text = "M" Then
        ZTipo = Tipo6.Text
        ZArticulo = Mp6.Text
        ZCantidad = Canti6.Text
        If ZArticulo <> "  -   -   " And Val(ZCantidad) <> 0 Then
            ZLugar = 6
            Call Inicio_Carga
        End If
            Else
        If Tipo6.Text = "T" Then
            ZTipo = Tipo6.Text
            ZTerminado = Pt6.Text
            ZCantidad = Canti6.Text
            If ZTerminado <> "  -   -   " And Val(ZCantidad) <> 0 Then
                ZLugar = 6
                Call Inicio_Carga
            End If
        End If
    End If

End Sub

Private Sub Canti7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Canti7.Text = Pusing("###,###.##", Canti7.Text)
        Call Canti7_DblClick
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti7_DblClick()
    
    If Tipo7.Text = "M" Then
        ZTipo = Tipo7.Text
        ZArticulo = Mp7.Text
        ZCantidad = Canti7.Text
        If ZArticulo <> "  -   -   " And Val(ZCantidad) <> 0 Then
            ZLugar = 7
            Call Inicio_Carga
        End If
            Else
        If Tipo7.Text = "T" Then
            ZTipo = Tipo7.Text
            ZTerminado = Pt7.Text
            ZCantidad = Canti7.Text
            If ZTerminado <> "  -   -   " And Val(ZCantidad) <> 0 Then
                ZLugar = 7
                Call Inicio_Carga
            End If
        End If
    End If

End Sub

Private Sub Canti8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Canti8.Text = Pusing("###,###.##", Canti8.Text)
        Call Canti8_DblClick
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti8_DblClick()
    
    If Tipo8.Text = "M" Then
        ZTipo = Tipo8.Text
        ZArticulo = Mp8.Text
        ZCantidad = Canti8.Text
        If ZArticulo <> "  -   -   " And Val(ZCantidad) <> 0 Then
            ZLugar = 8
            Call Inicio_Carga
        End If
            Else
        If Tipo8.Text = "T" Then
            ZTipo = Tipo8.Text
            ZTerminado = Pt8.Text
            ZCantidad = Canti8.Text
            If ZTerminado <> "  -   -   " And Val(ZCantidad) <> 0 Then
                ZLugar = 8
                Call Inicio_Carga
            End If
        End If
    End If

End Sub





Private Sub Tipo6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Tipo6.Text = "M" Or Tipo6.Text = "T" Then
            If Tipo6.Text = "M" Then
                Pt6.Text = "  -     -   "
                Mp6.SetFocus
                    Else
                Mp6.Text = "  -   -   "
                Pt6.SetFocus
            End If
                Else
            Tipo6.SetFocus
        End If
    End If
End Sub

Private Sub Pt6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pt6.Text = UCase(Pt6.Text)
        spTerminado = "ConsultaTerminado " + "'" + Pt6.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            Descri6.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            Canti6.SetFocus
                Else
            Pt6.SetFocus
        End If
    End If
End Sub

Private Sub Mp6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Mp6.Text = UCase(Mp6.Text)
        spArticulo = "ConsultaArticulo " + "'" + Mp6.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Descri6.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            Canti6.SetFocus
                Else
            Mp6.SetFocus
        End If
    End If
End Sub

Private Sub Tipo7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Tipo7.Text = "M" Or Tipo7.Text = "T" Then
            If Tipo7.Text = "M" Then
                Pt7.Text = "  -     -   "
                Mp7.SetFocus
                    Else
                Mp7.Text = "  -   -   "
                Pt7.SetFocus
            End If
                Else
            Tipo7.SetFocus
        End If
    End If
End Sub

Private Sub Pt7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pt7.Text = UCase(Pt7.Text)
        spTerminado = "ConsultaTerminado " + "'" + Pt7.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            Descri7.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            Canti7.SetFocus
                Else
            Pt7.SetFocus
        End If
    End If
End Sub

Private Sub Mp7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Mp7.Text = UCase(Mp7.Text)
        spArticulo = "ConsultaArticulo " + "'" + Mp7.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Descri7.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            Canti7.SetFocus
                Else
            Mp7.SetFocus
        End If
    End If
End Sub

Private Sub Tipo8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Tipo8.Text = "M" Or Tipo8.Text = "T" Then
            If Tipo8.Text = "M" Then
                Pt8.Text = "  -     -   "
                Mp8.SetFocus
                    Else
                Mp8.Text = "  -   -   "
                Pt8.SetFocus
            End If
                Else
            Tipo8.SetFocus
        End If
    End If
End Sub

Private Sub Pt8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pt8.Text = UCase(Pt8.Text)
        spTerminado = "ConsultaTerminado " + "'" + Pt8.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            Descri8.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            Canti8.SetFocus
                Else
            Pt8.SetFocus
        End If
    End If
End Sub

Private Sub Mp8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Mp8.Text = UCase(Mp8.Text)
        spArticulo = "ConsultaArticulo " + "'" + Mp8.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Descri8.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            Canti8.SetFocus
                Else
            Mp8.SetFocus
        End If
    End If
End Sub






