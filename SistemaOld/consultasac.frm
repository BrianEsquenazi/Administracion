VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsultaSac 
   Caption         =   "Consulta de SAC"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11775
   Begin VB.TextBox Referencia 
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
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   174
      Text            =   " "
      Top             =   1200
      Width           =   10455
   End
   Begin VB.TextBox Numero 
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
      Left            =   6120
      MaxLength       =   6
      TabIndex        =   163
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Centro 
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
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   162
      Text            =   " "
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Estado 
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
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   161
      Top             =   480
      Width           =   2775
   End
   Begin VB.ComboBox Origen 
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
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   160
      Top             =   480
      Width           =   2535
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
      Left            =   4320
      MaxLength       =   6
      TabIndex        =   159
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Tipo 
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
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox ResponsableEmisor 
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
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   142
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox ResponsableDestino 
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
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   141
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Titulo 
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
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   140
      Text            =   " "
      Top             =   1560
      Width           =   10455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11640
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   9128
      _Version        =   327680
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Descripcion de la NC"
      TabPicture(0)   =   "consultasac.frx":0000
      Tab(0).ControlCount=   4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "IngresoCausa"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "IngresoNoCon"
      Tab(0).Control(3).Enabled=   0   'False
      TabCaption(1)   =   "Acciones"
      TabPicture(1)   =   "consultasac.frx":001C
      Tab(1).ControlCount=   39
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DesResponsable1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DesResponsable2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DesResponsable3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "DesResponsable4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DesResponsable5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "DesResponsable6"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label31"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label30"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label29"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label28"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label27"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label26"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Plazo6"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Plazo5"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Plazo4"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Plazo3"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Plazo2"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Plazo1"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Accion11"
      Tab(1).Control(21).Enabled=   -1  'True
      Tab(1).Control(22)=   "Accion12"
      Tab(1).Control(22).Enabled=   -1  'True
      Tab(1).Control(23)=   "Accion21"
      Tab(1).Control(23).Enabled=   -1  'True
      Tab(1).Control(24)=   "Accion22"
      Tab(1).Control(24).Enabled=   -1  'True
      Tab(1).Control(25)=   "Accion31"
      Tab(1).Control(25).Enabled=   -1  'True
      Tab(1).Control(26)=   "Accion32"
      Tab(1).Control(26).Enabled=   -1  'True
      Tab(1).Control(27)=   "Accion41"
      Tab(1).Control(27).Enabled=   -1  'True
      Tab(1).Control(28)=   "Accion42"
      Tab(1).Control(28).Enabled=   -1  'True
      Tab(1).Control(29)=   "Accion51"
      Tab(1).Control(29).Enabled=   -1  'True
      Tab(1).Control(30)=   "Accion52"
      Tab(1).Control(30).Enabled=   -1  'True
      Tab(1).Control(31)=   "Accion61"
      Tab(1).Control(31).Enabled=   -1  'True
      Tab(1).Control(32)=   "Accion62"
      Tab(1).Control(32).Enabled=   -1  'True
      Tab(1).Control(33)=   "Responsable1"
      Tab(1).Control(33).Enabled=   -1  'True
      Tab(1).Control(34)=   "Responsable2"
      Tab(1).Control(34).Enabled=   -1  'True
      Tab(1).Control(35)=   "Responsable3"
      Tab(1).Control(35).Enabled=   -1  'True
      Tab(1).Control(36)=   "Responsable4"
      Tab(1).Control(36).Enabled=   -1  'True
      Tab(1).Control(37)=   "Responsable5"
      Tab(1).Control(37).Enabled=   -1  'True
      Tab(1).Control(38)=   "Responsable6"
      Tab(1).Control(38).Enabled=   -1  'True
      TabCaption(2)   =   "Implementacion"
      TabPicture(2)   =   "consultasac.frx":0038
      Tab(2).ControlCount=   59
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DesResponsable11"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label17"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label18"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label20"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "DesResponsable12"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "DesResponsable13"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "DesResponsable14"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "DesResponsable15"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "DesResponsable16"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label24"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label32"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label33"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label34"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label35"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label36"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label37"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Fecha6"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Fecha5"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Fecha4"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Fecha3"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Fecha2"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Fecha1"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Accion111"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Accion112"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Accion121"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Accion122"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Accion131"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Accion132"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Accion141"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Accion142"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Accion151"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Accion152"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Accion161"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Accion162"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Responsable11"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Comentario62"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Comentario61"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Comentario52"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Comentario51"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Comentario42"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Comentario41"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Comentario32"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Comentario31"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Comentario22"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Comentario21"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "Comentario12"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "Comentario11"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "Responsable12"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "Responsable13"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "Responsable14"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "Responsable16"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "Responsable15"
      Tab(2).Control(52).Enabled=   0   'False
      Tab(2).Control(53)=   "Estado16"
      Tab(2).Control(53).Enabled=   0   'False
      Tab(2).Control(54)=   "Estado15"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).Control(55)=   "Estado14"
      Tab(2).Control(55).Enabled=   0   'False
      Tab(2).Control(56)=   "Estado13"
      Tab(2).Control(56).Enabled=   0   'False
      Tab(2).Control(57)=   "Estado12"
      Tab(2).Control(57).Enabled=   0   'False
      Tab(2).Control(58)=   "Estado11"
      Tab(2).Control(58).Enabled=   0   'False
      TabCaption(3)   =   "Verificacion"
      TabPicture(3)   =   "consultasac.frx":0054
      Tab(3).ControlCount=   58
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label8"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label21"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "DesResponsable21"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label23"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "DesResponsable22"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "DesResponsable23"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "DesResponsable24"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "DesResponsable25"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "DesResponsable26"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label38"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label39"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label40"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label41"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Label42"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Label43"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Fecha26"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Fecha25"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "Fecha24"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "Fecha23"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "Fecha22"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "Fecha21"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "Responsable21"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "Accion262"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "Accion261"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "Accion252"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "Accion251"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "Accion242"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "Accion241"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "Accion232"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "Accion231"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "Accion222"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "Accion221"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "Accion212"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "Accion211"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "Comentario262"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).Control(36)=   "Comentario261"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "Comentario252"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "Comentario251"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "Comentario242"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).Control(40)=   "Comentario241"
      Tab(3).Control(40).Enabled=   0   'False
      Tab(3).Control(41)=   "Comentario232"
      Tab(3).Control(41).Enabled=   0   'False
      Tab(3).Control(42)=   "Comentario231"
      Tab(3).Control(42).Enabled=   0   'False
      Tab(3).Control(43)=   "Comentario222"
      Tab(3).Control(43).Enabled=   0   'False
      Tab(3).Control(44)=   "Comentario221"
      Tab(3).Control(44).Enabled=   0   'False
      Tab(3).Control(45)=   "Comentario212"
      Tab(3).Control(45).Enabled=   0   'False
      Tab(3).Control(46)=   "Comentario211"
      Tab(3).Control(46).Enabled=   0   'False
      Tab(3).Control(47)=   "Responsable22"
      Tab(3).Control(47).Enabled=   0   'False
      Tab(3).Control(48)=   "Responsable23"
      Tab(3).Control(48).Enabled=   0   'False
      Tab(3).Control(49)=   "Responsable24"
      Tab(3).Control(49).Enabled=   0   'False
      Tab(3).Control(50)=   "Responsable25"
      Tab(3).Control(50).Enabled=   0   'False
      Tab(3).Control(51)=   "Responsable26"
      Tab(3).Control(51).Enabled=   0   'False
      Tab(3).Control(52)=   "Estado1"
      Tab(3).Control(52).Enabled=   0   'False
      Tab(3).Control(53)=   "Estado2"
      Tab(3).Control(53).Enabled=   0   'False
      Tab(3).Control(54)=   "Estado3"
      Tab(3).Control(54).Enabled=   0   'False
      Tab(3).Control(55)=   "Estado4"
      Tab(3).Control(55).Enabled=   0   'False
      Tab(3).Control(56)=   "Estado5"
      Tab(3).Control(56).Enabled=   0   'False
      Tab(3).Control(57)=   "Estado6"
      Tab(3).Control(57).Enabled=   0   'False
      TabCaption(4)   =   "Comentario"
      TabPicture(4)   =   "consultasac.frx":0070
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Comentario"
      Tab(4).Control(0).Enabled=   0   'False
      Begin VB.TextBox Comentario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   188
         Top             =   480
         Width           =   11175
      End
      Begin VB.ComboBox Estado11 
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
         Left            =   -68160
         Locked          =   -1  'True
         TabIndex        =   157
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Estado12 
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
         Left            =   -68160
         Locked          =   -1  'True
         TabIndex        =   156
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox Estado13 
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
         Left            =   -68160
         Locked          =   -1  'True
         TabIndex        =   155
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox Estado14 
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
         Left            =   -68160
         Locked          =   -1  'True
         TabIndex        =   154
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox Estado15 
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
         Left            =   -68160
         Locked          =   -1  'True
         TabIndex        =   153
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox Estado16 
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
         Left            =   -68160
         Locked          =   -1  'True
         TabIndex        =   152
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox IngresoNoCon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   149
         Top             =   840
         Width           =   11175
      End
      Begin VB.TextBox IngresoCausa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   148
         Top             =   3120
         Width           =   11175
      End
      Begin VB.ComboBox Estado6 
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
         Left            =   -68880
         TabIndex        =   139
         Top             =   4320
         Width           =   1335
      End
      Begin VB.ComboBox Estado5 
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
         Left            =   -68880
         TabIndex        =   138
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox Estado4 
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
         Left            =   -68880
         TabIndex        =   137
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox Estado3 
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
         Left            =   -68880
         TabIndex        =   136
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox Estado2 
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
         Left            =   -68880
         TabIndex        =   135
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox Estado1 
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
         Left            =   -68880
         TabIndex        =   134
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Responsable26 
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   131
         Text            =   " "
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox Responsable25 
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   128
         Text            =   " "
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Responsable24 
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   125
         Text            =   " "
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Responsable23 
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   122
         Text            =   " "
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Responsable22 
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   119
         Text            =   " "
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Comentario211 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   118
         Text            =   " "
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox Comentario212 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   117
         Text            =   " "
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Comentario221 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   116
         Text            =   " "
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Comentario222 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   115
         Text            =   " "
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox Comentario231 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   114
         Text            =   " "
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox Comentario232 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   113
         Text            =   " "
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox Comentario241 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   112
         Text            =   " "
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox Comentario242 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   111
         Text            =   " "
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox Comentario251 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   110
         Text            =   " "
         Top             =   3600
         Width           =   3615
      End
      Begin VB.TextBox Comentario252 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   109
         Text            =   " "
         Top             =   3840
         Width           =   3615
      End
      Begin VB.TextBox Comentario261 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   108
         Text            =   " "
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox Comentario262 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   107
         Text            =   " "
         Top             =   4560
         Width           =   3615
      End
      Begin VB.TextBox Accion211 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   100
         Text            =   " "
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox Accion212 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   99
         Text            =   " "
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Accion221 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   98
         Text            =   " "
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox Accion222 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   97
         Text            =   " "
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Accion231 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   96
         Text            =   " "
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox Accion232 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   95
         Text            =   " "
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox Accion241 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   94
         Text            =   " "
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox Accion242 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   93
         Text            =   " "
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox Accion251 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   92
         Text            =   " "
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox Accion252 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   91
         Text            =   " "
         Top             =   3840
         Width           =   3375
      End
      Begin VB.TextBox Accion261 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   90
         Text            =   " "
         Top             =   4320
         Width           =   3375
      End
      Begin VB.TextBox Accion262 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   89
         Text            =   " "
         Top             =   4560
         Width           =   3375
      End
      Begin VB.TextBox Responsable21 
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   88
         Text            =   " "
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Responsable15 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   85
         Text            =   " "
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Responsable16 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   82
         Text            =   " "
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox Responsable14 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   79
         Text            =   " "
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Responsable13 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   76
         Text            =   " "
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Responsable12 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   73
         Text            =   " "
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Comentario11 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   72
         Text            =   " "
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox Comentario12 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   71
         Text            =   " "
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Comentario21 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   70
         Text            =   " "
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Comentario22 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   69
         Text            =   " "
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Comentario31 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   68
         Text            =   " "
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox Comentario32 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   67
         Text            =   " "
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Comentario41 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   66
         Text            =   " "
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox Comentario42 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   65
         Text            =   " "
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Comentario51 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   64
         Text            =   " "
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox Comentario52 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   63
         Text            =   " "
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox Comentario61 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   62
         Text            =   " "
         Top             =   4320
         Width           =   2895
      End
      Begin VB.TextBox Comentario62 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   61
         Text            =   " "
         Top             =   4560
         Width           =   2895
      End
      Begin VB.TextBox Responsable11 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   54
         Text            =   " "
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Accion162 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   53
         Text            =   " "
         Top             =   4560
         Width           =   3015
      End
      Begin VB.TextBox Accion161 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   52
         Text            =   " "
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox Accion152 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         Text            =   " "
         Top             =   3840
         Width           =   3015
      End
      Begin VB.TextBox Accion151 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         Text            =   " "
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox Accion142 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Text            =   " "
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox Accion141 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   48
         Text            =   " "
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox Accion132 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   47
         Text            =   " "
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox Accion131 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   46
         Text            =   " "
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox Accion122 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   45
         Text            =   " "
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox Accion121 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   44
         Text            =   " "
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox Accion112 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   43
         Text            =   " "
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Accion111 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   42
         Text            =   " "
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Responsable6 
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
         Left            =   -66700
         MaxLength       =   6
         TabIndex        =   39
         Text            =   " "
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox Responsable5 
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
         Left            =   -66700
         MaxLength       =   6
         TabIndex        =   36
         Text            =   " "
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Responsable4 
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
         Left            =   -66700
         MaxLength       =   6
         TabIndex        =   33
         Text            =   " "
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Responsable3 
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
         Left            =   -66700
         MaxLength       =   6
         TabIndex        =   30
         Text            =   " "
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Responsable2 
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
         Left            =   -66700
         MaxLength       =   6
         TabIndex        =   27
         Text            =   " "
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Responsable1 
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
         Left            =   -66700
         MaxLength       =   6
         TabIndex        =   21
         Text            =   " "
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Accion62 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   20
         Text            =   " "
         Top             =   4560
         Width           =   8000
      End
      Begin VB.TextBox Accion61 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   19
         Text            =   " "
         Top             =   4320
         Width           =   8000
      End
      Begin VB.TextBox Accion52 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   18
         Text            =   " "
         Top             =   3840
         Width           =   8000
      End
      Begin VB.TextBox Accion51 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   17
         Text            =   " "
         Top             =   3600
         Width           =   8000
      End
      Begin VB.TextBox Accion42 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   16
         Text            =   " "
         Top             =   3120
         Width           =   8000
      End
      Begin VB.TextBox Accion41 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   15
         Text            =   " "
         Top             =   2880
         Width           =   8000
      End
      Begin VB.TextBox Accion32 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   14
         Text            =   " "
         Top             =   2400
         Width           =   8000
      End
      Begin VB.TextBox Accion31 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   13
         Text            =   " "
         Top             =   2160
         Width           =   8000
      End
      Begin VB.TextBox Accion22 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   12
         Text            =   " "
         Top             =   1800
         Width           =   8000
      End
      Begin VB.TextBox Accion21 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   11
         Text            =   " "
         Top             =   1560
         Width           =   8000
      End
      Begin VB.TextBox Accion12 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   10
         Text            =   " "
         Top             =   1080
         Width           =   8000
      End
      Begin VB.TextBox Accion11 
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
         Left            =   -74760
         MaxLength       =   60
         TabIndex        =   9
         Text            =   " "
         Top             =   840
         Width           =   8000
      End
      Begin VB.TextBox WTexto2555 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
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
         Left            =   1320
         TabIndex        =   8
         Top             =   -360
         Width           =   375
      End
      Begin VB.TextBox WTexto1555 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   -360
         Width           =   375
      End
      Begin MSMask.MaskEdBox WTexto3555 
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   -360
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16711935
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
      Begin MSMask.MaskEdBox Plazo1 
         Height          =   285
         Left            =   -64700
         TabIndex        =   24
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Plazo2 
         Height          =   285
         Left            =   -64700
         TabIndex        =   29
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Plazo3 
         Height          =   285
         Left            =   -64700
         TabIndex        =   32
         Top             =   2160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Plazo4 
         Height          =   285
         Left            =   -64700
         TabIndex        =   35
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Plazo5 
         Height          =   285
         Left            =   -64700
         TabIndex        =   38
         Top             =   3600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Plazo6 
         Height          =   285
         Left            =   -64700
         TabIndex        =   41
         Top             =   4320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha1 
         Height          =   285
         Left            =   -69600
         TabIndex        =   57
         Top             =   840
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
      Begin MSMask.MaskEdBox Fecha2 
         Height          =   285
         Left            =   -69600
         TabIndex        =   74
         Top             =   1560
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
      Begin MSMask.MaskEdBox Fecha3 
         Height          =   285
         Left            =   -69600
         TabIndex        =   77
         Top             =   2160
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
      Begin MSMask.MaskEdBox Fecha4 
         Height          =   285
         Left            =   -69600
         TabIndex        =   80
         Top             =   2880
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
      Begin MSMask.MaskEdBox Fecha5 
         Height          =   285
         Left            =   -69600
         TabIndex        =   83
         Top             =   3600
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
      Begin MSMask.MaskEdBox Fecha6 
         Height          =   285
         Left            =   -69600
         TabIndex        =   86
         Top             =   4320
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
      Begin MSMask.MaskEdBox Fecha21 
         Height          =   285
         Left            =   -71280
         TabIndex        =   101
         Top             =   1080
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
      Begin MSMask.MaskEdBox Fecha22 
         Height          =   285
         Left            =   -71280
         TabIndex        =   120
         Top             =   1800
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
      Begin MSMask.MaskEdBox Fecha23 
         Height          =   285
         Left            =   -71280
         TabIndex        =   123
         Top             =   2400
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
      Begin MSMask.MaskEdBox Fecha24 
         Height          =   285
         Left            =   -71280
         TabIndex        =   126
         Top             =   3120
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
      Begin MSMask.MaskEdBox Fecha25 
         Height          =   285
         Left            =   -71280
         TabIndex        =   129
         Top             =   3840
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
      Begin MSMask.MaskEdBox Fecha26 
         Height          =   285
         Left            =   -71280
         TabIndex        =   132
         Top             =   4560
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
      Begin VB.Label Label43 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   194
         Top             =   4560
         Width           =   200
      End
      Begin VB.Label Label42 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   193
         Top             =   3840
         Width           =   200
      End
      Begin VB.Label Label41 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   192
         Top             =   3120
         Width           =   200
      End
      Begin VB.Label Label40 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   191
         Top             =   2400
         Width           =   200
      End
      Begin VB.Label Label39 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   190
         Top             =   1800
         Width           =   200
      End
      Begin VB.Label Label38 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   189
         Top             =   1080
         Width           =   200
      End
      Begin VB.Label Label37 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   187
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label36 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   186
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label35 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   185
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label34 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   184
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label33 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   183
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label32 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   182
         Top             =   4320
         Width           =   135
      End
      Begin VB.Label Label26 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   181
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label27 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   180
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label28 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   179
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label29 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   178
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label30 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   177
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label31 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74920
         TabIndex        =   176
         Top             =   4320
         Width           =   135
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
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
         Left            =   -68160
         TabIndex        =   158
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Causas que lo originaron"
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
         Height          =   300
         Left            =   120
         TabIndex        =   151
         Top             =   2760
         Width           =   11175
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion de la No Conformidad"
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
         Height          =   300
         Left            =   120
         TabIndex        =   150
         Top             =   480
         Width           =   11175
      End
      Begin VB.Label DesResponsable26 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   133
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label DesResponsable25 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   130
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label DesResponsable24 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   127
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label DesResponsable23 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   124
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label DesResponsable22 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   121
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
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
         Height          =   285
         Left            =   -71280
         TabIndex        =   106
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label DesResponsable21 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   105
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
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
         Left            =   -68880
         TabIndex        =   104
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acciones Correctivas"
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
         Height          =   285
         Left            =   -74760
         TabIndex        =   103
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comentarios"
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
         Left            =   -67440
         TabIndex        =   102
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label DesResponsable16 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   87
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label DesResponsable15 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   84
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label DesResponsable14 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   81
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label DesResponsable13 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   78
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label DesResponsable12 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   75
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comentarios"
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
         Left            =   -66720
         TabIndex        =   60
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acciones Correctivas"
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
         Height          =   285
         Left            =   -74760
         TabIndex        =   59
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -69600
         TabIndex        =   58
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label DesResponsable11 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   56
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
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
         Height          =   285
         Left            =   -71640
         TabIndex        =   55
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label DesResponsable6 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66000
         TabIndex        =   40
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label DesResponsable5 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66000
         TabIndex        =   37
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label DesResponsable4 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66000
         TabIndex        =   34
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label DesResponsable3 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66000
         TabIndex        =   31
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label DesResponsable2 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66000
         TabIndex        =   28
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acciones Correctivas"
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
         Height          =   285
         Left            =   -74760
         TabIndex        =   26
         Top             =   480
         Width           =   7935
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plazo"
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
         Left            =   -64700
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label DesResponsable1 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66000
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
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
         Height          =   285
         Left            =   -66720
         TabIndex        =   22
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.TextBox Ayuda 
      BackColor       =   &H00FFFFC0&
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
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox Pantalla 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "consultasac.frx":008C
      Left            =   360
      List            =   "consultasac.frx":0093
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox Opcion 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   1440
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1320
      TabIndex        =   164
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.Label Label25 
      Caption         =   "Referencia"
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
      TabIndex        =   175
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label22 
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
      Left            =   3600
      TabIndex        =   173
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label15 
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
      TabIndex        =   172
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "Centro"
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
      Left            =   7320
      TabIndex        =   171
      Top             =   120
      Width           =   855
   End
   Begin VB.Label DesCentro 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   170
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Estado"
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
      Left            =   8040
      TabIndex        =   169
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Origen"
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
      Left            =   4080
      TabIndex        =   168
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Numero"
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
      Left            =   5280
      TabIndex        =   167
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo"
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
      TabIndex        =   166
      Top             =   120
      Width           =   735
   End
   Begin VB.Label DesTipo 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   165
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      Left            =   240
      TabIndex        =   147
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Resp. Inv."
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
      Left            =   5520
      TabIndex        =   146
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label DesResponsableEmisor 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   145
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label DesResponsableDestino 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   144
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Titulo"
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
      TabIndex        =   143
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   4680
      MouseIcon       =   "consultasac.frx":00A1
      MousePointer    =   99  'Custom
      Picture         =   "consultasac.frx":03AB
      ToolTipText     =   "Consulta de Datos"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6360
      MouseIcon       =   "consultasac.frx":0BED
      MousePointer    =   99  'Custom
      Picture         =   "consultasac.frx":0EF7
      ToolTipText     =   "Salida"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   8760
      MouseIcon       =   "consultasac.frx":1739
      MousePointer    =   99  'Custom
      Picture         =   "consultasac.frx":1A43
      ToolTipText     =   "Elimina el Registro"
      Top             =   7800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "consultasac.frx":2285
      MousePointer    =   99  'Custom
      Picture         =   "consultasac.frx":258F
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "consultasac.frx":2DD1
      MousePointer    =   99  'Custom
      Picture         =   "consultasac.frx":30DB
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7320
      Width           =   480
   End
End
Attribute VB_Name = "PrgConsultaSac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTipoSac As Recordset
Dim spTipoSac As String
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCargaSacII As Recordset
Dim spCargaSacII As String
Dim rstCargaSacIII As Recordset
Dim spCargaSacIII As String
Dim rstCargaSacIV As Recordset
Dim spCargaSacIV As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String

Dim XParam As String
Dim ZZLugar As Integer

Sub Imprime_Descripcion()
    
    Sql1 = "Select *"
    Sql2 = " FROM TipoSac"
    Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
    spTipoSac = Sql1 + Sql2 + Sql3
    Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoSac.RecordCount > 0 Then
        DesTipo.Caption = Trim(rstTipoSac!Descripcion)
        rstTipoSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM CentroSac"
    Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
    spCentroSac = Sql1 + Sql2 + Sql3
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        DesCentro.Caption = Trim(rstCentroSac!Descripcion)
        rstCentroSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableEmisor.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableEmisor.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableDestino.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableDestino.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    






    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable1.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable1.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable2.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable2.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable3.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable3.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable4.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable4.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable5.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable5.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable6.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable6.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable11.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable11.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable12.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable12.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable13.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable13.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable14.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable14.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable15.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable15.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable16.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable16.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    
    
    
    
    
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable21.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable21.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable22.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable22.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable23.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable23.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable24.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable24.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable25.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable25.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable26.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable26.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM CentroSac"
    Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
    spCentroSac = Sql1 + Sql2 + Sql3
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        DesCentro.Caption = Trim(rstCentroSac!Descripcion)
        rstCentroSac.Close
    End If
    
    
    
    
    
    
    

End Sub

Sub Imprime_Datos()

    On Error GoTo WError

    ZTipo = Tipo.Text
    ZAno = Ano.Text
    ZNumero = Numero.Text
    
    Call CmdLimpiar_Click
    
    ZExiste = "N"
    
    Tipo.Text = ZTipo
    Ano.Text = ZAno
    Numero.Text = ZNumero
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
    
        Centro.Text = rstCargaSac!Centro
        Fecha.Text = rstCargaSac!Fecha
        Origen.ListIndex = rstCargaSac!Origen
        Estado.ListIndex = rstCargaSac!Estado
        ResponsableEmisor.Text = rstCargaSac!ResponsableEmisor
        ResponsableDestino.Text = rstCargaSac!ResponsableDestino
        Referencia.Text = Trim(rstCargaSac!Referencia)
        Titulo.Text = Trim(rstCargaSac!Titulo)
        IngresoNoCon.Text = IIf(IsNull(rstCargaSac!IngresoNoCon), "", rstCargaSac!IngresoNoCon)
        IngresoCausa.Text = IIf(IsNull(rstCargaSac!IngresoCausa), "", rstCargaSac!IngresoCausa)
        
        rstCargaSac.Close
    End If
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacII"
    ZSql = ZSql + " Where CargaSacII.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacII.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacII.Numero = " + "'" + Numero.Text + "'"
    spCargaSacII = ZSql
    Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacII.RecordCount > 0 Then
    
        Accion11.Text = Trim(rstCargaSacII!Accion11)
        Accion12.Text = Trim(rstCargaSacII!Accion12)
        Accion21.Text = Trim(rstCargaSacII!Accion21)
        Accion22.Text = Trim(rstCargaSacII!Accion22)
        Accion31.Text = Trim(rstCargaSacII!Accion31)
        Accion32.Text = Trim(rstCargaSacII!Accion32)
        Accion41.Text = Trim(rstCargaSacII!Accion41)
        Accion42.Text = Trim(rstCargaSacII!Accion42)
        Accion51.Text = Trim(rstCargaSacII!Accion51)
        Accion52.Text = Trim(rstCargaSacII!Accion52)
        Accion61.Text = Trim(rstCargaSacII!Accion61)
        Accion62.Text = Trim(rstCargaSacII!Accion62)
        
        Responsable1.Text = rstCargaSacII!Responsable1
        Responsable2.Text = rstCargaSacII!Responsable2
        Responsable3.Text = rstCargaSacII!Responsable3
        Responsable4.Text = rstCargaSacII!Responsable4
        Responsable5.Text = rstCargaSacII!Responsable5
        Responsable6.Text = rstCargaSacII!Responsable6
        
        Plazo1.Text = rstCargaSacII!Plazo1
        Plazo2.Text = rstCargaSacII!Plazo2
        Plazo3.Text = rstCargaSacII!Plazo3
        Plazo4.Text = rstCargaSacII!Plazo4
        Plazo5.Text = rstCargaSacII!Plazo5
        Plazo6.Text = rstCargaSacII!Plazo6
        
        rstCargaSacII.Close
    End If
    
    
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacIII"
    ZSql = ZSql + " Where CargaSacIII.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacIII.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacIII.Numero = " + "'" + Numero.Text + "'"
    spCargaSacIII = ZSql
    Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacIII.RecordCount > 0 Then
    
        Responsable11.Text = rstCargaSacIII!Responsable1
        Responsable12.Text = rstCargaSacIII!Responsable2
        Responsable13.Text = rstCargaSacIII!Responsable3
        Responsable14.Text = rstCargaSacIII!Responsable4
        Responsable15.Text = rstCargaSacIII!Responsable5
        Responsable16.Text = rstCargaSacIII!Responsable6
        
        Fecha1.Text = rstCargaSacIII!Fecha1
        Fecha2.Text = rstCargaSacIII!Fecha2
        Fecha3.Text = rstCargaSacIII!Fecha3
        Fecha4.Text = rstCargaSacIII!Fecha4
        Fecha5.Text = rstCargaSacIII!Fecha5
        Fecha6.Text = rstCargaSacIII!Fecha6
        
        Comentario11.Text = Trim(rstCargaSacIII!Comentario11)
        Comentario12.Text = Trim(rstCargaSacIII!Comentario12)
        Comentario21.Text = Trim(rstCargaSacIII!Comentario21)
        Comentario22.Text = Trim(rstCargaSacIII!Comentario22)
        Comentario31.Text = Trim(rstCargaSacIII!Comentario31)
        Comentario32.Text = Trim(rstCargaSacIII!Comentario32)
        Comentario41.Text = Trim(rstCargaSacIII!Comentario41)
        Comentario42.Text = Trim(rstCargaSacIII!Comentario42)
        Comentario51.Text = Trim(rstCargaSacIII!Comentario51)
        Comentario52.Text = Trim(rstCargaSacIII!Comentario52)
        Comentario61.Text = Trim(rstCargaSacIII!Comentario61)
        Comentario62.Text = Trim(rstCargaSacIII!Comentario62)
        
        Estado11.ListIndex = rstCargaSacIII!Estado1
        Estado12.ListIndex = rstCargaSacIII!Estado2
        Estado13.ListIndex = rstCargaSacIII!Estado3
        Estado14.ListIndex = rstCargaSacIII!Estado4
        Estado15.ListIndex = rstCargaSacIII!Estado5
        Estado16.ListIndex = rstCargaSacIII!Estado6
        
        rstCargaSacIII.Close
    End If
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacIV"
    ZSql = ZSql + " Where CargaSacIV.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacIV.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacIV.Numero = " + "'" + Numero.Text + "'"
    spCargaSacIV = ZSql
    Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacIV.RecordCount > 0 Then
    
        Responsable21.Text = rstCargaSacIV!Responsable1
        Responsable22.Text = rstCargaSacIV!Responsable2
        Responsable23.Text = rstCargaSacIV!Responsable3
        Responsable24.Text = rstCargaSacIV!Responsable4
        Responsable25.Text = rstCargaSacIV!Responsable5
        Responsable26.Text = rstCargaSacIV!Responsable6
        
        Fecha21.Text = rstCargaSacIV!Fecha1
        Fecha22.Text = rstCargaSacIV!Fecha2
        Fecha23.Text = rstCargaSacIV!Fecha3
        Fecha24.Text = rstCargaSacIV!Fecha4
        Fecha25.Text = rstCargaSacIV!Fecha5
        Fecha26.Text = rstCargaSacIV!Fecha6
        
        Comentario211.Text = Trim(rstCargaSacIV!Comentario11)
        Comentario212.Text = Trim(rstCargaSacIV!Comentario12)
        Comentario221.Text = Trim(rstCargaSacIV!Comentario21)
        Comentario222.Text = Trim(rstCargaSacIV!Comentario22)
        Comentario231.Text = Trim(rstCargaSacIV!Comentario31)
        Comentario232.Text = Trim(rstCargaSacIV!Comentario32)
        Comentario241.Text = Trim(rstCargaSacIV!Comentario41)
        Comentario242.Text = Trim(rstCargaSacIV!Comentario42)
        Comentario251.Text = Trim(rstCargaSacIV!Comentario51)
        Comentario252.Text = Trim(rstCargaSacIV!Comentario52)
        Comentario261.Text = Trim(rstCargaSacIV!Comentario61)
        Comentario262.Text = Trim(rstCargaSacIV!Comentario62)
        
        Estado1.ListIndex = rstCargaSacIV!Estado1
        Estado2.ListIndex = rstCargaSacIV!Estado2
        Estado3.ListIndex = rstCargaSacIV!Estado3
        Estado4.ListIndex = rstCargaSacIV!Estado4
        Estado5.ListIndex = rstCargaSacIV!Estado5
        Estado6.ListIndex = rstCargaSacIV!Estado6
        
        rstCargaSacIV.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacAdicional"
    ZSql = ZSql + " Where CargaSacAdicional.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacAdicional.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacAdicional.Numero = " + "'" + Numero.Text + "'"
    spCargaSacAdicional = ZSql
    Set rstCargaSacAdicional = db.OpenRecordset(spCargaSacAdicional, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacAdicional.RecordCount > 0 Then
        Comentario.Text = rstCargaSacAdicional!Dato1
        rstCargaSacAdicional.Close
    End If
    
    
    
    
    
    
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Tipo.Text = "1"
    DesTipo.Caption = "SAC"
    Ano.Text = "2010"
    Numero.Text = ""
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    DesResponsableEmisor.Caption = ""
    DesResponsableDestino.Caption = ""
    Referencia.Text = ""
    Titulo.Text = ""
    IngresoNoCon.Text = ""
    IngresoCausa.Text = ""
    Comentario.Text = ""
    
    Origen.ListIndex = 0
    Estado.ListIndex = 0
    
    Ano.SetFocus
    
    Accion11.Text = ""
    Accion12.Text = ""
    Accion21.Text = ""
    Accion22.Text = ""
    Accion31.Text = ""
    Accion32.Text = ""
    Accion41.Text = ""
    Accion42.Text = ""
    Accion51.Text = ""
    Accion52.Text = ""
    Accion61.Text = ""
    Accion62.Text = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""

    Plazo1.Text = "  /  /    "
    Plazo2.Text = "  /  /    "
    Plazo3.Text = "  /  /    "
    Plazo4.Text = "  /  /    "
    Plazo5.Text = "  /  /    "
    Plazo6.Text = "  /  /    "
    
    Accion111.Text = ""
    Accion112.Text = ""
    Accion121.Text = ""
    Accion122.Text = ""
    Accion131.Text = ""
    Accion132.Text = ""
    Accion141.Text = ""
    Accion142.Text = ""
    Accion151.Text = ""
    Accion152.Text = ""
    Accion151.Text = ""
    Accion162.Text = ""
    
    Responsable11.Text = ""
    Responsable12.Text = ""
    Responsable13.Text = ""
    Responsable14.Text = ""
    Responsable15.Text = ""
    Responsable16.Text = ""
    
    DesResponsable11.Caption = ""
    DesResponsable12.Caption = ""
    DesResponsable13.Caption = ""
    DesResponsable14.Caption = ""
    DesResponsable15.Caption = ""
    DesResponsable16.Caption = ""

    Fecha1.Text = "  /  /    "
    Fecha2.Text = "  /  /    "
    Fecha3.Text = "  /  /    "
    Fecha4.Text = "  /  /    "
    Fecha5.Text = "  /  /    "
    Fecha6.Text = "  /  /    "
    
    Fecha21.Text = "  /  /    "
    Fecha22.Text = "  /  /    "
    Fecha23.Text = "  /  /    "
    Fecha24.Text = "  /  /    "
    Fecha25.Text = "  /  /    "
    Fecha26.Text = "  /  /    "
    
    Comentario11.Text = ""
    Comentario12.Text = ""
    Comentario21.Text = ""
    Comentario22.Text = ""
    Comentario31.Text = ""
    Comentario32.Text = ""
    Comentario41.Text = ""
    Comentario42.Text = ""
    Comentario51.Text = ""
    Comentario52.Text = ""
    Comentario61.Text = ""
    Comentario62.Text = ""
    
    Estado11.ListIndex = 0
    Estado12.ListIndex = 0
    Estado13.ListIndex = 0
    Estado14.ListIndex = 0
    Estado15.ListIndex = 0
    Estado16.ListIndex = 0
    
    Accion211.Text = ""
    Accion212.Text = ""
    Accion221.Text = ""
    Accion222.Text = ""
    Accion231.Text = ""
    Accion232.Text = ""
    Accion241.Text = ""
    Accion242.Text = ""
    Accion251.Text = ""
    Accion252.Text = ""
    Accion251.Text = ""
    Accion262.Text = ""
    
    Responsable21.Text = ""
    Responsable22.Text = ""
    Responsable23.Text = ""
    Responsable24.Text = ""
    Responsable25.Text = ""
    Responsable26.Text = ""
    
    DesResponsable21.Caption = ""
    DesResponsable22.Caption = ""
    DesResponsable23.Caption = ""
    DesResponsable24.Caption = ""
    DesResponsable25.Caption = ""
    DesResponsable26.Caption = ""
    
    Estado1.ListIndex = 0
    Estado2.ListIndex = 0
    Estado3.ListIndex = 0
    Estado4.ListIndex = 0
    Estado5.ListIndex = 0
    Estado6.ListIndex = 0
    
    Comentario211.Text = ""
    Comentario212.Text = ""
    Comentario221.Text = ""
    Comentario222.Text = ""
    Comentario231.Text = ""
    Comentario232.Text = ""
    Comentario241.Text = ""
    Comentario242.Text = ""
    Comentario251.Text = ""
    Comentario252.Text = ""
    Comentario261.Text = ""
    Comentario262.Text = ""
    
    
    Estado1.ListIndex = 0
    Estado2.ListIndex = 0
    Estado3.ListIndex = 0
    Estado4.ListIndex = 0
    Estado5.ListIndex = 0
    Estado6.ListIndex = 0
    
    Tablas.Tab = 0
    
    Centro.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    PrgConsultaSac.Hide
    Unload Me
    Menu.Show
    
End Sub



Sub Form_Load()

    Tipo.Text = "1"
    DesTipo.Caption = "SAC"
    Ano.Text = "2010"
    Numero.Text = ""
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    DesResponsableEmisor.Caption = ""
    DesResponsableDestino.Caption = ""
    Referencia.Text = ""
    Titulo.Text = ""
    IngresoNoCon.Text = ""
    IngresoCausa.Text = ""
    Comentario.Text = ""
    
    Accion11.Text = ""
    Accion12.Text = ""
    Accion21.Text = ""
    Accion22.Text = ""
    Accion31.Text = ""
    Accion32.Text = ""
    Accion41.Text = ""
    Accion42.Text = ""
    Accion51.Text = ""
    Accion52.Text = ""
    Accion61.Text = ""
    Accion62.Text = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""

    Plazo1.Text = "  /  /    "
    Plazo2.Text = "  /  /    "
    Plazo3.Text = "  /  /    "
    Plazo4.Text = "  /  /    "
    Plazo5.Text = "  /  /    "
    Plazo6.Text = "  /  /    "
    
    Accion111.Text = ""
    Accion112.Text = ""
    Accion121.Text = ""
    Accion122.Text = ""
    Accion131.Text = ""
    Accion132.Text = ""
    Accion141.Text = ""
    Accion142.Text = ""
    Accion151.Text = ""
    Accion152.Text = ""
    Accion151.Text = ""
    Accion162.Text = ""
    
    Responsable11.Text = ""
    Responsable12.Text = ""
    Responsable13.Text = ""
    Responsable14.Text = ""
    Responsable15.Text = ""
    Responsable16.Text = ""
    
    DesResponsable11.Caption = ""
    DesResponsable12.Caption = ""
    DesResponsable13.Caption = ""
    DesResponsable14.Caption = ""
    DesResponsable15.Caption = ""
    DesResponsable16.Caption = ""

    Fecha1.Text = "  /  /    "
    Fecha2.Text = "  /  /    "
    Fecha3.Text = "  /  /    "
    Fecha4.Text = "  /  /    "
    Fecha5.Text = "  /  /    "
    Fecha6.Text = "  /  /    "
    
    Comentario11.Text = ""
    Comentario12.Text = ""
    Comentario21.Text = ""
    Comentario22.Text = ""
    Comentario31.Text = ""
    Comentario32.Text = ""
    Comentario41.Text = ""
    Comentario42.Text = ""
    Comentario51.Text = ""
    Comentario52.Text = ""
    Comentario61.Text = ""
    Comentario62.Text = ""
    
    Accion211.Text = ""
    Accion212.Text = ""
    Accion221.Text = ""
    Accion222.Text = ""
    Accion231.Text = ""
    Accion232.Text = ""
    Accion241.Text = ""
    Accion242.Text = ""
    Accion251.Text = ""
    Accion252.Text = ""
    Accion251.Text = ""
    Accion262.Text = ""
    
    Responsable21.Text = ""
    Responsable22.Text = ""
    Responsable23.Text = ""
    Responsable24.Text = ""
    Responsable25.Text = ""
    Responsable26.Text = ""
    
    DesResponsable21.Caption = ""
    DesResponsable22.Caption = ""
    DesResponsable23.Caption = ""
    DesResponsable24.Caption = ""
    DesResponsable25.Caption = ""
    DesResponsable26.Caption = ""
    
    Comentario211.Text = ""
    Comentario212.Text = ""
    Comentario221.Text = ""
    Comentario222.Text = ""
    Comentario231.Text = ""
    Comentario232.Text = ""
    Comentario241.Text = ""
    Comentario242.Text = ""
    Comentario251.Text = ""
    Comentario252.Text = ""
    Comentario261.Text = ""
    Comentario262.Text = ""
    
    Estado.Clear
    
    Estado.AddItem ""
    Estado.AddItem "INICIADA"
    Estado.AddItem "INVESTIGACION"
    Estado.AddItem "IMPLEMENTACION"
    Estado.AddItem "IMPLEMENTACION A VERIFICAR"
    Estado.AddItem "IMPLEMENTACION VERIFICADA"
    Estado.AddItem "CERRADA"
    Estado.AddItem "ANULADA"
    
    Estado.ListIndex = 0
    
    Origen.Clear
    
    Origen.AddItem ""
    Origen.AddItem "Auditoria"
    Origen.AddItem "Reclamo"
    Origen.AddItem "I. No Conformidad"
    Origen.AddItem "Proceso/Sist"
    Origen.AddItem "Otro"
    
    Origen.ListIndex = 0
    
    
    
    Estado11.Clear
    
    Estado11.AddItem ""
    Estado11.AddItem "Imple."
    Estado11.AddItem "Nula"
    
    Estado11.ListIndex = 0
    
    Estado12.Clear
    
    Estado12.AddItem ""
    Estado12.AddItem "Imple."
    Estado12.AddItem "Nula"
    
    Estado12.ListIndex = 0
    
    Estado13.Clear
    
    Estado13.AddItem ""
    Estado13.AddItem "Imple."
    Estado13.AddItem "Nula"
    
    Estado13.ListIndex = 0
    
    Estado14.Clear
    
    Estado14.AddItem ""
    Estado14.AddItem "Imple."
    Estado14.AddItem "Nula"
    
    Estado14.ListIndex = 0
    
    Estado15.Clear
    
    Estado15.AddItem ""
    Estado15.AddItem "Imple."
    Estado15.AddItem "Nula"
 
    Estado15.ListIndex = 0
    
    Estado16.Clear
    
    Estado16.AddItem ""
    Estado16.AddItem "Imple."
    Estado16.AddItem "Nula"
    
    Estado16.ListIndex = 0
    
    
    
    
    
    Estado1.Clear
    
    Estado1.AddItem "No Imple."
    Estado1.AddItem "Imple."
    Estado1.AddItem "Nula"
    Estado1.AddItem "Cerrada"
    Estado1.AddItem ""
    
    Estado1.ListIndex = 0
    
    Estado2.Clear
    
    Estado2.AddItem "No Imple."
    Estado2.AddItem "Imple."
    Estado2.AddItem "Nula"
    Estado2.AddItem "Cerrada"
    Estado2.AddItem ""
    
    Estado2.ListIndex = 0
    
    Estado3.Clear
    
    Estado3.AddItem "No Imple."
    Estado3.AddItem "Imple."
    Estado3.AddItem "Nula"
    Estado3.AddItem "Cerrada"
    Estado3.AddItem ""
    
    Estado3.ListIndex = 0
    
    Estado4.Clear
    
    Estado4.AddItem "No Imple."
    Estado4.AddItem "Imple."
    Estado4.AddItem "Nula"
    Estado4.AddItem "Cerrada"
    Estado4.AddItem ""
    
    Estado4.ListIndex = 0
    
    Estado5.Clear
    
    Estado5.AddItem "No Imple."
    Estado5.AddItem "Imple."
    Estado5.AddItem "Nula"
    Estado5.AddItem "Cerrada"
    Estado5.AddItem ""
 
    Estado5.ListIndex = 0
    
    Estado6.Clear
    
    Estado6.AddItem "No Imple."
    Estado6.AddItem "Imple."
    Estado6.AddItem "Nula"
    Estado6.AddItem "Cerrada"
    Estado6.AddItem ""
    
    Estado6.ListIndex = 0
    
End Sub


Private Sub Tablas_Click(PreviousTab As Integer)
    
    Select Case Tablas.Tab
        Case 2
            Accion111.Text = Accion11.Text
            Accion112.Text = Accion12.Text
            Accion121.Text = Accion21.Text
            Accion122.Text = Accion22.Text
            Accion131.Text = Accion31.Text
            Accion132.Text = Accion32.Text
            Accion141.Text = Accion41.Text
            Accion142.Text = Accion42.Text
            Accion151.Text = Accion51.Text
            Accion152.Text = Accion52.Text
            Accion161.Text = Accion61.Text
            Accion162.Text = Accion62.Text
        Case 3
            Accion211.Text = Accion11.Text
            Accion212.Text = Accion12.Text
            Accion221.Text = Accion21.Text
            Accion222.Text = Accion22.Text
            Accion231.Text = Accion31.Text
            Accion232.Text = Accion32.Text
            Accion241.Text = Accion41.Text
            Accion242.Text = Accion42.Text
            Accion251.Text = Accion51.Text
            Accion252.Text = Accion52.Text
            Accion261.Text = Accion61.Text
            Accion262.Text = Accion62.Text
        Case Else
    End Select
End Sub

Private Sub Tipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM TipoSac"
        Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
        spTipoSac = Sql1 + Sql2 + Sql3
        Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoSac.RecordCount > 0 Then
            DesTipo.Caption = Trim(rstTipoSac!Descripcion)
            rstTipoSac.Close
            Ano.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Tipo.Text = ""
        DesTipo.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Numero.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
                
                Centro.Text = rstCargaSac!Centro
                Fecha.Text = rstCargaSac!Fecha
                Origen.ListIndex = rstCargaSac!Origen
                Estado.ListIndex = rstCargaSac!Estado
                ResponsableEmisor.Text = rstCargaSac!ResponsableEmisor
                ResponsableDestino.Text = rstCargaSac!ResponsableDestino
                Referencia.Text = Trim(rstCargaSac!Referencia)
                Titulo.Text = rstCargaSac!Titulo
                IngresoNoCon.Text = IIf(IsNull(rstCargaSac!IngresoNoCon), "", rstCargaSac!IngresoNoCon)
                IngresoCausa.Text = IIf(IsNull(rstCargaSac!IngresoCausa), "", rstCargaSac!IngresoCausa)
                Rem dada
                rstCargaSac.Close
                
                Call Imprime_Datos
                Call Imprime_Descripcion
                
                    Else
                    
                WTipo = Tipo.Text
                WAno = Ano.Text
                WNumero = Numero.Text
                CmdLimpiar_Click
                Ano.Text = WAno
                Numero.Text = WNumero
                Tipo.Text = WTipo
                Sql1 = "Select *"
                Sql2 = " FROM TipoSac"
                Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
                spTipoSac = Sql1 + Sql2 + Sql3
                Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstTipoSac.RecordCount > 0 Then
                    DesTipo.Caption = Trim(rstTipoSac!Descripcion)
                    rstTipoSac.Close
                    Ano.SetFocus
                End If
                Tipo.SetFocus
                
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
