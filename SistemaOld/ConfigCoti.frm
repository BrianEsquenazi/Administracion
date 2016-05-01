VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PrgConfigCoti 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreos de Atributos de Cotiza"
   ClientHeight    =   8400
   ClientLeft      =   225
   ClientTop       =   390
   ClientWidth     =   11535
   LinkTopic       =   "Form2"
   ScaleHeight     =   8400
   ScaleWidth      =   11535
   Begin VB.TextBox Operador 
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
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Salida 
      Caption         =   "Salida"
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
      Left            =   6480
      TabIndex        =   3
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   7320
      Width           =   1215
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   11668
      _Version        =   327680
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "ConfigCoti.frx":0000
      Tab(0).ControlCount=   17
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Titulo1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Titulo2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Titulo3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Titulo4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Titulo7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Titulo6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Titulo5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "titulo8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Opcion1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Opcion2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Opcion3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Opcion4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Opcion8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Opcion7"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Opcion6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Opcion5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Opcion9"
      Tab(0).Control(16).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "ConfigCoti.frx":001C
      Tab(1).ControlCount=   62
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Opcion131"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "Opcion130"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "Opcion129"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "Opcion128"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "Opcion127"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "Opcion124"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "Opcion125"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "Opcion126"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "Opcion123"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Opcion122"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "Opcion121"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "Opcion120"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "Opcion119"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "Opcion118"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "Opcion117"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "Opcion116"
      Tab(1).Control(15).Enabled=   -1  'True
      Tab(1).Control(16)=   "Opcion113"
      Tab(1).Control(16).Enabled=   -1  'True
      Tab(1).Control(17)=   "Opcion114"
      Tab(1).Control(17).Enabled=   -1  'True
      Tab(1).Control(18)=   "Opcion115"
      Tab(1).Control(18).Enabled=   -1  'True
      Tab(1).Control(19)=   "Opcion19"
      Tab(1).Control(19).Enabled=   -1  'True
      Tab(1).Control(20)=   "Opcion110"
      Tab(1).Control(20).Enabled=   -1  'True
      Tab(1).Control(21)=   "Opcion111"
      Tab(1).Control(21).Enabled=   -1  'True
      Tab(1).Control(22)=   "Opcion112"
      Tab(1).Control(22).Enabled=   -1  'True
      Tab(1).Control(23)=   "Opcion15"
      Tab(1).Control(23).Enabled=   -1  'True
      Tab(1).Control(24)=   "Opcion16"
      Tab(1).Control(24).Enabled=   -1  'True
      Tab(1).Control(25)=   "Opcion17"
      Tab(1).Control(25).Enabled=   -1  'True
      Tab(1).Control(26)=   "Opcion18"
      Tab(1).Control(26).Enabled=   -1  'True
      Tab(1).Control(27)=   "Opcion14"
      Tab(1).Control(27).Enabled=   -1  'True
      Tab(1).Control(28)=   "Opcion13"
      Tab(1).Control(28).Enabled=   -1  'True
      Tab(1).Control(29)=   "Opcion12"
      Tab(1).Control(29).Enabled=   -1  'True
      Tab(1).Control(30)=   "Opcion11"
      Tab(1).Control(30).Enabled=   -1  'True
      Tab(1).Control(31)=   "Titulo131"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Titulo130"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Titulo129"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Titulo128"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Titulo127"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Titulo124"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Titulo125"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Titulo126"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Titulo123"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Titulo122"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Titulo121"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Titulo120"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Titulo119"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Titulo118"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Titulo117"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Titulo116"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Titulo113"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Titulo114"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Titulo115"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Titulo19"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Titulo110"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Titulo111"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Titulo112"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Titulo15"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Titulo16"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "Titulo17"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "Titulo18"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "Titulo14"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "Titulo13"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Titulo12"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "Titulo11"
      Tab(1).Control(61).Enabled=   0   'False
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "ConfigCoti.frx":0038
      Tab(2).ControlCount=   59
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Opcion228"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "Opcion229"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "Opcion230"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "Opcion227"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "Opcion226"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "Opcion225"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "Opcion216"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "Opcion215"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "Opcion214"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "Opcion213"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "Opcion220"
      Tab(2).Control(10).Enabled=   -1  'True
      Tab(2).Control(11)=   "Opcion219"
      Tab(2).Control(11).Enabled=   -1  'True
      Tab(2).Control(12)=   "Opcion218"
      Tab(2).Control(12).Enabled=   -1  'True
      Tab(2).Control(13)=   "Opcion217"
      Tab(2).Control(13).Enabled=   -1  'True
      Tab(2).Control(14)=   "Opcion224"
      Tab(2).Control(14).Enabled=   -1  'True
      Tab(2).Control(15)=   "Opcion223"
      Tab(2).Control(15).Enabled=   -1  'True
      Tab(2).Control(16)=   "Opcion222"
      Tab(2).Control(16).Enabled=   -1  'True
      Tab(2).Control(17)=   "Opcion221"
      Tab(2).Control(17).Enabled=   -1  'True
      Tab(2).Control(18)=   "Opcion29"
      Tab(2).Control(18).Enabled=   -1  'True
      Tab(2).Control(19)=   "Opcion210"
      Tab(2).Control(19).Enabled=   -1  'True
      Tab(2).Control(20)=   "Opcion211"
      Tab(2).Control(20).Enabled=   -1  'True
      Tab(2).Control(21)=   "Opcion212"
      Tab(2).Control(21).Enabled=   -1  'True
      Tab(2).Control(22)=   "Opcion25"
      Tab(2).Control(22).Enabled=   -1  'True
      Tab(2).Control(23)=   "Opcion26"
      Tab(2).Control(23).Enabled=   -1  'True
      Tab(2).Control(24)=   "Opcion27"
      Tab(2).Control(24).Enabled=   -1  'True
      Tab(2).Control(25)=   "Opcion28"
      Tab(2).Control(25).Enabled=   -1  'True
      Tab(2).Control(26)=   "opcion21"
      Tab(2).Control(26).Enabled=   -1  'True
      Tab(2).Control(27)=   "Opcion22"
      Tab(2).Control(27).Enabled=   -1  'True
      Tab(2).Control(28)=   "Opcion23"
      Tab(2).Control(28).Enabled=   -1  'True
      Tab(2).Control(29)=   "Opcion24"
      Tab(2).Control(29).Enabled=   -1  'True
      Tab(2).Control(30)=   "Titulo228"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Titulo229"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Titulo227"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Titulo226"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Titulo225"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Titulo216"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Titulo215"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Titulo214"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Titulo213"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Titulo220"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Titulo219"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Titulo218"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Titulo217"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Titulo224"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Titulo223"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Titulo222"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "Titulo221"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "Titulo29"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "Titulo210"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "Titulo211"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "Titulo212"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "Titulo25"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "Titulo26"
      Tab(2).Control(52).Enabled=   0   'False
      Tab(2).Control(53)=   "Titulo27"
      Tab(2).Control(53).Enabled=   0   'False
      Tab(2).Control(54)=   "Titulo28"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).Control(55)=   "Titulo21"
      Tab(2).Control(55).Enabled=   0   'False
      Tab(2).Control(56)=   "Titulo22"
      Tab(2).Control(56).Enabled=   0   'False
      Tab(2).Control(57)=   "Titulo23"
      Tab(2).Control(57).Enabled=   0   'False
      Tab(2).Control(58)=   "Titulo24"
      Tab(2).Control(58).Enabled=   0   'False
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "ConfigCoti.frx":0054
      Tab(3).ControlCount=   56
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Opcion331"
      Tab(3).Control(0).Enabled=   -1  'True
      Tab(3).Control(1)=   "Opcion329"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "Opcion330"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "Opcion328"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "Opcion327"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "Opcion326"
      Tab(3).Control(5).Enabled=   -1  'True
      Tab(3).Control(6)=   "Opcion325"
      Tab(3).Control(6).Enabled=   -1  'True
      Tab(3).Control(7)=   "Opcion316"
      Tab(3).Control(7).Enabled=   -1  'True
      Tab(3).Control(8)=   "Opcion315"
      Tab(3).Control(8).Enabled=   -1  'True
      Tab(3).Control(9)=   "Opcion314"
      Tab(3).Control(9).Enabled=   -1  'True
      Tab(3).Control(10)=   "Opcion313"
      Tab(3).Control(10).Enabled=   -1  'True
      Tab(3).Control(11)=   "Opcion320"
      Tab(3).Control(11).Enabled=   -1  'True
      Tab(3).Control(12)=   "Opcion319"
      Tab(3).Control(12).Enabled=   -1  'True
      Tab(3).Control(13)=   "Opcion318"
      Tab(3).Control(13).Enabled=   -1  'True
      Tab(3).Control(14)=   "Opcion317"
      Tab(3).Control(14).Enabled=   -1  'True
      Tab(3).Control(15)=   "Opcion324"
      Tab(3).Control(15).Enabled=   -1  'True
      Tab(3).Control(16)=   "Opcion323"
      Tab(3).Control(16).Enabled=   -1  'True
      Tab(3).Control(17)=   "Opcion322"
      Tab(3).Control(17).Enabled=   -1  'True
      Tab(3).Control(18)=   "Opcion321"
      Tab(3).Control(18).Enabled=   -1  'True
      Tab(3).Control(19)=   "Opcion34"
      Tab(3).Control(19).Enabled=   -1  'True
      Tab(3).Control(20)=   "Opcion33"
      Tab(3).Control(20).Enabled=   -1  'True
      Tab(3).Control(21)=   "Opcion32"
      Tab(3).Control(21).Enabled=   -1  'True
      Tab(3).Control(22)=   "Opcion31"
      Tab(3).Control(22).Enabled=   -1  'True
      Tab(3).Control(23)=   "Opcion38"
      Tab(3).Control(23).Enabled=   -1  'True
      Tab(3).Control(24)=   "Opcion37"
      Tab(3).Control(24).Enabled=   -1  'True
      Tab(3).Control(25)=   "Opcion36"
      Tab(3).Control(25).Enabled=   -1  'True
      Tab(3).Control(26)=   "Opcion35"
      Tab(3).Control(26).Enabled=   -1  'True
      Tab(3).Control(27)=   "Opcion312"
      Tab(3).Control(27).Enabled=   -1  'True
      Tab(3).Control(28)=   "Opcion311"
      Tab(3).Control(28).Enabled=   -1  'True
      Tab(3).Control(29)=   "Opcion310"
      Tab(3).Control(29).Enabled=   -1  'True
      Tab(3).Control(30)=   "Opcion39"
      Tab(3).Control(30).Enabled=   -1  'True
      Tab(3).Control(31)=   "Titulo325"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "Titulo316"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "Titulo315"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "Titulo314"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "Titulo313"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).Control(36)=   "Titulo320"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "Titulo319"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "Titulo318"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "Titulo317"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).Control(40)=   "Titulo324"
      Tab(3).Control(40).Enabled=   0   'False
      Tab(3).Control(41)=   "Titulo323"
      Tab(3).Control(41).Enabled=   0   'False
      Tab(3).Control(42)=   "Titulo322"
      Tab(3).Control(42).Enabled=   0   'False
      Tab(3).Control(43)=   "Titulo321"
      Tab(3).Control(43).Enabled=   0   'False
      Tab(3).Control(44)=   "Titulo34"
      Tab(3).Control(44).Enabled=   0   'False
      Tab(3).Control(45)=   "Titulo33"
      Tab(3).Control(45).Enabled=   0   'False
      Tab(3).Control(46)=   "Titulo32"
      Tab(3).Control(46).Enabled=   0   'False
      Tab(3).Control(47)=   "Titulo31"
      Tab(3).Control(47).Enabled=   0   'False
      Tab(3).Control(48)=   "Titulo38"
      Tab(3).Control(48).Enabled=   0   'False
      Tab(3).Control(49)=   "Titulo37"
      Tab(3).Control(49).Enabled=   0   'False
      Tab(3).Control(50)=   "Titulo36"
      Tab(3).Control(50).Enabled=   0   'False
      Tab(3).Control(51)=   "Titulo35"
      Tab(3).Control(51).Enabled=   0   'False
      Tab(3).Control(52)=   "Titulo312"
      Tab(3).Control(52).Enabled=   0   'False
      Tab(3).Control(53)=   "Titulo311"
      Tab(3).Control(53).Enabled=   0   'False
      Tab(3).Control(54)=   "Titulo310"
      Tab(3).Control(54).Enabled=   0   'False
      Tab(3).Control(55)=   "Titulo39"
      Tab(3).Control(55).Enabled=   0   'False
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "ConfigCoti.frx":0070
      Tab(4).ControlCount=   31
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Titulo44"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Titulo43"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Titulo42"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Titulo41"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Titulo48"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Titulo47"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Titulo46"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Titulo45"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Titulo49"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Titulo410"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Titulo415"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Opcion44"
      Tab(4).Control(11).Enabled=   -1  'True
      Tab(4).Control(12)=   "Opcion43"
      Tab(4).Control(12).Enabled=   -1  'True
      Tab(4).Control(13)=   "Opcion42"
      Tab(4).Control(13).Enabled=   -1  'True
      Tab(4).Control(14)=   "Opcion41"
      Tab(4).Control(14).Enabled=   -1  'True
      Tab(4).Control(15)=   "Opcion48"
      Tab(4).Control(15).Enabled=   -1  'True
      Tab(4).Control(16)=   "Opcion47"
      Tab(4).Control(16).Enabled=   -1  'True
      Tab(4).Control(17)=   "Opcion46"
      Tab(4).Control(17).Enabled=   -1  'True
      Tab(4).Control(18)=   "Opcion45"
      Tab(4).Control(18).Enabled=   -1  'True
      Tab(4).Control(19)=   "Opcion49"
      Tab(4).Control(19).Enabled=   -1  'True
      Tab(4).Control(20)=   "Opcion410"
      Tab(4).Control(20).Enabled=   -1  'True
      Tab(4).Control(21)=   "Opcion411"
      Tab(4).Control(21).Enabled=   -1  'True
      Tab(4).Control(22)=   "Opcion412"
      Tab(4).Control(22).Enabled=   -1  'True
      Tab(4).Control(23)=   "Opcion413"
      Tab(4).Control(23).Enabled=   -1  'True
      Tab(4).Control(24)=   "Opcion414"
      Tab(4).Control(24).Enabled=   -1  'True
      Tab(4).Control(25)=   "Opcion417"
      Tab(4).Control(25).Enabled=   -1  'True
      Tab(4).Control(26)=   "Opcion416"
      Tab(4).Control(26).Enabled=   -1  'True
      Tab(4).Control(27)=   "Opcion415"
      Tab(4).Control(27).Enabled=   -1  'True
      Tab(4).Control(28)=   "Opcion418"
      Tab(4).Control(28).Enabled=   -1  'True
      Tab(4).Control(29)=   "Opcion419"
      Tab(4).Control(29).Enabled=   -1  'True
      Tab(4).Control(30)=   "Opcion420"
      Tab(4).Control(30).Enabled=   -1  'True
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "ConfigCoti.frx":008C
      Tab(5).ControlCount=   23
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Opcion514"
      Tab(5).Control(0).Enabled=   -1  'True
      Tab(5).Control(1)=   "Opcion513"
      Tab(5).Control(1).Enabled=   -1  'True
      Tab(5).Control(2)=   "Opcion54"
      Tab(5).Control(2).Enabled=   -1  'True
      Tab(5).Control(3)=   "Opcion53"
      Tab(5).Control(3).Enabled=   -1  'True
      Tab(5).Control(4)=   "Opcion52"
      Tab(5).Control(4).Enabled=   -1  'True
      Tab(5).Control(5)=   "Opcion51"
      Tab(5).Control(5).Enabled=   -1  'True
      Tab(5).Control(6)=   "Opcion58"
      Tab(5).Control(6).Enabled=   -1  'True
      Tab(5).Control(7)=   "Opcion57"
      Tab(5).Control(7).Enabled=   -1  'True
      Tab(5).Control(8)=   "Opcion56"
      Tab(5).Control(8).Enabled=   -1  'True
      Tab(5).Control(9)=   "Opcion55"
      Tab(5).Control(9).Enabled=   -1  'True
      Tab(5).Control(10)=   "Opcion512"
      Tab(5).Control(10).Enabled=   -1  'True
      Tab(5).Control(11)=   "Opcion511"
      Tab(5).Control(11).Enabled=   -1  'True
      Tab(5).Control(12)=   "Opcion510"
      Tab(5).Control(12).Enabled=   -1  'True
      Tab(5).Control(13)=   "Opcion59"
      Tab(5).Control(13).Enabled=   -1  'True
      Tab(5).Control(14)=   "Titulo54"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "Titulo53"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "Titulo52"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "Titulo51"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "Titulo58"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "Titulo57"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "Titulo56"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "Titulo55"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "Titulo59"
      Tab(5).Control(22).Enabled=   0   'False
      Begin VB.CheckBox Opcion420 
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
         Left            =   -64800
         TabIndex        =   252
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Opcion419 
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
         Left            =   -64800
         TabIndex        =   251
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Opcion418 
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
         Left            =   -64800
         TabIndex        =   250
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Opcion9 
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
         Left            =   4680
         TabIndex        =   249
         Top             =   3480
         Width           =   615
      End
      Begin VB.CheckBox Opcion331 
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
         Left            =   -64680
         TabIndex        =   248
         Top             =   5880
         Width           =   615
      End
      Begin VB.CheckBox Opcion514 
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
         Left            =   -69600
         TabIndex        =   247
         Top             =   5280
         Width           =   615
      End
      Begin VB.CheckBox Opcion415 
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
         Left            =   -64800
         TabIndex        =   245
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Opcion416 
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
         Left            =   -64800
         TabIndex        =   244
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Opcion417 
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
         Left            =   -64800
         TabIndex        =   243
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Opcion329 
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
         Left            =   -64680
         TabIndex        =   242
         Top             =   5160
         Width           =   615
      End
      Begin VB.CheckBox Opcion330 
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
         Left            =   -64680
         TabIndex        =   241
         Top             =   5520
         Width           =   615
      End
      Begin VB.CheckBox Opcion228 
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
         Left            =   -64800
         TabIndex        =   238
         Top             =   4680
         Width           =   615
      End
      Begin VB.CheckBox Opcion229 
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
         Left            =   -64800
         TabIndex        =   237
         Top             =   5040
         Width           =   615
      End
      Begin VB.CheckBox Opcion230 
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
         Left            =   -64800
         TabIndex        =   236
         Top             =   5400
         Width           =   615
      End
      Begin VB.CheckBox Opcion5 
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
         Left            =   4680
         TabIndex        =   232
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox Opcion6 
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
         Left            =   4680
         TabIndex        =   231
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox Opcion7 
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
         Left            =   4680
         TabIndex        =   230
         Top             =   2760
         Width           =   615
      End
      Begin VB.CheckBox Opcion8 
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
         Left            =   4680
         TabIndex        =   229
         Top             =   3120
         Width           =   615
      End
      Begin VB.CheckBox Opcion131 
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
         Left            =   -64800
         TabIndex        =   227
         Top             =   6120
         Width           =   615
      End
      Begin VB.CheckBox Opcion414 
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
         Left            =   -69960
         TabIndex        =   226
         Top             =   5400
         Width           =   615
      End
      Begin VB.CheckBox Opcion413 
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
         Left            =   -69960
         TabIndex        =   225
         Top             =   5040
         Width           =   615
      End
      Begin VB.CheckBox Opcion412 
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
         Left            =   -69960
         TabIndex        =   224
         Top             =   4680
         Width           =   615
      End
      Begin VB.CheckBox Opcion328 
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
         Left            =   -64680
         TabIndex        =   223
         Top             =   4680
         Width           =   615
      End
      Begin VB.CheckBox Opcion227 
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
         Left            =   -64800
         TabIndex        =   221
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox Opcion226 
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
         Left            =   -64800
         TabIndex        =   219
         Top             =   3960
         Width           =   615
      End
      Begin VB.CheckBox Opcion130 
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
         Left            =   -64800
         TabIndex        =   217
         Top             =   5760
         Width           =   615
      End
      Begin VB.CheckBox Opcion129 
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
         Left            =   -64800
         TabIndex        =   215
         Top             =   5400
         Width           =   615
      End
      Begin VB.CheckBox Opcion411 
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
         Left            =   -69960
         TabIndex        =   214
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox Opcion128 
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
         Left            =   -64800
         TabIndex        =   212
         Top             =   5040
         Width           =   615
      End
      Begin VB.CheckBox Opcion327 
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
         Left            =   -64680
         TabIndex        =   211
         Top             =   4260
         Width           =   615
      End
      Begin VB.CheckBox Opcion326 
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
         Left            =   -64680
         TabIndex        =   210
         Top             =   3840
         Width           =   615
      End
      Begin VB.CheckBox Opcion410 
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
         Left            =   -69960
         TabIndex        =   208
         Top             =   3960
         Width           =   615
      End
      Begin VB.CheckBox Opcion325 
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
         Left            =   -64680
         TabIndex        =   206
         Top             =   3480
         Width           =   615
      End
      Begin VB.CheckBox Opcion127 
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
         Left            =   -64800
         TabIndex        =   204
         Top             =   4680
         Width           =   615
      End
      Begin VB.CheckBox Opcion124 
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
         Left            =   -64800
         TabIndex        =   200
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Opcion125 
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
         Left            =   -64800
         TabIndex        =   199
         Top             =   3960
         Width           =   615
      End
      Begin VB.CheckBox Opcion126 
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
         Left            =   -64800
         TabIndex        =   198
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox Opcion123 
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
         Left            =   -64800
         TabIndex        =   196
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox Opcion49 
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
         Left            =   -69960
         TabIndex        =   194
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Opcion122 
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
         Left            =   -64800
         TabIndex        =   192
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Opcion121 
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
         Left            =   -64800
         TabIndex        =   190
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Opcion225 
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
         Left            =   -64800
         TabIndex        =   188
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Opcion513 
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
         Left            =   -69600
         TabIndex        =   187
         Top             =   4920
         Width           =   615
      End
      Begin VB.CheckBox Opcion54 
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
         Left            =   -69600
         TabIndex        =   177
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox Opcion53 
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
         Left            =   -69600
         TabIndex        =   176
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox Opcion52 
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
         Left            =   -69600
         TabIndex        =   175
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Opcion51 
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
         Left            =   -69600
         TabIndex        =   174
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox Opcion58 
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
         Left            =   -69600
         TabIndex        =   173
         Top             =   3120
         Width           =   615
      End
      Begin VB.CheckBox Opcion57 
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
         Left            =   -69600
         TabIndex        =   172
         Top             =   2760
         Width           =   615
      End
      Begin VB.CheckBox Opcion56 
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
         Left            =   -69600
         TabIndex        =   171
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox Opcion55 
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
         Left            =   -69600
         TabIndex        =   170
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox Opcion512 
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
         Left            =   -69600
         TabIndex        =   169
         Top             =   4560
         Width           =   615
      End
      Begin VB.CheckBox Opcion511 
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
         Left            =   -69600
         TabIndex        =   168
         Top             =   4200
         Width           =   615
      End
      Begin VB.CheckBox Opcion510 
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
         Left            =   -69600
         TabIndex        =   167
         Top             =   3840
         Width           =   615
      End
      Begin VB.CheckBox Opcion59 
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
         Left            =   -69600
         TabIndex        =   166
         Top             =   3480
         Width           =   615
      End
      Begin VB.CheckBox Opcion120 
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
         Left            =   -64800
         TabIndex        =   164
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Opcion119 
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
         Left            =   -64800
         TabIndex        =   162
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Opcion118 
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
         Left            =   -64800
         TabIndex        =   160
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Opcion117 
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
         Left            =   -64800
         TabIndex        =   158
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Opcion116 
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
         Left            =   -64800
         TabIndex        =   156
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Opcion45 
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
         Left            =   -69960
         TabIndex        =   147
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Opcion46 
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
         Left            =   -69960
         TabIndex        =   146
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Opcion47 
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
         Left            =   -69960
         TabIndex        =   145
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Opcion48 
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
         Left            =   -69960
         TabIndex        =   144
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox Opcion41 
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
         Left            =   -69960
         TabIndex        =   143
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Opcion42 
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
         Left            =   -69960
         TabIndex        =   142
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Opcion43 
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
         Left            =   -69960
         TabIndex        =   141
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Opcion44 
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
         Left            =   -69960
         TabIndex        =   140
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Opcion113 
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
         Left            =   -70320
         TabIndex        =   136
         Top             =   4980
         Width           =   615
      End
      Begin VB.CheckBox Opcion114 
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
         Left            =   -70320
         TabIndex        =   135
         Top             =   5340
         Width           =   615
      End
      Begin VB.CheckBox Opcion115 
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
         Left            =   -70320
         TabIndex        =   134
         Top             =   5700
         Width           =   615
      End
      Begin VB.CheckBox Opcion316 
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
         Left            =   -69840
         TabIndex        =   121
         Top             =   6060
         Width           =   615
      End
      Begin VB.CheckBox Opcion315 
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
         Left            =   -69840
         TabIndex        =   120
         Top             =   5700
         Width           =   615
      End
      Begin VB.CheckBox Opcion314 
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
         Left            =   -69840
         TabIndex        =   119
         Top             =   5340
         Width           =   615
      End
      Begin VB.CheckBox Opcion313 
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
         Left            =   -69840
         TabIndex        =   118
         Top             =   4980
         Width           =   615
      End
      Begin VB.CheckBox Opcion320 
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
         Left            =   -64680
         TabIndex        =   117
         Top             =   1740
         Width           =   615
      End
      Begin VB.CheckBox Opcion319 
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
         Left            =   -64680
         TabIndex        =   116
         Top             =   1380
         Width           =   615
      End
      Begin VB.CheckBox Opcion318 
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
         Left            =   -64680
         TabIndex        =   115
         Top             =   1020
         Width           =   615
      End
      Begin VB.CheckBox Opcion317 
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
         Left            =   -64680
         TabIndex        =   114
         Top             =   660
         Width           =   615
      End
      Begin VB.CheckBox Opcion324 
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
         Left            =   -64680
         TabIndex        =   113
         Top             =   3180
         Width           =   615
      End
      Begin VB.CheckBox Opcion323 
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
         Left            =   -64680
         TabIndex        =   112
         Top             =   2820
         Width           =   615
      End
      Begin VB.CheckBox Opcion322 
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
         Left            =   -64680
         TabIndex        =   111
         Top             =   2460
         Width           =   615
      End
      Begin VB.CheckBox Opcion321 
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
         Left            =   -64680
         TabIndex        =   110
         Top             =   2100
         Width           =   615
      End
      Begin VB.CheckBox Opcion34 
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
         Left            =   -69840
         TabIndex        =   97
         Top             =   1740
         Width           =   615
      End
      Begin VB.CheckBox Opcion33 
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
         Left            =   -69840
         TabIndex        =   96
         Top             =   1380
         Width           =   615
      End
      Begin VB.CheckBox Opcion32 
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
         Left            =   -69840
         TabIndex        =   95
         Top             =   1020
         Width           =   615
      End
      Begin VB.CheckBox Opcion31 
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
         Left            =   -69840
         TabIndex        =   94
         Top             =   660
         Width           =   615
      End
      Begin VB.CheckBox Opcion38 
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
         Left            =   -69840
         TabIndex        =   93
         Top             =   3180
         Width           =   615
      End
      Begin VB.CheckBox Opcion37 
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
         Left            =   -69840
         TabIndex        =   92
         Top             =   2820
         Width           =   615
      End
      Begin VB.CheckBox Opcion36 
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
         Left            =   -69840
         TabIndex        =   91
         Top             =   2460
         Width           =   615
      End
      Begin VB.CheckBox Opcion35 
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
         Left            =   -69840
         TabIndex        =   90
         Top             =   2100
         Width           =   615
      End
      Begin VB.CheckBox Opcion312 
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
         Left            =   -69840
         TabIndex        =   89
         Top             =   4620
         Width           =   615
      End
      Begin VB.CheckBox Opcion311 
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
         Left            =   -69840
         TabIndex        =   88
         Top             =   4260
         Width           =   615
      End
      Begin VB.CheckBox Opcion310 
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
         Left            =   -69840
         TabIndex        =   87
         Top             =   3900
         Width           =   615
      End
      Begin VB.CheckBox Opcion39 
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
         Left            =   -69840
         TabIndex        =   86
         Top             =   3540
         Width           =   615
      End
      Begin VB.CheckBox Opcion216 
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
         Left            =   -70320
         TabIndex        =   73
         Top             =   6060
         Width           =   615
      End
      Begin VB.CheckBox Opcion215 
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
         Left            =   -70320
         TabIndex        =   72
         Top             =   5760
         Width           =   615
      End
      Begin VB.CheckBox Opcion214 
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
         Left            =   -70320
         TabIndex        =   71
         Top             =   5340
         Width           =   615
      End
      Begin VB.CheckBox Opcion213 
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
         Left            =   -70320
         TabIndex        =   70
         Top             =   4980
         Width           =   615
      End
      Begin VB.CheckBox Opcion220 
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
         Left            =   -64800
         TabIndex        =   69
         Top             =   1740
         Width           =   615
      End
      Begin VB.CheckBox Opcion219 
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
         Left            =   -64800
         TabIndex        =   68
         Top             =   1380
         Width           =   615
      End
      Begin VB.CheckBox Opcion218 
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
         Left            =   -64800
         TabIndex        =   67
         Top             =   1020
         Width           =   615
      End
      Begin VB.CheckBox Opcion217 
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
         Left            =   -64800
         TabIndex        =   66
         Top             =   660
         Width           =   615
      End
      Begin VB.CheckBox Opcion224 
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
         Left            =   -64800
         TabIndex        =   65
         Top             =   3180
         Width           =   615
      End
      Begin VB.CheckBox Opcion223 
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
         Left            =   -64800
         TabIndex        =   64
         Top             =   2820
         Width           =   615
      End
      Begin VB.CheckBox Opcion222 
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
         Left            =   -64800
         TabIndex        =   63
         Top             =   2460
         Width           =   615
      End
      Begin VB.CheckBox Opcion221 
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
         Left            =   -64800
         TabIndex        =   62
         Top             =   2100
         Width           =   615
      End
      Begin VB.CheckBox Opcion29 
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
         Left            =   -70320
         TabIndex        =   57
         Top             =   3540
         Width           =   615
      End
      Begin VB.CheckBox Opcion210 
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
         Left            =   -70320
         TabIndex        =   56
         Top             =   3900
         Width           =   615
      End
      Begin VB.CheckBox Opcion211 
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
         Left            =   -70320
         TabIndex        =   55
         Top             =   4260
         Width           =   615
      End
      Begin VB.CheckBox Opcion212 
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
         Left            =   -70320
         TabIndex        =   54
         Top             =   4620
         Width           =   615
      End
      Begin VB.CheckBox Opcion25 
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
         Left            =   -70320
         TabIndex        =   49
         Top             =   2100
         Width           =   615
      End
      Begin VB.CheckBox Opcion26 
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
         Left            =   -70320
         TabIndex        =   48
         Top             =   2460
         Width           =   615
      End
      Begin VB.CheckBox Opcion27 
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
         Left            =   -70320
         TabIndex        =   47
         Top             =   2820
         Width           =   615
      End
      Begin VB.CheckBox Opcion28 
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
         Left            =   -70320
         TabIndex        =   46
         Top             =   3180
         Width           =   615
      End
      Begin VB.CheckBox opcion21 
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
         Left            =   -70320
         TabIndex        =   41
         Top             =   660
         Width           =   615
      End
      Begin VB.CheckBox Opcion22 
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
         Left            =   -70320
         TabIndex        =   40
         Top             =   1020
         Width           =   615
      End
      Begin VB.CheckBox Opcion23 
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
         Left            =   -70320
         TabIndex        =   39
         Top             =   1380
         Width           =   615
      End
      Begin VB.CheckBox Opcion24 
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
         Left            =   -70320
         TabIndex        =   38
         Top             =   1740
         Width           =   615
      End
      Begin VB.CheckBox Opcion19 
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
         Left            =   -70320
         TabIndex        =   33
         Top             =   3540
         Width           =   615
      End
      Begin VB.CheckBox Opcion110 
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
         Left            =   -70320
         TabIndex        =   32
         Top             =   3900
         Width           =   615
      End
      Begin VB.CheckBox Opcion111 
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
         Left            =   -70320
         TabIndex        =   31
         Top             =   4260
         Width           =   615
      End
      Begin VB.CheckBox Opcion112 
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
         Left            =   -70320
         TabIndex        =   30
         Top             =   4620
         Width           =   615
      End
      Begin VB.CheckBox Opcion15 
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
         Left            =   -70320
         TabIndex        =   25
         Top             =   2100
         Width           =   615
      End
      Begin VB.CheckBox Opcion16 
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
         Left            =   -70320
         TabIndex        =   24
         Top             =   2460
         Width           =   615
      End
      Begin VB.CheckBox Opcion17 
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
         Left            =   -70320
         TabIndex        =   23
         Top             =   2820
         Width           =   615
      End
      Begin VB.CheckBox Opcion18 
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
         Left            =   -70320
         TabIndex        =   22
         Top             =   3180
         Width           =   615
      End
      Begin VB.CheckBox Opcion4 
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
         Left            =   4680
         TabIndex        =   21
         Top             =   1740
         Width           =   615
      End
      Begin VB.CheckBox Opcion3 
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
         Left            =   4680
         TabIndex        =   19
         Top             =   1380
         Width           =   615
      End
      Begin VB.CheckBox Opcion2 
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
         Left            =   4680
         TabIndex        =   17
         Top             =   1020
         Width           =   615
      End
      Begin VB.CheckBox Opcion1 
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
         Left            =   4680
         TabIndex        =   15
         Top             =   660
         Width           =   615
      End
      Begin VB.CheckBox Opcion14 
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
         Left            =   -70320
         TabIndex        =   13
         Top             =   1740
         Width           =   615
      End
      Begin VB.CheckBox Opcion13 
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
         Left            =   -70320
         TabIndex        =   11
         Top             =   1380
         Width           =   615
      End
      Begin VB.CheckBox Opcion12 
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
         Left            =   -70320
         TabIndex        =   9
         Top             =   1020
         Width           =   615
      End
      Begin VB.CheckBox Opcion11 
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
         Left            =   -70320
         TabIndex        =   7
         Top             =   660
         Width           =   615
      End
      Begin VB.Label titulo8 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   253
         Top             =   3120
         Width           =   3135
      End
      Begin VB.Label Titulo415 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -69360
         TabIndex        =   246
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Titulo228 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   240
         Top             =   4680
         Width           =   4335
      End
      Begin VB.Label Titulo229 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -69240
         TabIndex        =   239
         Top             =   5040
         Width           =   4335
      End
      Begin VB.Label Titulo5 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   235
         Top             =   2100
         Width           =   3855
      End
      Begin VB.Label Titulo6 
         Caption         =   "Ingreso de "
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
         Left            =   480
         TabIndex        =   234
         Top             =   2460
         Width           =   3975
      End
      Begin VB.Label Titulo7 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         TabIndex        =   233
         Top             =   2800
         Width           =   3855
      End
      Begin VB.Label Titulo131 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   228
         Top             =   6120
         Width           =   4695
      End
      Begin VB.Label Titulo227 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   222
         Top             =   4320
         Width           =   4335
      End
      Begin VB.Label Titulo226 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   220
         Top             =   3960
         Width           =   4335
      End
      Begin VB.Label Titulo130 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   218
         Top             =   5760
         Width           =   4695
      End
      Begin VB.Label Titulo129 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   216
         Top             =   5400
         Width           =   4695
      End
      Begin VB.Label Titulo128 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   213
         Top             =   5040
         Width           =   4695
      End
      Begin VB.Label Titulo410 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74520
         TabIndex        =   209
         Top             =   3960
         Width           =   4455
      End
      Begin VB.Label Titulo325 
         Caption         =   "Ingreso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   -69240
         TabIndex        =   207
         Top             =   3480
         Width           =   4455
      End
      Begin VB.Label Titulo127 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   205
         Top             =   4680
         Width           =   4695
      End
      Begin VB.Label Titulo124 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   203
         Top             =   3600
         Width           =   4935
      End
      Begin VB.Label Titulo125 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   202
         Top             =   3960
         Width           =   4695
      End
      Begin VB.Label Titulo126 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   201
         Top             =   4320
         Width           =   4695
      End
      Begin VB.Label Titulo123 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   197
         Top             =   3240
         Width           =   4695
      End
      Begin VB.Label Titulo49 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   195
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Label Titulo122 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   193
         Top             =   2880
         Width           =   4695
      End
      Begin VB.Label Titulo121 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   191
         Top             =   2520
         Width           =   4935
      End
      Begin VB.Label Titulo225 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   189
         Top             =   3600
         Width           =   4335
      End
      Begin VB.Label Titulo54 
         Caption         =   "Ingreso de "
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
         Left            =   -74640
         TabIndex        =   186
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Titulo53 
         Caption         =   "Ingrso de "
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
         Left            =   -74640
         TabIndex        =   185
         Top             =   1320
         Width           =   4935
      End
      Begin VB.Label Titulo52 
         Caption         =   "Ingreso de "
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
         Left            =   -74640
         TabIndex        =   184
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label Titulo51 
         Caption         =   "Ingrso de "
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
         Left            =   -74640
         TabIndex        =   183
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Titulo58 
         Caption         =   "Ingreso de "
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
         Left            =   -74640
         TabIndex        =   182
         Top             =   3120
         Width           =   4935
      End
      Begin VB.Label Titulo57 
         Caption         =   "Ingrso de "
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
         Left            =   -74640
         TabIndex        =   181
         Top             =   2760
         Width           =   4935
      End
      Begin VB.Label Titulo56 
         Caption         =   "Ingreso de "
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
         Left            =   -74640
         TabIndex        =   180
         Top             =   2400
         Width           =   4935
      End
      Begin VB.Label Titulo55 
         Caption         =   "Ingrso de "
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
         Left            =   -74640
         TabIndex        =   179
         Top             =   2040
         Width           =   4935
      End
      Begin VB.Label Titulo59 
         Caption         =   "Ingrso de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -74640
         TabIndex        =   178
         Top             =   3480
         Width           =   4935
      End
      Begin VB.Label Titulo120 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   165
         Top             =   2160
         Width           =   4935
      End
      Begin VB.Label Titulo119 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   163
         Top             =   1800
         Width           =   4935
      End
      Begin VB.Label Titulo118 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   161
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Label Titulo117 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   159
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label Titulo116 
         Caption         =   "Ingrso de "
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
         Left            =   -69720
         TabIndex        =   157
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label Titulo45 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   155
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label Titulo46 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   154
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label Titulo47 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   153
         Top             =   2880
         Width           =   4455
      End
      Begin VB.Label Titulo48 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   152
         Top             =   3240
         Width           =   4455
      End
      Begin VB.Label Titulo41 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   151
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Titulo42 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   150
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label Titulo43 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   149
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label Titulo44 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   148
         Top             =   1800
         Width           =   4455
      End
      Begin VB.Label Titulo113 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   139
         Top             =   4980
         Width           =   3975
      End
      Begin VB.Label Titulo114 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   138
         Top             =   5340
         Width           =   4095
      End
      Begin VB.Label Titulo115 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   137
         Top             =   5700
         Width           =   4095
      End
      Begin VB.Label Titulo316 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   133
         Top             =   6060
         Width           =   4455
      End
      Begin VB.Label Titulo315 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   132
         Top             =   5700
         Width           =   4575
      End
      Begin VB.Label Titulo314 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   131
         Top             =   5340
         Width           =   4455
      End
      Begin VB.Label Titulo313 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   130
         Top             =   4980
         Width           =   4335
      End
      Begin VB.Label Titulo320 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   129
         Top             =   1740
         Width           =   4455
      End
      Begin VB.Label Titulo319 
         Caption         =   "Ingrso de "
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
         Left            =   -69240
         TabIndex        =   128
         Top             =   1380
         Width           =   4455
      End
      Begin VB.Label Titulo318 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   127
         Top             =   1020
         Width           =   4455
      End
      Begin VB.Label Titulo317 
         Caption         =   "Ingrso de "
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
         Left            =   -69240
         TabIndex        =   126
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Titulo324 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   125
         Top             =   3180
         Width           =   4455
      End
      Begin VB.Label Titulo323 
         Caption         =   "Ingrso de "
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
         Left            =   -69240
         TabIndex        =   124
         Top             =   2820
         Width           =   4455
      End
      Begin VB.Label Titulo322 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   123
         Top             =   2460
         Width           =   4455
      End
      Begin VB.Label Titulo321 
         Caption         =   "Ingrso de "
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
         Left            =   -69240
         TabIndex        =   122
         Top             =   2100
         Width           =   4455
      End
      Begin VB.Label Titulo34 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   109
         Top             =   1740
         Width           =   4695
      End
      Begin VB.Label Titulo33 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   108
         Top             =   1380
         Width           =   4695
      End
      Begin VB.Label Titulo32 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   107
         Top             =   1020
         Width           =   4575
      End
      Begin VB.Label Titulo31 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   106
         Top             =   660
         Width           =   4575
      End
      Begin VB.Label Titulo38 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   105
         Top             =   3180
         Width           =   4695
      End
      Begin VB.Label Titulo37 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   104
         Top             =   2820
         Width           =   4695
      End
      Begin VB.Label Titulo36 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   103
         Top             =   2460
         Width           =   4815
      End
      Begin VB.Label Titulo35 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   102
         Top             =   2100
         Width           =   4695
      End
      Begin VB.Label Titulo312 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   101
         Top             =   4620
         Width           =   4695
      End
      Begin VB.Label Titulo311 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   100
         Top             =   4260
         Width           =   4815
      End
      Begin VB.Label Titulo310 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   99
         Top             =   3900
         Width           =   4695
      End
      Begin VB.Label Titulo39 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   98
         Top             =   3540
         Width           =   4695
      End
      Begin VB.Label Titulo216 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   85
         Top             =   6060
         Width           =   4335
      End
      Begin VB.Label Titulo215 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   84
         Top             =   5700
         Width           =   4335
      End
      Begin VB.Label Titulo214 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   83
         Top             =   5340
         Width           =   4335
      End
      Begin VB.Label Titulo213 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   82
         Top             =   4980
         Width           =   4335
      End
      Begin VB.Label Titulo220 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   81
         Top             =   1740
         Width           =   4335
      End
      Begin VB.Label Titulo219 
         Caption         =   "Ingrso de "
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
         Left            =   -69240
         TabIndex        =   80
         Top             =   1380
         Width           =   4215
      End
      Begin VB.Label Titulo218 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   79
         Top             =   1020
         Width           =   4335
      End
      Begin VB.Label Titulo217 
         Caption         =   "Ingrso de "
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
         Left            =   -69240
         TabIndex        =   78
         Top             =   660
         Width           =   4335
      End
      Begin VB.Label Titulo224 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   77
         Top             =   3180
         Width           =   4335
      End
      Begin VB.Label Titulo223 
         Caption         =   "Ingrso de "
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
         Left            =   -69240
         TabIndex        =   76
         Top             =   2820
         Width           =   4335
      End
      Begin VB.Label Titulo222 
         Caption         =   "Ingreso de "
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
         Left            =   -69240
         TabIndex        =   75
         Top             =   2460
         Width           =   4335
      End
      Begin VB.Label Titulo221 
         Caption         =   "Ingrso de "
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
         Left            =   -69240
         TabIndex        =   74
         Top             =   2100
         Width           =   4335
      End
      Begin VB.Label Titulo29 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   61
         Top             =   3540
         Width           =   3975
      End
      Begin VB.Label Titulo210 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   60
         Top             =   3900
         Width           =   4095
      End
      Begin VB.Label Titulo211 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   59
         Top             =   4260
         Width           =   4095
      End
      Begin VB.Label Titulo212 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   58
         Top             =   4620
         Width           =   3975
      End
      Begin VB.Label Titulo25 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   53
         Top             =   2100
         Width           =   4095
      End
      Begin VB.Label Titulo26 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   52
         Top             =   2460
         Width           =   3975
      End
      Begin VB.Label Titulo27 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   51
         Top             =   2820
         Width           =   4095
      End
      Begin VB.Label Titulo28 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   50
         Top             =   3180
         Width           =   4095
      End
      Begin VB.Label Titulo21 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   45
         Top             =   660
         Width           =   4095
      End
      Begin VB.Label Titulo22 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   44
         Top             =   1020
         Width           =   4095
      End
      Begin VB.Label Titulo23 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   43
         Top             =   1380
         Width           =   4095
      End
      Begin VB.Label Titulo24 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   42
         Top             =   1740
         Width           =   4095
      End
      Begin VB.Label Titulo19 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   37
         Top             =   3540
         Width           =   4095
      End
      Begin VB.Label Titulo110 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   36
         Top             =   3900
         Width           =   4095
      End
      Begin VB.Label Titulo111 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   35
         Top             =   4260
         Width           =   3855
      End
      Begin VB.Label Titulo112 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   34
         Top             =   4620
         Width           =   3975
      End
      Begin VB.Label Titulo15 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   29
         Top             =   2100
         Width           =   3855
      End
      Begin VB.Label Titulo16 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   28
         Top             =   2460
         Width           =   3975
      End
      Begin VB.Label Titulo17 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   27
         Top             =   2820
         Width           =   3735
      End
      Begin VB.Label Titulo18 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   26
         Top             =   3180
         Width           =   3975
      End
      Begin VB.Label Titulo4 
         Caption         =   "Ingreso de "
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
         Left            =   480
         TabIndex        =   20
         Top             =   1740
         Width           =   3975
      End
      Begin VB.Label Titulo3 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   18
         Top             =   1380
         Width           =   3855
      End
      Begin VB.Label Titulo2 
         Caption         =   "Ingreso de "
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
         Left            =   480
         TabIndex        =   16
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Titulo1 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   14
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Titulo14 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   12
         Top             =   1740
         Width           =   3855
      End
      Begin VB.Label Titulo13 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   10
         Top             =   1380
         Width           =   3975
      End
      Begin VB.Label Titulo12 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   8
         Top             =   1020
         Width           =   3855
      End
      Begin VB.Label Titulo11 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   6
         Top             =   660
         Width           =   3975
      End
   End
   Begin VB.Label DesOperador 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Operador"
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
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "PrgConfigCoti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOperador As Recordset
Dim spOperador As String
Dim rstAtributos As Recordset
Dim spAtributos As String
Dim XParam As String


Sub Form_Load()

    Operador.Text = ""
    DesOperador.Caption = ""
    
    Tablas.TabCaption(0) = "Archivos"
    Tablas.TabCaption(1) = "Novedades"
    Tablas.TabCaption(2) = "Listados"
    Tablas.TabCaption(3) = "Listados"
    Tablas.TabCaption(4) = "Listados DY"
    Tablas.TabCaption(5) = "Procesos"
    
    Rem titulo1
    
    Titulo1.Caption = "Ingreso de Envases"
    Titulo2.Caption = "Ingreso de Materias Primas"
    Titulo3.Caption = "Ingreso de Producto Terminado"
    Titulo4.Caption = "Ingreso de Proveedores"
    Titulo5.Caption = "Ingreso de Efluentes de Lavado"
    Titulo6.Caption = "Homologacion de Muestras de M.P."
    Titulo7.Caption = "Consulta de Evaluacion Semestral" _
            + Chr$(13) _
            + Chr$(13) _
            + "Actualizacion de Evaluacion Semestral" _
            + Chr$(13) _
            + Chr$(13) _
            + "Planilla de Evaluacion Semestra"
Rem by nan
    
    
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    Opcion5.Value = 0
    Opcion6.Value = 0
    Opcion7.Value = 0
    Opcion8.Value = 0
    Opcion9.Value = 0
    
    
    Rem titulo2
    
    Titulo11.Caption = "Ingreso de Cotizaciones"
    Titulo12.Caption = "Emision de Ordenes de Compra"
    Titulo13.Caption = "Ingreso de Informe de Recepcion"
    Titulo14.Caption = "Ingreso de Laudo de Liberacion"
    Titulo15.Caption = "Ingreso de Hoja de Produccion"
    Titulo16.Caption = "Ingreso de Movimientos Varios"
    Titulo17.Caption = "Ingreso y Egreso de Envases"
    Titulo18.Caption = "Emision de Etiquetas de Exportacion"
    Titulo19.Caption = "Ingreso de Guias de Traslado Interno"
    Titulo110.Caption = "Prestamos entre plantas"
    Titulo111.Caption = "Actualizacion de Pedidos"
    Titulo112.Caption = "Actualizacion de Pedidos de Exportacion"
    Titulo113.Caption = "Ingreso de Solicitud de Pedido de Compra"
    Titulo114.Caption = "Consulta de Solicitiudes de Pedido de Compra"
    Titulo115.Caption = "Entrada de Devolucion de Mercaderia"
    Titulo116.Caption = "Ingreso de Solicitud de Hoja de Produccion"
    Titulo117.Caption = "Actualizacion de Pedidos de Colorantes"
    Titulo118.Caption = "Verificacion de Pedidos"
    Titulo119.Caption = "Ingreso de Solidictud de Guia de Traslado Interno"
    Titulo120.Caption = "Consultas de Guias de Traslado Interno"
    Titulo121.Caption = "Actualizacion de Informes de Recepcion"
    Titulo122.Caption = "Depuracion de Saldos de Materia Prima"
    Titulo123.Caption = "Depuracion de Saldos de Producto Terminado"
    Titulo124.Caption = "Ingreso de Proyeccion de Ventas Anuales de Dy"
    Titulo125.Caption = "Ingreso de Solicitud de Compras de Insumos"
    Titulo126.Caption = "Consulta de Solicitudes de Compras de Insumos"
    Titulo127.Caption = "Consulta de Solicitudes por Origen"
    Titulo128.Caption = "Verificacion de Pedidos Pendientes"
    Titulo129.Caption = "Carga de Solicitud de Produccion"
    Titulo130.Caption = "Recepcion de Produccion"
    Titulo131.Caption = "Actualizacion de Hojas de Produccion"
    
    Opcion11.Value = 0
    Opcion12.Value = 0
    Opcion13.Value = 0
    Opcion14.Value = 0
    Opcion15.Value = 0
    Opcion16.Value = 0
    Opcion17.Value = 0
    Opcion18.Value = 0
    Opcion19.Value = 0
    Opcion110.Value = 0
    Opcion111.Value = 0
    Opcion112.Value = 0
    Opcion113.Value = 0
    Opcion114.Value = 0
    Opcion115.Value = 0
    Opcion116.Value = 0
    Opcion117.Value = 0
    Opcion118.Value = 0
    Opcion119.Value = 0
    Opcion120.Value = 0
    Opcion121.Value = 0
    Opcion122.Value = 0
    Opcion123.Value = 0
    Opcion124.Value = 0
    Opcion125.Value = 0
    Opcion126.Value = 0
    Opcion127.Value = 0
    Opcion128.Value = 0
    Opcion129.Value = 0
    Opcion130.Value = 0
    Opcion131.Value = 0

    
    Rem titulo3
    
    Titulo21.Caption = "Listado de Cotizaciones"
    Titulo22.Caption = "Listado de Ordenes de Compra"
    Titulo23.Caption = "Listado de Cotizaciones por Proveedor"
    Titulo24.Caption = "Listado de Cotizaciones por Articulo"
    Titulo25.Caption = "Listado de O/C Pend. por Proveedor"
    Titulo26.Caption = "Listado de O/C Pend. por Articulos"
    Titulo27.Caption = "Listado de Materia Prima"
    Titulo28.Caption = "Listado de Materia Prima ( Stock )"
    Titulo29.Caption = "Listado de Producto Terminado (Stock)"
    Titulo210.Caption = "Listado de Valuacion de Materia Prima"
    Titulo211.Caption = "Listado de Valuacion de Producto Terminado"
    Titulo212.Caption = "Listado de Materia Prima (Minimo)"
    Titulo213.Caption = "Listado de Producto Terminado (Minimo)"
    Titulo214.Caption = "Listado de Composicion"
    Titulo215.Caption = "Listado de Proyeccion de Entradas"
    Titulo216.Caption = "Listado de Ficha de Stock de M.P."
    Titulo217.Caption = "Listado de Ficha de Stock de P.T."
    Titulo218.Caption = "Listado de Movimientos Varios de Materia Prima"
    Titulo219.Caption = "Listado de Movimientos Varios de P.Terminado"
    Titulo220.Caption = "Listado de Hojas de Produccion"
    Titulo221.Caption = "Listado de Materia Prima (Minimo Consolidado)"
    Titulo222.Caption = "Listado de P.Terminado (Minimo Consolidado)"
    Titulo223.Caption = "Listado de Consumo de Productos Terminados"
    Titulo224.Caption = "Listado de Consumo de Materia Prima"
    Titulo225.Caption = "Listado de Disponibilidad de Producto Terminado"
    Titulo226.Caption = "Listado de Solicitudes de Produccion Pendientes"
    Titulo227.Caption = "Listado de Disponibilidad de P.T. (Pellital)"
    Titulo228.Caption = "Analisis de Materia Prima"
    Titulo229.Caption = "Analisis de Producto Terminado" + Chr$(13) + _
                         "" + Chr$(13) + _
                         "Listado de Valorizacion de Importaciones a Fecha"

    opcion21.Value = 0
    Opcion22.Value = 0
    Opcion23.Value = 0
    Opcion24.Value = 0
    Opcion25.Value = 0
    Opcion26.Value = 0
    Opcion27.Value = 0
    Opcion28.Value = 0
    Opcion29.Value = 0
    Opcion210.Value = 0
    Opcion211.Value = 0
    Opcion212.Value = 0
    Opcion213.Value = 0
    Opcion214.Value = 0
    Opcion215.Value = 0
    Opcion216.Value = 0
    Opcion217.Value = 0
    Opcion218.Value = 0
    Opcion219.Value = 0
    Opcion220.Value = 0
    Opcion221.Value = 0
    Opcion222.Value = 0
    Opcion223.Value = 0
    Opcion224.Value = 0
    Opcion225.Value = 0
    Opcion226.Value = 0
    Opcion227.Value = 0
    Opcion228.Value = 0
    Opcion229.Value = 0
    Opcion230.Value = 0
    
    
    
    Rem titulo4
    
    Titulo31.Caption = "Listado de Compras por Proveedor"
    Titulo32.Caption = "Listado de Compras por Materia Prima"
    Titulo33.Caption = "Listado de Compras por Materia Prima Concolidada"
    Titulo34.Caption = "Listado de Informe de Recepcion"
    Titulo35.Caption = "Listado de Informe de Recepcion Pend. Aprobacion"
    Titulo36.Caption = "Listado de Control de Ordenes"
    Titulo37.Caption = "Consulta de Ficha de Stock M.P."
    Titulo38.Caption = "Consulta de Ficha de Stock P.T."
    Titulo39.Caption = "Listado de Ultima Compra de Materia Prima"
    Titulo310.Caption = "Emision de Etiquetas"
    Titulo311.Caption = "Emision de Etiquetas DY"
    Titulo312.Caption = "Listado de Envases por Cliente"
    Titulo313.Caption = "Listado de Envases por Envases"
    Titulo314.Caption = "Listado de Verificacion de Correlatividades"
    Titulo315.Caption = "Listado de Componentes de Formulas (M.P.)"
    Titulo316.Caption = "Listado de Componentes de Formulas (P.T.)"
    Titulo317.Caption = "Listado de Ficha de Lote de Materia Prima"
    Titulo318.Caption = "Listado de Ficha de Lote de Producto Terminado"
    Titulo319.Caption = "Listado de Pedidos Pendientes"
    Titulo320.Caption = "Listado de Costos por Partida"
    Titulo321.Caption = "Listado de Valorizacion de Stock de M.P. a fecha"
    Titulo322.Caption = "Listado de Valorizacion de Stock de P.T. a fecha"
    Titulo323.Caption = "Consulta de Ficha de Materia Prima Historica"
    Titulo324.Caption = "Consulta de Ficha de Producto Terminado Historico"
    Titulo325.Caption = "Pedidos Pendientes por Producto Terminado" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Analisis de Consumo de Materia Prima" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Analisis de Consumo de Producto Termnado" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Listado de Analisis de Ordenes de Compra de Importacion" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Listado de Verificacion de Ultimos Movimientos de P.T." + Chr$(13) + _
                        "Listado de Verificacion de Ultimos Movimientos de M.P." + Chr$(13) + _
                        "Emision de Etiquetas Verdes"

    Opcion31.Value = 0
    Opcion32.Value = 0
    Opcion33.Value = 0
    Opcion34.Value = 0
    Opcion35.Value = 0
    Opcion36.Value = 0
    Opcion37.Value = 0
    Opcion38.Value = 0
    Opcion39.Value = 0
    Opcion310.Value = 0
    Opcion311.Value = 0
    Opcion312.Value = 0
    Opcion313.Value = 0
    Opcion314.Value = 0
    Opcion315.Value = 0
    Opcion316.Value = 0
    Opcion317.Value = 0
    Opcion318.Value = 0
    Opcion319.Value = 0
    Opcion320.Value = 0
    Opcion321.Value = 0
    Opcion322.Value = 0
    Opcion323.Value = 0
    Opcion324.Value = 0
    Opcion325.Value = 0
    Opcion326.Value = 0
    Opcion327.Value = 0
    Opcion328.Value = 0
    Opcion329.Value = 0
    Opcion330.Value = 0
    Opcion331.Value = 0
   
    
    
    
    Rem titulo5
    
    Titulo41.Caption = "Consulta de Fichas de Stock"
    Titulo42.Caption = "Listado de Fichas de Stock"
    Titulo43.Caption = "Listado de Stock"
    Titulo44.Caption = "Listado de Pedido Pendiente M.P."
    Titulo45.Caption = "Listado de Stock por Familia"
    Titulo46.Caption = "Listado de Disponibilidad de Stock"
    Titulo47.Caption = "Listado de Stock Minimo"
    Titulo48.Caption = "Proyeccion de Stock"
    Titulo49.Caption = "Listado de Ordenes Compras Pendientes de M.P."
    Titulo410.Caption = "Listado de Proyeccion de M.P." + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Listado de Ordenes de Compra Anuales" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Listado de Proyeccion de Ordenes de Compra" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Actualizacion de Proyeccion de Ordenes de Compra" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Listado de Ordenes Compras Pendientes de Dy"
    Titulo415.Caption = "Ingreso de Solicitud de Mecaderia a Zona Franca" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Listado de Stock por Familia DW" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Listado de Pedidos de Importacion" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Listado de Analisis de Cortes de Pedidos" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Listado de Stock por Familia DS" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Listado de Stock por Familia DQ"

    
    Opcion41.Value = 0
    Opcion42.Value = 0
    Opcion43.Value = 0
    Opcion44.Value = 0
    Opcion45.Value = 0
    Opcion46.Value = 0
    Opcion47.Value = 0
    Opcion48.Value = 0
    Opcion49.Value = 0
    Opcion410.Value = 0
    Opcion411.Value = 0
    Opcion412.Value = 0
    Opcion413.Value = 0
    Opcion414.Value = 0
    Opcion415.Value = 0
    Opcion416.Value = 0
    Opcion417.Value = 0
    Opcion418.Value = 0
    Opcion419.Value = 0
    Opcion420.Value = 0

    
    Rem titulo6
    
    Titulo51.Caption = "Cierre del Stock"
    Titulo52.Caption = "Reproceso de Materia Prima"
    Titulo53.Caption = "Reproceso de Producto Terminado"
    Titulo54.Caption = "Reproceso de Pedidos"
    Titulo55.Caption = "Pasa de CO a DY"
    Titulo56.Caption = "Reproceso de Fechas de Laudos y Hojas"
    Titulo57.Caption = "Generacion de NK y Re"
    Titulo58.Caption = "Verificacion de MP"
    Titulo59.Caption = "Veriricacion de PT" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Verificacion de Lotes de M.P." + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Verificacion de Lotes de P.T." + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Verificacion de Hojas de Produccion" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Fin del Sistema" + Chr$(13) + _
                        "" + Chr$(13) + _
                        "Pasa Pt a DW"
    
    Opcion51.Value = 0
    Opcion52.Value = 0
    Opcion53.Value = 0
    Opcion54.Value = 0
    Opcion55.Value = 0
    Opcion56.Value = 0
    Opcion57.Value = 0
    Opcion58.Value = 0
    Opcion59.Value = 0
    Opcion510.Value = 0
    Opcion511.Value = 0
    Opcion512.Value = 0
    Opcion513.Value = 0
    Opcion514.Value = 0

End Sub


Private Sub Graba_Click()

    XParam = "'" + Operador.Text + "','" _
                 + "1" + "'"
    spAtributos = "BorrarAtributos " + XParam
    Set rstAtributos = db.OpenRecordset(spAtributos, dbOpenSnapshot, dbSQLPassThrough)
    
    WAtributo1 = ""
    WAtributo2 = ""
    WAtributo3 = ""
    WAtributo4 = ""
    WAtributo5 = ""
    WAtributo6 = ""
    WAtributo7 = ""
    WAtributo8 = ""
    WAtributo9 = ""
    WAtributo10 = ""
    
    
    
    If Opcion1.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion2.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion3.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion4.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion5.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion6.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion7.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion8.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion9.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    
    
    
    If Opcion11.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion12.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion13.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion14.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion15.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion16.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion17.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion18.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion19.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion110.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion111.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion112.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion113.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion114.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion115.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion116.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion117.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion118.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion119.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion120.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion121.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion122.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion123.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion124.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion125.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion126.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion127.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion128.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion129.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion130.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion131.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    
    
    
    
    If opcion21.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion22.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion23.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion24.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion25.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion26.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion27.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion28.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion29.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion210.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion211.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion212.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion213.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion214.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion215.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion216.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion217.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion218.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion219.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion220.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion221.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion222.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion223.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion224.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion225.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion226.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion227.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion228.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion229.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion230.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    
    
    
    If Opcion31.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion32.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion33.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion34.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion35.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion36.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion37.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion38.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion39.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion310.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion311.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion312.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion313.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion314.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion315.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion316.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion317.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion318.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion319.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion320.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion321.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion322.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion323.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion324.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion325.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion326.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion327.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion328.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion329.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion330.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion331.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    
    
    
    If Opcion41.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion42.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion43.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion44.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion45.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion46.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion47.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion48.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion49.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion410.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion411.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion412.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion413.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion414.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion415.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion416.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion417.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion418.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion419.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    If Opcion420.Value = 0 Then
        WAtributo5 = WAtributo5 + "0"
            Else
        WAtributo5 = WAtributo5 + "1"
    End If
    
    
    
    If Opcion51.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion52.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion53.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion54.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion55.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion56.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion57.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion58.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion59.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion510.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion511.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion512.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion513.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    If Opcion514.Value = 0 Then
        WAtributo6 = WAtributo6 + "0"
            Else
        WAtributo6 = WAtributo6 + "1"
    End If
    
    
    WProceso = "1"
                                       
    XParam = "'" + Operador.Text + "','" _
                 + WProceso + "','" _
                 + WAtributo1 + "','" _
                 + WAtributo2 + "','" _
                 + WAtributo3 + "','" _
                 + WAtributo4 + "','" _
                 + WAtributo5 + "','" _
                 + WAtributo6 + "','" _
                 + WAtributo7 + "','" _
                 + WAtributo8 + "','" _
                 + WAtributo9 + "','" _
                 + WAtributo10 + "'"
                    
    spAtributos = "AltaAtributos " + XParam
    Set rstAtributos = db.OpenRecordset(spAtributos, dbOpenSnapshot, dbSQLPassThrough)
    
    Operador.Text = ""
    DesOperador.Caption = ""
    
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    Opcion5.Value = 0
    Opcion6.Value = 0
    Opcion7.Value = 0
    Opcion8.Value = 0
    Opcion9.Value = 0
                    
    Opcion11.Value = 0
    Opcion12.Value = 0
    Opcion13.Value = 0
    Opcion14.Value = 0
    Opcion15.Value = 0
    Opcion16.Value = 0
    Opcion17.Value = 0
    Opcion18.Value = 0
    Opcion19.Value = 0
    Opcion110.Value = 0
    Opcion111.Value = 0
    Opcion112.Value = 0
    Opcion113.Value = 0
    Opcion114.Value = 0
    Opcion115.Value = 0
    Opcion116.Value = 0
    Opcion117.Value = 0
    Opcion118.Value = 0
    Opcion119.Value = 0
    Opcion120.Value = 0
    Opcion121.Value = 0
    Opcion122.Value = 0
    Opcion123.Value = 0
    Opcion124.Value = 0
    Opcion125.Value = 0
    Opcion126.Value = 0
    Opcion127.Value = 0
    Opcion128.Value = 0
    Opcion129.Value = 0
    Opcion130.Value = 0
    Opcion131.Value = 0

    opcion21.Value = 0
    Opcion22.Value = 0
    Opcion23.Value = 0
    Opcion24.Value = 0
    Opcion25.Value = 0
    Opcion26.Value = 0
    Opcion27.Value = 0
    Opcion28.Value = 0
    Opcion29.Value = 0
    Opcion210.Value = 0
    Opcion211.Value = 0
    Opcion212.Value = 0
    Opcion213.Value = 0
    Opcion214.Value = 0
    Opcion215.Value = 0
    Opcion216.Value = 0
    Opcion217.Value = 0
    Opcion218.Value = 0
    Opcion219.Value = 0
    Opcion220.Value = 0
    Opcion221.Value = 0
    Opcion222.Value = 0
    Opcion223.Value = 0
    Opcion224.Value = 0
    Opcion225.Value = 0
    Opcion226.Value = 0
    Opcion227.Value = 0
    Opcion228.Value = 0
    Opcion229.Value = 0
    Opcion230.Value = 0
                    
    Opcion31.Value = 0
    Opcion32.Value = 0
    Opcion33.Value = 0
    Opcion34.Value = 0
    Opcion35.Value = 0
    Opcion36.Value = 0
    Opcion37.Value = 0
    Opcion38.Value = 0
    Opcion39.Value = 0
    Opcion310.Value = 0
    Opcion311.Value = 0
    Opcion312.Value = 0
    Opcion313.Value = 0
    Opcion314.Value = 0
    Opcion315.Value = 0
    Opcion316.Value = 0
    Opcion317.Value = 0
    Opcion318.Value = 0
    Opcion319.Value = 0
    Opcion320.Value = 0
    Opcion321.Value = 0
    Opcion322.Value = 0
    Opcion323.Value = 0
    Opcion324.Value = 0
    Opcion325.Value = 0
    Opcion326.Value = 0
    Opcion327.Value = 0
    Opcion328.Value = 0
    Opcion329.Value = 0
    Opcion330.Value = 0
    Opcion331.Value = 0
                    
                    
    Opcion41.Value = 0
    Opcion42.Value = 0
    Opcion43.Value = 0
    Opcion44.Value = 0
    Opcion45.Value = 0
    Opcion46.Value = 0
    Opcion47.Value = 0
    Opcion48.Value = 0
    Opcion49.Value = 0
    Opcion410.Value = 0
    Opcion411.Value = 0
    Opcion412.Value = 0
    Opcion413.Value = 0
    Opcion414.Value = 0
    Opcion415.Value = 0
    Opcion416.Value = 0
    Opcion417.Value = 0
    Opcion418.Value = 0
    Opcion419.Value = 0
    Opcion420.Value = 0
                    
    Opcion51.Value = 0
    Opcion52.Value = 0
    Opcion53.Value = 0
    Opcion54.Value = 0
    Opcion55.Value = 0
    Opcion56.Value = 0
    Opcion57.Value = 0
    Opcion58.Value = 0
    Opcion59.Value = 0
    Opcion510.Value = 0
    Opcion511.Value = 0
    Opcion512.Value = 0
    Opcion513.Value = 0
    Opcion514.Value = 0
    
    Operador.SetFocus
    
    Tablas.Tab = 0

End Sub

Private Sub Operador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Operador.Text <> "" Then
        
            spOperador = "ConsultaOperador " + "'" + Operador.Text + "'"
            Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperador.RecordCount > 0 Then
                DesOperador.Caption = rstOperador!Descripcion
                rstOperador.Close
                
                Opcion1.Value = 0
                Opcion2.Value = 0
                Opcion3.Value = 0
                Opcion4.Value = 0
                Opcion5.Value = 0
                Opcion6.Value = 0
                Opcion7.Value = 0
                Opcion8.Value = 0
                Opcion9.Value = 0
                    
                Opcion11.Value = 0
                Opcion12.Value = 0
                Opcion13.Value = 0
                Opcion14.Value = 0
                Opcion15.Value = 0
                Opcion16.Value = 0
                Opcion17.Value = 0
                Opcion18.Value = 0
                Opcion19.Value = 0
                Opcion110.Value = 0
                Opcion111.Value = 0
                Opcion112.Value = 0
                Opcion113.Value = 0
                Opcion114.Value = 0
                Opcion115.Value = 0
                Opcion116.Value = 0
                Opcion117.Value = 0
                Opcion118.Value = 0
                Opcion119.Value = 0
                Opcion120.Value = 0
                Opcion121.Value = 0
                Opcion122.Value = 0
                Opcion123.Value = 0
                Opcion124.Value = 0
                Opcion125.Value = 0
                Opcion126.Value = 0
                Opcion127.Value = 0
                Opcion128.Value = 0
                Opcion129.Value = 0
                Opcion130.Value = 0
                Opcion131.Value = 0

                opcion21.Value = 0
                Opcion22.Value = 0
                Opcion23.Value = 0
                Opcion24.Value = 0
                Opcion25.Value = 0
                Opcion26.Value = 0
                Opcion27.Value = 0
                Opcion28.Value = 0
                Opcion29.Value = 0
                Opcion210.Value = 0
                Opcion211.Value = 0
                Opcion212.Value = 0
                Opcion213.Value = 0
                Opcion214.Value = 0
                Opcion215.Value = 0
                Opcion216.Value = 0
                Opcion217.Value = 0
                Opcion218.Value = 0
                Opcion219.Value = 0
                Opcion220.Value = 0
                Opcion221.Value = 0
                Opcion222.Value = 0
                Opcion223.Value = 0
                Opcion224.Value = 0
                Opcion225.Value = 0
                Opcion226.Value = 0
                Opcion227.Value = 0
                Opcion228.Value = 0
                Opcion229.Value = 0
                Opcion230.Value = 0
                    
                Opcion31.Value = 0
                Opcion32.Value = 0
                Opcion33.Value = 0
                Opcion34.Value = 0
                Opcion35.Value = 0
                Opcion36.Value = 0
                Opcion37.Value = 0
                Opcion38.Value = 0
                Opcion39.Value = 0
                Opcion310.Value = 0
                Opcion311.Value = 0
                Opcion312.Value = 0
                Opcion313.Value = 0
                Opcion314.Value = 0
                Opcion315.Value = 0
                Opcion316.Value = 0
                Opcion317.Value = 0
                Opcion318.Value = 0
                Opcion319.Value = 0
                Opcion320.Value = 0
                Opcion321.Value = 0
                Opcion322.Value = 0
                Opcion323.Value = 0
                Opcion324.Value = 0
                Opcion325.Value = 0
                Opcion326.Value = 0
                Opcion327.Value = 0
                Opcion328.Value = 0
                Opcion329.Value = 0
                Opcion330.Value = 0
                Opcion331.Value = 0
                    
                    
                Opcion41.Value = 0
                Opcion42.Value = 0
                Opcion43.Value = 0
                Opcion44.Value = 0
                Opcion45.Value = 0
                Opcion46.Value = 0
                Opcion47.Value = 0
                Opcion48.Value = 0
                Opcion49.Value = 0
                Opcion410.Value = 0
                Opcion411.Value = 0
                Opcion412.Value = 0
                Opcion413.Value = 0
                Opcion414.Value = 0
                Opcion415.Value = 0
                Opcion416.Value = 0
                Opcion417.Value = 0
                Opcion418.Value = 0
                Opcion419.Value = 0
                Opcion420.Value = 0
                    
                Opcion51.Value = 0
                Opcion52.Value = 0
                Opcion53.Value = 0
                Opcion54.Value = 0
                Opcion55.Value = 0
                Opcion56.Value = 0
                Opcion57.Value = 0
                Opcion58.Value = 0
                Opcion59.Value = 0
                Opcion510.Value = 0
                Opcion511.Value = 0
                Opcion512.Value = 0
                Opcion513.Value = 0
                Opcion514.Value = 0
                
                
                XParam = "'" + Operador.Text + "','" _
                             + "1" + "'"
                spAtributos = "ConsultaAtributo " + XParam
                Set rstAtributos = db.OpenRecordset(spAtributos, dbOpenSnapshot, dbSQLPassThrough)
                If rstAtributos.RecordCount > 0 Then
                
                    Opcion1.Value = Val(Mid$(rstAtributos!atributo1, 1, 1))
                    Opcion2.Value = Val(Mid$(rstAtributos!atributo1, 2, 1))
                    Opcion3.Value = Val(Mid$(rstAtributos!atributo1, 3, 1))
                    Opcion4.Value = Val(Mid$(rstAtributos!atributo1, 4, 1))
                    Opcion5.Value = Val(Mid$(rstAtributos!atributo1, 5, 1))
                    Opcion6.Value = Val(Mid$(rstAtributos!atributo1, 6, 1))
                    Opcion7.Value = Val(Mid$(rstAtributos!atributo1, 7, 1))
                    Opcion8.Value = Val(Mid$(rstAtributos!atributo1, 8, 1))
                    Opcion9.Value = Val(Mid$(rstAtributos!atributo1, 9, 1))
                    
                    Opcion11.Value = Val(Mid$(rstAtributos!atributo2, 1, 1))
                    Opcion12.Value = Val(Mid$(rstAtributos!atributo2, 2, 1))
                    Opcion13.Value = Val(Mid$(rstAtributos!atributo2, 3, 1))
                    Opcion14.Value = Val(Mid$(rstAtributos!atributo2, 4, 1))
                    Opcion15.Value = Val(Mid$(rstAtributos!atributo2, 5, 1))
                    Opcion16.Value = Val(Mid$(rstAtributos!atributo2, 6, 1))
                    Opcion17.Value = Val(Mid$(rstAtributos!atributo2, 7, 1))
                    Opcion18.Value = Val(Mid$(rstAtributos!atributo2, 8, 1))
                    Opcion19.Value = Val(Mid$(rstAtributos!atributo2, 9, 1))
                    Opcion110.Value = Val(Mid$(rstAtributos!atributo2, 10, 1))
                    Opcion111.Value = Val(Mid$(rstAtributos!atributo2, 11, 1))
                    Opcion112.Value = Val(Mid$(rstAtributos!atributo2, 12, 1))
                    Opcion113.Value = Val(Mid$(rstAtributos!atributo2, 13, 1))
                    Opcion114.Value = Val(Mid$(rstAtributos!atributo2, 14, 1))
                    Opcion115.Value = Val(Mid$(rstAtributos!atributo2, 15, 1))
                    Opcion116.Value = Val(Mid$(rstAtributos!atributo2, 16, 1))
                    Opcion117.Value = Val(Mid$(rstAtributos!atributo2, 17, 1))
                    Opcion118.Value = Val(Mid$(rstAtributos!atributo2, 18, 1))
                    Opcion119.Value = Val(Mid$(rstAtributos!atributo2, 19, 1))
                    Opcion120.Value = Val(Mid$(rstAtributos!atributo2, 20, 1))
                    Opcion121.Value = Val(Mid$(rstAtributos!atributo2, 21, 1))
                    Opcion122.Value = Val(Mid$(rstAtributos!atributo2, 22, 1))
                    Opcion123.Value = Val(Mid$(rstAtributos!atributo2, 23, 1))
                    Opcion124.Value = Val(Mid$(rstAtributos!atributo2, 24, 1))
                    Opcion125.Value = Val(Mid$(rstAtributos!atributo2, 25, 1))
                    Opcion126.Value = Val(Mid$(rstAtributos!atributo2, 26, 1))
                    Opcion127.Value = Val(Mid$(rstAtributos!atributo2, 27, 1))
                    Opcion128.Value = Val(Mid$(rstAtributos!atributo2, 28, 1))
                    Opcion129.Value = Val(Mid$(rstAtributos!atributo2, 29, 1))
                    Opcion130.Value = Val(Mid$(rstAtributos!atributo2, 30, 1))
                    Opcion131.Value = Val(Mid$(rstAtributos!atributo2, 31, 1))
                    
                    opcion21.Value = Val(Mid$(rstAtributos!atributo3, 1, 1))
                    Opcion22.Value = Val(Mid$(rstAtributos!atributo3, 2, 1))
                    Opcion23.Value = Val(Mid$(rstAtributos!atributo3, 3, 1))
                    Opcion24.Value = Val(Mid$(rstAtributos!atributo3, 4, 1))
                    Opcion25.Value = Val(Mid$(rstAtributos!atributo3, 5, 1))
                    Opcion26.Value = Val(Mid$(rstAtributos!atributo3, 6, 1))
                    Opcion27.Value = Val(Mid$(rstAtributos!atributo3, 7, 1))
                    Opcion28.Value = Val(Mid$(rstAtributos!atributo3, 8, 1))
                    Opcion29.Value = Val(Mid$(rstAtributos!atributo3, 9, 1))
                    Opcion210.Value = Val(Mid$(rstAtributos!atributo3, 10, 1))
                    Opcion211.Value = Val(Mid$(rstAtributos!atributo3, 11, 1))
                    Opcion212.Value = Val(Mid$(rstAtributos!atributo3, 12, 1))
                    Opcion213.Value = Val(Mid$(rstAtributos!atributo3, 13, 1))
                    Opcion214.Value = Val(Mid$(rstAtributos!atributo3, 14, 1))
                    Opcion215.Value = Val(Mid$(rstAtributos!atributo3, 15, 1))
                    Opcion216.Value = Val(Mid$(rstAtributos!atributo3, 16, 1))
                    Opcion217.Value = Val(Mid$(rstAtributos!atributo3, 17, 1))
                    Opcion218.Value = Val(Mid$(rstAtributos!atributo3, 18, 1))
                    Opcion219.Value = Val(Mid$(rstAtributos!atributo3, 19, 1))
                    Opcion220.Value = Val(Mid$(rstAtributos!atributo3, 20, 1))
                    Opcion221.Value = Val(Mid$(rstAtributos!atributo3, 21, 1))
                    Opcion222.Value = Val(Mid$(rstAtributos!atributo3, 22, 1))
                    Opcion223.Value = Val(Mid$(rstAtributos!atributo3, 23, 1))
                    Opcion224.Value = Val(Mid$(rstAtributos!atributo3, 24, 1))
                    Opcion225.Value = Val(Mid$(rstAtributos!atributo3, 25, 1))
                    Opcion226.Value = Val(Mid$(rstAtributos!atributo3, 26, 1))
                    Opcion227.Value = Val(Mid$(rstAtributos!atributo3, 27, 1))
                    Opcion228.Value = Val(Mid$(rstAtributos!atributo3, 28, 1))
                    Opcion229.Value = Val(Mid$(rstAtributos!atributo3, 29, 1))
                    Opcion230.Value = Val(Mid$(rstAtributos!atributo3, 30, 1))
                    
                    Opcion31.Value = Val(Mid$(rstAtributos!atributo4, 1, 1))
                    Opcion32.Value = Val(Mid$(rstAtributos!atributo4, 2, 1))
                    Opcion33.Value = Val(Mid$(rstAtributos!atributo4, 3, 1))
                    Opcion34.Value = Val(Mid$(rstAtributos!atributo4, 4, 1))
                    Opcion35.Value = Val(Mid$(rstAtributos!atributo4, 5, 1))
                    Opcion36.Value = Val(Mid$(rstAtributos!atributo4, 6, 1))
                    Opcion37.Value = Val(Mid$(rstAtributos!atributo4, 7, 1))
                    Opcion38.Value = Val(Mid$(rstAtributos!atributo4, 8, 1))
                    Opcion39.Value = Val(Mid$(rstAtributos!atributo4, 9, 1))
                    Opcion310.Value = Val(Mid$(rstAtributos!atributo4, 10, 1))
                    Opcion311.Value = Val(Mid$(rstAtributos!atributo4, 11, 1))
                    Opcion312.Value = Val(Mid$(rstAtributos!atributo4, 12, 1))
                    Opcion313.Value = Val(Mid$(rstAtributos!atributo4, 13, 1))
                    Opcion314.Value = Val(Mid$(rstAtributos!atributo4, 14, 1))
                    Opcion315.Value = Val(Mid$(rstAtributos!atributo4, 15, 1))
                    Opcion316.Value = Val(Mid$(rstAtributos!atributo4, 16, 1))
                    Opcion317.Value = Val(Mid$(rstAtributos!atributo4, 17, 1))
                    Opcion318.Value = Val(Mid$(rstAtributos!atributo4, 18, 1))
                    Opcion319.Value = Val(Mid$(rstAtributos!atributo4, 19, 1))
                    Opcion320.Value = Val(Mid$(rstAtributos!atributo4, 20, 1))
                    Opcion321.Value = Val(Mid$(rstAtributos!atributo4, 21, 1))
                    Opcion322.Value = Val(Mid$(rstAtributos!atributo4, 22, 1))
                    Opcion323.Value = Val(Mid$(rstAtributos!atributo4, 23, 1))
                    Opcion324.Value = Val(Mid$(rstAtributos!atributo4, 24, 1))
                    Opcion325.Value = Val(Mid$(rstAtributos!atributo4, 25, 1))
                    Opcion326.Value = Val(Mid$(rstAtributos!atributo4, 26, 1))
                    Opcion327.Value = Val(Mid$(rstAtributos!atributo4, 27, 1))
                    Opcion328.Value = Val(Mid$(rstAtributos!atributo4, 28, 1))
                    Opcion329.Value = Val(Mid$(rstAtributos!atributo4, 29, 1))
                    Opcion330.Value = Val(Mid$(rstAtributos!atributo4, 30, 1))
                    Opcion331.Value = Val(Mid$(rstAtributos!atributo4, 31, 1))
                    
                    
                    Opcion41.Value = Val(Mid$(rstAtributos!atributo5, 1, 1))
                    Opcion42.Value = Val(Mid$(rstAtributos!atributo5, 2, 1))
                    Opcion43.Value = Val(Mid$(rstAtributos!atributo5, 3, 1))
                    Opcion44.Value = Val(Mid$(rstAtributos!atributo5, 4, 1))
                    Opcion45.Value = Val(Mid$(rstAtributos!atributo5, 5, 1))
                    Opcion46.Value = Val(Mid$(rstAtributos!atributo5, 6, 1))
                    Opcion47.Value = Val(Mid$(rstAtributos!atributo5, 7, 1))
                    Opcion48.Value = Val(Mid$(rstAtributos!atributo5, 8, 1))
                    Opcion49.Value = Val(Mid$(rstAtributos!atributo5, 9, 1))
                    Opcion410.Value = Val(Mid$(rstAtributos!atributo5, 10, 1))
                    Opcion411.Value = Val(Mid$(rstAtributos!atributo5, 11, 1))
                    Opcion412.Value = Val(Mid$(rstAtributos!atributo5, 12, 1))
                    Opcion413.Value = Val(Mid$(rstAtributos!atributo5, 13, 1))
                    Opcion414.Value = Val(Mid$(rstAtributos!atributo5, 14, 1))
                    Opcion415.Value = Val(Mid$(rstAtributos!atributo5, 15, 1))
                    Opcion416.Value = Val(Mid$(rstAtributos!atributo5, 16, 1))
                    Opcion417.Value = Val(Mid$(rstAtributos!atributo5, 17, 1))
                    Opcion418.Value = Val(Mid$(rstAtributos!atributo5, 18, 1))
                    Opcion419.Value = Val(Mid$(rstAtributos!atributo5, 19, 1))
                    Opcion420.Value = Val(Mid$(rstAtributos!atributo5, 20, 1))
                    
                    Opcion51.Value = Val(Mid$(rstAtributos!atributo6, 1, 1))
                    Opcion52.Value = Val(Mid$(rstAtributos!atributo6, 2, 1))
                    Opcion53.Value = Val(Mid$(rstAtributos!atributo6, 3, 1))
                    Opcion54.Value = Val(Mid$(rstAtributos!atributo6, 4, 1))
                    Opcion55.Value = Val(Mid$(rstAtributos!atributo6, 5, 1))
                    Opcion56.Value = Val(Mid$(rstAtributos!atributo6, 6, 1))
                    Opcion57.Value = Val(Mid$(rstAtributos!atributo6, 7, 1))
                    Opcion58.Value = Val(Mid$(rstAtributos!atributo6, 8, 1))
                    Opcion59.Value = Val(Mid$(rstAtributos!atributo6, 9, 1))
                    Opcion510.Value = Val(Mid$(rstAtributos!atributo6, 10, 1))
                    Opcion511.Value = Val(Mid$(rstAtributos!atributo6, 11, 1))
                    Opcion512.Value = Val(Mid$(rstAtributos!atributo6, 12, 1))
                    Opcion513.Value = Val(Mid$(rstAtributos!atributo6, 13, 1))
                    Opcion514.Value = Val(Mid$(rstAtributos!atributo6, 14, 1))
                    
                    rstAtributos.Close
                End If
                
            End If
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Salida_Click()
    Operador.SetFocus
    PrgConfigCoti.Hide
    Unload Me
    Menu.Show
End Sub

