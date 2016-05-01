VERSION 5.00
Begin VB.Form PrgAgendaTotal 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Evaluacion Semestral Actual de Proveedores"
   ClientHeight    =   7350
   ClientLeft      =   285
   ClientTop       =   705
   ClientWidth     =   11610
   LinkTopic       =   "Form2"
   ScaleHeight     =   7350
   ScaleWidth      =   11610
   Begin VB.TextBox Filtro 
      Height          =   285
      Left            =   240
      TabIndex        =   271
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   9480
      TabIndex        =   269
      Top             =   0
      Width           =   255
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
      Height          =   735
      Left            =   5040
      TabIndex        =   268
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Color2 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6720
      TabIndex        =   264
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox Color3 
      BackColor       =   &H00FFFF00&
      Height          =   495
      Left            =   8280
      TabIndex        =   263
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox Color1 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   5520
      TabIndex        =   262
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   11175
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   42
         Left            =   10665
         TabIndex        =   261
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   41
         Left            =   9105
         TabIndex        =   260
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   40
         Left            =   7545
         TabIndex        =   259
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   39
         Left            =   5970
         TabIndex        =   258
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   38
         Left            =   4410
         TabIndex        =   257
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   37
         Left            =   2865
         TabIndex        =   256
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   36
         Left            =   1320
         TabIndex        =   255
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   35
         Left            =   10665
         TabIndex        =   254
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   34
         Left            =   9105
         TabIndex        =   253
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   33
         Left            =   7545
         TabIndex        =   252
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   32
         Left            =   5970
         TabIndex        =   251
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   31
         Left            =   4410
         TabIndex        =   250
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   30
         Left            =   2865
         TabIndex        =   249
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   29
         Left            =   1320
         TabIndex        =   248
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   28
         Left            =   10665
         TabIndex        =   247
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   27
         Left            =   9105
         TabIndex        =   246
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   26
         Left            =   7545
         TabIndex        =   245
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   25
         Left            =   5970
         TabIndex        =   244
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   24
         Left            =   4410
         TabIndex        =   243
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   23
         Left            =   2865
         TabIndex        =   242
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   22
         Left            =   1320
         TabIndex        =   241
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   21
         Left            =   10665
         TabIndex        =   240
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   20
         Left            =   9105
         TabIndex        =   239
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   19
         Left            =   7545
         TabIndex        =   238
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   18
         Left            =   5970
         TabIndex        =   237
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   17
         Left            =   4410
         TabIndex        =   236
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   16
         Left            =   2865
         TabIndex        =   235
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   15
         Left            =   1320
         TabIndex        =   234
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   14
         Left            =   10665
         TabIndex        =   233
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   13
         Left            =   9105
         TabIndex        =   232
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   12
         Left            =   7545
         TabIndex        =   231
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   11
         Left            =   5970
         TabIndex        =   230
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   10
         Left            =   4410
         TabIndex        =   229
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   9
         Left            =   2865
         TabIndex        =   228
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   8
         Left            =   1320
         TabIndex        =   227
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   7
         Left            =   10665
         TabIndex        =   226
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   6
         Left            =   9105
         TabIndex        =   225
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   5
         Left            =   7560
         TabIndex        =   224
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   4
         Left            =   5970
         TabIndex        =   223
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   3
         Left            =   4410
         TabIndex        =   222
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   2
         Left            =   2865
         TabIndex        =   221
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia5 
         BackColor       =   &H000000FF&
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   220
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   42
         Left            =   10395
         TabIndex        =   219
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   41
         Left            =   8835
         TabIndex        =   218
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   40
         Left            =   7275
         TabIndex        =   217
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   39
         Left            =   5700
         TabIndex        =   216
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   38
         Left            =   4140
         TabIndex        =   215
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   37
         Left            =   2595
         TabIndex        =   214
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   36
         Left            =   1050
         TabIndex        =   213
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   35
         Left            =   10395
         TabIndex        =   212
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   34
         Left            =   8835
         TabIndex        =   211
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   33
         Left            =   7275
         TabIndex        =   210
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   32
         Left            =   5700
         TabIndex        =   209
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   31
         Left            =   4140
         TabIndex        =   208
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   30
         Left            =   2595
         TabIndex        =   207
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   29
         Left            =   1050
         TabIndex        =   206
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   28
         Left            =   10395
         TabIndex        =   205
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   27
         Left            =   8835
         TabIndex        =   204
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   26
         Left            =   7275
         TabIndex        =   203
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   25
         Left            =   5700
         TabIndex        =   202
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   24
         Left            =   4140
         TabIndex        =   201
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   23
         Left            =   2595
         TabIndex        =   200
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   22
         Left            =   1050
         TabIndex        =   199
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   21
         Left            =   10395
         TabIndex        =   198
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   20
         Left            =   8835
         TabIndex        =   197
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   19
         Left            =   7275
         TabIndex        =   196
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   18
         Left            =   5700
         TabIndex        =   195
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   17
         Left            =   4140
         TabIndex        =   194
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   16
         Left            =   2595
         TabIndex        =   193
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   15
         Left            =   1050
         TabIndex        =   192
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   14
         Left            =   10395
         TabIndex        =   191
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   13
         Left            =   8835
         TabIndex        =   190
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   12
         Left            =   7275
         TabIndex        =   189
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   11
         Left            =   5700
         TabIndex        =   188
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   10
         Left            =   4140
         TabIndex        =   187
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   9
         Left            =   2595
         TabIndex        =   186
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   8
         Left            =   1050
         TabIndex        =   185
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   7
         Left            =   10395
         TabIndex        =   184
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   6
         Left            =   8835
         TabIndex        =   183
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   5
         Left            =   7275
         TabIndex        =   182
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   4
         Left            =   5700
         TabIndex        =   181
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   3
         Left            =   4140
         TabIndex        =   180
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   2
         Left            =   2595
         TabIndex        =   179
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia4 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Index           =   1
         Left            =   1050
         TabIndex        =   178
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   42
         Left            =   10125
         TabIndex        =   177
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   41
         Left            =   8580
         TabIndex        =   176
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   40
         Left            =   7020
         TabIndex        =   175
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   39
         Left            =   5445
         TabIndex        =   174
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   38
         Left            =   3870
         TabIndex        =   173
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   37
         Left            =   2325
         TabIndex        =   172
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   36
         Left            =   780
         TabIndex        =   171
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   35
         Left            =   10125
         TabIndex        =   170
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   34
         Left            =   8580
         TabIndex        =   169
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   33
         Left            =   7020
         TabIndex        =   168
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   32
         Left            =   5445
         TabIndex        =   167
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   31
         Left            =   3870
         TabIndex        =   166
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   30
         Left            =   2325
         TabIndex        =   165
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   29
         Left            =   780
         TabIndex        =   164
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   28
         Left            =   10125
         TabIndex        =   163
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   27
         Left            =   8580
         TabIndex        =   162
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   26
         Left            =   7020
         TabIndex        =   161
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   25
         Left            =   5445
         TabIndex        =   160
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   24
         Left            =   3870
         TabIndex        =   159
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   23
         Left            =   2325
         TabIndex        =   158
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   22
         Left            =   780
         TabIndex        =   157
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   21
         Left            =   10125
         TabIndex        =   156
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   20
         Left            =   8580
         TabIndex        =   155
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   19
         Left            =   7020
         TabIndex        =   154
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   18
         Left            =   5445
         TabIndex        =   153
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   17
         Left            =   3870
         TabIndex        =   152
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   16
         Left            =   2325
         TabIndex        =   151
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   15
         Left            =   780
         TabIndex        =   150
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   14
         Left            =   10125
         TabIndex        =   149
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   13
         Left            =   8580
         TabIndex        =   148
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   12
         Left            =   7020
         TabIndex        =   147
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   11
         Left            =   5445
         TabIndex        =   146
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   10
         Left            =   3870
         TabIndex        =   145
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   9
         Left            =   2325
         TabIndex        =   144
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   8
         Left            =   780
         TabIndex        =   143
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   7
         Left            =   10125
         TabIndex        =   142
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   6
         Left            =   8580
         TabIndex        =   141
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   5
         Left            =   7020
         TabIndex        =   140
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   4
         Left            =   5445
         TabIndex        =   139
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   3
         Left            =   3870
         TabIndex        =   138
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   2
         Left            =   2325
         TabIndex        =   137
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia3 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Index           =   1
         Left            =   780
         TabIndex        =   136
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   42
         Left            =   9855
         TabIndex        =   128
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   41
         Left            =   8310
         TabIndex        =   127
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   40
         Left            =   6750
         TabIndex        =   126
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   39
         Left            =   5175
         TabIndex        =   125
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   38
         Left            =   3600
         TabIndex        =   124
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   37
         Left            =   2055
         TabIndex        =   123
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   36
         Left            =   510
         TabIndex        =   122
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   35
         Left            =   9855
         TabIndex        =   121
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   34
         Left            =   8310
         TabIndex        =   120
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   33
         Left            =   6750
         TabIndex        =   119
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   32
         Left            =   5175
         TabIndex        =   118
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   31
         Left            =   3600
         TabIndex        =   117
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   30
         Left            =   2055
         TabIndex        =   116
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   29
         Left            =   510
         TabIndex        =   115
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   28
         Left            =   9855
         TabIndex        =   114
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   27
         Left            =   8310
         TabIndex        =   113
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   26
         Left            =   6750
         TabIndex        =   112
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   25
         Left            =   5175
         TabIndex        =   111
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   24
         Left            =   3600
         TabIndex        =   110
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   23
         Left            =   2055
         TabIndex        =   109
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   22
         Left            =   510
         TabIndex        =   108
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   21
         Left            =   9855
         TabIndex        =   107
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   20
         Left            =   8310
         TabIndex        =   106
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   19
         Left            =   6750
         TabIndex        =   105
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   18
         Left            =   5175
         TabIndex        =   104
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   17
         Left            =   3600
         TabIndex        =   103
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   16
         Left            =   2055
         TabIndex        =   102
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   15
         Left            =   510
         TabIndex        =   101
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   14
         Left            =   9855
         TabIndex        =   100
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   13
         Left            =   8310
         TabIndex        =   99
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   12
         Left            =   6750
         TabIndex        =   98
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   11
         Left            =   5175
         TabIndex        =   97
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   10
         Left            =   3600
         TabIndex        =   96
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   9
         Left            =   2055
         TabIndex        =   95
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   8
         Left            =   510
         TabIndex        =   94
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   7
         Left            =   9855
         TabIndex        =   93
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   6
         Left            =   8310
         TabIndex        =   92
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   5
         Left            =   6750
         TabIndex        =   91
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   4
         Left            =   5175
         TabIndex        =   90
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   3
         Left            =   3600
         TabIndex        =   89
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   2
         Left            =   2055
         TabIndex        =   88
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia2 
         BackColor       =   &H0000FF00&
         Height          =   495
         Index           =   1
         Left            =   510
         TabIndex        =   87
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   42
         Left            =   9600
         TabIndex        =   86
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   41
         Left            =   8040
         TabIndex        =   85
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   40
         Left            =   6480
         TabIndex        =   84
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   39
         Left            =   4920
         TabIndex        =   83
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   38
         Left            =   3360
         TabIndex        =   82
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   37
         Left            =   1800
         TabIndex        =   81
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   36
         Left            =   240
         TabIndex        =   80
         Top             =   4560
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   35
         Left            =   9600
         TabIndex        =   79
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   34
         Left            =   8040
         TabIndex        =   78
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   33
         Left            =   6480
         TabIndex        =   77
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   32
         Left            =   4920
         TabIndex        =   76
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   31
         Left            =   3360
         TabIndex        =   75
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   30
         Left            =   1800
         TabIndex        =   74
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   29
         Left            =   240
         TabIndex        =   73
         Top             =   3840
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   28
         Left            =   9600
         TabIndex        =   72
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   27
         Left            =   8040
         TabIndex        =   71
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   26
         Left            =   6480
         TabIndex        =   70
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   25
         Left            =   4920
         TabIndex        =   69
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   24
         Left            =   3360
         TabIndex        =   68
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   23
         Left            =   1800
         TabIndex        =   67
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   22
         Left            =   240
         TabIndex        =   66
         Top             =   3120
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   21
         Left            =   9600
         TabIndex        =   65
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   20
         Left            =   8040
         TabIndex        =   64
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   19
         Left            =   6480
         TabIndex        =   63
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   18
         Left            =   4920
         TabIndex        =   62
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   17
         Left            =   3360
         TabIndex        =   61
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   16
         Left            =   1800
         TabIndex        =   60
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   15
         Left            =   240
         TabIndex        =   59
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   14
         Left            =   9600
         TabIndex        =   58
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   13
         Left            =   8040
         TabIndex        =   57
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   12
         Left            =   6480
         TabIndex        =   56
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   11
         Left            =   4920
         TabIndex        =   55
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   10
         Left            =   3360
         TabIndex        =   54
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   9
         Left            =   1800
         TabIndex        =   53
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   8
         Left            =   240
         TabIndex        =   52
         Top             =   1680
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   7
         Left            =   9600
         TabIndex        =   51
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   6
         Left            =   8040
         TabIndex        =   50
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   5
         Left            =   6480
         TabIndex        =   49
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   4
         Left            =   4920
         TabIndex        =   48
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   3
         Left            =   3360
         TabIndex        =   47
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   2
         Left            =   1800
         TabIndex        =   46
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   280
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   42
         Left            =   9600
         TabIndex        =   44
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   41
         Left            =   8040
         TabIndex        =   43
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   40
         Left            =   6480
         TabIndex        =   42
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   39
         Left            =   4920
         TabIndex        =   41
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   38
         Left            =   3360
         TabIndex        =   40
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   37
         Left            =   1800
         TabIndex        =   39
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   36
         Left            =   240
         TabIndex        =   38
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   35
         Left            =   9600
         TabIndex        =   37
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   34
         Left            =   8040
         TabIndex        =   36
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   33
         Left            =   6480
         TabIndex        =   35
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   32
         Left            =   4920
         TabIndex        =   34
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   31
         Left            =   3360
         TabIndex        =   33
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   30
         Left            =   1800
         TabIndex        =   32
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   29
         Left            =   240
         TabIndex        =   31
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   28
         Left            =   9600
         TabIndex        =   30
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   27
         Left            =   8040
         TabIndex        =   29
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   26
         Left            =   6480
         TabIndex        =   28
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   25
         Left            =   4920
         TabIndex        =   27
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   24
         Left            =   3360
         TabIndex        =   26
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   23
         Left            =   1800
         TabIndex        =   25
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   22
         Left            =   240
         TabIndex        =   24
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   21
         Left            =   9600
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   20
         Left            =   8040
         TabIndex        =   22
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   19
         Left            =   6480
         TabIndex        =   21
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   18
         Left            =   4920
         TabIndex        =   20
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   17
         Left            =   3360
         TabIndex        =   19
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   16
         Left            =   1800
         TabIndex        =   18
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   15
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   14
         Left            =   9600
         TabIndex        =   16
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   13
         Left            =   8040
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   12
         Left            =   6480
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   11
         Left            =   4920
         TabIndex        =   13
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   10
         Left            =   3360
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   9
         Left            =   1800
         TabIndex        =   11
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   8
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   7
         Left            =   9600
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   6
         Left            =   8040
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   6480
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   4920
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   3
         Left            =   3360
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   2
         Left            =   1800
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Dia 
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.Line Line16 
         X1              =   3240
         X2              =   3240
         Y1              =   240
         Y2              =   5280
      End
      Begin VB.Line Line15 
         X1              =   1680
         X2              =   1680
         Y1              =   240
         Y2              =   5280
      End
      Begin VB.Line Line14 
         X1              =   4800
         X2              =   4800
         Y1              =   240
         Y2              =   5280
      End
      Begin VB.Line Line13 
         X1              =   6360
         X2              =   6360
         Y1              =   240
         Y2              =   5280
      End
      Begin VB.Line Line12 
         X1              =   7920
         X2              =   7920
         Y1              =   240
         Y2              =   5280
      End
      Begin VB.Line Line11 
         X1              =   9480
         X2              =   9480
         Y1              =   240
         Y2              =   5280
      End
      Begin VB.Line Line10 
         X1              =   120
         X2              =   11040
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   120
         Y1              =   5280
         Y2              =   240
      End
      Begin VB.Line Line8 
         X1              =   11040
         X2              =   11040
         Y1              =   5280
         Y2              =   240
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   11040
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   11040
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   11040
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   11040
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   11040
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   11040
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   11040
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label DTitulo1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Domingo"
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
         Left            =   240
         TabIndex        =   135
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lunes"
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
         Left            =   1800
         TabIndex        =   134
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Martes"
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
         Left            =   3360
         TabIndex        =   133
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Miercoles"
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
         Left            =   4920
         TabIndex        =   132
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jueves"
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
         Left            =   6480
         TabIndex        =   131
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Viernes"
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
         Left            =   8040
         TabIndex        =   130
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sabado"
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
         Left            =   9600
         TabIndex        =   129
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox XAno 
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox XMes 
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Capacitacion"
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
      Left            =   9840
      TabIndex        =   270
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Propio"
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
      Left            =   8640
      TabIndex        =   267
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Asignacion"
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
      Left            =   7080
      TabIndex        =   266
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "SAC"
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
      Left            =   5880
      TabIndex        =   265
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Siguiente 
      Height          =   480
      Left            =   4920
      MouseIcon       =   "AgendaTotal.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "AgendaTotal.frx":030A
      ToolTipText     =   "Registro Posterior"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Anterior 
      Height          =   480
      Left            =   2760
      MouseIcon       =   "AgendaTotal.frx":074C
      MousePointer    =   99  'Custom
      Picture         =   "AgendaTotal.frx":0A56
      ToolTipText     =   "Registro Anterior"
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "PrgAgendaTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstCronogramaII As Recordset
Dim spCronogramaII As String
Dim rstCurso As Recordset
Dim spCurso As String
    
Dim ZZMes As String
Dim ZZDia As String
Dim ZZAno As String
Dim ZZFecha As String
Dim ZZError As String
Dim ZFecha As String
Dim ZVto As String
Dim SumaDia As Integer
Dim ZAno As String
Dim ZTipo As String
Dim ZNumero As String

Dim CargaEmpresa(10, 10) As String
Dim ZVector(100) As String

Dim ZVector2(40, 100) As String
Dim ZVector3(40, 100) As String
Dim ZVector4(40, 100) As String
Dim ZVector5(40, 1000, 2) As String

Dim ZSuma2(100) As Integer
Dim ZSuma3(100) As Integer
Dim ZSuma4(100) As Integer
Dim ZSuma5(100) As Integer

Dim ZSac(1000, 10) As String
Dim ZImple(10, 3) As String
Dim ZZCurso(1000) As String

Dim ZZVector(1000, 20) As String
Dim ZZVectorII(1000, 5) As String
Dim ZZVectorIII(1000, 2) As String

Private Sub Acepta_Click()

    XEmpresa = WEmpresa
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0002"
    CargaEmpresa(2, 2) = "Empresa02"
    CargaEmpresa(3, 1) = "0003"
    CargaEmpresa(3, 2) = "Empresa03"
    CargaEmpresa(4, 1) = "0004"
    CargaEmpresa(4, 2) = "Empresa04"
    CargaEmpresa(5, 1) = "0005"
    CargaEmpresa(5, 2) = "Empresa05"
    CargaEmpresa(6, 1) = "0006"
    CargaEmpresa(6, 2) = "Empresa06"
    CargaEmpresa(7, 1) = "0007"
    CargaEmpresa(7, 2) = "Empresa07"
    CargaEmpresa(8, 1) = "0008"
    CargaEmpresa(8, 2) = "Empresa08"
    CargaEmpresa(9, 1) = "0009"
    CargaEmpresa(9, 2) = "Empresa09"

End Sub

Private Sub Mes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ano.SetFocus
    End If
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Calendario_Click
    End If
End Sub

Private Sub Cancela_click()
    PrgAgendaTotal.Hide
    Unload Me
    Menu.Show
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



Private Sub Anterior_Click()

    If XMes.Text > 1 Then
        XMes.Text = Str$(Val(XMes.Text) - 1)
            Else
        XMes.Text = "12"
        XAno.Text = Str$(Val(XAno.Text) - 1)
    End If
    Call Calcula_Calendario_Click

End Sub

Private Sub Calcula_Calendario_Click()

    Rem On Error GoTo WError
    
    ZZError = "N"
    
    ZZMes = XMes.Text
    ZZAno = XAno.Text
    
    Call Ceros(ZZMes, 2)
    Call Ceros(ZZAno, 4)
    
    For Ciclo = 1 To 42
        ZVector(Ciclo) = "N"
        Dia(Ciclo).Text = ""
        Dia(Ciclo).Visible = False
        Dia1(Ciclo).Text = ""
        Dia1(Ciclo).Visible = False
        Dia2(Ciclo).Text = ""
        Dia2(Ciclo).Visible = False
        Dia3(Ciclo).Text = ""
        Dia3(Ciclo).Visible = False
        Dia4(Ciclo).Text = ""
        Dia4(Ciclo).Visible = False
        Dia5(Ciclo).Text = ""
        Dia5(Ciclo).Visible = False
    Next Ciclo
    
    
    ZPasa = 0
    ZLugar = 0
    ZLugarII = 0
    
    Erase ZSuma2
    Erase ZSuma3
    Erase ZSuma4
    Erase ZSuma5
    
    Erase ZVector
    Erase ZVector2
    Erase ZVector3
    Erase ZVector4
    Erase ZVector5
    
    Erase ZSac
    
    For Ciclo = 1 To 31
    
        ZZDia = Ciclo
        Call Ceros(ZZDia, 2)
        ZZFecha = ZZDia + "/" + ZZMes + "/" + ZZAno
        DiaSemana = Format(ZZFecha, "w")
        
        If ZZError = "N" Then
        
            If ZPasa = 0 Then
                For Baja = 1 To Val(DiaSemana) - 1
                  ZLugar = ZLugar + 1
                Next Baja
                ZPasa = 1
            End If
            
            ZLugar = ZLugar + 1
            Rem Dia(ZLugar).Visible = True
            Dia1(ZLugar).Visible = True
            Dia1(ZLugar).Text = Trim(Str$(Ciclo))
            Rem Dia2(ZLugar).Visible = True
            Dia2(ZLugar).Text = ""
            Rem Dia3(ZLugar).Visible = True
            Dia3(ZLugar).Text = ""
            Rem Dia4(ZLugar).Visible = True
            Dia4(ZLugar).Text = ""
            Rem Dia5(ZLugar).Visible = True
            Dia5(ZLugar).Text = ""
            ZVector(ZLugar) = ZZFecha
               
        End If
        
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "Select CargaSac.Tipo, CargaSac.Ano, CargaSac.Numero, CargaSac.Fecha "
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.ResponsableDestino = " + "'" + Str$(ZZOperadorResponsable) + "'"
    ZSql = ZSql + " AND CargaSac.Estado < 3"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
        With rstCargaSac
        .MoveFirst
        Do
            If .EOF = False Then
            
                If Val(Filtro.Text) = 0 Or Val(Filtro.Text) = rstCargaSac!Ano Then
                
                    ZLugarII = ZLugarII + 1
                    
                    ZSac(ZLugarII, 1) = rstCargaSac!Tipo
                    ZSac(ZLugarII, 2) = rstCargaSac!Ano
                    ZSac(ZLugarII, 3) = rstCargaSac!Numero
                    ZSac(ZLugarII, 4) = rstCargaSac!Fecha
                
                End If
                
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
            For CicloII = 1 To 40
                ZFechaI = ZVector(CicloII)
                WFechaordI = Right$(ZFechaI, 4) + Mid$(ZFechaI, 4, 2) + Left$(ZFechaI, 2)
                If Trim(ZFechaI) <> "" Then
                    If ZPasa = 0 Then
                        If WFechaOrdII <= WFechaordI Then
                            ZSuma2(CicloII) = ZSuma2(CicloII) + 1
                            Dia2(CicloII).Text = Trim(Str$(ZSuma2(CicloII)))
                            ZVector2(CicloII, ZSuma2(CicloII)) = ZTipo + ZAno + ZNumero
                            Exit For
                        End If
                        ZPasa = 1
                            Else
                        If ZFechaII = ZFechaI Then
                            ZSuma2(CicloII) = ZSuma2(CicloII) + 1
                            Dia2(CicloII).Text = Trim(Str$(ZSuma2(CicloII)))
                            ZVector2(CicloII, ZSuma2(CicloII)) = ZTipo + ZAno + ZNumero
                            Exit For
                        End If
                    End If
                End If
            Next CicloII
            
        End If
        
    Next Ciclo
    
    
    
    
    
    ZLugarII = 0
    Erase ZSac
    
    ZSql = ""
    ZSql = ZSql + "Select CargaSacII.Tipo, CargaSacII.Ano, CargaSacII.Numero "
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
            
                If Val(Filtro.Text) = 0 Or Val(Filtro.Text) = rstCargaSacII!Ano Then
                
                    ZLugarII = ZLugarII + 1
                    ZSac(ZLugarII, 1) = rstCargaSacII!Tipo
                    ZSac(ZLugarII, 2) = rstCargaSacII!Ano
                    ZSac(ZLugarII, 3) = rstCargaSacII!Numero
                    
                End If
                
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
        ZSql = ZSql + "Select CargaSac.Fecha, CargaSac.Estado"
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
        
            If ZEstado <= 3 Then
            
                ZSql = ""
                ZSql = ZSql + "Select CargaSacII.Responsable1, CargaSacII.Responsable2, CargaSacII.Responsable3, CargaSacII.Responsable4, CargaSacII.Responsable5, CargaSacII.Responsable6, CargaSacII.Plazo1, CargaSacII.Plazo2, CargaSacII.Plazo3, CargaSacII.Plazo4, CargaSacII.Plazo5, CargaSacII.Plazo6 "
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
                ZSql = ZSql + "Select CargaSacIII.Estado1, CargaSacIII.Estado2, CargaSacIII.Estado3, CargaSacIII.Estado4, CargaSacIII.Estado5, CargaSacIII.Estado6 "
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
                    
                    If Val(ZResponsable1) = Val(ZZOperadorResponsable) And Val(ZEstado1) < 1 Then
                        ZEntra = "S"
                        ZPlazo = ZPlazo1
                    End If
            
                    If ZEntra = "S" Then
                    
                        Rem SumaDia = 31
                        Rem Call Calcula_vencimiento(ZFecha, SumaDia, ZVto)
                        ZPasa = 0
                        ZFechaII = ZPlazo
                        WFechaOrdII = Right$(ZFechaII, 4) + Mid$(ZFechaII, 4, 2) + Left$(ZFechaII, 2)
                        For CicloII = 1 To 40
                            ZFechaI = ZVector(CicloII)
                            WFechaordI = Right$(ZFechaI, 4) + Mid$(ZFechaI, 4, 2) + Left$(ZFechaI, 2)
                            If Trim(ZFechaI) <> "" Then
                                If ZPasa = 0 Then
                                    If WFechaOrdII <= WFechaordI Then
                                        ZSuma2(CicloII) = ZSuma2(CicloII) + 1
                                        Dia2(CicloII).Text = Trim(Str$(ZSuma2(CicloII)))
                                        ZVector2(CicloII, ZSuma2(CicloII)) = ZTipo + ZAno + ZNumero
                                        Exit For
                                    End If
                                    ZPasa = 1
                                        Else
                                    If ZFechaII = ZFechaI Then
                                        ZSuma2(CicloII) = ZSuma2(CicloII) + 1
                                        Dia2(CicloII).Text = Trim(Str$(ZSuma2(CicloII)))
                                        ZVector2(CicloII, ZSuma2(CicloII)) = ZTipo + ZAno + ZNumero
                                        Exit For
                                    End If
                                End If
                            End If
                        Next CicloII
                        
                    End If
            
                Next CicloRes
                
            End If
            
        End If
        
    Next Ciclo
    
    
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
                    
                    For CicloII = 1 To 40
                    
                        ZFechaI = ZVector(CicloII)
                        WFechaordI = Right$(ZFechaI, 4) + Mid$(ZFechaI, 4, 2) + Left$(ZFechaI, 2)
                        
                        If Trim(ZFechaI) <> "" Then
                            If ZPasa = 0 Then
                                If WFechaOrdII <= WFechaordI Then
                                    ZSuma3(CicloII) = ZSuma3(CicloII) + 1
                                    Dia3(CicloII).Text = Trim(Str$(ZSuma3(CicloII)))
                                    ZVector3(CicloII, ZSuma3(CicloII)) = rstPlanifica!Clave
                                    Exit For
                                End If
                                ZPasa = 1
                                    Else
                                If ZFechaII = ZFechaI Then
                                    ZSuma3(CicloII) = ZSuma3(CicloII) + 1
                                    Dia3(CicloII).Text = Trim(Str$(ZSuma3(CicloII)))
                                    ZVector3(CicloII, ZSuma3(CicloII)) = rstPlanifica!Clave
                                    Exit For
                                End If
                            End If
                        End If
                        
                    Next CicloII
                    
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
                    
                    For CicloII = 1 To 40
                    
                        ZFechaI = ZVector(CicloII)
                        WFechaordI = Right$(ZFechaI, 4) + Mid$(ZFechaI, 4, 2) + Left$(ZFechaI, 2)
                        
                        If Trim(ZFechaI) <> "" Then
                            If ZPasa = 0 Then
                                If WFechaOrdII <= WFechaordI Then
                                    ZSuma4(CicloII) = ZSuma4(CicloII) + 1
                                    Dia4(CicloII).Text = Trim(Str$(ZSuma4(CicloII)))
                                    ZVector4(CicloII, ZSuma4(CicloII)) = rstPlanifica!Clave
                                    Exit For
                                End If
                                ZPasa = 1
                                    Else
                                If ZFechaII = ZFechaI Then
                                    ZSuma4(CicloII) = ZSuma4(CicloII) + 1
                                    Dia4(CicloII).Text = Trim(Str$(ZSuma4(CicloII)))
                                    ZVector4(CicloII, ZSuma4(CicloII)) = rstPlanifica!Clave
                                    Exit For
                                End If
                            End If
                        End If
                        
                    Next CicloII
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPlanifica.Close
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    Erase ZZVectorII
    Erase ZZVectorIII
    WRenglonII = 0
    WRenglonIII = 0
    
    ZSql = ""
    ZSql = ZSql + "Select Cronograma.Horas, Cronograma.Realizado, Cronograma.Legajo, Cronograma.Curso"
    ZSql = ZSql + " FROM Cronograma"
    ZSql = ZSql + " Where Cronograma.Ano = " + "'" + XAno.Text + "'"
    ZSql = ZSql + " and Cronograma.Horas > Cronograma.Realizado"
    
    Rem ZSql = ZSql + " Order by Cronograma.Legajo, Cronograma.Curso"
    rsCronograma = ZSql
    Set rstCronograma = db.OpenRecordset(rsCronograma, dbOpenSnapshot, dbSQLPassThrough)
    If rstCronograma.RecordCount > 0 Then
        With rstCronograma
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstCronograma!Horas > rstCronograma!Realizado Then
                        WRenglonII = WRenglonII + 1
                        ZZVectorII(WRenglonII, 1) = rstCronograma!Legajo
                        ZZVectorII(WRenglonII, 2) = rstCronograma!Curso
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCronograma.Close
    End If
        
    For Ciclo = 1 To WRenglonII
        
        WLegajo = ZZVectorII(Ciclo, 1)
        WCurso = ZZVectorII(Ciclo, 2)
        WPerfil = 0

        Rem WPerfil = ZZVectorII(Ciclo, 3)
        Rem WResponsable = ZZVectorII(Ciclo, 3)
        Rem WResponsableII = ZZVectorII(Ciclo, 3)
        
        Rem ZZResponsable = WResponsable
        Rem ZZResponsableII = WResponsableII
    
        ZSql = ""
        ZSql = ZSql + "Select Legajo.Perfil"
        ZSql = ZSql + " FROM Legajo"
        ZSql = ZSql + " Where Legajo.Codigo = " + "'" + WLegajo + "'"
        spLegajo = ZSql
        Set rstlegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstlegajo.RecordCount > 0 Then
            WPerfil = rstlegajo!Perfil
            rstlegajo.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select Tarea.Responsable, Tarea.ResponsableII"
        ZSql = ZSql + " FROM Tarea"
        ZSql = ZSql + " Where Tarea.Codigo = " + "'" + Str$(WPerfil) + "'"
        spTarea = ZSql
        Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
        If rstTarea.RecordCount > 0 Then
            ZZResponsable = IIf(IsNull(rstTarea!Responsable), "0", rstTarea!Responsable)
            ZZResponsableII = IIf(IsNull(rstTarea!ResponsableII), "0", rstTarea!ResponsableII)
            rstTarea.Close
        End If
        
        If ZZResponsable = Val(ZZOperadorResponsable) Or ZZResponsableII = Val(ZZOperadorResponsable) Then
        
            ZSql = ""
            ZSql = ZSql + "Select CronogramaII.Mes1, CronogramaII.Mes2, CronogramaII.Mes3, CronogramaII.Mes4, CronogramaII.Mes5, CronogramaII.Mes6, CronogramaII.Mes7, CronogramaII.Mes8, CronogramaII.Mes9, CronogramaII.Mes10, CronogramaII.Mes11, CronogramaII.Mes12 "
            ZSql = ZSql + " FROM CronogramaII"
            ZSql = ZSql + " Where CronogramaII.Ano = " + "'" + XAno.Text + "'"
            ZSql = ZSql + " and CronogramaII.Curso = " + "'" + WCurso + "'"
            
            rsCronogramaII = ZSql
            Set rstCronogramaII = db.OpenRecordset(rsCronogramaII, dbOpenSnapshot, dbSQLPassThrough)
            If rstCronogramaII.RecordCount > 0 Then
                With rstCronogramaII
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            ZZMes1 = Trim(rstCronogramaII!MES1)
                            ZZMes2 = Trim(rstCronogramaII!Mes2)
                            ZZMes3 = Trim(rstCronogramaII!Mes3)
                            ZZMes4 = Trim(rstCronogramaII!Mes4)
                            ZZMes5 = Trim(rstCronogramaII!Mes5)
                            ZZMes6 = Trim(rstCronogramaII!Mes6)
                            ZZMes7 = Trim(rstCronogramaII!Mes7)
                            ZZMes8 = Trim(rstCronogramaII!Mes8)
                            ZZMes9 = Trim(rstCronogramaII!Mes9)
                            ZZMes10 = Trim(rstCronogramaII!Mes10)
                            ZZMes11 = Trim(rstCronogramaII!Mes11)
                            ZZMes12 = Trim(rstCronogramaII!Mes12)
                            
                            If UCase(ZZMes12) = "X" Then
                                ZZZMes = 12
                            End If
                            If UCase(ZZMes11) = "X" Then
                                ZZZMes = 11
                            End If
                            If UCase(ZZMes10) = "X" Then
                                ZZZMes = 10
                            End If
                            If UCase(ZZMes9) = "X" Then
                                ZZZMes = 9
                            End If
                            If UCase(ZZMes8) = "X" Then
                                ZZZMes = 8
                            End If
                            If UCase(ZZMes7) = "X" Then
                                ZZZMes = 7
                            End If
                            If UCase(ZZMes6) = "X" Then
                                ZZZMes = 6
                            End If
                            If UCase(ZZMes5) = "X" Then
                                ZZZMes = 5
                            End If
                            If UCase(ZZMes4) = "X" Then
                                ZZZMes = 4
                            End If
                            If UCase(ZZMes3) = "X" Then
                                ZZZMes = 3
                            End If
                            If UCase(ZZMes2) = "X" Then
                                ZZZMes = 2
                            End If
                            If UCase(ZZMes1) = "X" Then
                                ZZZMes = 1
                            End If
                            
                            If Val(XMes.Text) >= ZZZMes Then
                            
                                For CicloII = 1 To 40
                                    ZFechaI = ZVector(CicloII)
                                    If Trim(ZFechaI) <> "" Then
                                        ZSuma5(CicloII) = ZSuma5(CicloII) + 1
                                        Dia5(CicloII).Text = Trim(Str$(ZSuma5(CicloII)))
                                        ZVector5(CicloII, ZSuma5(CicloII), 1) = WLegajo
                                        ZVector5(CicloII, ZSuma5(CicloII), 2) = WCurso
                                        Exit For
                                    End If
                                Next CicloII
                                
                            End If
                    
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCronogramaII.Close
            End If
        
        End If
        
    Next Ciclo
        
    
    For Ciclo = 1 To 40
        If Val(Dia2(Ciclo).Text) <> 0 Then
            Dia2(Ciclo).Visible = True
        End If
        If Val(Dia3(Ciclo).Text) <> 0 Then
            Dia3(Ciclo).Visible = True
        End If
        If Val(Dia4(Ciclo).Text) <> 0 Then
            Dia4(Ciclo).Visible = True
        End If
        If Val(Dia5(Ciclo).Text) <> 0 Then
            Dia5(Ciclo).Visible = True
        End If
    Next Ciclo
    
    Exit Sub
    
WError:
    ZZError = "S"
    Resume Next


End Sub

Private Sub Dia2_dblclick(Index As Integer)
    Erase ZZPasaDatos
    If Val(Dia2(Index).Text) <> 0 Then
        For Ciclo = 1 To Val(Dia2(Index).Text)
            ZZPasaDatos(Ciclo, 1) = ZVector2(Index, Ciclo)
        Next Ciclo
    End If
    PrgAgendaIndiceSac.Show
End Sub

Private Sub Dia3_dblclick(Index As Integer)
    Erase ZZPasaDatos
    If Val(Dia3(Index).Text) <> 0 Then
        For Ciclo = 1 To Val(Dia3(Index).Text)
            ZZPasaDatos(Ciclo, 1) = ZVector3(Index, Ciclo)
        Next Ciclo
    End If
    PrgAgendaPlanificaI.Show
End Sub

Private Sub Dia4_dblclick(Index As Integer)
    Erase ZZPasaDatos
    If Val(Dia4(Index).Text) <> 0 Then
        For Ciclo = 1 To Val(Dia4(Index).Text)
            ZZPasaDatos(Ciclo, 1) = ZVector4(Index, Ciclo)
        Next Ciclo
    End If
    PrgAgendaPlanificaI.Show
End Sub

Private Sub Dia5_dblclick(Index As Integer)
    Erase ZZPasaDatos
    If Val(Dia5(Index).Text) <> 0 Then
        For Ciclo = 1 To Val(Dia5(Index).Text)
            ZZPasaDatos(Ciclo, 1) = ZVector5(Index, Ciclo, 1)
            ZZPasaDatos(Ciclo, 2) = ZVector5(Index, Ciclo, 2)
        Next Ciclo
    End If
    PrgConsultaCursos.Show
End Sub

Private Sub Form_Activate()
    Call Calcula_Calendario_Click
End Sub

Private Sub Form_Load()
    XMes.Text = Left$(Date$, 2)
    XAno.Text = Right$(Date$, 4)
    Filtro.Text = "2013"
    
    PrgAgendaTotal.Caption = "Agenda de : " + ZZOperadorResponsableNombre
    
End Sub

Private Sub Siguiente_Click()

    If XMes.Text < 12 Then
        XMes.Text = Str$(Val(XMes.Text) + 1)
            Else
        XMes.Text = "1"
        XAno.Text = Str$(Val(XAno.Text) + 1)
    End If
    Call Calcula_Calendario_Click

End Sub

