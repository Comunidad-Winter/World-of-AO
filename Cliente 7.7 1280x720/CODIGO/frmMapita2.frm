VERSION 5.00
Begin VB.Form frmMap 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Mapa del Mundo AodraG"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   DrawStyle       =   1  'Dash
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10440
      Picture         =   "frmMapita.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   276
      Top             =   420
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   267
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   275
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   266
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   274
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   265
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   273
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   264
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   272
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   263
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   271
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   262
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   270
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   261
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   269
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   260
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   268
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   259
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   267
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   258
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   266
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   257
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   265
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   256
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   264
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   255
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   263
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   254
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   262
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   253
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   261
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   252
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   260
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   251
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   259
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   250
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   258
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   249
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   257
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   248
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   256
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   247
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   255
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   246
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   254
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   245
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   253
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   244
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   252
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   249
      Top             =   8280
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8280
      Picture         =   "frmMapita.frx":0CB5
      Style           =   1  'Graphical
      TabIndex        =   244
      Top             =   7740
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8760
      TabIndex        =   243
      Top             =   7080
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   77
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   242
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   243
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   241
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   242
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   240
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   241
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   239
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   240
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   238
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   239
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   237
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   238
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   236
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   237
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   235
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   236
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   234
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   235
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   233
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   234
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   232
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   233
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   231
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   232
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   230
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   231
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   229
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   230
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   228
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   229
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   227
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   228
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   226
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   227
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   225
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   226
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   224
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   225
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   223
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   224
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   222
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   223
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   221
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   222
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   220
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   221
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   219
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   220
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   218
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   219
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   217
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   218
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   216
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   217
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   215
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   216
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   214
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   215
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   213
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   214
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   212
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   213
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   211
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   212
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   210
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   211
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   209
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   210
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   208
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   209
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   207
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   208
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   206
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   207
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   205
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   206
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   204
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   205
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   203
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   204
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   202
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   203
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   201
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   202
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   200
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   201
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   199
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   200
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   198
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   199
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   197
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   198
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   196
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   197
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   195
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   196
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   194
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   195
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   193
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   194
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   192
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   193
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   191
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   192
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   190
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   191
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   189
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   190
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   188
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   189
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   187
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   188
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   186
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   187
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   185
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   186
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   184
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   185
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   183
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   184
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   182
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   183
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   181
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   182
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   180
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   181
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   179
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   180
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   178
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   179
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   177
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   178
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   176
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   177
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   175
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   176
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   174
      Top             =   6960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   175
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   173
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   174
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   172
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   173
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   172
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   170
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   171
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   169
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   170
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   168
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   169
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   168
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   167
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   166
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   165
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   164
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   162
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   163
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   161
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   162
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   160
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   161
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   159
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   160
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   158
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   159
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   157
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   158
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   157
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   156
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   155
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   153
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   154
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   152
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   153
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   151
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   152
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   151
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   150
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   149
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   148
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   146
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   147
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   145
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   146
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   144
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   145
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   143
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   144
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   143
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   142
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   141
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   140
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   139
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   138
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   137
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   136
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   135
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   134
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   133
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   132
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   131
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   129
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   130
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   129
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   128
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   127
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   126
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   125
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   124
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   123
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   122
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   121
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   120
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   119
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   118
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   117
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   116
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   115
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   114
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   113
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   112
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   111
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   110
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   109
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   108
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   107
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   106
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   105
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   104
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   103
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   102
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   101
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   100
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   99
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   98
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   97
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   96
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   95
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   94
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   93
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   92
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   91
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   90
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   89
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   88
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   87
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Enabled         =   0   'False
      Height          =   255
      Index           =   86
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   360
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   85
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   84
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   83
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   82
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   81
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   80
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   79
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   78
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   76
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   75
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   74
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   73
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   72
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   71
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   70
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   69
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   68
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   67
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   66
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   65
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   64
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   63
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   62
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   61
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   60
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   59
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   58
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   57
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   56
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   55
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   54
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   53
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   52
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   51
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   50
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   49
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   48
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   47
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   46
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   45
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   44
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   43
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   42
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   41
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   40
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   39
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   38
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   37
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   36
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   35
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "34"
      Height          =   255
      Index           =   34
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   33
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   32
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   31
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   30
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   29
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   28
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   27
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   26
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   25
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   24
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   23
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   22
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   21
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   20
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   19
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   18
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   17
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   16
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   15
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   14
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   13
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   12
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   11
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   10
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   9
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   8
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   7
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   6
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   5
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   4
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   3
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   2
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   1
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   7440
      TabIndex        =   251
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   7440
      TabIndex        =   250
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   7440
      TabIndex        =   248
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa:"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   247
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Criaturas del Mapa:"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   246
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Info del Mapa:"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   245
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
busca7 = Text3.Text
Mapa8 = "Mapa" & busca7
Label6.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "bichos")
Label5.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "Info")
Label4.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "nombre")
Option1(busca7).Value = True
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmMap.Picture = LoadPicture(DirGraficos & "pergamino.bmp")

Dim n As Integer
Dim colortipo As String
Mapa8 = "Mapa" & n
Dir7 = App.Path
For n = 1 To 267
    Mapa8 = "Mapa" & n
    Option1(n).Caption = n
    colortipo = GetVar(Dir7 & "\Data.txt", Mapa8, "tipo")
    Select Case colortipo
        Case "Agua"
            Option1(n).BackColor = RGB(0, 128, 192)
        Case "Bosque"
            Option1(n).BackColor = RGB(0, 128, 128)
        Case "Poblado"
            Option1(n).BackColor = RGB(204, 102, 204)
            'Option1(n).Picture = LoadPicture(App.Path & "\Graficos\azul.bmp")
        Case "Costa"
            Option1(n).BackColor = RGB(150, 200, 150)
        Case "Dungeon"
            Option1(n).BackColor = RGB(200, 140, 130)
        Case "Ciudad"
        
         If n = 1 And Due�oUlla = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        If n = 1 And Due�oUlla = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        If n = 20 And Due�oDesierto = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        If n = 20 And Due�oDesierto = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        
        If n = 59 And Due�oBander = 1 Then
        Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        Option1(58).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        Option1(60).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        Option1(61).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        Option1(66).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
 
        End If
        
        If n = 59 And Due�oBander = 2 Then
        Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        Option1(58).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        Option1(60).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        Option1(61).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        Option1(66).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")

        End If
        
        If n = 62 And Due�oLindos = 1 Then
        Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        Option1(63).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        Option1(64).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        End If
        
        If n = 62 And Due�oLindos = 2 Then
        Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        Option1(63).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        Option1(64).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        End If
        
        '[Tite]Nix neutral
        If n = 34 Then Option1(n).BackColor = RGB(255, 170, 85)
        'If n = 34 And Due�oNix = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        'If n = 34 And Due�oNix = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        '[\Tite]
        
        If n = 81 And Due�oDescanso = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        If n = 81 And Due�oDescanso = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        
        If n = 84 And Due�oAtlantis = 1 Then
        Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        Option1(83).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        Option1(85).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        End If
        
        If n = 84 And Due�oAtlantis = 2 Then
        Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        Option1(83).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        Option1(85).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        End If
        
        If (n = 111 Or n = 112) And Due�oEsperanza = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        If (n = 111 Or n = 112) And Due�oEsperanza = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")

        If (n = 150 Or n = 151) And Due�oArghal = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        If (n = 150 Or n = 151) And Due�oArghal = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
        
        If n = 157 And Due�oQuest = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        If n = 157 And Due�oQuest = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")


        If n = 170 And Due�oCaos = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        If n = 170 And Due�oCaos = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")


If n = 20 And Due�oDesierto = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
If n = 20 And Due�oDesierto = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")


         If (n = 183 Or n = 184) And Due�oLaurana = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        If (n = 183 Or n = 184) And Due�oLaurana = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")


                Option1(n).ForeColor = vbWhite
       
        Case "Desierto"
            Option1(n).BackColor = &H80FFFF
        Case "Castillo"
            Option1(n).BackColor = RGB(205, 128, 0)
        Case "Quest"
            Option1(n).BackColor = &HFF00FF
        Case "Mina"
            Option1(n).BackColor = RGB(192, 192, 192)
        Case "Piramide"
            Option1(n).BackColor = RGB(200, 140, 130)
        Case "Nieve"
            Option1(n).BackColor = &HFFFFFF
        Case "Isla"
            Option1(n).BackColor = RGB(220, 185, 185)
        Case "Encantada"
            Option1(n).BackColor = RGB(200, 140, 130)
        Case Else
            Option1(n).Enabled = False
            Option1(n).Visible = False
    End Select
Next
End Sub

Private Sub Option1_Click(Index As Integer)

busca7 = Index
Mapa8 = "Mapa" & busca7
Label6.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "bichos")
Label5.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "Info")
Label4.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "nombre")
Text3.Text = busca7
End Sub


