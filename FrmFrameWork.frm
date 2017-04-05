VERSION 5.00
Begin VB.Form FrmFrameWork 
   Caption         =   "MachiMatch"
   ClientHeight    =   6645
   ClientLeft      =   2430
   ClientTop       =   2265
   ClientWidth     =   12435
   Icon            =   "FrmFrameWork.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   12435
   Begin VB.PictureBox PctOrdenar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   6000
      ScaleHeight     =   945
      ScaleWidth      =   1905
      TabIndex        =   97
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton BtnLstOrdenar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Por %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton BtnLstOrdenar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Por Archivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton BtnLstOrdenar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Por nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton BtnLstOrdenar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ninguno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.CommandButton BtnBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   96
      Top             =   5880
      Width           =   735
   End
   Begin VB.PictureBox PctFiltros 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   3960
      ScaleHeight     =   1425
      ScaleWidth      =   1905
      TabIndex        =   87
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton BtnLstFiltros 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Marcados para usar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton BtnLstFiltros 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sin usar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton BtnLstFiltros 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Repetidos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton BtnLstFiltros 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sin match"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton BtnLstFiltros 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Solo con error"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton BtnLstFiltros 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.CommandButton BtnFiltros 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   5880
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   -3120
      ScaleHeight     =   2385
      ScaleWidth      =   3345
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label LblRenombrando 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   37
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RENOMBRANDO..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.PictureBox PctCalculando 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   -3120
      ScaleHeight     =   2385
      ScaleWidth      =   3345
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   18
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CALCULANDO ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.PictureBox PctLstArchivos 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   11880
      ScaleHeight     =   5745
      ScaleWidth      =   5025
      TabIndex        =   78
      Top             =   -5160
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton BtnCancelarLstArchivos 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   81
         Top             =   5280
         Width           =   1695
      End
      Begin VB.ListBox LstArchivos 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4680
         Left            =   120
         TabIndex        =   80
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label LblTituloLstArchivos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lista de archivos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   0
         TabIndex        =   79
         Top             =   0
         Width           =   5055
      End
   End
   Begin VB.PictureBox PctBotones 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   -5760
      ScaleHeight     =   4185
      ScaleWidth      =   5985
      TabIndex        =   22
      Top             =   3960
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Frame frmExportarTexto 
         Caption         =   "Exportar a archivos de texto"
         Height          =   1815
         Left            =   2160
         TabIndex        =   51
         Top             =   1920
         Width           =   3735
         Begin VB.CommandButton BtnExportarArchivos 
            Caption         =   "Todos los archivos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   57
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton BtnExportarArchivosNOusados 
            Caption         =   "Archivos no usados"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   56
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton BtnExportarArchivosUsados 
            Caption         =   "Archivos usados"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   55
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton BtnExportarNombres 
            Caption         =   "Todos los nombres"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton BtnExportarNombresNOUsados 
            Caption         =   "Nombres no usados"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton BtnExportarNombresUsados 
            Caption         =   "Nombres usados"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame FrmAcciones 
         Caption         =   "Acciones"
         Height          =   1935
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   1935
         Begin VB.CheckBox Check4 
            Caption         =   "Mover también las ROM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   65
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Mantener copia de los archivos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton BtnMoverSNAP 
            Caption         =   "Mover archivos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton BtnRenombrarSNAP 
            Caption         =   "Renombrar archivos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton BtnTerminar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   30
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label lblSNAPTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   63
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LblSNAPSinUsar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   62
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Totales"
         Height          =   255
         Left            =   4440
         TabIndex        =   61
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "sin coincidencia"
         Height          =   255
         Left            =   2400
         TabIndex        =   60
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblROMTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   59
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Totales"
         Height          =   255
         Left            =   4440
         TabIndex        =   58
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "sin coincidencia"
         Height          =   255
         Left            =   2400
         TabIndex        =   50
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "a usar"
         Height          =   255
         Left            =   1080
         TabIndex        =   49
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "a usar"
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblROMusadas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.Label LblSNAPRepetidas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblSNAPusadas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LblROMSinUsar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "total ARCHIVOS a ser renombrados más de una vez"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label19 
         Caption         =   "ARCHIVOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
      Begin VB.Label LblRom 
         Caption         =   "NOMBRES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   6015
      End
   End
   Begin VB.Frame FrameResultados 
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   3840
      TabIndex        =   66
      Top             =   120
      Width           =   8535
      Begin VB.PictureBox PctTblResultados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5055
         Left            =   120
         ScaleHeight     =   5025
         ScaleWidth      =   8265
         TabIndex        =   67
         Top             =   240
         Width           =   8295
         Begin VB.ComboBox CmbB 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            ItemData        =   "FrmFrameWork.frx":0E92
            Left            =   3240
            List            =   "FrmFrameWork.frx":0E94
            TabIndex        =   77
            Top             =   240
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.CheckBox ChekUsar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   270
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   5055
            Left            =   8040
            TabIndex        =   72
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LblP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   6000
            TabIndex        =   76
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label LblNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   75
            Top             =   240
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label LblCheck 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   73
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label LblTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   6000
            TabIndex        =   71
            Top             =   0
            Width           =   735
         End
         Begin VB.Label LblTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Archivos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   70
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label LblTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombres"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   69
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label LblTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Usar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   68
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.Timer TimerResize 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   5640
   End
   Begin VB.CommandButton BtnRenombrar 
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Frame FrameArchivo 
      Caption         =   "Directorio de archivos a renombrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   3615
      Begin VB.TextBox TxtExtSNAP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Text            =   ".png|.jpg"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtPathSNAP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton BtnSelectSNAP 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Extensión:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Ruta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton BtnCalcular 
      Caption         =   "Calcular"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   3615
      Begin VB.CheckBox ChkTurbo 
         Height          =   255
         Left            =   3120
         TabIndex        =   83
         Top             =   1560
         Width           =   375
      End
      Begin VB.ComboBox CmbAlgoritmo 
         Height          =   315
         ItemData        =   "FrmFrameWork.frx":0E96
         Left            =   840
         List            =   "FrmFrameWork.frx":0EA3
         TabIndex        =   82
         Text            =   "Ratcliff/Obershelp"
         Top             =   360
         Width           =   2655
      End
      Begin VB.CheckBox CheckRepetida 
         Height          =   255
         Left            =   3120
         TabIndex        =   34
         Top             =   1320
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.CheckBox CheckRepetidaCalcular 
         Height          =   255
         Left            =   3120
         TabIndex        =   31
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TxtSimilMin 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Text            =   "60"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label LblCheckTurbo 
         Caption         =   "Acelerar busqueda en listas ordenadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label lblAlgoritmo 
         Caption         =   "Algoritmo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Permitir SNAP repetidas al renombrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label14 
         Caption         =   "Repetir SNAP al calcular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Similitud Min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame FrameNombres 
      Caption         =   "Obtner lista de nombres"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox ChkExtension 
         Height          =   255
         Left            =   2760
         TabIndex        =   85
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CmbDesde 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "FrmFrameWork.frx":0ED6
         Left            =   1320
         List            =   "FrmFrameWork.frx":0EE3
         TabIndex        =   40
         Text            =   "Combo1"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox TxtPropiedad 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   39
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox TxtNodo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox TxtExtROM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Text            =   ".*"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton BtnSelectROM 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox TxtPathROM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label22 
         Caption         =   "Propiedad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Nodo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Obtener desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Ruta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Extensión:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   2535
      End
   End
   Begin VB.CommandButton BtnOrdenar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label LblOrdenar 
      Caption         =   "Ordenar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   102
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label LblFiltros 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   94
      Top             =   5520
      Width           =   735
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu MnuIdioma 
         Caption         =   "Idioma"
         Begin VB.Menu MnuEspanol 
            Caption         =   "Español"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuIngles 
            Caption         =   "English"
         End
      End
      Begin VB.Menu MnuGuion01 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu MnuAcercade 
         Caption         =   "Acerca de ..."
      End
   End
End
Attribute VB_Name = "FrmFrameWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

'Función para abrir el openfile dialog
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long


Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Dim Foco As Integer
Dim FactorScroll As Long
Public Filtro As Integer '0=todo, 1=con error, 2=solo missmatch, 3= solo repetidas
Public Ordenar As Integer '0=Ninguno, 1 = ROM, 2 = SNAP, 3 = PORCENTAJE, 4 = ROM INVERSO, 5 = SNAP INVERSA, 6 = PORCENTAJE INVERSO
'Dim Ordenar As Integer '0=nada 1=check usar, 2=nombre ROM, 3=nombre SNAP 4=porcentaje similitud

Dim indexRomSeleccionManual As Integer 'Almacena el indice de la rom a la cual se le está agregando una SNAP de forma manual

Private Sub CmbB_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
Exit Sub
End Sub



'************************************************************************************
'Forulario
'************************************************************************************
Private Sub Form_Load()
Dim i As Long

Me.Caption = "MachiMatch  v:" & App.Major & "." & App.Minor & "." & App.Revision

'Posicionamos el formulario al centro
If Me.Width > Screen.Width Then Me.Width = Screen.Width
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

'Cargamos valores por defecto
BtnRenombrar.Enabled = False
VScroll1.Enabled = False
LblFiltros.Enabled = False
BtnFiltros.Enabled = False
LblOrdenar.Enabled = False
BtnOrdenar.Enabled = False
BtnBuscar.Enabled = False
'fraFiltro.Enabled = False
FactorScroll = 1

CmbDesde.ListIndex = 0
TxtSimilMin = Format(60, "00.0")

ReDim lstIndex(0)


iIdioma = CInt(IniGet(App.Path & "\config.ini", "general", "idioma", "2"))
If iIdioma = 2 And ModVariables.Get_locale() = 1033 Then 'Idioma automático a ingles
    iIdioma = 1
ElseIf iIdioma = 2 Then 'idioma automático a español
    iIdioma = 0
End If

If iIdioma = 0 Then
    'Idioma español
    Me.MnuEspanol.Checked = True
    Me.MnuIngles.Checked = False
    CambiarIdioma 0
Else
    'Idioma ingles
    Me.MnuEspanol.Checked = False
    Me.MnuIngles.Checked = True
    CambiarIdioma 1
End If

'Cargar rutas
CmbDesde.ListIndex = CInt(IniGet(App.Path & "/config.ini", "last", "romtipo", 0))
TxtNodo.Text = IniGet(App.Path & "/config.ini", "last", "romnodo", "")
TxtPropiedad.Text = IniGet(App.Path & "/config.ini", "last", "rompropiedad", "")
TxtExtROM.Text = IniGet(App.Path & "/config.ini", "last", "romext", ".*")
TxtPathROM.Text = IniGet(App.Path & "/config.ini", "last", "path_rom", App.Path)
    
TxtExtSNAP.Text = IniGet(App.Path & "/config.ini", "last", "snapext", ".*")
TxtPathSNAP.Text = IniGet(App.Path & "/config.ini", "last", "path_snap", App.Path)
    
CmbAlgoritmo.ListIndex = CInt(IniGet(App.Path & "/config.ini", "last", "algoritmo", 0))
TxtSimilMin.Text = IniGet(App.Path & "/config.ini", "last", "simil", "60")
CheckRepetidaCalcular.Value = CInt(IniGet(App.Path & "/config.ini", "last", "repetirsnap", 0))
CheckRepetida.Value = CInt(IniGet(App.Path & "/config.ini", "last", "permitirsnap", 1))
ChkTurbo.Value = CInt(IniGet(App.Path & "/config.ini", "last", "acelerarsnap", 0))

'Cargar lista de commandos a reemplazar en XML
ReDim lstCharSearch(1)
ReDim lstCharReplace(1)

lstCharSearch(0) = "&amp;"
lstCharReplace(0) = "&"
lstCharSearch(1) = "&apos;"
lstCharReplace(1) = "'"

Dim auxSearch As String
Dim auxReplace As String

auxSearch = IniGet(App.Path & "/config.ini", "xml", "search_0", "")
auxReplace = GetAscii(IniGet(App.Path & "/config.ini", "xml", "Replace_0", ""))
i = 0
Do While auxSearch <> ""
    'Incrementamos en uno el valor
    ReDim Preserve lstCharSearch(2 + i)
    ReDim Preserve lstCharReplace(2 + i)
    
    lstCharSearch(2 + i) = auxSearch
    lstCharReplace(2 + i) = auxReplace
    
    i = i + 1
    
    auxSearch = IniGet(App.Path & "/config.ini", "xml", "search_" & i, "")
    auxReplace = GetAscii(IniGet(App.Path & "/config.ini", "xml", "Replace_" & i, ""))
Loop

'Centrar ventanas
PctCalculando.Left = (Me.Width - PctCalculando.Width) / 2
PctCalculando.Top = (Me.Height - PctCalculando.Height) / 2

Picture1.Left = (Me.Width - Picture1.Width) / 2
Picture1.Top = (Me.Height - Picture1.Height) / 2

End Sub

Private Sub MnuAcercade_Click()
Call MsgBox("Programa desarrollado por:" & vbCrLf & _
"Hechelion (Alberto Ortiz)" & vbCrLf & _
"hechelion@gmail.com" & vbCrLf & _
"Para la comunidad de arcadespain.info" & vbCrLf & _
vbCrLf & _
"Agradecimientos a:" & vbCrLf & _
"Machiminax" & vbCrLf & _
"gucaza" & vbCrLf & _
"getterrobot" & vbCrLf & _
"Pevalle" & vbCrLf & _
"empardopo" & vbCrLf & _
"onofer" & vbCrLf & _
"" & vbCrLf & _
"Licencia:" & vbCrLf & _
"This work is licensed under the Creative Commons Attribution 4.0 International License." & vbCrLf & _
"To view a copy of this license," & vbCrLf & _
"visit http://creativecommons.org/licenses/by/4.0/" & vbCrLf & _
"or send a letter to Creative Commons, PO Box 1866, Mountain View, CA 94042, USA.", vbOKOnly, "Acerca de...")
End Sub

Private Sub MnuEspanol_Click()
MnuEspanol.Checked = True
MnuIngles.Checked = False
IniWrite App.Path & "\config.ini", "general", "idioma", "0"
CambiarIdioma 0
End Sub

Private Sub MnuIngles_Click()
MnuEspanol.Checked = False
MnuIngles.Checked = True
IniWrite App.Path & "\config.ini", "general", "idioma", "1"
CambiarIdioma 1
End Sub

Private Sub MnuSalir_Click()
End
End Sub


Private Sub BtnSelectROM_Click()
If CmbDesde.ListIndex = 0 Then
    'Opens a Treeview control that displays the directories in a computer

   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo

   szTitle = "Seleccionar directorio ROM"
   With tBrowseInfo
      .hwndOwner = Me.hWnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
   End With

   lpIDList = SHBrowseForFolder(tBrowseInfo)

   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      TxtPathROM.Text = sBuffer
   End If
Else 'Abrir open file dialog desde API
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim sFilter As String
    
    'Seteamos el filtro según sea la opción seleccionada
    If CmbDesde.ListIndex = 1 Then
        sFilter = "Archivo de texto (*.txt)" & Chr(0) & "*.TXT" & Chr(0)
    Else
        sFilter = "Archivo XML (*.xml)" & Chr(0) & "*.XML" & Chr(0)
    End If
    
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = Me.hWnd
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = App.Path
    OpenFile.lpstrTitle = "Seleccione archivo para cargar nombres"
    OpenFile.flags = 0
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
       'MsgBox "The User pressed the Cancel Button"
    Else
       TxtPathROM.Text = Trim(OpenFile.lpstrFile)
    End If
End If
End Sub

Private Sub BtnSelectSNAP_Click()
'Opens a Treeview control that displays the directories in a computer

   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo

   szTitle = "Seleccionar directorio SNAP"
   With tBrowseInfo
      .hwndOwner = Me.hWnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
   End With

   lpIDList = SHBrowseForFolder(tBrowseInfo)

   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      TxtPathSNAP.Text = sBuffer
   End If
End Sub

Private Sub CheckRepetida_Click()
If CheckRepetida.Value = 0 Then CheckRepetidaCalcular.Value = 0
End Sub

Private Sub CheckRepetidaCalcular_Click()
If CheckRepetidaCalcular.Value = 1 Then CheckRepetida.Value = 1
End Sub

Private Sub CmbAlgoritmo_Click()
If CmbAlgoritmo.ListIndex = 1 Then
    Label10.Enabled = False
    TxtSimilMin.Enabled = False
    LblCheckTurbo.Enabled = True
    ChkTurbo.Enabled = True
Else
    Label10.Enabled = True
    TxtSimilMin.Enabled = True
    LblCheckTurbo.Enabled = False
    ChkTurbo.Value = 0
    ChkTurbo.Enabled = False
End If
End Sub

'************************************************************************************
'Formulario cambio de tamaño
'************************************************************************************
Private Sub Form_Resize()
TimerResize.Enabled = False

'cambiar tamaño, alto
If FrmFrameWork.Height > 7000 Then
    BtnCalcular.Top = FrmFrameWork.Height - 1350
    BtnRenombrar.Top = FrmFrameWork.Height - 1350
    LblFiltros.Top = FrmFrameWork.Height - 1470
    BtnFiltros.Top = FrmFrameWork.Height - 1230
    PctFiltros.Top = FrmFrameWork.Height - 2415
    LblOrdenar.Top = FrmFrameWork.Height - 1470
    BtnOrdenar.Top = FrmFrameWork.Height - 1230
    PctOrdenar.Top = FrmFrameWork.Height - 2175
    BtnBuscar.Top = FrmFrameWork.Height - 1230
    
    FrameResultados.Height = FrmFrameWork.Height - 1590
    PctTblResultados.Height = CInt((FrameResultados.Height - 360) / LblCheck(0).Height) * LblCheck(0).Height
    VScroll1.Height = PctTblResultados.ScaleHeight
    
    PctBotones.Top = FrameResultados.Top + (FrameResultados.Height - PctBotones.Height) / 2
    PctLstArchivos.Top = FrameResultados.Top + (FrameResultados.Height - PctLstArchivos.Height) / 2


End If
'Cambiar tamaño, ancho
If FrmFrameWork.Width > 10950 Then
    BtnRenombrar.Left = FrmFrameWork.Width - 1995
    
    'frameFiltros.Width = FrmFrameWork.Width - 5940
    FrameResultados.Width = FrmFrameWork.Width - 4020
    PctTblResultados.Width = FrameResultados.Width - 240
    VScroll1.Left = PctTblResultados.ScaleWidth - VScroll1.Width

    PctBotones.Left = FrameResultados.Left + (FrameResultados.Width - PctBotones.Width) / 2
    PctLstArchivos.Left = FrameResultados.Left + (FrameResultados.Width - PctLstArchivos.Width) / 2
    
    LblTitulo(0).Left = 0
    LblTitulo(1).Left = LblTitulo(0).Width
    LblTitulo(1).Width = (PctTblResultados.Width - (LblTitulo(0).Width + LblTitulo(3).Width + VScroll1.Width)) / 2
    LblTitulo(2).Left = LblTitulo(1).Left + LblTitulo(1).Width
    LblTitulo(2).Width = LblTitulo(1).Width
    LblTitulo(3).Left = LblTitulo(2).Left + LblTitulo(2).Width
    
End If

TimerResize.Enabled = True
End Sub

Private Sub TimerResize_Timer()
TimerResize.Enabled = False
ResizeForm
End Sub

Private Sub ResizeForm()
Dim i As Long
'Adaptamos el ancho
For i = 0 To LblCheck.Count - 1
    LblCheck(i).Width = LblTitulo(0).Width
    LblNombre(i).Left = LblTitulo(1).Left
    LblNombre(i).Width = LblTitulo(1).Width
    CmbB(i).Left = LblTitulo(2).Left
    CmbB(i).Width = LblTitulo(2).Width
    LblP(i).Left = LblTitulo(3).Left
    LblP(i).Width = LblTitulo(3).Width
Next

'Adaptamos el alto y lo refrescamos
If UBound(lstIndex) > 0 Then
    movRoll = True
    Call llenarResultado
    movRoll = False
End If
End Sub

'************************************************************************************
'Calcular
'************************************************************************************

Private Sub BtnCalcular_Click()
Dim i As Long
Dim Comprobar As Boolean

'On Error GoTo msgError

Comprobar = True

'Borramos cualquier resultado previo de la tabla de resultados
Call BorrarResultados

If Right(TxtPathROM.Text, 1) = "\" Then TxtPathROM.Text = Left(TxtPathROM.Text, Len(TxtPathROM.Text) - 1)

'Comprobar directorios
If CmbDesde.ListIndex = 0 Then

    If Dir(TxtPathROM.Text, vbDirectory) = "" Then
        Comprobar = False
        Call MsgBox(lstTextos(0), vbCritical, "ERROR")
    ElseIf Dir(TxtPathROM.Text, vbArchive) <> "" Then
        Comprobar = False
        Call MsgBox(lstTextos(0), vbCritical, "ERROR")
    End If
ElseIf CmbDesde.ListIndex = 1 Then
    If Dir(TxtPathROM.Text, vbArchive) = "" Then
        Comprobar = False
        Call MsgBox(lstTextos(1), vbCritical, "ERROR")
    ElseIf LCase(Right(TxtPathROM, 4)) <> ".txt" Then
        Comprobar = False
        Call MsgBox(lstTextos(1), vbCritical, "ERROR")
    End If
Else
    If Dir(TxtPathROM.Text, vbArchive) = "" Then
        Comprobar = False
        Call MsgBox(lstTextos(1), vbCritical, "ERROR")
    ElseIf LCase(Right(TxtPathROM, 4)) <> ".xml" Then
        Comprobar = False
        Call MsgBox(lstTextos(1), vbCritical, "ERROR")
    End If
    If CmbDesde.ListIndex = 2 And TxtNodo = "" Then
        Comprobar = False
        Call MsgBox(lstTextos(2), vbCritical, "ERROR")
    End If
End If

If Dir(TxtPathSNAP.Text, vbDirectory) = "" Then
    Comprobar = False
    Call MsgBox(lstTextos(3), vbCritical, "ERROR")
End If

If IsNumeric(TxtSimilMin.Text) Then
    If TxtSimilMin.Text < 0 Or TxtSimilMin.Text > 100 Then
        Comprobar = False
        Call MsgBox(lstTextos(4), vbCritical, "ERROR")
    End If
Else
    Comprobar = False
    Call MsgBox(lstTextos(4), vbCritical, "ERROR")
End If

If Comprobar Then
    Call IniWrite(App.Path & "/config.ini", "last", "romtipo", CmbDesde.ListIndex)
    Call IniWrite(App.Path & "/config.ini", "last", "romnodo", TxtNodo.Text)
    Call IniWrite(App.Path & "/config.ini", "last", "rompropiedad", TxtPropiedad.Text)
    Call IniWrite(App.Path & "/config.ini", "last", "romext", TxtExtROM.Text)
    Call IniWrite(App.Path & "/config.ini", "last", "path_rom", TxtPathROM.Text)
    
    Call IniWrite(App.Path & "/config.ini", "last", "snapext", TxtExtSNAP.Text)
    Call IniWrite(App.Path & "/config.ini", "last", "path_snap", TxtPathSNAP.Text)
    
    Call IniWrite(App.Path & "/config.ini", "last", "algoritmo", CmbAlgoritmo.ListIndex)
    Call IniWrite(App.Path & "/config.ini", "last", "simil", TxtSimilMin.Text)
    Call IniWrite(App.Path & "/config.ini", "last", "repetirsnap", CheckRepetidaCalcular.Value)
    Call IniWrite(App.Path & "/config.ini", "last", "permitirsnap", CheckRepetida.Value)
    Call IniWrite(App.Path & "/config.ini", "last", "acelerarsnap", ChkTurbo.Value)
    
    PctCalculando.Visible = True
    
    If CmbDesde.ListIndex = 0 Then
        sinROM = False
    Else
        sinROM = True
    End If
    
    'Antes de calcular reiniciamos algunos valores
    Filtro = 0
    Ordenar = 0
    BtnOrdenar.Caption = BtnLstOrdenar(0).Caption
    'chkFiltro(0).Value = True
    Call Calcular
End If

Exit Sub
msgError:
If Err.Number = 52 Then
    MsgBox lstTextos(5), vbCritical, "ERROR"
Else
    MsgBox Err.Description & vbCrLf & Err.Number, vbCritical, "ERROR"
End If
End Sub

Public Sub Calcular()
Dim TotalRom As Long
Dim TotalSnap As Long
Dim i As Long
Dim e As Long
Dim auxUltimaSnap As Long
Dim auxIndex As Long

Dim total As Double
Dim contador As Double
Dim valorRatCliff As Double
Dim SimilMin As Double

Dim extROM As String
Dim extSNAP As String

Dim auxPos As Integer
Dim auxNombre As String
Dim auxExtension As String

Dim nFic
Dim Linea
Dim sFileName
Dim Texto As String

extROM = LCase(TxtExtROM.Text)
extSNAP = LCase(TxtExtSNAP.Text)

'Limpiamos las listas de rom y snap
ReDim lstROM(0)
ReDim lstSNAP(0)

'Obtenemos la lista de nombres
If CmbDesde.ListIndex = 0 Then 'DESDE ARCHIVO
    'Obtenemos la lista de archivo ROM
    If Right(TxtPathROM.Text, 1) <> "\" Then TxtPathROM.Text = TxtPathROM.Text & "\"
    sFileName = Dir(TxtPathROM.Text)
    
    'Contamos total de archivos existentes para redimencionar lista de ROM
    TotalRom = 0
    Do While sFileName > ""
        TotalRom = TotalRom + 1
        sFileName = Dir()
    Loop
    ReDim lstROM(TotalRom - 1)
    
    'Comenzamos a filtrar cada archivos existente en la ruta, según sea el filtro de extensión
    i = 0
    TotalRom = 0
    sFileName = Dir(TxtPathROM.Text)
    Do While sFileName > ""
        auxPos = InStrRev(sFileName, ".") - 1
        If auxPos = -1 Then
            'No se encontro extensión
            auxNombre = Trim(LCase(sFileName))
            auxExtension = ""
        ElseIf auxPos = 0 Then
            'El archivo no tiene nombre, solo extensión
            auxNombre = ""
            auxExtension = Trim(LCase(sFileName))
        Else
            'Todo normal
            auxExtension = Trim(LCase(Right(sFileName, Len(sFileName) - (auxPos + 1))))
            auxNombre = Trim(Left(sFileName, auxPos))
        End If
        
        If InStr(extROM, "." & auxExtension) > 0 Or extROM = ".*" Then
            lstROM(i).full_name = sFileName
            lstROM(i).lcase_name = LCase(auxNombre)
            lstROM(i).name_SinExtension = auxNombre
            lstROM(i).Extension = auxExtension
            TotalRom = 1 + TotalRom
            i = i + 1
        End If
        sFileName = Dir()
    Loop
ElseIf CmbDesde.ListIndex = 1 Then 'DESDE TEXTO
    'Obtenemos la lista de nombres desde un archivo de texto
    nFic = FreeFile
    Open TxtPathROM.Text For Input As nFic
        Do While Not EOF(nFic)
            Line Input #nFic, Linea
            
            If ChkExtension.Value = 1 Then 'Si hay que buscar extensión
                auxPos = InStrRev(Linea, ".") - 1
                If auxPos = -1 Then
                    'No se encontro extensión
                    auxNombre = Trim(Linea)
                    auxExtension = ""
                ElseIf auxPos = 0 Then
                    'El archivo no tiene nombre, solo extensión
                    auxNombre = ""
                    auxExtension = Trim(Linea)
                Else
                    'Todo normal
                    auxExtension = Trim(Right(Linea, Len(Linea) - (auxPos + 1)))
                    auxNombre = Trim(Left(Linea, auxPos))
                End If
            Else 'Si no hay que buscar extensión
                auxNombre = Trim(Linea)
                auxExtension = ""
            End If
            
            ReDim Preserve lstROM(TotalRom)
            
            lstROM(TotalRom).full_name = Linea
            lstROM(TotalRom).lcase_name = LCase(auxNombre)
            lstROM(TotalRom).name_SinExtension = auxNombre
            lstROM(TotalRom).Extension = auxExtension
            
            TotalRom = 1 + TotalRom
        Loop
    Close nFic
    
Else 'DESDE XML
    'Obtenemos la lista de nombres desde un archivo XML

    nFic = FreeFile
    Open TxtPathROM.Text For Input As nFic
        Texto = Input(LOF(nFic), nFic)
    Close nFic
    
    Dim nTexto As String
    Dim nodos() As String
    Dim auxNodos() As String
    Dim auxValor As Integer
    Dim auxInicio As Integer
    Dim auxFin As Integer
    'Dim propiedades() As String
    Dim valores() As String
    
    nTexto = Replace(Texto, vbCrLf, "")
    nodos = Split(nTexto, "<" & TxtNodo.Text)
    For i = 1 To UBound(nodos)
        auxNodos = Split(nodos(i), ">")
        If TxtPropiedad.Text = "" Then
            'Si es un valor de nodo
            valores = Split(auxNodos(1), "<")
            Linea = Trim(valores(0))
        Else
            'Si es una propiedad
            auxValor = InStr(LCase(auxNodos(0)), (Trim(LCase(TxtPropiedad.Text) & "=")))
            If auxValor = 0 Then auxValor = InStr(auxNodos(0), CStr(Trim(TxtPropiedad.Text) & " ="))
            
            If auxValor > 0 Then
                auxInicio = InStr(auxValor, auxNodos(0), """")
                auxFin = InStr(auxInicio + 1, auxNodos(0), """")
                Linea = Mid(auxNodos(0), auxInicio + 1, auxFin - auxInicio - 1)
            Else
                Linea = ""
            End If
        End If
        
        If Linea <> "" Then
            Linea = ModHTML.decodingXML(CStr(Linea)) 'Si leemos desde un archivo XML entonces debemos cambiar los signos copdificados en XML
            If ChkExtension.Value = 1 Then 'Si hay que buscar extensión
                auxPos = InStrRev(Linea, ".") - 1
                If auxPos = -1 Then
                    'No se encontro extensión
                    auxNombre = Trim(Linea)
                    auxExtension = ""
                ElseIf auxPos = 0 Then
                    'El archivo no tiene nombre, solo extensión
                    auxNombre = ""
                    auxExtension = Trim(Linea)
                Else
                    'Todo normal
                    auxExtension = Trim(Right(Linea, Len(Linea) - (auxPos + 1)))
                    auxNombre = Trim(Left(Linea, auxPos))
                End If
            Else 'Si no hay que buscar extensión
                auxNombre = Trim(Linea)
                auxExtension = ""
            End If
            
            ReDim Preserve lstROM(TotalRom)
            
            lstROM(TotalRom).full_name = Linea
            lstROM(TotalRom).lcase_name = LCase(auxNombre)
            lstROM(TotalRom).name_SinExtension = auxNombre
            lstROM(TotalRom).Extension = auxExtension
            
            TotalRom = 1 + TotalRom
        End If
    Next

End If
'Acomodamos el largo de la lista de rom al total de rom que hayan pasado el filtro
If TotalRom > 0 Then ReDim Preserve lstROM(TotalRom - 1)


If TotalRom > 0 Then
    'Obtenemos la lista de archivos SNAP, solo si primero encontramos nombres validos
    If Right(TxtPathSNAP.Text, 1) <> "\" Then TxtPathSNAP.Text = TxtPathSNAP.Text & "\"
    sFileName = Dir(TxtPathSNAP.Text)
    
    'Contamos el número de archivos en la ruta de SNAP
    TotalSnap = 0
    Do While sFileName > ""
        TotalSnap = TotalSnap + 1
        sFileName = Dir()
    Loop
    ReDim lstSNAP(TotalSnap)
    
    
    i = 0
    TotalSnap = 0
    LstArchivos.Clear 'Limpiamos la lista de archivos
    sFileName = Dir(TxtPathSNAP.Text)
    Do While sFileName > ""
        auxPos = InStrRev(sFileName, ".") - 1
        If auxPos = -1 Then
            'No se encontro extensión
            auxNombre = Trim(LCase(sFileName))
            auxExtension = ""
        ElseIf auxPos = 0 Then
            'El archivo no tiene nombre, solo extensión
            auxNombre = ""
            auxExtension = Trim(LCase(sFileName))
        Else
            'Todo normal
            auxExtension = Trim(LCase(Right(sFileName, Len(sFileName) - (auxPos + 1))))
            auxNombre = Trim(Left(sFileName, auxPos))
        End If
        If InStr(extSNAP, "." & auxExtension) > 0 Or extSNAP = ".*" Then
            lstSNAP(i).full_name = sFileName
            lstSNAP(i).lcase_name = LCase(auxNombre)
            lstSNAP(i).name_SinExtension = auxNombre
            lstSNAP(i).Extension = auxExtension
            i = i + 1
            TotalSnap = TotalSnap + 1
            
            'Agregamos la SNAP al listbox lstarchivos
            LstArchivos.AddItem sFileName
        End If
        sFileName = Dir()
    Loop
    
    
    If TotalSnap > 0 Then
        ReDim Preserve lstSNAP(TotalSnap - 1)
    Else
        MsgBox lstTextos(6), vbOKOnly, lstTextos(7)
    End If
Else
    MsgBox lstTextos(8), vbOKOnly, lstTextos(7)
End If

'Procedemos a evaluar cada nombre de rom sobre las snap
SimilMin = TxtSimilMin.Text
total = CDbl(TotalRom) * TotalSnap 'calculamos el total de cálculos que serán necesarios para comprar el 100% de los nombres

If total > 0 Then
    i = 0
    Do While i < TotalRom
        e = 0
        If ChkTurbo.Value = 1 And CmbAlgoritmo.ListIndex = 1 Then e = auxUltimaSnap: contador = contador + e
        Do While e < TotalSnap
            If lstSNAP(e).have100 = False Or CheckRepetida.Value = 1 Then
                If CmbAlgoritmo.ListIndex = 0 Then
                    'Busqueda normal, usando RatCliff
                    valorRatCliff = comparar(lstROM(i).lcase_name, lstSNAP(e).lcase_name)
                    valorRatCliff = (2 * valorRatCliff / (Len(lstROM(i).lcase_name) + Len(lstSNAP(e).lcase_name))) * 100
                ElseIf CmbAlgoritmo.ListIndex = 1 Then
                    'Busqueda exacta
                    If lstROM(i).lcase_name = lstSNAP(e).lcase_name Then
                        valorRatCliff = 100
                    Else
                        valorRatCliff = 0
                    End If
                ElseIf CmbAlgoritmo.ListIndex = 2 Then
                    'Distancia de Levenshtein
                    valorRatCliff = CompararLevenshtein(lstROM(i).lcase_name, lstSNAP(e).lcase_name)
                End If
            Else
                'Si la SNAP ya está asignada a una ROM con un 100% de similitud y no se permiten SNAP repetidas en los nombres
                valorRatCliff = 0
            End If
            
            auxUltimaSnap = e 'Memorizamos la posición de E en la cual nos encontramos
            
            If valorRatCliff >= SimilMin Then Call agregar_resultado(e, i, valorRatCliff)
            If valorRatCliff = 100 Then 'Si las cadenas son exactamente iguales salimos del bucle para ganar tiempo
                contador = contador + (TotalSnap - 1 - e)
                Exit Do
            End If
            contador = contador + 1
            'Actualizamos el indicador de avance
            
            DoEvents
            Label15.Caption = Format((contador / total * 100), "00.00")
            e = e + 1
        Loop
        i = i + 1
    Loop
    
    'Calculamos si hay coincidencias repetidas
    'Limpiamos el contador de referencias
    For i = 0 To UBound(lstSNAP)
        lstSNAP(i).NumReferencias = 0
    Next
    
    'Contamos las referencias
    For i = 0 To UBound(lstROM)
        auxIndex = lstROM(i).new_name(lstROM(i).index_newName) - 1
        If auxIndex > -1 Then lstSNAP(auxIndex).NumReferencias = lstSNAP(auxIndex).NumReferencias + 1
    Next
    
    'Procesamos las repetidas
    For i = 0 To UBound(lstROM)
        auxIndex = lstROM(i).new_name(lstROM(i).index_newName) - 1
        If auxIndex > -1 Then 'Si tiene una SNAP asociada, revisamos el número de repeticiones que tiene dicha SNAP
            If lstSNAP(auxIndex).NumReferencias > 1 Then
                If CheckRepetidaCalcular.Value = 0 Then
                    lstROM(i).used = 0
                    lstROM(i).conError = 1
                    lstROM(i).NumError = SNAP_Repetida
                Else
                    lstROM(i).used = 1
                    lstROM(i).conError = 1
                    lstROM(i).NumError = SNAP_Repetida
                End If
                'accuRepetidas = accuRepetidas + 1
            End If
        Else 'Si no tiene asociada ninguna SNAP, marcamos el error
            lstROM(i).conError = 1
            lstROM(i).NumError = sin_SNAP
        End If
    Next
    
    BtnRenombrar.Enabled = True
    'fraFiltro.Enabled = True
    LblFiltros.Enabled = True
    BtnFiltros.Enabled = True
    LblOrdenar.Enabled = True
    BtnOrdenar.Enabled = True
    BtnBuscar.Enabled = True
    Call FiltrarResultado
End If
Foco = 0
PctCalculando.Visible = False
End Sub

Public Function Ratcliff(Cadena1 As String, Cadena2 As String) As Double
Dim Len1 As Integer
Dim Len2 As Integer

Len1 = Len(Cadena1)
Len2 = Len(Cadena2)

Ratcliff = 0
End Function

Public Function comparar(Cadena1 As String, Cadena2 As String) As Integer
Dim Len1 As Integer
Dim Len2 As Integer
Dim pos2 As Integer
Dim i As Long

Dim subCadenaL1 As String
Dim subCadenaL2 As String
Dim subCadenaR1 As String
Dim subCadenaR2 As String
Dim resultado As Integer

Dim TotalMatch As Integer

Len1 = Len(Cadena1)
Len2 = Len(Cadena2)

'Buscmos la subcadena de mayor tamaño
Do While Len1 > 0
    For i = 0 To (Len(Cadena1) - Len1)
        'gato = Mid(Cadena1, i + 1, Len1)
        pos2 = InStr(Cadena2, Mid(Cadena1, i + 1, Len1))
        If pos2 > 0 Then
            resultado = Len1
            subCadenaL1 = Left(Cadena1, i)
            subCadenaL2 = Left(Cadena2, pos2 - 1)
            subCadenaR1 = Right(Cadena1, Len(Cadena1) - (i + Len1))
            subCadenaR2 = Right(Cadena2, Len(Cadena2) - (pos2 - 1 + Len1))
            resultado = resultado + comparar(subCadenaL1, subCadenaL2)
            resultado = resultado + comparar(subCadenaR1, subCadenaR2)
            comparar = resultado
            Exit Function
        End If
    Next
    Len1 = Len1 - 1
Loop
comparar = 0
End Function

Public Function CompararLevenshtein(Cadena1 As String, Cadena2 As String) As Double
Dim coste As Integer
Dim n1 As Integer
Dim n2 As Integer
Dim m() As Integer
Dim numCambios As Integer
Dim largo As Integer
Dim i, i1, i2 As Long

coste = 0
n1 = Len(Cadena1)
n2 = Len(Cadena2)
ReDim m(n1, n2)

For i = 0 To n1
    m(i, 0) = i
Next
For i = 1 To n2
    m(0, i) = i
Next
For i1 = 1 To n1
    For i2 = 1 To n2
        coste = IIf(Mid(Cadena1, (i1), 1) = Mid(Cadena2, (i2), 1), 0, 1)
        m(i1, i2) = Min(Min(m(i1 - 1, i2) + 1, m(i1, i2 - 1) + 1), m(i1 - 1, i2 - 1) + coste)
    Next
Next

numCambios = m(n1, n2)
largo = IIf((n1 > n2), n1, n2)
CompararLevenshtein = 100# - (CLng(numCambios) * 100# / CLng(largo))
End Function

Public Function Min(valorA As Integer, valorB As Integer) As Integer
If valorA <= valorB Then Min = valorA Else Min = valorB
End Function


Public Sub agregar_resultado(numSNAP As Long, numROM As Long, simil As Double, Optional ForcePos As Integer = -1)
Dim A, B As Integer
lstROM(numROM).used = 1

If ForcePos = -1 Then
    For A = 0 To 9
        If lstROM(numROM).similitud(A) < simil Then Exit For
    Next
Else
    A = ForcePos
End If

'Solo si A se detuvo en una posición valida
If A < 10 Then
    For B = 9 To A + 1 Step -1
        lstROM(numROM).new_name(B) = lstROM(numROM).new_name(B - 1)
        lstROM(numROM).similitud(B) = lstROM(numROM).similitud(B - 1)
    Next
    lstROM(numROM).new_name(A) = numSNAP + 1
    lstROM(numROM).similitud(A) = simil
    If simil = 100 Then 'Si la similitud es igual al 100% entonces marcamos los archivos
        lstROM(numROM).have100 = True
        lstSNAP(numSNAP).have100 = True
    End If
End If

End Sub
'************************************************************************************
'Edición
'************************************************************************************
'*** selección manual
Private Sub LblNombre_DblClick(Index As Integer)
'LstArchivos.ListIndex = 0
indexRomSeleccionManual = lstIndex(Index + VScroll1.Value)
LblTituloLstArchivos.Caption = LblNombre(Index).Caption
PctLstArchivos.Visible = True
End Sub
Private Sub BtnCancelarLstArchivos_Click()
PctLstArchivos.Visible = False
End Sub

Private Sub LstArchivos_Click()
Dim indexRom As Long
Dim indexSnap_OLD As Integer
Dim indexSnap_NEW As Integer
Dim i As Long

indexRom = indexRomSeleccionManual 'Obtenemos el indice a la ROM que abrió esta ventana

'Optememos la refeencia a la SNAP previa si existe y a la que acabamos de agregar
indexSnap_OLD = lstROM(indexRom).new_name(lstROM(indexRom).index_newName) - 1
indexSnap_NEW = LstArchivos.ListIndex

'Procedemos a agregar el archivo seleccionado a la lista de de resultados posibles para la ROM
Call agregar_resultado(LstArchivos.ListIndex, indexRom, 110#, 0)

        
'Liberamos la vieja referencia si esta existe
If indexSnap_OLD > -1 Then lstSNAP(indexSnap_OLD).NumReferencias = lstSNAP(indexSnap_OLD).NumReferencias - 1

'incrementamos la referencia nueva
lstSNAP(indexSnap_NEW).NumReferencias = lstSNAP(indexSnap_NEW).NumReferencias + 1
        
'Asignamos el valor nuevo
lstROM(indexRom).index_newName = 0
        
'Evaluamos la referencia antigua
If indexSnap_OLD > -1 Then
    If lstSNAP(indexSnap_OLD).NumReferencias = 1 Then
        'Si luego de liberar la referenci antigua. Esta es referida por una única ROM. Entonces buscamos esa rom y desmarcamos la opción de repetido
        For i = 0 To UBound(lstROM)
            If lstROM(i).new_name(lstROM(i).index_newName) = indexSnap_OLD + 1 Then
                lstROM(i).used = 1
                lstROM(i).conError = 0
                lstROM(i).NumError = sin_error
                Exit For
            End If
        Next
    End If
End If

'Evaluamos la referencia nueva
If lstSNAP(indexSnap_NEW).NumReferencias >= 2 Then
    'Si la nueva referencia se encuentra repetida, entonces buscamos donde están y las marcamos como repetidas
    For i = 0 To UBound(lstROM)
        If lstROM(i).new_name(lstROM(i).index_newName) = indexSnap_NEW + 1 Then
            If CheckRepetida.Value = 0 Then lstROM(i).used = 0
            lstROM(i).conError = 1
            lstROM(i).NumError = SNAP_Repetida
            'Exit For
        End If
    Next
Else
    'Si la nueva SNAP seleccionada NO está repetida, entonces marcamos esta ROM como sin problema
    lstROM(indexRom).used = 1
    lstROM(indexRom).conError = 0
    lstROM(indexRom).NumError = sin_error
End If

Call llenarResultado

PctLstArchivos.Visible = False
End Sub


'*** Edición manual
Private Sub BtnFiltros_Click()
PctFiltros.Visible = True
End Sub

Private Sub BtnLstFiltros_Click(Index As Integer)
movRoll = True
Filtro = Index
BtnFiltros.Caption = BtnLstFiltros(Index).Caption
PctFiltros.Visible = False
Call BorrarResultados
Call FiltrarResultado
movRoll = False
End Sub

Private Sub BtnLstOrdenar_Click(Index As Integer)
movRoll = True
PctOrdenar.Visible = False
Me.Ordenar = IIf(Index > 0, IIf(Me.Ordenar = Index, Index + 3, Index), 0)
BtnOrdenar.Caption = BtnLstOrdenar(Index).Caption
If Ordenar > 3 Then BtnOrdenar.Caption = BtnOrdenar.Caption & "(I)"
Call BorrarResultados
Call FiltrarResultado
movRoll = False
End Sub

Private Sub BtnOrdenar_Click()
PctOrdenar.Visible = True
End Sub

Private Sub BtnBuscar_Click()
'aquí buscar
Dim i As Long
Dim auxValor As String
Dim auxIndice As Integer
Dim Rows As Integer 'Número de filas que se muestran en pantalla
Static InicioAnterior As Integer
Static BusquedaAnterior As String

auxValor = InputBox(lstTextos(28), lstTextos(29), BusquedaAnterior)
Rows = LblNombre.Count

If auxValor <> "" And DatosFiltrados > 0 Then
    If auxValor <> BusquedaAnterior Then
        BusquedaAnterior = auxValor
        InicioAnterior = 0
    End If
    For i = InicioAnterior To UBound(lstIndex)
        If InStr(LCase(lstROM(lstIndex(i)).name_SinExtension), LCase(auxValor)) > 0 Then
            'Si encontramos una coincidencia
            'Primero tratamos de mover el scroll para que el primer valor calce con el valor que buscamos
            If VScroll1.Max >= i / FactorScroll Then
                VScroll1.Value = i / FactorScroll
            Else
                VScroll1.Value = VScroll1.Max
            End If
            
            auxIndice = i - (VScroll1.Value * FactorScroll)
            'Actualizamos la visualización
            movRoll = True
            Call llenarResultado
            LblNombre(auxIndice).BackColor = &H80FF80
            movRoll = False
            Exit For
        End If
    Next
    InicioAnterior = i + 1
    If i > UBound(lstIndex) Then InicioAnterior = 0: MsgBox lstTextos(30)
End If

End Sub

Private Sub ChekUsar_Click(Index As Integer)
Dim indexRom As Long
If movRoll = False Then
    indexRom = lstIndex(Index + VScroll1.Value)
    'Si el cuadro no tiene coincidencia entonces no lo dejamos marcar
    If lstROM(indexRom).new_name(lstROM(indexRom).index_newName) = 0 Then
        ChekUsar(Index).Value = 0
    End If
    
    'If ChekUsar(Index).Value = 1 Then TxtA(Index).BackColor = &H80000005 Else TxtA(Index).BackColor = &H8080FF
    lstROM(indexRom).used = ChekUsar(Index).Value
    'If ChekUsar(Index).Value = 1 Then lstROM(indexROM).conError = 0 'Si lo vamos a usar, desmarcamos el flag de error
    If ChekUsar(Index).Value = 0 And lstROM(indexRom).NumError <> sin_error Then lstROM(indexRom).conError = 1 'Si lo desmarcamos, pero hay registrado algún tipo de error, marcamos el flag de error
 
    'lstROM(indexROM).conError = 0
    Call llenarResultado
End If
End Sub

Private Sub chkFiltro_Click(Index As Integer)
    Filtro = Index
    movRoll = True
    Call BorrarResultados
    Call FiltrarResultado
    movRoll = False
End Sub

Private Sub VScroll1_Change()
If Not movRoll Then
    movRoll = True
    Call llenarResultado
    movRoll = False
End If
End Sub

Private Sub VScroll1_Scroll()
movRoll = True
Call llenarResultado
movRoll = False
End Sub

Private Sub CmbB_Click(Index As Integer)
Dim indexSnap_OLD As Long
Dim indexSnap_NEW As Long
Dim i As Long
Dim indexRom As Long

indexRom = lstIndex(Index + VScroll1.Value)

    LblP(Index).Caption = Format(lstROM(indexRom).similitud(CmbB(Index).ListIndex), "00.00")
    'Si se ha producido un cambio...
    If lstROM(indexRom).index_newName <> CmbB(Index).ListIndex Then
        indexSnap_OLD = lstROM(indexRom).new_name(lstROM(indexRom).index_newName) - 1
        indexSnap_NEW = lstROM(indexRom).new_name(CmbB(Index).ListIndex) - 1
        
        'Liberamos la vieja referencia
        lstSNAP(indexSnap_OLD).NumReferencias = lstSNAP(indexSnap_OLD).NumReferencias - 1
        
        'incrementamos la referencia nueva
        lstSNAP(indexSnap_NEW).NumReferencias = lstSNAP(indexSnap_NEW).NumReferencias + 1
        
        'Asignamos el valor nuevo
        lstROM(indexRom).index_newName = CmbB(Index).ListIndex
        
        'Evaluamos la referencia antigua
        If lstSNAP(indexSnap_OLD).NumReferencias = 1 Then
            'Si luego de liberar la referenci antigua. Esta es referida por una única ROM. Entonces buscamos esa rom y desmarcamos la opción de repetido
            For i = 0 To UBound(lstROM)
                If lstROM(i).new_name(lstROM(i).index_newName) = indexSnap_OLD + 1 Then
                    lstROM(i).used = 1
                    lstROM(i).conError = 0
                    lstROM(i).NumError = sin_error
                    Exit For
                End If
            Next
        End If
        
        'Evaluamos la referencia nueva
        If lstSNAP(indexSnap_NEW).NumReferencias >= 2 Then
            'Si la nueva referencia se encuentra repetida, entonces buscamos donde están y las marcamos como repetidas
            For i = 0 To UBound(lstROM)
                If lstROM(i).new_name(lstROM(i).index_newName) = indexSnap_NEW + 1 Then
                    If CheckRepetida.Value = 0 Then lstROM(i).used = 0
                    lstROM(i).conError = 1
                    lstROM(i).NumError = SNAP_Repetida
                    'Exit For
                End If
            Next
        Else
            'Si la nueva SNAP seleccionada NO está repetida, entonces marcamos esta ROM como sin problema
            lstROM(indexRom).used = 1
            lstROM(indexRom).conError = 0
            lstROM(indexRom).NumError = sin_error
        End If
        
        Call llenarResultado(True)
    End If
End Sub

Private Sub FiltrarResultado()
'Función encargada de filtrar la lista de rom según sea el criterio a usar.
'Contamos el total de coincidencias que existen aplicando el filtro correspondiente
Dim TotalDatos As Long 'total de rom que cumplen con los requisitos del filtro
Dim e, i As Long
If Filtro = 0 Then 'Si no hay filro activo
    TotalDatos = UBound(lstROM) + 1
Else
    For i = 0 To UBound(lstROM)
        Select Case Filtro
            Case 1: If lstROM(i).conError Then TotalDatos = TotalDatos + 1 'Filtror solo con errores
            Case 2: If lstROM(i).NumError = sin_SNAP Then TotalDatos = TotalDatos + 1 'Filtrar missmatch
            Case 3: If lstROM(i).NumError = SNAP_Repetida Then TotalDatos = TotalDatos + 1 'Filtrar repetidas
            Case 4: If lstROM(i).used = 0 Then TotalDatos = TotalDatos + 1 'Filtrar sin usar
            Case 5: If lstROM(i).used = 1 Then TotalDatos = TotalDatos + 1 'Filtrar solo con usar
        End Select
    Next
End If
DatosFiltrados = TotalDatos
If TotalDatos = 0 Then Call LimpiarResultados: Exit Sub
ReDim lstIndex(TotalDatos - 1) 'Redimencionamos la lista intermedia


'lleamos la lstIndex con los indices a lstrom que cumplen el criterio de filtro
e = 0
For i = 0 To UBound(lstROM)
    Select Case Filtro
        Case 0
            lstIndex(e) = i 'guardamos el indice a lstrom que se corresponde al lstindex
            e = e + 1
        Case 1
            If lstROM(i).conError Then
                lstIndex(e) = i 'guardamos el indice a lstrom que se corresponde al lstindex
                e = e + 1
            End If
        Case 2
            If lstROM(i).NumError = sin_SNAP Then
                lstIndex(e) = i 'guardamos el indice a lstrom que se corresponde al lstindex
                e = e + 1
            End If
        Case 3
            If lstROM(i).NumError = SNAP_Repetida Then
                lstIndex(e) = i 'guardamos el indice a lstrom que se corresponde al lstindex
                e = e + 1
            End If
        Case 4
            If lstROM(i).used = 0 Then
                lstIndex(e) = i 'guardamos el indice a lstrom que se corresponde al lstindex
                e = e + 1
            End If
        Case 5
            If lstROM(i).used = 1 Then
                lstIndex(e) = i 'guardamos el indice a lstrom que se corresponde al lstindex
                e = e + 1
            End If
    End Select
Next

'Colocar aquí código para ordenar la lista
Call OrdenarResultado
Call llenarResultado
End Sub

Private Sub LblTitulo_DblClick(Index As Integer)
If Index > 0 Then
    Call BtnLstOrdenar_Click(Index)
End If
End Sub

Public Sub OrdenarResultado()
Dim i As Integer
Dim Max As Long
Dim Aux As Long

For i = 1 To 3
    LblTitulo(i).Caption = Replace(LblTitulo(i).Caption, "*", "")
Next

If Ordenar > 0 And DatosFiltrados > 0 Then
    If Ordenar = 1 Or Ordenar = 4 Then
        Call ModVariables.QSROM(lstIndex, LBound(lstIndex), UBound(lstIndex))
    ElseIf Ordenar = 2 Or Ordenar = 5 Then
        Call ModVariables.QSSNAP(lstIndex, LBound(lstIndex), UBound(lstIndex))
    ElseIf Ordenar = 3 Or Ordenar = 6 Then
        Call ModVariables.QSPor(lstIndex, LBound(lstIndex), UBound(lstIndex))
    End If
    
    If Ordenar > 3 Then 'Si es orden inverso, entonces damos vuelta la lista
        Max = UBound(lstIndex)
        For i = 0 To Max / 2
            Aux = lstIndex(i)
            lstIndex(i) = lstIndex(Max - i)
            lstIndex(Max - i) = Aux
        Next
    End If
    
    If Ordenar = 1 Then LblTitulo(1).Caption = LblTitulo(1).Caption & "*"
    If Ordenar = 2 Then LblTitulo(2).Caption = LblTitulo(2).Caption & "*"
    If Ordenar = 3 Then LblTitulo(3).Caption = LblTitulo(3).Caption & "*"
    If Ordenar = 4 Then LblTitulo(1).Caption = LblTitulo(1).Caption & "**"
    If Ordenar = 5 Then LblTitulo(2).Caption = LblTitulo(2).Caption & "**"
    If Ordenar = 6 Then LblTitulo(3).Caption = LblTitulo(3).Caption & "**"
    
End If
End Sub

Public Sub llenarResultado(Optional OmitirCMB As Boolean = False)
Dim i As Long
Dim e As Long
Dim indexRom As Long 'indice a la lista de rom del dato actualmente activo
Dim MaxRows As Integer 'Máximo de filas que podemos llegar a tener dentro del contenedor pctTblResultados
Dim TotalDatos As Long

movRoll = True

If Not OmitirCMB Then Call LimpiarResultados

MaxRows = CInt(PctTblResultados.Height / LblCheck(0).Height) - 1
TotalDatos = UBound(lstIndex) + 1

'Si el número de filas (rows) que se pueden tener dentro del contenedor es diferente del actual número de filas lo adaptamos
If MaxRows > LblCheck.Count Then
    'Si hay que agregar filas
    For i = LblCheck.Count To MaxRows - 1
        Load LblCheck(i)
        Set LblCheck(i).Container = PctTblResultados
        LblCheck(i).Left = LblCheck(0).Left
        LblCheck(i).Width = LblCheck(0).Width
        LblCheck(i).Top = (i + 1) * LblCheck(0).Height
        
        Load ChekUsar(i)
        Set ChekUsar(i).Container = PctTblResultados
        ChekUsar(i).Left = ChekUsar(0).Left
        ChekUsar(i).Width = ChekUsar(0).Width
        ChekUsar(i).Top = ((i + 1) * LblCheck(0).Height) + 30
        
        Load LblNombre(i)
        Set LblNombre(i).Container = PctTblResultados
        LblNombre(i).Left = LblNombre(0).Left
        LblNombre(i).Width = LblNombre(0).Width
        LblNombre(i).Top = (i + 1) * LblNombre(0).Height
        
        Load CmbB(i)
        Set CmbB(i).Container = PctTblResultados
        CmbB(i).Left = CmbB(0).Left
        CmbB(i).Width = CmbB(0).Width
        CmbB(i).Top = (i + 1) * LblCheck(0).Height
        
        Load LblP(i)
        Set LblP(i).Container = PctTblResultados
        LblP(i).Left = LblP(0).Left
        LblP(i).Width = LblP(0).Width
        LblP(i).Top = (i + 1) * LblP(0).Height
    Next
ElseIf MaxRows < LblCheck.Count Then 'Si hay que quitar filas
    For i = LblCheck.Count - 1 To MaxRows Step -1
        Unload LblCheck(i)
        Unload ChekUsar(i)
        Unload LblNombre(i)
        Unload CmbB(i)
        Unload LblP(i)
    Next
End If

'Comprovamos si el el total de filas que podemos mostras es menor que el total de datos y habilitamos el scroll vertical
If TotalDatos > MaxRows Then
    VScroll1.Visible = True
    Dim auxScrollMax As Long
    auxScrollMax = TotalDatos - MaxRows
        
    If auxScrollMax < 32000 Then
        FactorScroll = 1
    ElseIf FactorScroll < 320000 Then
        FactorScroll = 10
    ElseIf FactorScroll < 640000 Then
        FactorScroll = 20
    End If
    
    If VScroll1.Value > TotalDatos - MaxRows Then VScroll1.Value = 0
    VScroll1.Max = (TotalDatos - MaxRows) / FactorScroll
    VScroll1.Enabled = True
Else
    VScroll1.Visible = False
    VScroll1.Value = 0
End If

'llenamos la tabla con los valores de la matris
For i = 0 To MaxRows - 1
    If (i + VScroll1.Value) > UBound(lstIndex) Then Exit For 'Salida en caso que el último valor de la matris no sea el último valor de la tabla
    indexRom = lstIndex(i + (VScroll1.Value * FactorScroll)) 'Obtnemos el indice a la lista de rom que debe ser mostrada en esta fila
    
    ChekUsar(i).Value = lstROM(indexRom).used 'cargamos el valor del check
    'Control de colores
    If ChekUsar(i).Value = 1 Then
        LblNombre(i).BackColor = &H80000005
        'amarillo, marcada para usar, pero con SNAP repetida
        If lstROM(indexRom).conError And lstROM(indexRom).NumError = SNAP_Repetida Then LblNombre(i).BackColor = &H80FFFF
    Else
        If lstROM(indexRom).conError Then
            If lstROM(indexRom).NumError = sin_SNAP Then
                LblNombre(i).BackColor = &HFF& 'Rojo fuerte, ROM sin SNAP
            ElseIf lstROM(indexRom).NumError = SNAP_Repetida Then
                LblNombre(i).BackColor = &H8080FF 'Rojo debil, SNAP repetida al buscar
            Else
                LblNombre(i).BackColor = &H80FFFF    'amarillo debil, SNAP repetida por un cambio del usuario (repetida revisada)
            End If
        Else
            LblNombre(i).BackColor = &H80FF&         'Naranja. ROM sin error pero NO marcada para ser utilizada
        End If
    End If
    LblNombre(i).Caption = lstROM(indexRom).full_name
    
    If Not OmitirCMB Then
        CmbB(i).Clear
        If lstROM(indexRom).new_name(lstROM(indexRom).index_newName) > 0 Then
            CmbB(i).Text = lstSNAP(lstROM(indexRom).new_name(lstROM(indexRom).index_newName) - 1).full_name
            e = 0
            Do While lstROM(indexRom).new_name(e) > 0
                CmbB(i).AddItem lstSNAP(lstROM(indexRom).new_name(e) - 1).full_name
                e = e + 1
                If e = 10 Then Exit Do
            Loop
            LblP(i).Caption = Format(lstROM(indexRom).similitud(lstROM(indexRom).index_newName), "00.00")
        Else
            LblP(i).Caption = ""
        End If
    End If
    
    LblCheck(i).Visible = True
    ChekUsar(i).Visible = True
    LblNombre(i).Visible = True
    CmbB(i).Visible = True
    LblP(i).Visible = True
Next

movRoll = False
End Sub

Private Sub LimpiarResultados()
Dim i As Long
'Limpiar lista Resultado anterior
For i = 0 To LblCheck.Count - 1
    ChekUsar(i).Value = 0
    LblNombre(i).Caption = ""
    LblNombre(i).BackColor = &H80000005
    CmbB(i).Clear
    CmbB(i).Text = ""
    LblP(i).Caption = ""
Next
End Sub

Private Sub BorrarResultados()
Dim i As Long

movRoll = True

LblCheck(0).Visible = False
ChekUsar(0).Value = 0
ChekUsar(0).Visible = False
LblNombre(0).Caption = ""
LblNombre(0).BackColor = &H80000005
LblNombre(0).Visible = False
CmbB(0).Clear
CmbB(0).Text = ""
CmbB(0).Visible = False
LblP(0).Caption = ""
LblP(0).Visible = False
VScroll1.Value = 0
VScroll1.Enabled = False
'MsgBox LblCheck.Count
For i = LblCheck.Count - 1 To 1 Step -1
    Unload LblCheck(i)
    Unload ChekUsar(i)
    Unload LblNombre(i)
    Unload CmbB(i)
    Unload LblP(i)
Next

End Sub

'************************************************************************************
'Mover
'************************************************************************************

Private Sub BtnRenombrar_Click()
Dim auxIndex As Long
Dim accuRepetidas As Long
Dim ROMUsadas As Long
Dim ROMSinUsar As Long
Dim SNAPUsadas As Long
Dim SNAPSinUsar As Long
Dim i As Long

BtnRenombrar.Enabled = False
BtnCalcular.Enabled = False
'frameFiltros.Enabled = False

'Limpiamos el contador de referencias
For i = 0 To UBound(lstSNAP)
    lstSNAP(i).NumReferencias = 0
Next

'Contamos las referencias, pero solo de los archivos que si van a ser usados
For i = 0 To UBound(lstROM)
    If lstROM(i).used Then
        auxIndex = lstROM(i).new_name(lstROM(i).index_newName) - 1
        If auxIndex > -1 Then lstSNAP(auxIndex).NumReferencias = lstSNAP(auxIndex).NumReferencias + 1
        ROMUsadas = ROMUsadas + 1
    Else
        ROMSinUsar = ROMSinUsar + 1
    End If
Next

'Procesamos las SNAP repetidas
For i = 0 To UBound(lstSNAP)
    If lstSNAP(i).NumReferencias = 0 Then
        SNAPSinUsar = SNAPSinUsar + 1
    ElseIf lstSNAP(i).NumReferencias = 1 Then
        SNAPUsadas = SNAPUsadas + 1
    Else
        SNAPUsadas = SNAPUsadas + 1
        accuRepetidas = accuRepetidas + 1
    End If
Next

If sinROM Then
    Check4.Value = False
    Check4.Enabled = False
Else
    Check4.Enabled = True
End If

'llenamos los datos en la ventana de acción
If accuRepetidas = 0 Then BtnRenombrarSNAP.Enabled = True Else BtnRenombrarSNAP.Enabled = False
If accuRepetidas = 0 Then LblSNAPRepetidas.BackColor = &H80000005 Else LblSNAPRepetidas.BackColor = &H8080FF
LblSNAPRepetidas.Caption = accuRepetidas
lblROMusadas.Caption = ROMUsadas
LblROMSinUsar.Caption = ROMSinUsar
lblROMTotal.Caption = UBound(lstROM) + 1
lblSNAPusadas.Caption = SNAPUsadas
LblSNAPSinUsar.Caption = SNAPSinUsar
lblSNAPTotal.Caption = UBound(lstSNAP) + 1

PctBotones.Visible = True
'Call llenar_resultado
End Sub

Private Sub BtnExportarArchivos_Click()
Dim nFic
Dim i As Long
Dim numArchivo As Integer
Dim nombreArchivo As String

nombreArchivo = InputBox(lstTextos(17), lstTextos(18), "SNAP_")
If nombreArchivo = "" Then Exit Sub

i = 0
Do While Dir(App.Path & "\" & nombreArchivo & CStr(i) & ".txt", vbArchive) <> ""
    i = i + 1
Loop
numArchivo = i
    
nFic = FreeFile
Open App.Path & "\" & nombreArchivo & CStr(numArchivo) & ".txt" For Output As nFic
    For i = 0 To UBound(lstSNAP)
        Print #nFic, lstSNAP(i).name_SinExtension
    Next
Close nFic

MsgBox lstTextos(9) & vbCrLf & lstTextos(10) & vbCrLf & nombreArchivo & CStr(numArchivo) & ".txt", vbOKOnly, "Exportar"

End Sub

Private Sub BtnExportarArchivosNOusados_Click()
Dim nFic
Dim i As Long
Dim numArchivo As Integer
Dim nombreArchivo As String

nombreArchivo = InputBox(lstTextos(17), lstTextos(18), "SNAP_sinusar_")
If nombreArchivo = "" Then Exit Sub

i = 0
Do While Dir(App.Path & "\" & nombreArchivo & CStr(i) & ".txt", vbArchive) <> ""
    i = i + 1
Loop
numArchivo = i
    
nFic = FreeFile
Open App.Path & "\" & nombreArchivo & CStr(numArchivo) & ".txt" For Output As nFic
    For i = 0 To UBound(lstSNAP)
        If lstSNAP(i).NumReferencias = 0 Then Print #nFic, lstSNAP(i).name_SinExtension
    Next
Close nFic

MsgBox lstTextos(9) & vbCrLf & lstTextos(10) & vbCrLf & nombreArchivo & CStr(numArchivo) & ".txt", vbOKOnly, "Exportar"

End Sub

Private Sub BtnExportarArchivosUsados_Click()
Dim nFic
Dim i As Long
Dim numArchivo As Integer
Dim nombreArchivo As String

nombreArchivo = InputBox(lstTextos(17), lstTextos(18), "SNAP_usadas_")
If nombreArchivo = "" Then Exit Sub

i = 0
Do While Dir(App.Path & "\" & nombreArchivo & CStr(i) & ".txt", vbArchive) <> ""
    i = i + 1
Loop
numArchivo = i
    
nFic = FreeFile
Open App.Path & "\" & nombreArchivo & CStr(numArchivo) & ".txt" For Output As nFic
    For i = 0 To UBound(lstSNAP)
        If lstSNAP(i).NumReferencias > 0 Then Print #nFic, lstSNAP(i).name_SinExtension
    Next
Close nFic

MsgBox lstTextos(9) & vbCrLf & lstTextos(10) & vbCrLf & nombreArchivo & CStr(numArchivo) & ".txt", vbOKOnly, "Exportar"

End Sub

Private Sub BtnExportarNombres_Click()
Dim nFic
Dim i As Long
Dim numArchivo As Integer
Dim nombreArchivo As String

nombreArchivo = InputBox(lstTextos(17), lstTextos(18), "ROM_")
If nombreArchivo = "" Then Exit Sub

i = 0
Do While Dir(App.Path & "\" & nombreArchivo & CStr(i) & ".txt", vbArchive) <> ""
    i = i + 1
Loop
numArchivo = i
    
nFic = FreeFile
Open App.Path & "\" & nombreArchivo & CStr(numArchivo) & ".txt" For Output As nFic
    For i = 0 To UBound(lstROM)
        Print #nFic, lstROM(i).name_SinExtension
    Next
Close nFic

MsgBox lstTextos(9) & vbCrLf & lstTextos(10) & vbCrLf & nombreArchivo & CStr(numArchivo) & ".txt", vbOKOnly, "Exportar"

End Sub

Private Sub BtnExportarNombresNOUsados_Click()
Dim nFic
Dim i As Long
Dim numArchivo As Integer
Dim nombreArchivo As String

nombreArchivo = InputBox(lstTextos(17), lstTextos(18), "ROM_MISSMATCH_")
If nombreArchivo = "" Then Exit Sub

i = 0
Do While Dir(App.Path & "\" & nombreArchivo & CStr(i) & ".txt", vbArchive) <> ""
    i = i + 1
Loop
numArchivo = i
    
nFic = FreeFile
Open App.Path & "\" & nombreArchivo & CStr(numArchivo) & ".txt" For Output As nFic
    For i = 0 To UBound(lstROM)
        If lstROM(i).used = 0 Then Print #nFic, lstROM(i).name_SinExtension
    Next
Close nFic

MsgBox lstTextos(9) & vbCrLf & lstTextos(10) & vbCrLf & nombreArchivo & CStr(numArchivo) & ".txt", vbOKOnly, "Exportar"

End Sub

Private Sub BtnExportarNombresUsados_Click()
Dim nFic
Dim i As Long
Dim numArchivo As Integer
Dim nombreArchivo As String

nombreArchivo = InputBox(lstTextos(17), lstTextos(18), "ROM_MATCH_")
If nombreArchivo = "" Then Exit Sub

i = 0
Do While Dir(App.Path & "\" & nombreArchivo & CStr(i) & ".txt", vbArchive) <> ""
    i = i + 1
Loop
numArchivo = i
    
nFic = FreeFile
Open App.Path & "\" & nombreArchivo & CStr(numArchivo) & ".txt" For Output As nFic
    For i = 0 To UBound(lstROM)
        If lstROM(i).used Then Print #nFic, lstROM(i).name_SinExtension
    Next
Close nFic

MsgBox lstTextos(9) & vbCrLf & lstTextos(10) & vbCrLf & nombreArchivo & CStr(numArchivo) & ".txt", vbOKOnly, "Exportar"

End Sub

Private Sub BtnRenombrarSNAP_Click()
Call MoverSnap(True)
End Sub

Private Sub BtnMoverSNAP_Click()
Call MoverSnap(False)
End Sub

Private Sub MoverSnap(nRenombrar As Boolean)
Dim MaxRom As Long
Dim RutaDestino As String
Dim i As Long
Dim auxCopy As Boolean 'Flag que indica si debemos copiar y borrar (filecopy kill) los archivos o si podemos mover (name)

Dim FileInicial As String
Dim FileFinal As String

Dim nFic
Dim flagError As Boolean
Dim ErrorCount As Long

'On Error GoTo msgError

auxCopy = True
flagError = False
ErrorCount = 0

'borramos el log anterior
If Dir(App.Path & "\log.txt", vbArchive) <> "" Then Kill App.Path & "\log.txt"

'Abrimos el nuevo log en modo escritura
nFic = FreeFile
Open App.Path & "\log.txt" For Output As nFic

'Determianmos la ruta de destino
If nRenombrar Then 'Si vamos a renombrar
    RutaDestino = TxtPathSNAP
Else 'Si vamos a mover a un subdirectorio
    i = 0
    Do While Dir(TxtPathSNAP.Text & "machimax_" & i & "\", vbDirectory) <> ""
        i = i + 1
    Loop
    'creamos el nuevo subdirectorio
    RutaDestino = TxtPathSNAP.Text & "machimax_" & i & "\"
    MkDir (RutaDestino)
End If

Picture1.Visible = True 'cuadro de avance de renombrado
MaxRom = UBound(lstROM)
If nRenombrar And LblSNAPRepetidas.Caption <> "0" Then
    MsgBox lstTextos(11), vbOKOnly, "Error"
ElseIf nRenombrar Then 'RENOMBRAR
    'Guardamos el dato en el log
    'Write #nFic, "*********************"
    Print #nFic, "*********************"
    Print #nFic, "*****RENOMBRANDO*****"
    Print #nFic, "*********************"
    Print #nFic, " "
    
    For i = 0 To MaxRom
        If lstROM(i).new_name(lstROM(i).index_newName) > 0 And lstROM(i).used = 1 Then
            FileInicial = TxtPathSNAP.Text & lstSNAP(lstROM(i).new_name(lstROM(i).index_newName) - 1).full_name
            FileFinal = RutaDestino & lstROM(i).name_SinExtension & "." & lstSNAP(lstROM(i).new_name(lstROM(i).index_newName) - 1).Extension
            Name FileInicial As FileFinal
            'Agregamos al log la acción
            Print #nFic, FileInicial & "  -->  " & FileFinal
        End If
        LblRenombrando.Caption = Format(i / MaxRom * 100, "00.00")
        DoEvents
    Next
Else 'MOVER
    'Guardamos el dato en el log
    Print #nFic, "*********************"
    Print #nFic, "******Moviendo*******"
    Print #nFic, "*********************"
    Print #nFic, " "
    
    'Detectamos si podemos mover (más rápido) o si debemos copiar y pegar
    If LblSNAPRepetidas.Caption = "0" And Check5.Value = 0 Then 'Si no hay snap repetidas y si no se ha marcado la opción de dejar copia.
        auxCopy = False 'Opción de mover
    Else
        auxCopy = True 'Opción de copiar
    End If
    
    For i = 0 To MaxRom
        If lstROM(i).new_name(lstROM(i).index_newName) > 0 And lstROM(i).used = 1 Then
            FileInicial = TxtPathSNAP.Text & lstSNAP(lstROM(i).new_name(lstROM(i).index_newName) - 1).full_name
            FileFinal = RutaDestino & lstROM(i).name_SinExtension & "." & lstSNAP(lstROM(i).new_name(lstROM(i).index_newName) - 1).Extension
            If auxCopy Then FileCopy FileInicial, FileFinal
            If Not auxCopy Then Name FileInicial As FileFinal
            
            If Not flagError Then 'Si no hay error al momento de copiar los archivos
                'If Check5.Value = 0 Then Kill FileInicial
                lstSNAP(lstROM(i).new_name(lstROM(i).index_newName) - 1).movida = 1
                'Agregamos al log la acción
                Print #nFic, FileInicial & "  -->  " & FileFinal
                'Si hay que mover la ROM también
                If Check4.Value = 1 Then
                    FileInicial = TxtPathROM & lstROM(i).full_name
                    FileFinal = RutaDestino & lstROM(i).full_name
                    Name FileInicial As FileFinal
                    lstROM(i).movida = 1
                    'Agregamos al log la acción
                    Print #nFic, FileInicial & "  -->  " & FileFinal
                End If
            Else
                lstROM(i).conError = 1
                lstROM(i).NumError = Error_copia
            End If
            flagError = False
        End If
        If MaxRom > 0 Then
            LblRenombrando.Caption = Format(i / MaxRom * 100, "00.00")
            DoEvents
        End If
    Next
End If
Picture1.Visible = False

'Lista de SNAP que se borraron si correspondía
If Not nRenombrar And Check5.Value = 0 And auxCopy Then 'Si hay que mover las snap , pero solo si no hay que dejar copia
    Print #nFic, " " & vbCrLf & vbCrLf & "*******************************" & vbCrLf & "Lista de SNAP borradas" & vbCrLf & "*******************************" & vbCrLf
    
    For i = 0 To UBound(lstSNAP) - 1
        If lstSNAP(i).movida = 1 Then
            Kill (TxtPathSNAP & lstSNAP(i).full_name)
            Print #nFic, lstSNAP(i).full_name
        End If
    Next
End If

'Lista de rom que se borraron si correspondía
'If Not nRenombrar And Check4.Value = 1 And Check5.Value = 0 Then 'Si hay que mover las snap y también las rom, pero solo si no hay que dejar copia
'    Print #nFic, " " & vbCrLf & vbCrLf & "*******************************" & vbCrLf & "Lista de ROM borradas" & vbCrLf & "*******************************" & vbCrLf
'
'    For i = 0 To MaxRom
''        If lstROM(i).movida = 1 Then
'            Kill (TxtPathROM & lstROM(i).full_name)
'            Print #nFic, lstROM(i).full_name
'        End If
'    Next
'End If

'Lista de rom que no se movieron por que se encontró algún error
Print #nFic, " " & vbCrLf & vbCrLf & "*******************************" & vbCrLf & "Lista de ROM donde se detectó un error" & vbCrLf & "*******************************" & vbCrLf

For i = 0 To MaxRom
    If lstROM(i).conError = 1 Then
        If lstROM(i).NumError = Error_copia Then
            Print #nFic, ("ERROR copiado/renombrado :  " & lstROM(i).full_name)
        ElseIf lstROM(i).NumError = sin_SNAP Then
            Print #nFic, ("Sin SNAP :  " & lstROM(i).full_name)
        ElseIf lstROM(i).NumError = SNAP_Repetida Then
            Print #nFic, ("SNAP repetida :  " & lstROM(i).full_name)
        ElseIf lstROM(i).NumError = SNAP_Repetida Then
            Print #nFic, ("SNAP repetida por usuario :  " & lstROM(i).full_name)
        End If
    End If
Next


'Lista de rom que no fueron utilizadas
Print #nFic, " " & vbCrLf & vbCrLf & "*******************************" & vbCrLf & "Lista de ROM que no se movieron" & vbCrLf & "*******************************" & vbCrLf

For i = 0 To MaxRom
    If lstROM(i).used = 0 Then
        Print #nFic, lstROM(i).full_name
    End If
Next

Close nFic

If ErrorCount = 0 Then
    MsgBox lstTextos(12), vbOKOnly, lstTextos(13)
Else
    MsgBox lstTextos(14), vbOKOnly, lstTextos(13)
End If
PctBotones.Visible = False
BtnCalcular.Enabled = True
BtnRenombrar.Enabled = True

If MsgBox(lstTextos(15), vbYesNo, lstTextos(16)) = vbYes Then
    Call BtnCalcular_Click
End If

Exit Sub

msgError:
    Write #nFic, "ERROR: " & _
    Err.Number & "-" & Err.Description & vbCrLf & _
    "Archivo inicial: " & FileInicial & vbCrLf & _
    "Archivo final: " & FileFinal
    
    ErrorCount = ErrorCount + 1
    flagError = True
    Resume Next
End Sub

Private Sub BtnTerminar_Click()
PctBotones.Visible = False
BtnRenombrar.Enabled = True
BtnCalcular.Enabled = True
End Sub

Private Sub ChkEsXML_Click()
'If ChkEsXML.Value Then
'
'Else'
'
'End If
End Sub

Public Sub CmbDesde_Click()
Select Case CmbDesde.ListIndex
    Case 0 'Directorio
        TxtNodo.Enabled = False
        TxtPropiedad.Enabled = False
        TxtExtROM.Visible = True
        Label12.Enabled = False
        Label22.Enabled = False
        Label1.Caption = lstTextos(20)
        ChkExtension.Visible = False
    Case 1 'TXT
        TxtNodo.Enabled = False
        TxtPropiedad.Enabled = False
        TxtExtROM.Visible = False
        Label12.Enabled = False
        Label22.Enabled = False
        Label1.Caption = lstTextos(21)
        ChkExtension.Visible = True
    Case 2 'XML
        TxtNodo.Enabled = True
        TxtPropiedad.Enabled = True
        TxtExtROM.Visible = False
        Label12.Enabled = True
        Label22.Enabled = True
        Label1.Caption = lstTextos(21)
        ChkExtension.Visible = True
End Select
End Sub






