VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EtaComp"
   ClientHeight    =   6000
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10575
   Enabled         =   0   'False
   Icon            =   "FrmGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   5040
      Index           =   2
      Left            =   120
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   10455
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexSerial 
         Height          =   3855
         Left            =   0
         TabIndex        =   41
         Top             =   1080
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6800
         _Version        =   393216
         BackColorSel    =   -2147483643
         ForeColorSel    =   65535
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame FrameRS232 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   10455
         Begin VB.Frame FrameDifference 
            Caption         =   "Dernier écart calculé"
            Height          =   1095
            Left            =   3520
            TabIndex        =   39
            Top             =   0
            Width           =   3420
            Begin VB.TextBox TxtDifference 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008080&
               Height          =   675
               Left            =   120
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   360
               Width           =   3220
            End
         End
         Begin VB.Frame FrameWaitValue 
            Caption         =   "Prochaine course à régler"
            Height          =   1095
            Left            =   7040
            TabIndex        =   37
            Top             =   0
            Width           =   3420
            Begin VB.TextBox TxtWaitValue 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008080&
               Height          =   675
               Left            =   120
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   360
               Width           =   3220
            End
         End
         Begin VB.Frame FrameRS232Output 
            Caption         =   "Dernière valeur mesurée"
            Height          =   1095
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   3420
            Begin MSCommLib.MSComm MsComm 
               Left            =   1800
               Top             =   480
               _ExtentX        =   1005
               _ExtentY        =   1005
               _Version        =   393216
               DTREnable       =   -1  'True
               BaudRate        =   1200
               ParitySetting   =   2
               DataBits        =   7
            End
            Begin VB.Timer TimerRS232 
               Enabled         =   0   'False
               Interval        =   50
               Left            =   960
               Top             =   480
            End
            Begin VB.TextBox TxtRS232 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008080&
               Height          =   675
               Left            =   120
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   360
               Width           =   3220
            End
         End
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   5040
      Index           =   1
      Left            =   80
      TabIndex        =   15
      Top             =   720
      Width           =   10455
      Begin VB.Frame FrameCarac 
         Caption         =   "Déroulement"
         Height          =   2160
         Left            =   5880
         TabIndex        =   26
         Top             =   2880
         Width           =   4575
         Begin VB.CommandButton CmdModifyMeasures 
            Caption         =   "&Modifier ..."
            Height          =   320
            Left            =   3300
            TabIndex        =   11
            ToolTipText     =   "Modifier les valeurs de la série de mesures"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.ComboBox ComboSerialValues 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1800
            Width           =   3135
         End
         Begin MSComCtl2.UpDown UpDownMeasureCount 
            Height          =   280
            Left            =   4320
            TabIndex        =   9
            Top             =   1200
            Width           =   195
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   4
            Max             =   128
            Min             =   4
            Enabled         =   -1  'True
         End
         Begin VB.TextBox TxtMeasureCount 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   30
            Top             =   1206
            Width           =   4200
         End
         Begin MSComCtl2.UpDown UpDownSerialCount 
            Height          =   280
            Left            =   4320
            TabIndex        =   8
            Top             =   600
            Width           =   195
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   2
            Max             =   64
            Min             =   2
            Enabled         =   -1  'True
         End
         Begin VB.TextBox TxtSerialCount 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   28
            Top             =   612
            Width           =   4200
         End
         Begin VB.Label LblSerialValues 
            AutoSize        =   -1  'True
            Caption         =   "Valeurs de la série de mesures (mm):"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1548
            Width           =   2565
         End
         Begin VB.Label LblMeasures 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de mesures par série:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   954
            Width           =   2100
         End
         Begin VB.Label LblSerialCount 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de séries:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.Frame FrameComment 
         Caption         =   "Observations"
         Height          =   2880
         Left            =   5860
         TabIndex        =   25
         Top             =   0
         Width           =   4575
         Begin VB.TextBox TxtComment 
            Height          =   2480
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame FrameObject 
         Caption         =   "Comparateur"
         Height          =   2160
         Left            =   0
         TabIndex        =   22
         Top             =   2880
         Width           =   5775
         Begin VB.TextBox TxtRefObject 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   1206
            Width           =   5535
         End
         Begin VB.TextBox TxtSrcObject 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   1800
            Width           =   5535
         End
         Begin VB.TextBox TxtNumObject 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   612
            Width           =   5535
         End
         Begin VB.Label LblSrcObject 
            AutoSize        =   -1  'True
            Caption         =   "Fabriquant:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1548
            Width           =   795
         End
         Begin VB.Label LblRefObject 
            AutoSize        =   -1  'True
            Caption         =   "Référence du comparateur:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   954
            Width           =   1950
         End
         Begin VB.Label LblNumObject 
            AutoSize        =   -1  'True
            Caption         =   "N° du comparateur:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1380
         End
      End
      Begin VB.Frame FrameControl 
         Caption         =   "Conditions de contrôle"
         Height          =   2880
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   5775
         Begin VB.TextBox txtDetenteur 
            Height          =   285
            Left            =   3000
            TabIndex        =   70
            Top             =   1200
            Width           =   2655
         End
         Begin VB.TextBox TxtCelcius 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   1800
            Width           =   5535
         End
         Begin VB.TextBox TxtName 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   612
            Width           =   5535
         End
         Begin MSComCtl2.UpDown UpDownHumidity 
            Height          =   280
            Left            =   5460
            TabIndex        =   3
            Top             =   2400
            Width           =   195
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   50
            OrigLeft        =   5460
            OrigTop         =   2400
            OrigRight       =   5655
            OrigBottom      =   2680
            Max             =   100
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin VB.TextBox TxtWatter 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   2400
            Width           =   5340
         End
         Begin MSComCtl2.DTPicker DTPickDate 
            Height          =   285
            Left            =   120
            TabIndex        =   1
            Top             =   1206
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   503
            _Version        =   393216
            Format          =   56688641
            CurrentDate     =   37510
         End
         Begin VB.Label lblDetenteur 
            Caption         =   "Détenteur :"
            Height          =   255
            Left            =   3000
            TabIndex        =   69
            Top             =   960
            Width           =   975
         End
         Begin VB.Label LblName 
            AutoSize        =   -1  'True
            Caption         =   "Opérateur:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   750
         End
         Begin VB.Label LblCelcius 
            AutoSize        =   -1  'True
            Caption         =   "Température:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1548
            Width           =   945
         End
         Begin VB.Label LblWatter 
            AutoSize        =   -1  'True
            Caption         =   "Taux d'humidité:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   2142
            Width           =   1155
         End
         Begin VB.Label LblDate 
            AutoSize        =   -1  'True
            Caption         =   "Date du contrôle:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   954
            Width           =   1230
         End
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   5040
      Index           =   3
      Left            =   80
      TabIndex        =   42
      Top             =   720
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton CmdValidFidelity 
         Caption         =   "Valider"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   9120
         TabIndex        =   68
         ToolTipText     =   "Valider les mesures"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame FrameListOfMeasures 
         Caption         =   "Caractéristiques"
         Height          =   2520
         Left            =   0
         TabIndex        =   46
         Top             =   2520
         Width           =   10455
         Begin VB.TextBox TxtFDMeasureCount 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   55
            Top             =   600
            Width           =   4260
         End
         Begin VB.CommandButton CmdRemoveOne 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   52
            ToolTipText     =   "Retirer la mesure sélectionnée"
            Top             =   1920
            Width           =   855
         End
         Begin VB.CommandButton CmdAddOne 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   51
            ToolTipText     =   "Ajouter la mesure sélectionnée"
            Top             =   1320
            Width           =   855
         End
         Begin VB.ListBox ListEnabledMeasure 
            Height          =   1815
            Left            =   5880
            TabIndex        =   50
            Top             =   600
            Width           =   4440
         End
         Begin VB.ListBox ListDispo 
            Height          =   1230
            Left            =   120
            TabIndex        =   48
            Top             =   1200
            Width           =   4440
         End
         Begin MSComCtl2.UpDown UpDownFDMeasureCount 
            Height          =   285
            Left            =   4360
            TabIndex        =   56
            Top             =   600
            Width           =   195
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   5
            Max             =   128
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin VB.Label LblFDMeasureCount 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de mesures à effectuer pour chaque point:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   3600
         End
         Begin VB.Label LblMeasureEnabled 
            AutoSize        =   -1  'True
            Caption         =   "Points de mesures pris en compte:"
            Height          =   195
            Left            =   5880
            TabIndex        =   49
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label LblMeasurePoint 
            AutoSize        =   -1  'True
            Caption         =   "Points de mesures disponibles:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   2085
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexFidelity 
         Height          =   1215
         Left            =   0
         TabIndex        =   45
         Top             =   1200
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   2143
         _Version        =   393216
         AllowUserResizing=   3
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   5040
      Index           =   4
      Left            =   80
      TabIndex        =   53
      Top             =   720
      Visible         =   0   'False
      Width           =   10455
      Begin VB.Frame FrameLevels 
         Caption         =   "Classe de précision"
         Height          =   5040
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Width           =   5775
         Begin VB.ComboBox ComboLevel 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   560
            Width           =   5535
         End
         Begin VB.ComboBox ComboCapacity 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   1140
            Width           =   5535
         End
         Begin VB.Frame FrameClassResult 
            Caption         =   "Résultat"
            Height          =   2175
            Left            =   120
            TabIndex        =   58
            Top             =   2800
            Width           =   5535
            Begin VB.CommandButton CmdModifyResult 
               Caption         =   "Modifier ..."
               Height          =   375
               Left            =   4200
               TabIndex        =   67
               ToolTipText     =   "Modifier le résultat proposé par EtaComp"
               Top             =   1700
               Width           =   1215
            End
            Begin VB.TextBox TxtResult 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   120
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   59
               Top             =   1680
               Width           =   3975
            End
            Begin MSComctlLib.ImageList ImageListListView 
               Left            =   4200
               Top             =   360
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   3
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmGen.frx":0442
                     Key             =   "CHECKED"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmGen.frx":0894
                     Key             =   "WARNING"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmGen.frx":0CE6
                     Key             =   "FAILURE"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.ListView ListViewResult 
               Height          =   1395
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   2461
               View            =   3
               Arrange         =   1
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               Icons           =   "ImageListListView"
               SmallIcons      =   "ImageListListView"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Caractéristiques"
                  Object.Width           =   9243
               EndProperty
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexLimit 
            Height          =   1035
            Left            =   120
            TabIndex        =   61
            Top             =   1760
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1826
            _Version        =   393216
            Rows            =   3
            Cols            =   5
            Enabled         =   0   'False
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.Label LblEchelon 
            AutoSize        =   -1  'True
            Caption         =   "Valeur de l'échelon (mm):"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   320
            Width           =   1770
         End
         Begin VB.Label LblCOurse 
            AutoSize        =   -1  'True
            Caption         =   "Course (mm):"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   920
            Width           =   915
         End
         Begin VB.Label LblLimites 
            AutoSize        =   -1  'True
            Caption         =   "Limites d'usure (µm) d'après NFE 11-200:"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   1520
            Width           =   2880
         End
      End
   End
   Begin VB.Frame Frame 
      Height          =   5040
      Index           =   5
      Left            =   80
      TabIndex        =   43
      Top             =   720
      Visible         =   0   'False
      Width           =   10455
      Begin MSComDlg.CommonDialog Cmdlg 
         Left            =   2640
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin MSChart20Lib.MSChart MSChart 
         Height          =   5040
         Left            =   0
         OleObjectBlob   =   "FrmGen.frx":1138
         TabIndex        =   44
         Top             =   0
         Width           =   10455
      End
   End
   Begin MSComctlLib.ImageList ImageList16 
      Left            =   6840
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGen.frx":2F72
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGen.frx":3084
            Key             =   "OPTIONS"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGen.frx":34D6
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGen.frx":3630
            Key             =   "PREVIEW"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGen.frx":378A
            Key             =   "QUIT"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5745
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "RS232 KO"
            TextSave        =   "RS232 KO"
            Object.ToolTipText     =   "Etat de la liaison RS232 avec le comparateur"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Arrêté"
            TextSave        =   "Arrêté"
            Object.ToolTipText     =   "Indique si un contrôle est en cours"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Object.ToolTipText     =   "Sens de la course du comparateur à tester"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11615
            Object.ToolTipText     =   "Zone info"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "13:25"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "Nouveau contrôle"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "PRINT"
            Object.ToolTipText     =   "Imprimer"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "PREVIEW"
            Object.ToolTipText     =   "Aperçu avant impression"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPTIONS"
            Object.ToolTipText     =   "Options"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "QUIT"
            Object.ToolTipText     =   "Quitter"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   5340
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   9419
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Caractéristiques"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Séries de mesures"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ecarts de fidélité"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Finalisation"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Courbe d'étalonnage"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuFile_New 
         Caption         =   "Nouveau contrôle ..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile_Simul 
         Caption         =   "&Simuler un contrôle"
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTBefore_mnuFile_Print 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Preview 
         Caption         =   "Aperçu avant impression ..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "&Imprimer ..."
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Exporter au format Excel"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTBefore_mnuFile_Quit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Outils"
      Begin VB.Menu mnuTools_Options 
         Caption         =   "&Options ..."
      End
   End
   Begin VB.Menu mnuInterr 
      Caption         =   "&?"
      Begin VB.Menu mnuInterr_Help 
         Caption         =   "Sommaire de l'aide ..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuInterr_Tips 
         Caption         =   "Astuces du jour ..."
      End
      Begin VB.Menu mnuTBefore_mnuInterr_About 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInterr_About 
         Caption         =   "A propos de ..."
      End
   End
End
Attribute VB_Name = "FrmGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************
'FEUILLE PRINCIPALE DE L'APPLICATION
'********************************************************************************

'********************************************************************************
'Constantes
'********************************************************************************

'Constantes des bouttons de la barre d'outils
Private Const BUTTON_NEW = "NEW"
Private Const BUTTON_OPTIONS = "OPTIONS"
Private Const BUTTON_PREVIEW = "PREVIEW"
Private Const BUTTON_PRINT = "PRINT"
Private Const BUTTON_QUIT = "QUIT"

'Constantes des icones du ListView du panneau final
Private Const ICO_CHECKED = "CHECKED"
Private Const ICO_FAILURE = "FAILURE"
Private Const ICO_WARNING = "WARNING"

'Constantes de CommonDialog
Private Const CMDLG_CANCEL_ERROR = 32755

'Constantes des onglets du TabStrip
Private Const TAB_DESCRIPTION = 1
Private Const TAB_SERIAL = 2
Private Const TAB_DIFFERENCE = 3
Private Const TAB_FINAL = 4
Private Const TAB_GRAPH = 5

'Constantes de la barre d'état
Private Const PANEL_RS232 = 1
Private Const PANEL_STATUS = 2
Private Const PANEL_DIRECTION = 3
Private Const PANEL_INFO = 4

'********************************************************************************
'Données membres
'********************************************************************************
Private WithEvents objControl As ControlObject      'Objet "Contrôle d'un comparateur"
Attribute objControl.VB_VarHelpID = -1

Private m_IndxTab As Integer
Private m_objConfig As Configuration

'Sauvegarde
Private m_LastMSHFlexBackColor As Long
Private m_LastMSHFlexCol As Integer
Private m_LastMSHFlexRow As Integer

'utiliser pour une simulation de mesure
Private indexSimu As Integer

Private Sub CmdAddOne_Click()
'Clic sur le bouton "Ajouter un élément de série de mesures"

Dim bEnabled As Boolean
Dim strItem As String

'Sécurité
If ListDispo.ListIndex = -1 Then Exit Sub

'Transférer
strItem = ListDispo.List(ListDispo.ListIndex)
InsertValueOnListBox strItem, ListEnabledMeasure

'Ajouter
objControl.FidelityPoints.Add CCur(Left(strItem, Len(strItem) - 1)), (Right(strItem, 1) <> "D")

'Supprimer
ListDispo.RemoveItem ListDispo.ListIndex

'Enabled
If ListDispo.ListCount > 0 Then
    bEnabled = True
    ListDispo.ListIndex = 0
Else
    bEnabled = False
End If
CmdAddOne.Enabled = bEnabled
CmdRemoveOne.Enabled = True

'MSHFlex
Refresh_GUIInitFidelityTest

End Sub

Private Sub CmdModifyMeasures_Click()
'Clic sur le bouton "Modifier" (Valeurs de la série de mesures)

'Travail sur la feuille
With FrmSValues
    .SetReferenceToObject objControl.SerialValues
    .Show vbModal
    
    'Si modification
    If .Tag = TAG_CLIC_ON_VALID Then Refresh_GUISValuesCount
    
End With

'Décharger la feuille
Unload FrmSValues

End Sub

Private Sub CmdModifyResult_Click()
'Clic sur le bouton "Modifier" (Le résultat d'un contrôle)

Dim iClassBeforeShow

Dim iQBColor As Integer
Dim strText As String

'Affecter
iClassBeforeShow = objControl.RealClass

'Travail sur la feuille
With FrmEnd
    .SetReferenceToObject objControl
    .Show vbModal
End With

'Mise à jour si nécessaire
Select Case objControl.RealClass
    Case iClassBeforeShow
        Exit Sub
    Case 0
        iQBColor = QBCOLOR_GREEN
        strText = "CLASSE 0"
    Case 1
        iQBColor = QBCOLOR_YELLOW
        strText = "CLASSE 1"
    Case 2
        iQBColor = QBCOLOR_RED
        strText = "A REBUTER"
End Select

'Affichage
With TxtResult
    .ForeColor = QBColor(iQBColor)
    .Text = strText
End With

End Sub

Private Sub CmdRemoveOne_Click()
'Clic sur le bouton "Retirer une valeur de la liste de relevé d'écart de fidélité"

Dim strItem As String
Dim bEnabled As Boolean

'Sécurité
If ListEnabledMeasure.ListIndex = -1 Then Exit Sub

'Affecter
strItem = ListEnabledMeasure.List(ListEnabledMeasure.ListIndex)

'Supprimer de la ListBox des sélections
ListEnabledMeasure.RemoveItem ListEnabledMeasure.ListIndex

'Ajouter dans la ListBox des disponibilités
InsertValueOnListBox strItem, ListDispo

'Supprimer dans l'objet
objControl.FidelityPoints.RemoveItemByString strItem

'Enabled
If ListEnabledMeasure.ListCount > 1 Then
    bEnabled = True
Else
    bEnabled = False
End If
If ListEnabledMeasure.ListCount > 0 Then ListEnabledMeasure.ListIndex = 0
CmdRemoveOne.Enabled = bEnabled
CmdAddOne.Enabled = True

'MSHFlex
Refresh_GUIInitFidelityTest

End Sub

Private Sub InsertValueOnListBox(ByVal strValue As String, objListBox As ListBox)
'Insère une valeur de relevé d'écart de mesure dans un ListBox avec tri

Dim i As Integer
Dim currentItem As String
Dim cCurrentMeasure As Currency
Dim cCurrentListBoxMeasure As Currency
Dim bIsUp As Boolean

'Sécurités
strValue = Trim(UCase(strValue))
If strValue = "" Then Exit Sub
If objListBox Is Nothing Then Exit Sub

'Optimisation
If objListBox.ListCount = 0 Then

    'Ajouter
    objListBox.AddItem strValue
    
    'Sortir
    Exit Sub
    
End If

'Affecter les variables
bIsUp = (Right(strValue, 1) <> "D")
cCurrentMeasure = CCur(Left(strValue, Len(strValue) - 1))

'Rechercher la position
For i = objListBox.ListCount - 1 To 0 Step -1

    'Stocker
    currentItem = objListBox.List(i)
    cCurrentListBoxMeasure = CCur(Left(currentItem, Len(currentItem) - 1))
    
    'Conditions de sorties
    If cCurrentMeasure = cCurrentListBoxMeasure Then
        If Right(currentItem, 1) = "D" Then i = i - 1
        Exit For
    ElseIf cCurrentMeasure > cCurrentListBoxMeasure Then
        Exit For
    End If
    
Next

'Insérer
If i = objListBox.ListCount - 1 Then
    objListBox.AddItem strValue
Else
    objListBox.AddItem strValue, i + 1
End If

'Sélection auto
objListBox.ListIndex = i + 1

End Sub

Private Sub CmdValidFidelity_Click()
'Clic sur le bouton "Valider" des mesures d'écart de fidélité

'Valider les mesures
objControl.SetControlIsEnd

End Sub

Private Sub ComboCapacity_Click()
'Clic dans la zone de liste des courses

Dim currentClassification As Classification
Dim i As Integer

'Sauvegarder
m_objConfig.Result_LastCapacity = objControl.Levels(ComboLevel.ListIndex + 1).Capacitys(ComboCapacity.ListIndex + 1).Value

'Charger les classifications correspondantes
With MSHFlexLimit
    
    'Initialisation graphique
    .Visible = False
    
    'Ecraser les valeurs
    .Row = 1
    For Each currentClassification In objControl.Levels(ComboLevel.ListIndex + 1).Capacitys(ComboCapacity.ListIndex + 1).Classifications
        For i = 1 To 4
        
            'Affecter la colonne
            .Col = i
            
            'Afficher la valeur
            Select Case i
                Case 1
                    .Text = Format(currentClassification.TotalExactness)
                Case 2
                    'voir si affichage mesure locale nécessaire (selon la course)
                    If currentClassification.LocalExactness <= 100 Then
                        .Text = Format(currentClassification.LocalExactness)
                        .Row = 2
                        .Text = objControl.ExactnessLocalError
                        .Row = 1
                    Else
                        .Text = "Néant"
                        .Row = 2
                        .Text = "Néant"
                        .Row = 1
                        
                    End If
                Case 3
                    .Text = CHAR_PLUS_MINUS + Format(currentClassification.Fidelity)
                Case 4
                    .Text = CHAR_PLUS_MINUS + Format(currentClassification.Hysteresis)
            End Select
        Next
        
        'Passer à la ligne suivante
        .Row = .Row + 1
        
    Next
    
    'Rafraichir
    .Visible = True
    Refresh_GUIFrameFinal
    
    'Restauration graphique
    .Visible = True
    
End With

'Libérer
Set currentClassification = Nothing

End Sub

Private Sub ComboLevel_Click()
'Clic dans la zone de liste des échelons

Dim currentCapacity As Capacity

'Sauvegarder
m_objConfig.Result_LastLevel = objControl.Levels(ComboLevel.ListIndex + 1).Value

'Charger les courses correspondantes
With ComboCapacity
    
    'Initialisation graphique
    .Visible = False
    
    'Initialisation
    .Clear
    
    'Charger les courses
    For Each currentCapacity In objControl.Levels(ComboLevel.ListIndex + 1).Capacitys
        .AddItem "Jusqu'à" + Str(currentCapacity.Value)
    Next
    
    'Sélection auto
    If .ListCount > 0 Then .ListIndex = 0
    
    'Rafraichir
    Refresh_GUIFrameFinal
    
    'Restauration graphique
    .Visible = True
    
End With

'Libérer
Set currentCapacity = Nothing

End Sub

Private Sub Command1_Click()
mnuFile_Preview_Click
End Sub

Private Sub Form_Activate()
'Activation de la feuille

Static bAlreadyActivated As Boolean

Dim txt As String
Dim i As Integer

'Si première activation
If Not bAlreadyActivated Then
    
    'Enregistrer
    bAlreadyActivated = True
    
    'Les erreur sont gérées
    On Error GoTo GestErr
    
    ' Avec dll port
    Dim sComm As String
    sComm = "COM" & m_objConfig.Comm_CommPort & ":" & m_objConfig.Comm_Bauds & "," & m_objConfig.Comm_Parity & "," & _
             Format(m_objConfig.Comm_DataBits) & "," & Format(m_objConfig.Comm_StopBits)
   sComm = "COM1:4800,N,7,2" 'si utilise port.dll
   Ouvrir_Port sComm
    
    'Travail sur le port série
    With MsComm
        .Settings = Format(m_objConfig.Comm_Bauds) + "," + m_objConfig.Comm_Parity + "," + Format(m_objConfig.Comm_DataBits) + "," + _
        Format(m_objConfig.Comm_StopBits)
   '     .PortOpen = True
    End With
    
    'Afficher dans la barre d'état
    StatusBar.Panels(PANEL_RS232).Text = "RS232 OK"
    
    'Astuces du jours
    If m_objConfig.Tips_ShowTips Then mnuInterr_Tips_Click
    
End If

'Sortir normalement
Exit Sub

'Gestion des erreurs
GestErr:

'Traiter
TimerRS232.Enabled = False

'Afficher dans la barre d'état
StatusBar.Panels(PANEL_RS232).Text = "RS232 KO"

'Afficher un message
txt = "Interception d'une erreur imprévue lors de l'ouverture du port série" + vbCrLf + _
      "pour connecter le comparateur." + vbCrLf + vbCrLf + _
      "Code de l'erreur:" + Str(Err) + vbCrLf + _
      "Description: " + Err.Description
i = MsgBox(txt, vbCritical, "Erreur")

End Sub

Private Sub Form_Load()
'Chargement de la feuille

Dim currentButton As MSComctlLib.Button

Dim valKey As String
Dim varValue As Variant

'Travail sur la barre d'outils
With Toolbar
    .ImageList = ImageList16
    
    'Enumérer les boutons et affecter l'image
    For Each currentButton In .Buttons
        
        valKey = currentButton.Key
        If valKey <> "" Then currentButton.Image = valKey
    Next
    
End With

'Initialisation des données membres
m_IndxTab = 1
Set m_objConfig = New Configuration

Set objControl = New ControlObject
objControl.PrepareForANewControl

'Initialisation de l'interface graphique
InitDocument

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Demande de déchargement de la feuille

'Si la demande vient de l'utilisateur
If UnloadMode = vbFormControlMenu Then
    
    'Annuler si demande non-confirmée
    If Not UserWantsToQuit Then Cancel = 1
    
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Déchargement de la feuille

Dim i As Integer

'Fermer le port série si ouvert
'If MSComm.PortOpen Then MSComm.PortOpen = False
Fermer_Port

'Décharger toutes les feuilles
For i = 0 To Forms.Count - 1
    If Forms(i).Name <> "FrmGen" Then Unload Forms(i)
Next

'Enregistrer la configuration
m_objConfig.SaveConfig

'Libérer
Set m_objConfig = Nothing
Set objControl = Nothing
Set FrmGen = Nothing
End
End Sub

Private Function UserWantsToQuit() As Boolean
'Renvoie vrai si l'utilisateur veut quitter

Dim txt As String

'Si l'utilisateur doit confirmer
If m_objConfig.Gen_ConfirmExit Then

    'Affecter le message
    txt = "Confirmer l'arrêt de l'application ?"
    
    If objControl.ControlIsStart Then txt = "Un contrôle est en cours de réalisation." + vbCrLf + vbCrLf + txt
    
    'Demander confirmation
    If MsgBox(txt, vbQuestion + vbYesNo, "Quitter " + App.EXEName) = vbYes Then _
    UserWantsToQuit = True

Else
    
    'Confirmer par défaut
    UserWantsToQuit = True
    
End If

End Function

Private Sub mnuExport_Click()
Dim currentClassification As Classification
Set currentClassification = objControl.Levels(ComboLevel.ListIndex + 1).Capacitys(ComboCapacity.ListIndex + 1).Classifications(1)
objControl.Export_Excel m_objConfig, currentClassification
Set currentClassification = Nothing
End Sub

Private Sub mnuFile_New_Click()
'Clic sur le menu "Fichier->Nouveau contrôle ..."

Dim txt As String

'Demander confirmation si un conrôle est en cours
If objControl.ControlIsStart Then
    
    txt = "Les mesures actuellement enregistrées vont être perdues." + vbCrLf + vbCrLf + _
          "Confirmer la création d'un nouveau contrôle ?"
    If MsgBox(txt, vbQuestion + vbYesNo, "Nouveau contrôle") = vbNo Then Exit Sub
    
End If

'Nouveau contrôle
objControl.PrepareForANewControl
mnuFile_Preview.Enabled = False
mnuFile_Print.Enabled = False
mnuExport.Enabled = False
Toolbar.Buttons(BUTTON_PREVIEW).Enabled = False
Toolbar.Buttons(BUTTON_PRINT).Enabled = False

'Mise à jour de l'affichage
InitDocument

End Sub

Private Sub mnuFile_Preview_Click()
'Clic sur le menu "Fichier->Aperçu avant impression"

'Travail sur la feuille
With FrmPreview
    .SetRefToObjects objControl, m_objConfig
    .Show vbModal
End With

End Sub

Private Sub mnuFile_Print_Click()
'Clic sur le menu "Fichier->Imprimer"

Dim i As Integer
Dim txt As String

'Les erreurs sont gérées
On Error GoTo GestErr

'Travail sur la boîte de dialogue
With Cmdlg

    'Paramétrer
    .Flags = cdlPDHidePrintToFile + cdlPDReturnDC
    
    'Afficher
    .ShowPrinter
    
End With

'Rendre la main
DoEvents

'Imprimer
For i = 1 To Cmdlg.Copies
    objControl.PrintOn Printer, m_objConfig
Next

'Sortir normalement
Exit Sub

'Gestion des erreurs
GestErr:

'Si erreur <>"Annulation"
If Err <> CMDLG_CANCEL_ERROR Then
    
    'Afficher un message
    txt = "Interception d'une erreur imprevue lors de l'impression." + vbCrLf + vbCrLf + _
          "Code de l'erreur:" + Str(Err) + vbCrLf + _
          "Description: " + Err.Description
    i = MsgBox(txt, vbCritical, "Erreur")
    
End If

End Sub

Private Sub mnuFile_Quit_Click()
'Clic sur le menu "Fichier->Quitter"

'Quitter si demande confirmée
If UserWantsToQuit Then Unload Me

End Sub

Private Sub mnuFile_Simul_Click()
'Clic sur le menu "Simuler un contrôle"

Dim i As Integer
Dim txt As String

With objControl

    'Sécurité
    If Not .SerialValuesAreOkToStart Then
    
        'Afficher un message
        txt = "La simulation ne peut pas commencer car les valeurs de la" + vbCrLf + _
              "série de mesures ne sont pas cohérentes ou la série ne" + vbCrLf + _
              "commence pas par ""0""."
        i = MsgBox(txt, vbExclamation, "Simulation impossible")
    
        'Sortir
        Exit Sub
        
    End If
   
    'Préparer un nouveau contrôle
    mnuFile_New_Click
            
    'Initialiser le MSHFlexGrid
    MSHFlexSerial.Col = 1
    MSHFlexSerial.Row = 1
    m_LastMSHFlexCol = 1
    m_LastMSHFlexRow = 1
    m_LastMSHFlexBackColor = MSHFlexSerial.BackColor
    
    'TabStrip
    Set TabStrip.SelectedItem = TabStrip.Tabs(TAB_SERIAL)
    DoEvents
    
    'Simuler
    StatusBar.Panels(PANEL_INFO) = "Simulation en cours de réalisation ..."
    StatusBar.Panels(PANEL_STATUS) = "En cours"
    .CreateASimulation
    
End With

End Sub

Private Sub mnuInterr_About_Click()
'Clic sur le menu "?->A propos de ..."

'Afficher
FrmAbout.Show vbModal

End Sub

Private Sub mnuInterr_Help_Click()
'Clic sur le bouton "?->Sommaire de l'aide"

Dim i As Integer
Dim txt As String

'Tester la présence de la page Web
If Dir(App.Path + APP_WEB_PAGE) = "" Then

    'Afficher un message
    txt = "L'aide en ligne n'est pas disponible car au moins un" + vbCrLf + _
          "des fichiers d'aide est introuvable ou manquant."
    i = MsgBox(txt, vbCritical, "Aide en ligne")
    
    'Sortir
    Exit Sub
    
End If

'Afficher la feuille
FrmHelp.Show vbModeless

End Sub

Private Sub mnuInterr_Tips_Click()
'Clic sur le menu "?->Astuces du jour"

Dim i As Integer
Dim txt As String

'Si astuces disponibles
If FrmTips.IsOkToShow(m_objConfig) Then
    
    'Afficher la feuille
    FrmTips.Show vbModal
    
Else

    'Afficher un message
    txt = "A la suite d'une erreur imprévue, les astuces ne sont pas disponibles."
    i = MsgBox(txt, vbCritical, "Astuces du jour")
    
    'Décharger la feuille
    Unload FrmTips
    
End If

End Sub

Private Sub mnuTools_Options_Click()
'Clic sur le menu "Outils->Options ..."

'Travail sur la feuille
With FrmOptions
    .SetReferenceToConfig m_objConfig
    .Show vbModal
End With

End Sub

Private Sub MSHFlexFidelity_RowColChange()
'Modification de la colonne ou de la ligne active du MSHFlexFidelity

Dim iCol As Integer
Dim iRow As Integer
Dim txt As String

With MSHFlexFidelity

    iCol = .Col
    iRow = .Row
    
    'Si l'évènement survient à la suite d'une action volontaire
    If (iRow <> objControl.CurrentFidelityIndexOfPoint) Or (iCol <> objControl.CurrentFidelityIndexOfMeasure) Then
    
        'Si la cellule n'est pas sélectionnable
        If iCol > objControl.FidelityPoints.MeasuresCount Then
        
            'Afficher un message
            Select Case iCol
            
                'Case de la moyenne
                Case objControl.FidelityPoints.MeasuresCount + 1
                        
                    txt = "La moyenne n'est pas sélectionnable par l'utilisateur." + vbCrLf + vbCrLf + _
                          "Elle est calculée automatiquement par le programme."

                'Case de la fidélité
                Case objControl.FidelityPoints.MeasuresCount + 1
                
                    txt = "L'écart de fidélité n'est pas sélectionnable par l'utilisateur." + vbCrLf + vbCrLf + _
                          "Il est calculé automatiquement par le programme."
                    
            End Select
            iRow = MsgBox(txt, vbExclamation, "Sélection annulée")
            
            'Sélectionner la bonne cellule
            MSHFlexFidelity.Row = objControl.CurrentFidelityIndexOfPoint
            MSHFlexFidelity.Col = objControl.CurrentFidelityIndexOfMeasure
            
            'Sortir
            Exit Sub
            
        End If
                
        'Restaurer les couleurs de l'ancienne cellule
        With MSHFlexFidelity
            .Row = objControl.CurrentFidelityIndexOfPoint
            .Col = objControl.CurrentFidelityIndexOfMeasure
            .CellForeColor = QBColor(QBCOLOR_BLACK)
            .CellBackColor = .BackColor
        End With
            
        With objControl
            
            'Déclarer le mode de correction actif
            .CorrectionModeOfFidelityTest = _
            ((iRow <> objControl.CurrentFidelityIndexOfPoint) Or (iCol <> objControl.CurrentFidelityIndexOfMeasure))
            
            'Sélectionner la nouvelle cellule
            .CurrentFidelityIndexOfPoint = iRow
            .CurrentFidelityIndexOfMeasure = iCol
            
        End With
        
        With MSHFlexFidelity
        
            'Mettre en forme la cellule
            .Row = objControl.CurrentFidelityIndexOfPoint
            .Col = objControl.CurrentFidelityIndexOfMeasure
            .CellForeColor = QBColor(QBCOLOR_YELLOW)
            .CellBackColor = QBColor(QBCOLOR_BLUE)
            
        End With
        
    End If

End With

End Sub

Private Sub MSHFlexSerial_LeaveCell()
'Une autre cellule va devenir la cellule active

'Sécurité - On sort si le test n'est pas commencé
If Not objControl.ControlIsStart Then Exit Sub

'Travail sur le MSHFlex
With MSHFlexSerial
        
    'Sauvegarder l'ancienne position
    m_LastMSHFlexCol = .Col
    m_LastMSHFlexRow = .Row
    
    'Rétablir les couleurs
    .CellForeColor = .ForeColor
    .CellBackColor = m_LastMSHFlexBackColor
    
End With

End Sub

Private Sub MSHFlexSerial_RowColChange()
'Modification de la propriété Col ou Row

Dim txt As String

Dim iCol As Integer
Dim iNormalRow As Integer
Dim iRow As Integer

'Sécurité - On sort si le test n'est pas commencé
If Not objControl.ControlIsStart Then Exit Sub

With MSHFlexSerial
        
    iCol = .Col
    iRow = .Row
    iNormalRow = IIf(objControl.NormalSequenceIsUp, objControl.NormalIndexOfSequence, _
    objControl.NormalIndexOfSequence + 1 + objControl.SequencesCount)
    
    'Si l'évènement survient à la suite d'une action volontaire
    If (iCol <> objControl.NormalIndexOfSerialValue) Or (iRow <> iNormalRow) Then
        
        'Si la cellule n'est pas sélectionnable (Sélection sur la ligne des moyennes)
        If (iRow = objControl.SequencesCount + 1) Or (iRow = (objControl.SequencesCount + 1) * 2) Then
            
            'Sélectionner l'ancienne cellule
            .Col = objControl.CurrentIndexOfSerialValue
            .Row = iNormalRow
            
            'Afficher un message
            txt = "Le moyenne n'est pas sélectionnable par l'utilisateur." + vbCrLf + vbCrLf + _
                  "Elle est calculée automatiquement par le programme."
            iRow = MsgBox(txt, vbExclamation, "Sélection annulée")
            
        Else
            
            'La cellule est sélectionnable - Passer en mode modification
            objControl.IsOnCorrectionMode = True
            objControl.CurrentIndexOfSerialValue = iCol
            
            'Sélectionner le sens de course
            If iRow < objControl.SequencesCount + 1 Then
                objControl.CurrentSequenceIsUp = True
                objControl.CurrentIndexOfSequence = iRow
            Else
                objControl.CurrentSequenceIsUp = False
                objControl.CurrentIndexOfSequence = iRow - objControl.SequencesCount - 1
            End If
            
        End If
        
    Else
    
        'Retour sur la cellule origine
        objControl.IsOnCorrectionMode = False
        
    End If
    
    'Afficher le sens de la course
    StatusBar.Panels(PANEL_DIRECTION) = IIf(objControl.CurrentSequenceIsUp, "Montante", "Descendante")
    
    'Mémoriser l'ancienne couleur de fond
    m_LastMSHFlexBackColor = .CellBackColor
    
    'Colorier la nouvelle sélection
    .CellBackColor = QBColor(QBCOLOR_BLUE)                  'Fond bleu
    .CellForeColor = QBColor(QBCOLOR_YELLOW)                'Texte jaune
    
End With

End Sub

Private Sub objControl_AfterFidelityPointSave()
'Sauvegarde du point de fidélité avec calcul du nouveau pas

'Positionner la cellule
With MSHFlexFidelity
    .Row = objControl.CurrentFidelityIndexOfPoint
    .Col = objControl.CurrentFidelityIndexOfMeasure
    .CellForeColor = QBColor(QBCOLOR_YELLOW)
    .CellBackColor = QBColor(QBCOLOR_BLUE)
End With

'Prochaine course à afficher
TxtWaitValue = Format(objControl.FidelityPoints(objControl.CurrentFidelityIndexOfPoint).SerialValue, "0.00") + " mm"

End Sub

Private Sub objControl_AfterFidelityPointSaveWithoutStep()
'Sauvegarde du point de fidélité sans calcul du nouveau pas

Dim currentFidelityPoint As FidelityPoint

Dim iPoint As Integer
Dim iMeasure As Integer
Dim iMicrons As Integer
Dim cAverage As Currency

'Affecter
iPoint = objControl.CurrentFidelityIndexOfPoint
iMeasure = objControl.CurrentFidelityIndexOfMeasure
Set currentFidelityPoint = objControl.FidelityPoints(iPoint)

'Afficher
With MSHFlexFidelity

    'Différence mesurée
    .Row = iPoint
    .Col = iMeasure
    .CellForeColor = QBColor(QBCOLOR_BLACK)
    .CellBackColor = .BackColor
    iMicrons = objControl.GetFidelityMicronsDifference(iPoint, iMeasure)
    .Text = IIf(iMicrons > 0, "+" + Format(iMicrons), Format(iMicrons))
    
    'Panneau de mesure
    TxtDifference = IIf(iMicrons > 0, "+" + Format(iMicrons), Format(iMicrons)) + " " + CHAR_MICRON + "m"
    TxtRS232 = Format(currentFidelityPoint.Measures(iMeasure).Value, "0.000") + " mm"
    
    'Moyenne
    cAverage = currentFidelityPoint.GetAVerage
    .Col = objControl.FidelityPoints.MeasuresCount + 1
    .CellForeColor = QBColor(QBCOLOR_WHITE)
    .CellBackColor = QBColor(QBCOLOR_GREY)
    .Text = IIf(cAverage > 0, "+" + Format(cAverage, "0.00"), Format(cAverage, "0.00"))
     
    'Erreur de fidélité
    .Col = .Col + 1
    .CellForeColor = QBColor(QBCOLOR_WHITE)
    .CellBackColor = QBColor(QBCOLOR_GREY)
    .Text = Format(currentFidelityPoint.FidelityError, "0.00") + " " + CHAR_MICRON + "m"
    
End With

'Libérer
Set currentFidelityPoint = Nothing

End Sub

Private Sub objControl_ControlIsEnd()
'Le contrôle est terminé

'StatusBar
With StatusBar
    .Panels(PANEL_DIRECTION).Text = ""
    .Panels(PANEL_STATUS).Text = "Arrêté"
    .Panels(PANEL_INFO).Text = "Le contrôle est terminé."
End With

'Interface
mnuFile_Preview.Enabled = True
mnuFile_Print.Enabled = True
mnuExport.Enabled = True
Toolbar.Buttons(BUTTON_PREVIEW).Enabled = True
Toolbar.Buttons(BUTTON_PRINT).Enabled = True
MSHFlexFidelity.Enabled = False

'Préparer le panneau de finalisation
InitFrameFinal

'Affectation des valeurs par défaut
ComboLevel.Text = m_objConfig.Result_LastLevel
ComboCapacity.Text = "Jusqu'à" + Str(m_objConfig.Result_LastCapacity)

'Basculer le TabStrip
Set TabStrip.SelectedItem = TabStrip.Tabs(TAB_FINAL)

'Si mode modification
If CmdValidFidelity.Visible Then
    CmdValidFidelity.Visible = False
    MSHFlexFidelity.Width = 10440
End If

End Sub

Private Sub objControl_CorrectionModeChange()
'Le mode "Correction" du contrôle a changé

Dim txt As String

'Message
If objControl.IsOnCorrectionMode Then
    txt = "Correction d'une mesure en cours de réalisation ..."
Else
    txt = "Séries de mesures en cours de réalisation ..."
End If

'StatusBar
 StatusBar.Panels(PANEL_INFO).Text = txt
 
End Sub

Private Sub objControl_CorrectionModeOfFidelityTestChange()
'Modification du mode de correction pour le test des écarts de fidélité

Dim strTmp As String

'Zone info
If objControl.CorrectionModeOfFidelityTest Then
    strTmp = "Correction d'une mesure en cours de réalisation ..."
Else
    strTmp = "Relevé des écarts de fidélité en cours de réalisation ..."
End If

'StatusBar
StatusBar.Panels(PANEL_INFO).Text = strTmp

End Sub

Private Sub objControl_FidelityPointsTestStart()
'Début du relevé d'écarts de fidélité

'Enabled
GUI_SetEnabledOfFrameListOfMeasures False
MSHFlexSerial.Enabled = False
MSHFlexFidelity.SetFocus

'StatusBar
StatusBar.Panels(PANEL_INFO) = "Relevé des écarts de fidélité en cours de réalisation ..."

End Sub

Private Sub objControl_FidelityPointsTestTerminated()
'Fin de la phase de relevé d'écart de fidélité

Dim i As Integer
Dim txt As String

'StatusBar
With StatusBar
    .Panels(PANEL_DIRECTION).Text = ""
    .Panels(PANEL_INFO).Text = "Relevé des écarts de fidélité terminé."
End With

'Afficher un message
txt = "La phase de relevé des écarts de fidélité est terminée." + vbCrLf + vbCrLf + _
      "Valider les mesures et afficher le panneau de finalisation ?"

If MsgBox(txt, vbQuestion + vbYesNo, "Relevé des écarts de fidélité terminé") = vbYes Then
    
    'Le contrôle est terminé
    objControl.SetControlIsEnd
    
Else
    
    'Redimensionnment
    MSHFlexFidelity.Width = 9000
    CmdValidFidelity.Visible = True
    
End If

End Sub

Private Sub objControl_LeavePosition()
'Une mesure vient d'être effectuée - Elle est enregistrée
'La position en cours va être modifiée

Dim currentSequences As Sequences

Dim sDifference As Currency
Dim sMeasure As Currency

Dim i As Integer
Dim iAverage As Currency
Dim IndxOfMaxAverageCol As Integer

'Travail sur l'objet de contrôle
With objControl
    
    'Affecter la collection de séquences correspondante
    If .CurrentSequenceIsUp Then
        Set currentSequences = .UpSequences
    Else
        Set currentSequences = .DownSequences
    End If
        
    'Stocker la valeur mesurée
    sMeasure = currentSequences(.CurrentIndexOfSequence).Measures(.CurrentIndexOfSerialValue).Value
    
    'Afficher la valeur mesurée
    TxtRS232 = Format(sMeasure, "0.000") + " mm"
    
    'Stocker la différence
    sDifference = .GetMicronsDifference(.CurrentIndexOfSerialValue, .CurrentIndexOfSequence, .CurrentSequenceIsUp)
    
End With

'Travail sur le MSHFlexGrid
With MSHFlexSerial
    
    'Initialisation graphique
    .Visible = False
    
    'Afficher l'écart calculé dans le MSHFlex
    .CellForeColor = .ForeColor
    .CellBackColor = m_LastMSHFlexBackColor
    .Text = IIf(sDifference < 0, Format(sDifference), "+" + Format(sDifference))
    
    'Mettre à jour la zone de texte d'écart
    TxtDifference = IIf(sDifference < 0, Format(sDifference), "+" + Format(sDifference)) + " " + CHAR_MICRON + "m"
    
    'Pointer sur la ligne des moyennes
    .Row = IIf(objControl.CurrentSequenceIsUp, objControl.SequencesCount, (objControl.SequencesCount * 2) + 1) + 1
    
    'Afficher la valeur moyenne
    .CellBackColor = QBColor(QBCOLOR_GREY)
    iAverage = objControl.GetAVerage(objControl.CurrentIndexOfSerialValue, objControl.CurrentSequenceIsUp)
    .Text = IIf(iAverage < 0, Format(iAverage), "+" + Format(iAverage))
    
    'Afficher dans le MSChart
    MSChart.Column = IIf(objControl.CurrentSequenceIsUp, 1, 2)
    MSChart.Row = objControl.CurrentIndexOfSerialValue
    MSChart.Data = iAverage
    
    'Modifier la couleur d'affichage des cellules contenant les moyennes
    IndxOfMaxAverageCol = currentSequences.GetIndexOfMaxSerialValueAverage
    For i = 1 To .Cols - 2
        .Col = i
        .CellForeColor = IIf(IndxOfMaxAverageCol <> i, QBColor(QBCOLOR_WHITE), QBColor(QBCOLOR_RED))
    Next
    
    'Restauration graphique
    .Visible = True
    
End With

'Libérer
Set currentSequences = Nothing

End Sub

Private Sub objControl_PositionChange()
'Une mesure vient d'être effectuée
'La position en cours vient d'être modifiée

Dim iSeq As Integer

'Travail sur le MSHFlexGrid
With MSHFlexSerial

    'Initialisation graphique
    .Visible = False
    
    'Pointer sur la nouvelle cellule
    .Col = objControl.CurrentIndexOfSerialValue
    iSeq = objControl.CurrentIndexOfSequence
    .Row = IIf(objControl.CurrentSequenceIsUp, iSeq, iSeq + 1 + objControl.SequencesCount)
    
    'Colorier la nouvelle sélection
    .CellBackColor = QBColor(QBCOLOR_BLUE)                  'Fond bleu
    .CellForeColor = QBColor(QBCOLOR_YELLOW)                'Texte jaune

    'Restauration graphique
    .Visible = True
    
End With

'StatusBar
StatusBar.Panels(PANEL_DIRECTION).Text = IIf(objControl.CurrentSequenceIsUp, "Montante", "Descendante")

'Afficher la prochaine valeur à atteindre
TxtWaitValue = Format(objControl.SerialValues(objControl.CurrentIndexOfSerialValue), "0.00") + " mm"

End Sub

Private Sub CreateDefaultFidelityPoints()
'Créer les 2 points de relevés d'écart de fidélité par défaut

Dim currentFidelityPoints As FidelityPoints

Dim i As Integer
Dim strFD1 As String
Dim strFD2 As String

Dim valUp As String, valDown As String
Dim valeurUp As Currency, valeurDown As Currency
'Initialisation du relévé d'écart de fidélité
Set currentFidelityPoints = objControl.FidelityPoints
currentFidelityPoints.Clear
ListEnabledMeasure.Clear
'modification suite à évolution de la norme
valUp = objControl.SerialValues(objControl.UpSequences.GetIndexOfMaxSerialValueAverage)
valeurUp = objControl.UpSequences.GetAVerage(objControl.UpSequences.GetIndexOfMaxSerialValueAverage)
valDown = objControl.SerialValues(objControl.DownSequences.GetIndexOfMaxSerialValueAverage)
valeurDown = objControl.DownSequences.GetAVerage(objControl.DownSequences.GetIndexOfMaxSerialValueAverage)

With currentFidelityPoints
     If Abs(valeurUp) > Abs(valeurDown) Then
        .Add valUp, True
     Else
        .Add valDown, False
    End If
    .MeasuresCount = UpDownFDMeasureCount.Value
End With
'With currentFidelityPoints
'     If valUp > valDown Then
'        .Add valUp, True
'     Else
'        .Add valDown, False
'    End If
'    .MeasuresCount = UpDownFDMeasureCount.Value
'End With
Refresh_GUIInitFidelityTest

'Mise à jour du ListBox de sélection
strFD1 = Format(currentFidelityPoints(1).SerialValue, "0.00") + " " + IIf(currentFidelityPoints(1).IsUpDirection, "M", "D")
'strFD2 = Format(currentFidelityPoints(2).SerialValue, "0.00") + " " + IIf(currentFidelityPoints(2).IsUpDirection, "M", "D")

ListEnabledMeasure.AddItem strFD1
ListEnabledMeasure.AddItem strFD2
ListEnabledMeasure.ListIndex = 0

'Suppression du ListBox de disponibilité
For i = ListDispo.ListCount - 1 To 0 Step -1
    If ListDispo.List(i) = strFD1 Or ListDispo.List(i) = strFD2 Then ListDispo.RemoveItem i
Next

'Libérer
Set currentFidelityPoints = Nothing

End Sub

Private Sub objControl_SerialValueTestTerminated()
'La phase des séries de mesures est terminée

Dim i As Integer
Dim txt As String

'Affcetr le message
txt = "La phase des séries de mesures est terminée."

'StatusBar
With StatusBar
    .Panels(PANEL_DIRECTION).Text = ""
    .Panels(PANEL_INFO) = txt
End With

'Relevé d'écart de fidélité
CreateDefaultFidelityPoints

'Afficher un message
i = MsgBox(txt, vbInformation, "Séries de mesures terminée")

End Sub

Private Sub ClearTextBoxOfFrameRS232()
'Vider les zones de texte de mesures

TxtRS232 = ""
TxtDifference = ""
TxtWaitValue = ""

End Sub

Private Sub TabStrip_Click()
'Clic sur un onglet de TabStrip

Dim bEnabledRS232 As Boolean
Dim txt As String

Dim i As Integer
Dim iSelectedIndex As Integer

'Affectation
iSelectedIndex = TabStrip.SelectedItem.Index

With objControl

    'Action selon le Tab sélectionné
    Select Case iSelectedIndex
    
        'Optimisation
        Case m_IndxTab
        
            Exit Sub
            
        Case TAB_SERIAL
            
            'Interdire le basculement sur TAB_SERIAL si relevé d'écart est commencé
            If .FidelityTestIsStart Then
            
                'Afficher un message
                txt = "Le panneau des séries de mesures est verrouillé lorsque la phase" + vbCrLf + _
                      "de relevé des écarts de fidélité est en cours de réalisation."
                i = MsgBox(txt, vbExclamation, "Action incorrecte")
                
                'Basculer
                Set TabStrip.SelectedItem = TabStrip.Tabs(m_IndxTab)
                
                'Sortir
                Exit Sub
                
            End If
                    
        'Interdire le basculement sur TAB_DIFFERENCE si Séries non terminées
        Case TAB_DIFFERENCE
        
            If Not .SequencesAreEnd Then
            
                'Afficher un message
                txt = "Le relevé des écarts de fidélité ne peut être effectué" + vbCrLf + _
                      "que lorsque les séries de mesures sont terminées."
                i = MsgBox(txt, vbExclamation, "Action incorrecte")
                
                'Basculer
                Set TabStrip.SelectedItem = TabStrip.Tabs(m_IndxTab)
                
                'Sortir
                Exit Sub
            
            Else
            
                'Si le relevé n'est pas commencé
                If Not .FidelityTestIsStart Then
                    
                    'Vider les zones de texte de mesures RS232
                    ClearTextBoxOfFrameRS232
                    
                    'Pas de sens de course
                    StatusBar.Panels(PANEL_DIRECTION) = ""
                    
                End If
            
            End If
            
        'Interdire le basculement sur TAB_FINAL si le contrôle n'est pas terminé
        Case TAB_FINAL
             If Not .ControlIsEnd Then
                
                'Afficher un message
                txt = "Le panneau de finalisation ne peut être affiché" + vbCrLf + _
                      "que lorsqu'un contrôle est terminé."
                i = MsgBox(txt, vbExclamation, "Action incorrecte")
                
                'Basculer
                Set TabStrip.SelectedItem = TabStrip.Tabs(m_IndxTab)
                
                'Sortir
                Exit Sub
    
            End If
        
        'Interdire le basculement sur TAB_GRAPH si Séries non terminées
        Case TAB_GRAPH
        
            If Not .SequencesAreEnd Then
            
                'Afficher un message
                txt = "La courbe d'étalonnage ne peut être affichée que" + vbCrLf + _
                      "lorsque les séries de mesures sont terminées."
                i = MsgBox(txt, vbExclamation, "Action incorrecte")
                
                'Basculer
                Set TabStrip.SelectedItem = TabStrip.Tabs(m_IndxTab)
                
                'Sortir
                Exit Sub
            
            End If
            
    End Select

End With

'Masquer l'ancien frame
Frame(m_IndxTab).Visible = False
m_IndxTab = TabStrip.SelectedItem.Index

'Action selon le Tab sélectionné
Select Case m_IndxTab

    Case TAB_DESCRIPTION
        
        'Frame de commentaire
        Set FrameComment.Container = Frame(TAB_DESCRIPTION)
        FrameComment.Height = 2880
        TxtComment.Height = 2480
        
    Case TAB_DIFFERENCE

        'RS232
        If Not objControl.ControlIsEnd Then bEnabledRS232 = True

        'Frame de mesurage
        Set FrameRS232.Container = Frame(TAB_DIFFERENCE)

    Case TAB_SERIAL

        'RS232
        If Not objControl.ControlIsEnd Then bEnabledRS232 = True

        'Frame de mesurage
        Set FrameRS232.Container = Frame(TAB_SERIAL)
        
    Case TAB_FINAL
    
        'Frame de commentaire
        Set FrameComment.Container = Frame(TAB_FINAL)
        FrameComment.Height = 5040
        TxtComment.Height = 4720
        
End Select

'Vider le tampon du port série
'If bEnabledRS232 Then
'    DoEvents
'    Do While MSComm.Input <> ""
'    Loop
'End If

'Activer la lecture périodique selon paramétrage
TimerRS232.Enabled = bEnabledRS232

'Afficher le nouveau frame
Frame(m_IndxTab).Visible = True

End Sub
Private Function simulation1(s As Integer) As String
Dim valeurs As Variant
valeurs = Array(0, -0.9, -2.1, -2.9995, -4.0925, -4.821, -5.994, -7.0895, -7.7965, -8.9965, -9.9975, _
          -9.997, -8.9955, -7.7975, -7.089, -5.994, -4.803, -4.0945, -2.998, -2.0965, -0.9005, 0.001)
simulation1 = valeurs(s)
End Function

Private Function simulation2(s As Integer) As String
Dim valeurs As Variant
valeurs = Array(-4.798, -4.801, -4.801, -4.801, -4.802)
simulation2 = valeurs(s)
End Function

Private Sub TimerRS232_Timer()
'Minuterie de lecture du port série écoulée

Dim strInput As String

Dim iChar As Integer
Dim iLen As Integer

Dim cTest As Currency

Dim bModify As Boolean

'Lecture de la valeur
strInput = Trim(Reception) 'si utilise port.dll
 
'procédure de simulation valeurs dans 'simulation'
'If Not (m_IndxTab = TAB_DIFFERENCE) Then
'    strInput = (simulation1(indexSimu))
'    indexSimu = indexSimu + 1
'    If indexSimu > 21 Then indexSimu = 0
'Else
'    If indexSimu > 5 Then indexSimu = 0
'    strInput = (simulation2(indexSimu))
'    indexSimu = indexSimu + 1
'End If

'strInput = Trim(MsComm.Input)
iLen = Len(strInput)
'Optimisation - On sort si tampon vide
If strInput = "" Then Exit Sub

'Retirer les caractères indésirables
bModify = True
Do While (iLen > 0) And (bModify = True)
    
    'Travail à gauche
    iChar = Asc(Left(strInput, 1))
    If (iChar = 10) Or (iChar = 13) Then
        strInput = Right(strInput, iLen - 1)
        iLen = iLen - 1
    Else
        bModify = False
    End If
    
    'Sécurité
    If iLen = 0 Then Exit Sub
    
    'Travail à droite
    iChar = Asc(Right(strInput, 1))
    If (iChar = 10) Or (iChar = 13) Then
        strInput = Left(strInput, iLen - 1)
        iLen = iLen - 1
        bModify = True
    End If
    
Loop

'Ici, la chaîne est valide - On tente une conversion
On Error Resume Next
cTest = CCur(strInput)
If Err > 0 Then Exit Sub

'Travail sur l'objet de contrôle
With objControl

    'Si phase de relevé d'écart de fidélité
    If m_IndxTab = TAB_DIFFERENCE Then
    
       objControl.SaveFidelityPointMeasure CCur(strInput) * -1 '/ 1000
        'Debug.Print objControl.FidelityPoints(2).FidelityError
    Else
    
        'Si le contrôle n'est pas verrouillé
        If Not .ControlIsStart Then
        
            'Si le contrôle peut commencer
            If .SerialValuesAreOkToStart Then
            
                'Initialiser la matrice
                .InitializeData
                
                'Initialiser FrameCarac
                GUI_SetEnabledOfFrameCarac False
                
                'Verrouiller le contrôle
                .ControlIsStart = True
                StatusBar.Panels(PANEL_STATUS) = "En cours"
                
                'Initialiser le MSHFlexGrid
                MSHFlexSerial.Col = 1
                MSHFlexSerial.Row = 1
                MSHFlexSerial_RowColChange
                m_LastMSHFlexCol = 1
                m_LastMSHFlexRow = 1
                
                'StatusBar
                StatusBar.Panels(PANEL_INFO).Text = "Séries de mesures en cours de réalisation ..."
                
            Else
            
                'Le contrôle ne peut pas commencer
                
                'Afficher le premier panel du Tabstrip
                Set TabStrip.SelectedItem = TabStrip.Tabs(TAB_DESCRIPTION)
                
                'Le contrôle n'est pas prêt
                strInput = "Le contrôle du comparateur ne peut pas commencer car les valeurs" + vbCrLf + _
                           "de la série de mesures ne sont pas cohérentes ou la série ne" + vbCrLf + _
                           "commence pas par ""0""."
                iLen = MsgBox(strInput, vbExclamation, "Début de contrôle impossible")
                
                'Afficher la feuille de gestion
                CmdModifyMeasures_Click
                
                'Sortir
                Exit Sub
                
            End If
            
        End If
        
        'Sauvegarder la mesure
        .SaveMeasure CCur(strInput) * -1 '/ 1000
        
        
        'Prendre en compte pour le relevé d'écart de fidélité si série terminée
        If .SequencesAreEnd Then CreateDefaultFidelityPoints
        
    End If
    
End With

'Vider le tampon
'Do While MSComm.Input <> ""
'Loop

'Emettre un son
If m_objConfig.Control_BeepAfterMeasure Then Beep

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'Clic sur un boutton de la barre d'outils

'Action selon le boutton
Select Case Button.Key
    Case BUTTON_NEW
        mnuFile_New_Click
    Case BUTTON_PREVIEW
        mnuFile_Preview_Click
    Case BUTTON_PRINT
        mnuFile_Print_Click
    Case BUTTON_OPTIONS
        mnuTools_Options_Click
    Case BUTTON_QUIT
        mnuFile_Quit_Click
End Select

End Sub

Private Sub Refresh_GUISValuesCount()
'Affichier le nombre de valeurs de la série de mesures

Dim currentSerialValue As SerialValue

With objControl

    'Nombre de mesures dans la série
    UpDownMeasureCount.Value = .SerialValues.Count

    'Nombre de mesures dans la série
    UpDownMeasureCount.Value = .SerialValues.Count
    
    'Valeurs de la série de mesures
    ComboSerialValues.Clear
    For Each currentSerialValue In .SerialValues
        ComboSerialValues.AddItem Format(currentSerialValue.Value, "0.00")
    Next
    Set currentSerialValue = Nothing
    
    'Sélection auto
    If ComboSerialValues.ListCount > 0 Then ComboSerialValues.ListIndex = 0

End With

End Sub

Private Sub InitDocument()
'Initialisation du document en fonction des valeurs de l'objet de contrôle

Dim i As Integer

'TabStrip
Set TabStrip.SelectedItem = TabStrip.Tabs(TAB_DESCRIPTION)

With objControl
    
    'Valeurs de la série de mesures
    Refresh_GUISValuesCount
    
    'Nombre de séries (Entraine un CreateMSHFlexSerial)
    UpDownSerialCount.Value = .SequencesCount
    
    'Date du contrôle
    DTPickDate.Value = .DateOfControl
    
    'Zones de texte
    TxtName = m_objConfig.Last_OperatorName
    TxtCelcius = ""
    TxtNumObject = .EquipmentNumber
    TxtRefObject = .EquipmentReference
    TxtSrcObject = m_objConfig.Last_EquipmentManufacturer
    TxtComment = .ControlComment
    
    'Zone de mesure
    ClearTextBoxOfFrameRS232
    
    'Frame de déroulement
    GUI_SetEnabledOfFrameCarac True
    
    'UpDown
    UpDownHumidity.Value = .Humidity
    
    'Ecart de fidélité
    ListEnabledMeasure.Clear
    ListDispo.Clear
    RefreshGUIListBoxOfFidelity ListDispo
    GUI_SetEnabledOfFrameListOfMeasures True
    MSHFlexSerial.Enabled = True
    MSHFlexFidelity.Enabled = True
    
End With

'StatusBar
With StatusBar
    .Panels(PANEL_DIRECTION).Text = ""
    .Panels(PANEL_STATUS) = "Arrêté"
    .Panels(PANEL_INFO).Text = ""
End With

End Sub

Private Sub RefreshGUIListBoxOfFidelity(objListBox As ListBox)
'Rafraichir la liste des valeurs de mesure disponibles pour les écarts de fidélité

Dim i As Integer
Dim j As Integer

'Sécurité
If objListBox Is Nothing Then Exit Sub

'Initialiser
objListBox.Clear

'Remplissage
With objControl
    For i = 1 To .SerialValues.Count
        For j = 1 To 2
            objListBox.AddItem Format(.SerialValues(i), "0.00") + IIf(j = 1, " M", " D")
        Next
    Next
End With

'Sélection auto
If objListBox.ListCount > 0 Then objListBox.ListIndex = 0

End Sub

Private Sub TxtCelcius_LostFocus()
'La zone de texte contenant la température perd le focus

Dim cValue As Currency
Dim strValue As String

Dim i As Integer
Dim txt As String

'Optimisation
strValue = Trim(TxtCelcius)
If strValue = "" Then Exit Sub

'Gestion pas à pas des erreurs
On Error Resume Next

'Mise en forme
i = InStr(1, strValue, ".")
If i > 0 Then
    If i = 1 Then
        strValue = "0" + strValue
        i = i + 1
    End If
    strValue = Left(strValue, i - 1) + "," + Right(strValue, Len(strValue) - i)
End If

'Tentative de conversion
cValue = CCur(strValue)

'Si erreur
If Err > 0 Then
    
    'Ajuster l'affichage
    If m_IndxTab <> TAB_DESCRIPTION Then Set TabStrip.SelectedItem = TabStrip.Tabs(TAB_DESCRIPTION)
    
    'Afficher un message
    txt = "La valeur de température """ + strValue + """ n'est pas une température valide."
    i = MsgBox(txt, vbExclamation, "Température incorrecte")
    
    With TxtCelcius
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
Else

    'Température OK
    If TxtCelcius <> strValue Then TxtCelcius = strValue
    objControl.Temperature = cValue
    
End If

End Sub

Private Sub TxtComment_Change()
'Modification du commentaire

'Affecter
objControl.ControlComment = TxtComment.Text

End Sub

Private Sub txtDetenteur_LostFocus()
txtDetenteur.Text = UCase(txtDetenteur.Text)
objControl.Detenteur = txtDetenteur.Text
End Sub

Private Sub TxtMeasureCount_GotFocus()
'Le contrôle prend le focus

'Transmettre au UpDown associé
UpDownMeasureCount.SetFocus

End Sub

Private Sub TxtName_Change()
'Modification du nom de l'opérateur

'Affecter
objControl.OperatorName = Trim(TxtName)
m_objConfig.Last_OperatorName = objControl.OperatorName

End Sub

Private Sub TxtNumObject_Change()
'Modification du numéro du comparateur

'Affecter
objControl.EquipmentNumber = Trim(TxtNumObject)

End Sub

Private Sub TxtRefObject_Change()
'Modification de la référence du comparateur

'Affecter
objControl.EquipmentReference = Trim(TxtRefObject)

End Sub

Private Sub TxtSerialCount_GotFocus()
'Le contrôle prend le focus

'Transmettre au UpDown associé
UpDownSerialCount.SetFocus

End Sub

Private Sub TxtSrcObject_Change()
'Modification du fabriquant du comparateur

'Affecter
objControl.EquipmentManufacturer = Trim(TxtSrcObject)
m_objConfig.Last_EquipmentManufacturer = objControl.EquipmentManufacturer

End Sub

Private Sub UpDownFDMeasureCount_Change()
'Modification du nombre de mesures par point de relevé d'écart de fidélité

'Appliquer
TxtFDMeasureCount = Format(UpDownFDMeasureCount.Value)

'Mettre à jour
objControl.FidelityPoints.MeasuresCount = UpDownFDMeasureCount.Value
Refresh_GUIInitFidelityTest

End Sub

Private Sub UpDownHumidity_Change()
'Modification du taux d'himidité

'Afficher
TxtWatter = Format(UpDownHumidity.Value) + " %"

'Affecter
objControl.Humidity = UpDownHumidity.Value

End Sub

Private Sub UpDownMeasureCount_Change()
'Modification du nombre de mesures de la série

Dim nbValues As Integer
Dim i As Integer

'Appliquer
TxtMeasureCount = Format(UpDownMeasureCount.Value)

'Action selon l'écart calculé
nbValues = CInt(TxtMeasureCount)

'S'il faut modifier la collection
If objControl.SerialValues.Count <> nbValues Then

    'S'il faut ajouter des valeurs
    If objControl.SerialValues.Count < nbValues Then
    
        For i = objControl.SerialValues.Count + 1 To nbValues
            objControl.SerialValues.Add Int(objControl.SerialValues(i - 1).Value + 1)
        Next
        
    Else
        
        'Il faut retirer des valeurs
        For i = objControl.SerialValues.Count To nbValues + 1 Step -1
            objControl.SerialValues.Remove i
        Next
        
    End If
    
End If

'Préparer
CreateMSHFlexSerial
CreateMSChart

End Sub

Private Sub UpDownSerialCount_Change()
'Modification du nombre de séries de mesures

'Appliquer
TxtSerialCount = Format(UpDownSerialCount.Value)
objControl.SequencesCount = UpDownSerialCount.Value

'Préparer
CreateMSHFlexSerial
CreateMSChart

End Sub

Private Sub CreateMSChart()
'Préparer la structure du graphique

Dim varTitle As Variant

Dim i As Integer
Dim j As Integer

'Construction du tableau
varTitle = Array("Montante", "Descendante")

'Travail sur le graphique
With MSChart

    'Initialiser la taille
    .ColumnCount = 2
    .RowCount = objControl.SerialValues.Count
    
    'Initialiser les données
    For i = 1 To .ColumnCount
        .Column = i
        .ColumnLabel = varTitle(i - 1)
        For j = 1 To .RowCount
            .Row = j
            .RowLabel = Format(objControl.SerialValues(j), "0.00")
            .Data = 0
        Next
    Next
    
End With

'Libérer
Set varTitle = Nothing

End Sub

Private Sub CreateMSHFlexSerial()
'Préparer la structure du MSFlexSerial

Dim i As Integer
Dim SequencesCount As Integer
Dim SerialValuesCount As Integer

'Initialisation
With objControl
    SequencesCount = .SequencesCount
    SerialValuesCount = .SerialValues.Count
End With

With MSHFlexSerial

    'Supprimer les anciennes lignes si nécessaire
    .ClearStructure
    
    'Préparer
    .FixedCols = 1
    .FixedRows = 1
    .Rows = (SequencesCount * 2) + 3
    .Cols = SerialValuesCount + 1
    
    'Créer les titres verticaux
    .Col = 0
    For i = 1 To SequencesCount
        .Row = i
        .Text = "M" + Format(i)
        .Row = .Row + SequencesCount + 1
        .Text = "D" + Format(i)
    Next
    .Row = i
    .Text = "Moyenne"
    .Row = i * 2
    .Text = "Moyenne"
    
    'Créer les titres horizontaux
    .Row = 0
    For i = 1 To SerialValuesCount
        .Col = i
        .Text = Format(objControl.SerialValues(i).Value, "0.00") + " mm"
    Next
        
    'Largeur des colonnes
    For i = 1 To .Cols
        .ColWidth(i) = 780
        .ColAlignment(i - 1) = flexAlignLeftCenter
        .ColAlignmentFixed(i - 1) = flexAlignLeftCenter
    Next
    
End With

End Sub

Public Sub ConnectToMnuFile_Print_Slot()
'Accès public à mnuFile_Print_Click

'Pointer
mnuFile_Print_Click

End Sub

Private Sub Refresh_GUIInitFidelityTest()
'Rafraichir les composants graphiques concernés par le relevé d'écart de fidélité

Dim currentFidelityPoints As FidelityPoints

Dim iFidelityPointsCount As Integer
Dim iFDMeasuresCount As Integer
Dim i As Integer

'Initialisation graphique
Screen.MousePointer = vbHourglass

'Affecter
Set currentFidelityPoints = objControl.FidelityPoints
iFDMeasuresCount = currentFidelityPoints.MeasuresCount
iFidelityPointsCount = currentFidelityPoints.Count

'Travail sur le MSHFlexGrid
With MSHFlexFidelity
    
    'Masquer
    .Visible = False
    
    'Supprimer l'ancienne structure
    .ClearStructure
    
    'Initialiser
    .FixedCols = 1
    .FixedRows = 1
    .Rows = iFidelityPointsCount + 1
    .Cols = iFDMeasuresCount + 3
    
    'Créer les titres horizontaux
    .Row = 0
    For i = 1 To iFDMeasuresCount
        .Col = i
        .Text = "N°" + Str(i)
    Next
    .Col = i
    .Text = "Moyenne"
    .Col = i + 1
    .Text = "Fidélité"
    
    'Créer les titres verticaux
    .Col = 0
    For i = 1 To iFidelityPointsCount
        .Row = i
        .Text = Format(currentFidelityPoints(i).SerialValue, "0.00") + " " + IIf(currentFidelityPoints(i).IsUpDirection, "M", "D")
    Next
    
   'Largeur des colonnes
    For i = 1 To .Cols
        .ColWidth(i) = 780
        .ColAlignment(i - 1) = flexAlignLeftCenter
        .ColAlignmentFixed(i - 1) = flexAlignLeftCenter
    Next
    
End With

'Travail sur le TextBox
TxtFDMeasureCount = Format(objControl.FidelityPoints.MeasuresCount)

'Libérer
Set currentFidelityPoints = Nothing

'Restauration graphique
MSHFlexFidelity.Visible = True
Screen.MousePointer = vbNormal

End Sub

Private Sub InitFrameFinal()
'Initialise les données du frame final

Dim strVar As Variant
Dim i As Integer

'Charger les différents échelons
ComboLevel.Clear
For i = 1 To objControl.Levels.Count
    ComboLevel.AddItem Format(objControl.Levels(i).Value)
Next

'Intialiser le MSHFlexResult
With MSHFlexLimit
    
    'Initialiser
    .ClearStructure
    
    'Titres horizontaux
    strVar = Array("Justesse totale", "Justesse locale", "Fidélité", "Hystérésis", 1200, 1250, 700, 900)
    .Row = 0
    For i = 1 To .Cols - 1
        .Col = i
        .Text = strVar(i - 1)
        .ColWidth(i) = CInt(strVar(i + 3))
    Next
    
    'Alignement
    For i = 0 To .Cols - 1
        .ColAlignment(i - 1) = flexAlignLeftCenter
        .ColAlignmentFixed(i - 1) = flexAlignLeftCenter
    Next
    
    'Titres verticaux
    'strVar = Array("Classe 0", "Classe 1", "Mesures")
    strVar = Array("Normes", "Mesures")
    .Col = 0
    For i = 1 To UBound(strVar) + 1
        .Row = i
        .Text = strVar(i - 1)
    Next
    
    
    'Affichage des valeurs du comparateur
    strVar = _
    Array(objControl.ExactnessTotalError, objControl.ExactnessLocalError, objControl.FidelityError, objControl.HysteresisError)
    .Row = 2
    For i = 1 To 4
        .Col = i
        .CellBackColor = QBColor(QBCOLOR_GREY)
        .CellForeColor = QBColor(QBCOLOR_WHITE)
        .Text = Format(strVar(i - 1))
    Next
    
End With

'Initialisation du ListView
ListViewResult.ListItems.Clear

'Libérer
Set strVar = Nothing

End Sub

Private Sub Refresh_GUIFrameFinal()
'Rafaichir le panneau de finalisation en fonction des sélections et du résultat du contrôle

Dim currentClass0 As Classification
'Dim currentClass1 As Classification
Dim ListX As ListItem

Dim strItem As String
Dim strIco As String

'Affectation des classes courantes
Set currentClass0 = objControl.Levels(ComboLevel.ListIndex + 1).Capacitys(ComboCapacity.ListIndex + 1).Classifications(1)
'Set currentClass1 = objControl.Levels(ComboLevel.ListIndex + 1).Capacitys(ComboCapacity.ListIndex + 1).Classifications(2)

'Initialisation
ListViewResult.ListItems.Clear

'********************************************************************************
'Travail sur la justesse totale
'********************************************************************************
strItem = "L'erreur de justesse totale "

'Si "A REBUTER"
If objControl.ExactnessTotalError > currentClass0.TotalExactness Then

    strIco = ICO_FAILURE
    strItem = strItem + "est trop importante."
   
Else
    strIco = ICO_WARNING
    strItem = strItem + "est correcte."
End If

'Ajouter le ListItem
Set ListX = ListViewResult.ListItems.Add(, , strItem, strIco, strIco)

'********************************************************************************
'Travail sur la justesse locale
'********************************************************************************
strItem = "L'erreur de justesse locale "

'Si "A REBUTER"
If objControl.ExactnessLocalError > currentClass0.LocalExactness Then

    strIco = ICO_FAILURE
    strItem = strItem + "est trop importante."
   
Else
    strIco = ICO_WARNING
    strItem = strItem + "est correcte."
        
End If


'Ajouter le ListItem
Set ListX = ListViewResult.ListItems.Add(, , strItem, strIco, strIco)

'********************************************************************************
'Travail sur l'erreur de fidélité
'********************************************************************************
strItem = "L'erreur de fidélité "

'Si "A REBUTER"
If Abs(objControl.FidelityError) > currentClass0.Fidelity Then

    strIco = ICO_FAILURE
    strItem = strItem + "est trop importante."
Else
    strIco = ICO_WARNING
    strItem = strItem + "est correcte."
End If

'Ajouter le ListItem
Set ListX = ListViewResult.ListItems.Add(, , strItem, strIco, strIco)

'********************************************************************************
'Travail sur l'erreur d'Hystérésis
'********************************************************************************
strItem = "L'erreur d'hystérésis "

'Si "A REBUTER"
If Abs(objControl.HysteresisError) > currentClass0.Hysteresis Then
    strIco = ICO_FAILURE
    strItem = strItem + "est trop importante."
Else
    strIco = ICO_WARNING
    strItem = strItem + "est correcte."
End If

'Ajouter le ListItem
Set ListX = ListViewResult.ListItems.Add(, , strItem, strIco, strIco)

'********************************************************************************
'Travail de fin
'********************************************************************************
Select Case objControl.GetTheoreticalClass(ComboLevel.ListIndex + 1, ComboCapacity.ListIndex + 1, True)
    Case 0
        With TxtResult
            .ForeColor = QBColor(QBCOLOR_GREEN)
            .Text = "BON ETAT"
        End With
    'Case 1
    '    With TxtResult
    '        .ForeColor = QBColor(QBCOLOR_YELLOW)
    '        .Text = "CLASSE 1"
    '    End With
    Case 1
        With TxtResult
            .ForeColor = QBColor(QBCOLOR_RED)
            .Text = "A REBUTER"
        End With
End Select

'Libérer
Set currentClass0 = Nothing
'Set currentClass1 = Nothing
Set ListX = Nothing

End Sub

Private Sub GUI_SetEnabledOfFrameCarac(bIsEnabled As Boolean)
'Appliquer une valeur de Enabled sur les objets du FrameCarac

'Appliquer sur l'interface
LblSerialCount.Enabled = bIsEnabled
TxtSerialCount.Enabled = bIsEnabled
UpDownSerialCount.Enabled = bIsEnabled
LblMeasures.Enabled = bIsEnabled
TxtMeasureCount.Enabled = bIsEnabled
UpDownMeasureCount.Enabled = bIsEnabled
LblSerialValues.Enabled = bIsEnabled
ComboSerialValues.Enabled = bIsEnabled
CmdModifyMeasures.Enabled = bIsEnabled
FrameCarac.Enabled = bIsEnabled

End Sub

Private Sub GUI_SetEnabledOfFrameListOfMeasures(bIsEnabled As Boolean)
'Appliquer une valeur de Enabled sur les objets du FrameListOfMeasures

'Appliquer sur l'interface
LblFDMeasureCount.Enabled = bIsEnabled
TxtFDMeasureCount.Enabled = bIsEnabled
UpDownFDMeasureCount.Enabled = bIsEnabled
LblMeasurePoint.Enabled = bIsEnabled
ListDispo.Enabled = bIsEnabled
LblMeasureEnabled.Enabled = bIsEnabled
ListEnabledMeasure.Enabled = bIsEnabled
CmdAddOne.Enabled = bIsEnabled
CmdRemoveOne.Enabled = bIsEnabled
FrameListOfMeasures.Enabled = bIsEnabled

End Sub

