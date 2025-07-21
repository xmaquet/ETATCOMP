VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   3135
      Left            =   80
      TabIndex        =   1
      Top             =   80
      Width           =   4815
      Begin MSComCtl2.Animation Animation 
         Height          =   1140
         Left            =   165
         TabIndex        =   2
         Top             =   285
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   2011
         _Version        =   393216
         FullWidth       =   302
         FullHeight      =   76
      End
      Begin VB.PictureBox Picture 
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   4560
         TabIndex        =   3
         Top             =   240
         Width           =   4620
      End
      Begin VB.Label LblDescript 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Etalonnage de comparateurs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1535
         Width           =   4620
      End
      Begin VB.Label LblVersion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1870
         Width           =   4620
      End
      Begin VB.Label LblAuthor 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Développé en Visual Basic par le TSEF BLAUBLOMME (2002)"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2745
         Width           =   4620
      End
      Begin VB.Label LblWhere 
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
         Height          =   460
         Left            =   120
         TabIndex        =   4
         Top             =   2205
         Width           =   4620
      End
   End
   Begin VB.CommandButton CmdValid 
      Caption         =   "&Vu"
      Height          =   375
      Left            =   3680
      TabIndex        =   0
      Top             =   3280
      Width           =   1215
   End
   Begin VB.Timer TimerPlay 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   1440
      Top             =   480
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************
'FEUILLE DE DESCRIPTION DE L'APPLICATION
'********************************************************************************

'********************************************************************************
'Constantes
'********************************************************************************
Private Const AVI_FILE = "EtaComp.avi"

Private Sub CmdValid_Click()
'Clic sur le bouton "Vu"

'Décharger la  fenêtre
Unload Me

End Sub

Private Sub Form_Load()
'Chargement de la feuille

'Si le fichier existe
If Dir(App.Path + APP_FOLDER_SYSTEM + AVI_FILE) <> "" Then
    
    'Ouvrir
    Animation.Open App.Path + APP_FOLDER_SYSTEM + AVI_FILE
    
    'Activer la minuterie
    TimerPlay.Enabled = True
    
    'Jouer une première fois
    Animation.Play 1
    
End If

'Affecter les textes
Caption = "A propos de " + App.EXEName
LblVersion = "Version" + Str(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
LblWhere = "2ème Régiment du Matériel" + vbCrLf + "BRUZ"

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Déchargement de la feuille

'Arrêter l'animation
Animation.Stop

'Libérer
Set FrmAbout = Nothing

End Sub

Private Sub TimerPlay_Timer()
'Minuterie écoulée

Animation.Play 1

End Sub
