VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture 
      AutoSize        =   -1  'True
      Height          =   2400
      Left            =   0
      Picture         =   "FrmSplash.frx":0000
      ScaleHeight     =   2340
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.Timer TimerUnloadMe 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   5640
         Top             =   840
      End
      Begin VB.Timer TimerShowFrmGen 
         Interval        =   2000
         Left            =   3240
         Top             =   1080
      End
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************
'FEUILLE DE DEMARRAGE DE L'APPLICATION
'********************************************************************************

Private Sub Form_Unload(Cancel As Integer)
'Déchargement de la feuille

'Activer FrmGen
FrmGen.Enabled = True

'Libérer
Set FrmSplash = Nothing

End Sub

Private Sub TimerShowFrmGen_Timer()
'Minuterie de chargement de FrmGen écoulée

'Fin de la minuterie
TimerShowFrmGen.Enabled = False

'Charger FrmGen
FrmGen.Show vbModeless

'Repasser en 1er plan
ZOrder

'Activer la minuterie de déchargement du Splash
TimerUnloadMe.Enabled = True

End Sub

Private Sub TimerUnloadMe_Timer()
'Minuterie de déchargement écoulée

'Désactiver la minuterie
TimerUnloadMe.Enabled = False

'Décharger la feuille
Unload Me

End Sub
