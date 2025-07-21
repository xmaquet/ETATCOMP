VERSION 5.00
Begin VB.Form FrmTips 
   Caption         =   "Astuces du jour"
   ClientHeight    =   3285
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5220
   Icon            =   "FrmTips.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Afficher les astuces au démarrage"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2940
      Width           =   2895
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Suivante"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox PicLight 
      BackColor       =   &H80000005&
      Height          =   2715
      Left            =   120
      Picture         =   "FrmTips.frx":0442
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label LblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Le saviez-vous ..."
         Height          =   195
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H80000005&
         Height          =   1875
         Left            =   540
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.CommandButton CmdValid 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************
'FEUILLE D'AFFICHAGE DES ASTUCES DU JOURS
'********************************************************************************

'********************************************************************************
'Constantes
'********************************************************************************
Private Const TIPS_FILE = "Tips.dat"

'********************************************************************************
'Données membres
'********************************************************************************
Private mCol As Collection
Private mRefOfConfig As Configuration

Public Function IsOkToShow(objConfig As Configuration) As Boolean
'Renvoie vrai si la feuille peut-être affichée

Dim iFile As Integer
Dim strTip As String

'Les erreurs sont gérées
On Error GoTo GestErr

'Tester l'existence du fichier
If Dir(App.Path + APP_FOLDER_SYSTEM + TIPS_FILE) = "" Then Exit Function

'Initialisation
Set mCol = New Collection

'Ouverture du fichier
iFile = FreeFile
Open App.Path + APP_FOLDER_SYSTEM + TIPS_FILE For Input As #iFile

'Boucle de chargement des astuces
Line Input #iFile, strTip
Do While Not EOF(iFile)
    
    'Réduire
    strTip = Trim(strTip)
    
    'Ajouter
    If strTip <> "" Then mCol.Add strTip
    
    'Lire la prochaine ligne
    Line Input #iFile, strTip
    
Loop

'Fermeture du fichier
Close iFile

'Copie de la référence
Set mRefOfConfig = objConfig

'Mise à jour de la case à cocher
chkLoadTipsAtStartup.Value = IIf(mRefOfConfig.Tips_ShowTips, 1, 0)

'Affichage de l'astuce courante
ShowNextTip

'Chargement OK
IsOkToShow = True

'Sortir normalement
Exit Function

'Gestion des erreurs
GestErr:

Set mCol = Nothing

End Function

Private Sub ShowNextTip()
'Affiche l'astuce suivante et gère l'incrémentation

'Incrémenter
mRefOfConfig.Tips_CurrentTip = IIf(mRefOfConfig.Tips_CurrentTip = mCol.Count, 1, mRefOfConfig.Tips_CurrentTip + 1)

'Afficher
lblTipText = mCol(mRefOfConfig.Tips_CurrentTip)

End Sub

Private Sub chkLoadTipsAtStartup_Click()
'Clic dans la case à cocher

'Positionner
mRefOfConfig.Tips_ShowTips = CBool(chkLoadTipsAtStartup.Value)

End Sub

Private Sub cmdNextTip_Click()
'Clic sur le bouton "Suivante"

'Afficher l'asstuce suivante
ShowNextTip

End Sub

Private Sub CmdValid_Click()
'Clic sur le bouton "OK"

'Décharger la feuille
Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Déchargement de la feuille

'Libérer
Set mCol = Nothing
Set mRefOfConfig = Nothing
Set FrmTips = Nothing

End Sub
