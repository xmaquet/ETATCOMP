VERSION 5.00
Begin VB.Form FrmEnd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modifier le résultat d'un contrôle"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "FrmEnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdValid 
      Caption         =   "Valider"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame FrameResult 
      Caption         =   "Nouvelle valeur à appliquer"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton OptionResult 
         Caption         =   "A REBUTER"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton OptionResult 
         Caption         =   "CLASSE 1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptionResult 
         Caption         =   "CLASSE 0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************
'FEUILLE DE MODIFICATION DU RESULTAT D'UN CONTROLE
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private m_refOfControlObject As ControlObject
Private m_LastOptionSelected As Integer

Public Sub SetReferenceToObject(objControl As ControlObject)
'Affecter la référence à l'objet de contrôle

'Affecter
Set m_refOfControlObject = objControl

'Mettre à jour l'interface
OptionResult(m_refOfControlObject.RealClass).Value = True

End Sub

Private Sub CmdCancel_Click()
'Clic sur le bouton "Annuler"

'Décharger la feuille
Unload Me

End Sub

Private Sub CmdValid_Click()
'Clic sur le bouton "Valider"

Dim i As Integer

'Affecter la valeur
m_refOfControlObject.RealClass = m_LastOptionSelected

'Décharger la feuille
Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Déchargement de la feuille

'Libérer
Set m_refOfControlObject = Nothing
Set FrmEnd = Nothing

End Sub

Private Sub OptionResult_Click(Index As Integer)
'Clic sur une case d'option

'Stocker la valeur de l'index
m_LastOptionSelected = Index

End Sub
