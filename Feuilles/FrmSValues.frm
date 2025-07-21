VERSION 5.00
Begin VB.Form FrmSValues 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valeurs de la série de mesures"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "FrmSValues.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Sélection d'un comprateur"
      Height          =   2565
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   3495
      Begin VB.ListBox lstComparateurs 
         Height          =   2010
         Left            =   160
         TabIndex        =   10
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton CmdValid 
      Caption         =   "Valider"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame FrameList 
      Caption         =   "Liste des valeurs (mm)"
      Height          =   2295
      Left            =   80
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
      Begin VB.CommandButton CmdRemove 
         Caption         =   "&Supprimer"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         ToolTipText     =   "Supprimer la valeur sélectionnée"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ListBox ListValues 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Saisie d'une nouvelle valeur"
      Height          =   1000
      Left            =   80
      TabIndex        =   0
      Top             =   80
      Width           =   4455
      Begin VB.TextBox TxtNew 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Ajouter"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         ToolTipText     =   "Ajouter la valeur saisie"
         Top             =   560
         Width           =   1215
      End
      Begin VB.Label LblNewValue 
         AutoSize        =   -1  'True
         Caption         =   "Nouvelle valeur à ajouter :"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1860
      End
   End
End
Attribute VB_Name = "FrmSValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************
'FEUILLE DE GESTION DE LA SERIE DE MESURES
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mCopyOfSerialValues As SerialValues         'Copie de la collection
Private mRefOfOriginal As SerialValues              'Référence sur la copie originale


Private sComp() As String


Private Sub CmdAdd_Click()
'Clic sur le bouton "Ajouter"

Dim i As Integer
Dim txt As String

Dim cTest As Currency

'Sécurité
txt = Trim(TxtNew)
If txt = "" Then
    
    'Afficher un message
    txt = "Saisir une valeur numérique dans la zone de texte avant de cliquer sur ""Ajouter""."
    i = MsgBox(txt, vbExclamation, "Ajouter")
    
    'Sortir
    Exit Sub
    
End If

'Remplacer le "." par une ","
i = InStr(1, txt, ".")
If i > 0 Then txt = Left(txt, i - 1) + "," + Right(txt, Len(txt) - i)

'Tenter une conversion
On Error Resume Next
cTest = CCur(txt)

'Si conversion impossible
If Err > 0 Then
    
    'Afficher un message
    txt = """" + txt + """ n'est pas une valeur numérique valide !"
    i = MsgBox(txt, vbCritical, "Ajouter")
    
    'Sortir
    Exit Sub
    
End If

'Ajouter à la collection
mCopyOfSerialValues.Add cTest

'Vider la zone de texte
TxtNew = ""

'Rafraichir
RefreshGUI

'Marquer
Tag = TAG_CLIC_ON_VALID

End Sub

Private Sub CmdCancel_Click()
'Clic sur le bouton "Annuler"

'Masquer
Hide

End Sub

Public Sub SetReferenceToObject(objSValues As SerialValues)
'Copier la référence à la collection originale

'Affecter la référence
Set mRefOfOriginal = objSValues

'Affecter la copie
Set mCopyOfSerialValues = New SerialValues
mCopyOfSerialValues.Copy mRefOfOriginal

'Rafraichir
RefreshGUI
' Lire le fichier des comparateurs enregistrés
LireFichierComparateurs
End Sub

Private Sub RefreshGUI()
'Rafraichir l'interface graphique

Dim currentSValue As SerialValue

'Initialiser la zone de liste
With ListValues

    .Clear
    
    'Enumérer les valeurs
    For Each currentSValue In mCopyOfSerialValues
            
        'Ajouter à la zone de liste
        .AddItem Format(currentSValue.Value, "0.00")
        
    Next
    
    'Actions selon le nombre d'éléments
    If .ListCount > 0 Then
    
        .ListIndex = 0
        CmdRemove.Enabled = True
        
    Else
        
        CmdRemove.Enabled = False
        
    End If
    
End With
End Sub


Private Sub LireFichierComparateurs()
'Récupérer les valeurs des comprateurs enregistrés
Dim sCar As String, i As Integer, j As Integer
Open App.Path & "\Comparateurs.txt" For Input As #1
ReDim sComp(10, 12)
Do While Not (EOF(1))
    sCar = Input(1, #1)
    If Asc(sCar) <> 13 And Asc(sCar) <> Asc(";") Then
        sComp(i, j) = sComp(i, j) & sCar
    ElseIf Asc(sCar) = 13 Then
        lstComparateurs.AddItem sComp(i, 0)
        i = i + 1: j = 0
    ElseIf Asc(sCar) = Asc(";") Then
        j = j + 1
    End If
Loop
Close "1"

End Sub

Private Sub lstComparateurs_Click()
Dim i As Integer
If lstComparateurs.ListIndex > -1 Then
    ListValues.Clear
    For i = 1 To UBound(sComp, 2)
        If sComp(lstComparateurs.ListIndex, i) <> "" Then
            ListValues.AddItem sComp(lstComparateurs.ListIndex, i)
        End If
    Next i
End If
For i = 1 To mCopyOfSerialValues.Count
    mCopyOfSerialValues.Remove (1)
Next i
'Marquer
Tag = TAG_CLIC_ON_VALID
For i = 0 To ListValues.ListCount - 1
    mCopyOfSerialValues.Add ListValues.List(i)
Next i
End Sub


Private Sub CmdRemove_Click()
'Clic sur le bouton "Supprimer"

'Sécurité
If ListValues.ListIndex = -1 Then Exit Sub

'Demander confirmation
If MsgBox("Confirmer la suppression de la valeur """ + ListValues.List(ListValues.ListIndex) + """ ?", _
vbQuestion + vbYesNo, "Supprimer") = vbNo Then Exit Sub

'Supprimer
mCopyOfSerialValues.Remove ListValues.ListIndex + 1

'Rafraichir
RefreshGUI

'Marquer
Tag = TAG_CLIC_ON_VALID

End Sub

Private Sub CmdValid_Click()
'Clic sur le bouton "Valider"

' si la liste est une liste enregistrée dans le fichier
' alors supression des anciennes les valeurs et ajout des nouvelles

'Copier
mRefOfOriginal.Copy mCopyOfSerialValues

'Masquer
Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Déchargement de la feuille

'Libérer
Set mCopyOfSerialValues = Nothing
Set mRefOfOriginal = Nothing
Set FrmSValues = Nothing

End Sub

