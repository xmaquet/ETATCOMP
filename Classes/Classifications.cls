VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Classifications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Classification"
Attribute VB_Ext_KEY = "Member0" ,"Classification"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UNE COLLECTION DE CLASSIFICATIONS (CLASSE 0 OU 1)
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mCol As Collection

Public Function Add(ByVal TotalExactness As Integer, ByVal LocalExactness As Integer, ByVal Fidelity As Currency, _
ByVal Hysteresis As Currency, Optional sKey As String) As Classification
'Ajouter un objet à la collection

Dim objNewMember As Classification

'Créer un nouvel objet
Set objNewMember = New Classification

'Définir les propriétés
With objNewMember
    .Fidelity = Fidelity
    .Hysteresis = Hysteresis
    .LocalExactness = LocalExactness
    .TotalExactness = TotalExactness
End With

'Ajouter
If Len(sKey) = 0 Then
    mCol.Add objNewMember
Else
    mCol.Add objNewMember, sKey
End If

'Renvoyer l'objet créé
Set Add = objNewMember

'Libérer
Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Classification
Attribute Item.VB_UserMemId = 0
'Accès en lecture sur un élément de la collection

'Renvoyer
Set Item = mCol(vntIndexKey)

End Property

Public Property Get Count() As Long
'Accès en lecture sur le nombre d'éléments de la collection

'Renvoyer
Count = mCol.Count

End Property

Public Sub Remove(vntIndexKey As Variant)
'Retirer un élément de la collection

'Supprimer
mCol.Remove vntIndexKey

End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'Interface "For Each ..."

'Renvoyer
Set NewEnum = mCol.[_NewEnum]

End Property

Private Sub Class_Initialize()
'Constructeur

'Initialiser
Set mCol = New Collection

End Sub

Private Sub Class_Terminate()
'Destructeur

'Libérer
Set mCol = Nothing

End Sub
