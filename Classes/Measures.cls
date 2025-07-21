VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Measures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Measure"
Attribute VB_Ext_KEY = "Member0" ,"Measure"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UNE COLLECTION DE MESURES
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mCol As Collection

Public Sub Clear()
'Vider la collection

Dim i As Integer

'Dépiler
For i = mCol.Count To 1 Step -1
    mCol.Remove i
Next

End Sub

Public Function Add(Value As Currency, Optional sKey As String) As Measure
'Ajouter une mesure à la collection
'Renvoie le nouvel objet ou Nothing si erreur

Dim objNewMember As Measure
    
'Créer un nouvel objet
Set objNewMember = New Measure
    
'Définir les propriétés
objNewMember.Value = Value

'Gestion pas à pas des erreurs
On Error Resume Next

'Ajouter
sKey = Trim(sKey)
If Len(sKey) = 0 Then
    mCol.Add objNewMember
Else
    mCol.Add objNewMember, sKey
End If

'Sécurité
If Err > 0 Then Set objNewMember = Nothing

'Renvoyer
Set Add = objNewMember

'Libérer
Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Measure
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

'Retirer
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
