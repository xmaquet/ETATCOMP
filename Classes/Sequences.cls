VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sequences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Sequence"
Attribute VB_Ext_KEY = "Member0" ,"Sequence"
Option Explicit
Option Base 1

'********************************************************************************
'IMPLEMENTATION D'UNE COLLECTION DE SEQUENCES (DESCENDANTES OU MONTANTES)
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mAverage() As Currency
Private mCol As Collection

Public Sub Clear()
'Vider la collection

Dim i As Integer

'Libérer
ReDim mAverage(1)

'Dépiler
For i = mCol.Count To 1 Step -1
    mCol.Remove i
Next

End Sub

Public Function Add(Optional sKey As String) As Sequence
'Ajouter une séquence à la collection
'Renvoie la nouvelle séquence ou Nothing si erreur

Dim objNewMember As Sequence

'Créer un nouvel objet
Set objNewMember = New Sequence

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

Public Property Get Item(vntIndexKey As Variant) As Sequence
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
ReDim mAverage(1)
Set mCol = New Collection

End Sub

Private Sub Class_Terminate()
'Destructeur

'Libérer
Erase mAverage
Set mCol = Nothing

End Sub

Public Sub SetAverage(ByVal indexOfSerialValue As Integer, ByVal sAverage As Currency)
'Affecter la valeur moyenne pour une valeur de la série de mesure

'Redimensionner si taille insuffisante
If UBound(mAverage) < indexOfSerialValue Then ReDim Preserve mAverage(indexOfSerialValue)

'Enregistrer
mAverage(indexOfSerialValue) = sAverage

End Sub

Public Function GetAVerage(ByVal IndxOfSerialValue As Integer) As Currency
'Renvoie la moyenne d'écart d'une valeur de série de mesures

'Renvoyer
GetAVerage = mAverage(IndxOfSerialValue)

End Function

Public Function GetIndexOfMaxSerialValueAverage() As Integer
'Renvoie l'index de la SerialValue avec la plus
'grande moyenne d'écart

Dim i As Integer
Dim valMax As Currency
Dim indxMax As Integer

'Stocker la plus grande valeur
For i = 1 To UBound(mAverage)
    If Abs(mAverage(i)) >= valMax Then
        valMax = Abs(mAverage(i))
        indxMax = i
    End If
Next

'Renvoyer
GetIndexOfMaxSerialValueAverage = indxMax
 
End Function
