VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SerialValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"SerialValue"
Attribute VB_Ext_KEY = "Member0" ,"SerialValue"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UNE SERIE DE VALEURS DE MESURES
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

Public Function Add(ByVal Value As Currency, Optional sKey As String) As SerialValue
'Ajouter une nouvelle valeur à la série de mesure
'La collection est automatiquement triée par ordre croissant de valeur
'Renvoie le nouvel objet ou Nothing si erreur

Dim objNewMember As SerialValue
Dim currentSerialValue As SerialValue

Dim iAfter As Integer

'Créer le nouvel objet
Set objNewMember = New SerialValue
objNewMember.Value = Value

'Gestion pas à pas des erreur
On Error Resume Next

'Mise en forme de la clef
sKey = Trim(sKey)

'Si la collection est vide
If mCol.Count = 0 Then

    'Ajouter
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
Else

    'La collection n'est pas vide
    
    'Pointer sur l'index de la valeur immédiatement inférieure
    For iAfter = mCol.Count To 1 Step -1
        
        'Pointer sur l'objet courant
        Set currentSerialValue = mCol.Item(iAfter)
        
        'Sortir si la valeur est supérieure
        If currentSerialValue.Value < Value Then Exit For
        
    Next
    
    'Libérer
    Set currentSerialValue = Nothing
    
    'L'objet est en tête de pile
    If iAfter = 0 Then
        
        'Ajouter à la première position
        If Len(sKey) = 0 Then
            mCol.Add objNewMember, , 1
        Else
            mCol.Add objNewMember, sKey, 1
        End If
        
    Else
    
        'L'objet est dans la pile
        
        'Ajouter en insertion
        If Len(sKey) = 0 Then
            mCol.Add objNewMember, , , iAfter
        Else
            mCol.Add objNewMember, sKey, , iAfter
        End If
    
    End If
    
End If

'Contrôle d'erreur
If Err > 0 Then Set objNewMember = Nothing

'Renvoyer l'objet créé
Set Add = objNewMember

'Libérer
Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As SerialValue
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

Public Sub Copy(objSource As SerialValues)
'Copier une collection

Dim currentSValue As SerialValue
Dim newObject As SerialValue

'Vider
Clear

'Copier
For Each currentSValue In objSource
        
    'Créer une copie
    Set newObject = New SerialValue
    newObject.Value = currentSValue.Value
    
    'Ajouter
    mCol.Add newObject
    
Next

'Libérer
Set newObject = Nothing

End Sub
