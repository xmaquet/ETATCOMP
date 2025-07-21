VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Levels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Level"
Attribute VB_Ext_KEY = "Member0" ,"Level"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UNE COLLECTION D'ECHELONS
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mCol As Collection

Public Function Add(ByVal Value As Currency, Optional sKey As String) As Level
'Ajouter un objet à la collection

Dim objNewMember As Level

'Créer un nouvel objet
Set objNewMember = New Level

'Définir les propriétés
objNewMember.Value = Value

'Ajouter
If Len(sKey) = 0 Then
    mCol.Add objNewMember
Else
    mCol.Add objNewMember, sKey
End If

'Renvoyer
Set Add = objNewMember

'Libérer
Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Level
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

'Supression
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
'COnstructeur

'Initialiser
Set mCol = New Collection
LoadData

End Sub

Private Sub Class_Terminate()
'Destructeur

'Libérer
Set mCol = Nothing

End Sub

Private Sub LoadData()
'Chargement du tableau de données

Dim currentLevel As Level
Dim currentCapacity As Capacity

Dim varLevels As Variant
Dim varCapacitys As Variant
Dim varClassification As Variant

Dim i As Integer
Dim j As Integer
Dim k As Integer

'Initialiser le tableau des échelons
varLevels = Array(0.001, 0.01, 0.1)

'Créer les échelons
For i = LBound(varLevels) To UBound(varLevels)

    'Affecter l'échelon courant
    Set currentLevel = Add(varLevels(i))
    
    'Créer le tableau des courses
    Select Case currentLevel.Value
        Case 0.001
            varCapacitys = Array(1, 10)
        Case 0.01
            varCapacitys = Array(5, 10, 30)
        Case 0.1
            varCapacitys = Array(30)
    End Select
    
    'Créer les courses
    For j = LBound(varCapacitys) To UBound(varCapacitys)
        
        'Affecter la course courante
        Set currentCapacity = currentLevel.Capacitys.Add(CInt(varCapacitys(j)))
        
        'Créer le tableau des classes
        Select Case currentLevel.Value
            Case 0.001
                Select Case currentCapacity.Value
                    Case 1
                        varClassification = Array(5, 999, 0.25, 2) ', 7, 4, 0.7, 1.5)
                    Case 10
                        varClassification = Array(10, 3, 0.25, 2) ', 15, 4, 0.7, 1.5)
                End Select
            Case 0.01
                Select Case currentCapacity.Value
                    Case 5
                        varClassification = Array(15, 10, 1.5, 6) ', 22, 15, 4, 4)
                    Case 10
                        varClassification = Array(15, 10, 2.5, 10) ', 22, 15, 7, 7)
                    Case 30
                        varClassification = Array(20, 10, 5, 10) ', 30, 15, 15, 7)
                    End Select
            Case 0.1
                Select Case currentCapacity.Value
                    Case 30
                        varClassification = Array(150, 100, 15, 60) ', 220, 150, 45, 45)
                End Select
        End Select
        
        'Créer les classes
        For k = LBound(varClassification) To UBound(varClassification) Step 4
            currentCapacity.Classifications.Add _
            varClassification(k), varClassification(k + 1), varClassification(k + 2), varClassification(k + 3)
        Next
        
    Next
    
Next

'Libérer
Set currentCapacity = Nothing
Set currentLevel = Nothing
Set varCapacitys = Nothing
Set varClassification = Nothing
Set varLevels = Nothing

End Sub
