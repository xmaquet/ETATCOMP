VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FidelityPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UNE COLLECTION DE POINTS DE RELEVE D'ECART DE FIDELITE
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mCol As Collection
Private mvarMeasuresCount As Integer

Public Function Add(ByVal SerialValue As Currency, ByVal IsUpDirection As Boolean, Optional ByVal sKey As String) As FidelityPoint
'Ajouter un point de relevé d'écart de fidélité (Avec tri)
'Renvoie l'objet crée ou Nothing si erreur

Dim objNewMember As FidelityPoint
Dim currentFidelityPoint As FidelityPoint

Dim iAfter As Integer

'Gestion pas à pas des erreurs
On Error Resume Next

'Créer le nouvel objet
Set objNewMember = New FidelityPoint

'Affecter les propriétes
With objNewMember
    .IsUpDirection = IsUpDirection
    .SerialValue = SerialValue
End With

'Préparer
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
        Set currentFidelityPoint = mCol.Item(iAfter)
        
        'Sortir si la valeur est supérieure
        If currentFidelityPoint.SerialValue < SerialValue Then Exit For
        
        'Sortir si la valeur est égale
        If currentFidelityPoint.SerialValue = SerialValue Then
            
            '"M" prioritaire sur "D"
            If Not currentFidelityPoint.IsUpDirection Then iAfter = iAfter - 1
            
            'Sortir
            Exit For
            
        End If
        
    Next
    
    'Libérer
    Set currentFidelityPoint = Nothing
    
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

'Affecter la valeur de retour
If Err > 0 Then
    Set objNewMember = Nothing
Else
    'Créer la collection de mesures
    For iAfter = 1 To mvarMeasuresCount
        objNewMember.Measures.Add 0
    Next
End If

'Renvoyer
Set Add = objNewMember

'Libérer
Set objNewMember = Nothing

End Function

Public Property Get Item(ByVal IndxKey As Variant) As FidelityPoint
Attribute Item.VB_UserMemId = 0
'Accès en lecture sur un élément

'Renvoyer
Set Item = mCol(IndxKey)

End Property

Public Sub Remove(ByVal IndxKey As Variant)
'Retirer un élément

'Suppression
mCol.Remove IndxKey

End Sub

Public Function Clear()
'Vider la collection

Dim i As Integer

'Dépiler
For i = mCol.Count To 1 Step -1
    mCol.Remove i
Next

End Function

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

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'Interface "For Each ..."

'Renvoyer
Set NewEnum = mCol.[_NewEnum]

End Property

Public Property Get Count() As Integer
'Accès en lecture sur le nombre d'éléments

'Renvoyer
Count = mCol.Count

End Property

Public Property Let MeasuresCount(iMeasureCount As Integer)
'Accès en écriture sur le nombre de mesures a effectuer pour chaque point de relevé

Dim currentMeasures As Measures

Dim i As Integer
Dim j As Integer

'Sécurité
If iMeasureCount < 1 Then iMeasureCount = 1

'Affecter
mvarMeasuresCount = iMeasureCount

'Enumérer les points
For i = 1 To mCol.Count

    'Affecter
    Set currentMeasures = Item(i).Measures
    
    'Vider la collection
    currentMeasures.Clear
    
    'Créer les éléments
    For j = 1 To iMeasureCount
        currentMeasures.Add 0
    Next

Next

'Libérer
Set currentMeasures = Nothing

End Property

Public Property Get MeasuresCount() As Integer
'Accès en lecture sur le nombre de mesures par point de relevé d'écart

'Renvoyer
MeasuresCount = mvarMeasuresCount

End Property

Public Sub RemoveItemByString(ByVal strID As String)
'Retirer un élément identifié par une chaîne

Dim currentFidelityPoint As FidelityPoint

Dim cSerialValue As Currency
Dim bIsUp As Boolean
Dim i As Integer

'Optimisation
If mCol.Count = 0 Then Exit Sub

'Mise en forme - Sécurité
strID = UCase(Trim(strID))
If strID = "" Then Exit Sub

'Positionnement des variables
bIsUp = (Right(strID, 1) <> "D")
cSerialValue = CCur(Left(strID, Len(strID) - 1))

'Enumérer les points de fidélité
For i = 1 To mCol.Count

    'Affecter
    Set currentFidelityPoint = mCol(i)
    
    'Conditions de sortie
    If (currentFidelityPoint.IsUpDirection = bIsUp) And (currentFidelityPoint.SerialValue = cSerialValue) Then Exit For
        
Next

'Libérer
Set currentFidelityPoint = Nothing

'Suppression
If i < (mCol.Count + 1) Then mCol.Remove i

End Sub
