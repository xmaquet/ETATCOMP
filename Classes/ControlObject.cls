VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Measure"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Measures"
Attribute VB_Ext_KEY = "Member2" ,"Sequence"
Attribute VB_Ext_KEY = "Member3" ,"Sequences"
Attribute VB_Ext_KEY = "Member4" ,"SerialValues"
Attribute VB_Ext_KEY = "Member5" ,"FidelityPoints"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UN CONTROLE DE COMPARATEUR
'********************************************************************************

'********************************************************************************
'Types définis
'********************************************************************************
Private Type PrintData
    TextWidth As Integer
    ColWidth As Integer
End Type

'********************************************************************************
'Constantes
'********************************************************************************

'Marges verticales et horizontales pour l'impression
Const VERTICAL_MARGIN = 2

'********************************************************************************
'Evènements
'********************************************************************************

'Survient avant le passage à une nouvelle position dans la série de mesures
Public Event LeavePosition()

'Survient après un changement de mode du contrôle (Correction ou non)
Public Event CorrectionModeChange()

'Survient après un changement de mode de contrôle du test de fidélité
Public Event CorrectionModeOfFidelityTestChange()

'Survient après le passage à une nouvelle position dans la série de mesures
Public Event PositionChange()

'Survient lorsque la phase des séries de mesures est terminée
Public Event SerialValueTestTerminated()

'Survient lorsque la phase des rélévés d'écart de fidélité est commencée
Public Event FidelityPointsTestStart()

'Survient lorsque la phase des rélévés d'écart de fidélité est terminée
Public Event FidelityPointsTestTerminated()

'Survient après la sauvegarde du point de fidélité et avant le calcul de la nouvelle position
Public Event AfterFidelityPointSaveWithoutStep()

'Survient après la sauvegarde d'un point de relevé d'écart
Public Event AfterFidelityPointSave()

'Survient à la fin du contrôle
Public Event ControlIsEnd()

'********************************************************************************
'Données membres
'********************************************************************************
Private mvarSerialValues As SerialValues
Private mvarUpSequences As Sequences
Private mvarDownSequences As Sequences
Private mvarFidelityPoints As FidelityPoints
Private mvarLevels As Levels

Private mvarDateOfControl As Date
Private mvarTemperature As Currency
Private mvarDetenteur As String
Private mvarHumidity As Integer
Private mvarOperatorName As String
Private mvarEquipmentNumber As String
Private mvarEquipmentManufacturer As String
Private mvarEquipmentReference As String
Private mvarSequencesCount As Integer
Private mvarComment As String

'Etat du contrôle
Private mvarCorrectionMode As Boolean
Private mvarIsOkToStart As Boolean
Private mvarControlIsStart As Boolean
Private mvarSequencesAreEnd As Boolean
Private mvarFidelityTestIsStart As Boolean
Private mvarFidelityTestIsEnd As Boolean
Private mvarControlIsEnd As Boolean
Private mvarCorrectionModeOfFidelityTest As Boolean

'Positions dans l'ordre normal de mesure
Private mvarNormalIndexOfSerialValue As Integer
Private mvarNormalIndexOfSequence As Integer
Private mvarNormalSequenceIsUp As Boolean

'Positions courante
Private mvarCurrentIndexOfSerialValue As Integer
Private mvarCurrentIndexOfSequence As Integer
Private mvarCurrentSequenceIsUp As Boolean

'Positions dans l'ordre normal de relevé d'écarts
Private mvarNormalFidelityIndexOfPoint As Integer
Private mvarNormalFidelityIndexOfMeasure As Integer

'Positions courantes de relevé d'écarts
Private mvarCurrentFidelityIndexOfPoint As Integer
Private mvarCurrentFidelityIndexOfMeasure As Integer

'Résultat du contrôle
Private mvarRealClass As Integer
Private mvarIndexOfLevel As Integer
Private mvarIndexOfCapacity As Integer

Public Property Get CurrentFidelityIndexOfMeasure() As Integer
'Accès en lecture sur la mesure courante du point de fidélité courant

'Renvoyer
CurrentFidelityIndexOfMeasure = mvarCurrentFidelityIndexOfMeasure

End Property

Public Property Let CurrentFidelityIndexOfMeasure(ByVal iData As Integer)
'Accès en écriture sur la mesure courante du point de fidélité courant

'Affecter
mvarCurrentFidelityIndexOfMeasure = iData

End Property

Public Property Get CurrentFidelityIndexOfPoint() As Integer
'Accès en lecture sur la point de fidélité courant

'Renvoyer
CurrentFidelityIndexOfPoint = mvarCurrentFidelityIndexOfPoint

End Property

Public Property Let CurrentFidelityIndexOfPoint(ByVal iData As Integer)
'Accès en écriture sur la point de fidélité courant

'Affecter
mvarCurrentFidelityIndexOfPoint = iData

End Property

Public Property Get ControlIsEnd() As Boolean
'Accès en lecture sur l'état du contrôle

'Renvoyer
ControlIsEnd = mvarControlIsEnd

End Property

Public Property Get FidelityTestIsStart() As Boolean
'Accès en lecture sur l'état du contrôle

'Renvoyer
FidelityTestIsStart = mvarFidelityTestIsStart

End Property

Public Property Get ControlComment() As String
'Accès en lecture sur le commentaire

'Renvoyer
ControlComment = mvarComment

End Property

Public Property Let ControlComment(ByVal vData As String)
'Accès en écriture sur le commentaire

'Affecter
mvarComment = vData

End Property

Public Property Get SerialValuesAreOkToStart() As Boolean
'Accès en lecture sur l'état des valeurs de série de mesures

Dim currentSValue As SerialValue

Dim bZeroAlreadyLook As Boolean
Dim cLastValue As Currency
Dim cCurrentValue As Currency

'Vérifier qu'au moins une valeur est saisie
If mvarSerialValues.Count = 0 Then Exit Property

'Vérifier que la première valeur est 0
If mvarSerialValues(1) <> 0 Then Exit Property

'Initialiser sur une valeur impossible
cLastValue = -1000

'Vérifier que toutes les valeurs de la séries sont saisies
For Each currentSValue In mvarSerialValues
    
    'Stocker
    cCurrentValue = currentSValue.Value
    
    'Vérifier que des valeurs nulles ne se suivent pas
    If bZeroAlreadyLook And cCurrentValue = 0 Then Exit Property
    If cCurrentValue = 0 Then bZeroAlreadyLook = True
        
    'Vérifier que 2 valeurs ne se suivent pas
    If cCurrentValue = cLastValue Then Exit Property
    cLastValue = cCurrentValue
    
Next

'Renvoyer
SerialValuesAreOkToStart = True

End Property

Public Property Let IsOnCorrectionMode(ByVal bData As Boolean)
'Accès en écriture sur l'état du contrôle

'Si modification
If bData <> mvarCorrectionMode Then

    'Affecter
    mvarCorrectionMode = bData
    
    'Si mode normal
    If Not mvarCorrectionMode Then
        mvarCurrentIndexOfSequence = mvarNormalIndexOfSequence
        mvarCurrentIndexOfSerialValue = mvarNormalIndexOfSerialValue
        mvarCurrentSequenceIsUp = mvarNormalSequenceIsUp
    End If

    'Evènements
    RaiseEvent CorrectionModeChange
    
End If

End Property

Public Property Get IsOnCorrectionMode() As Boolean
'Accès en lecture sur l'état du contrôle

'Renvoyer
IsOnCorrectionMode = mvarCorrectionMode

End Property

Public Property Get SequencesAreEnd() As Boolean
'Accès en lecture sur l'état du Contrôle

'Renvoyer
SequencesAreEnd = mvarSequencesAreEnd

End Property

Public Property Get NormalSequenceIsUp() As Boolean
'Accès en lecture sur le sens norm de course

'Renvoyer
NormalSequenceIsUp = mvarNormalSequenceIsUp

End Property

Public Property Get NormalIndexOfSequence() As Integer
'Accès en lecture sur la séquence normale en cours

'Renvoyer
NormalIndexOfSequence = mvarNormalIndexOfSequence

End Property

Public Property Get NormalIndexOfSerialValue() As Integer
'Accès en lecture sur la mesure normale en cours

'Renvoyer
NormalIndexOfSerialValue = mvarNormalIndexOfSerialValue

End Property

Public Property Let CurrentSequenceIsUp(ByVal bData As Boolean)
'Accès en écriture sur le sens de la séquence en cours

'Affecter
mvarCurrentSequenceIsUp = bData

End Property

Public Property Get CurrentSequenceIsUp() As Boolean
'Accès en lecture sur le sens de la séquence en cours

'Renvoyer
CurrentSequenceIsUp = mvarCurrentSequenceIsUp

End Property

Public Property Let CurrentIndexOfSerialValue(ByVal iData As Integer)
'Accès en écriture sur la mesure en cours

'Affecter
mvarCurrentIndexOfSerialValue = iData

End Property

Public Property Get CurrentIndexOfSerialValue() As Integer
'Accès en lecture sur la mesure en cours

'Renvoyer
CurrentIndexOfSerialValue = mvarCurrentIndexOfSerialValue

End Property

Public Property Let CurrentIndexOfSequence(ByVal iData As Integer)
'Accès en écriture sur la séquence en cours

'Affecter
mvarCurrentIndexOfSequence = iData

End Property

Public Property Get CurrentIndexOfSequence() As Integer
'Accès en lecture sur la séquence en cours

'Renvoyer
CurrentIndexOfSequence = mvarCurrentIndexOfSequence

End Property

Public Sub InitializeData()
'Initialiser le contrôle (Ne modifie pas la matrice de données)

Dim currentSequence As Sequence

Dim i As Integer
Dim j As Integer

'Vider
mvarUpSequences.Clear
mvarDownSequences.Clear

'Initialiser
mvarControlIsStart = False
mvarSequencesAreEnd = False
mvarCorrectionMode = False
mvarFidelityTestIsStart = False
mvarControlIsEnd = False

'Libérer
Set mvarLevels = Nothing

'Mesures en série
mvarNormalIndexOfSequence = 1
mvarNormalIndexOfSerialValue = 1
mvarNormalSequenceIsUp = True

mvarCurrentIndexOfSequence = 1
mvarCurrentIndexOfSerialValue = 1
mvarCurrentSequenceIsUp = True

'Ecart de fidélité
mvarNormalFidelityIndexOfMeasure = 1
mvarNormalFidelityIndexOfPoint = 1

mvarCurrentFidelityIndexOfMeasure = 1
mvarCurrentFidelityIndexOfPoint = 1

'Propriétés du contrôle
mvarComment = ""
mvarDateOfControl = Now
'mvarEquipmentNumber = ""
'mvarEquipmentReference = ""
'If mvarHumidity = 0 Then mvarHumidity = 60
'mvarTemperature = 0

'Créer les objets
For i = 1 To mvarSequencesCount
    
    'Créer la séquence montante
    Set currentSequence = mvarUpSequences.Add()
    
    'Créer la collection de Mesures
    For j = 1 To mvarSerialValues.Count
        currentSequence.Measures.Add 0
    Next
    
    'Créer la séquence descendante
    Set currentSequence = mvarDownSequences.Add()
    
    'Créer la collection de Mesures
    For j = 1 To mvarSerialValues.Count
        currentSequence.Measures.Add 0
    Next
    
Next

'Libérer
Set currentSequence = Nothing

End Sub

Public Property Get ControlIsStart() As Boolean
'Accès en lecture sur l'état du contrôle

'Renvoyer
ControlIsStart = mvarControlIsStart

End Property

Public Sub SaveFidelityPointMeasure(sMeasure As Currency)
'Sauvegarder une mesure d'écart de fidélité

Dim currentFidelityPoint As FidelityPoint

'Si première sauvegarde
If Not mvarFidelityTestIsStart Then

    'Positionner
    mvarFidelityTestIsStart = True
            
    'Retour au début
    mvarNormalFidelityIndexOfMeasure = 1
    mvarNormalFidelityIndexOfPoint = 1
    mvarCurrentFidelityIndexOfMeasure = 1
    mvarCurrentFidelityIndexOfPoint = 1
    
    'Générer l'évènement
    RaiseEvent FidelityPointsTestStart

End If

'Affecter
Set currentFidelityPoint = mvarFidelityPoints(mvarCurrentFidelityIndexOfPoint)

'Sauvegarder la mesure
With currentFidelityPoint.Measures(mvarCurrentFidelityIndexOfMeasure)
    .Enabled = True
    .Value = sMeasure
End With

'Avant calcul du nouveau pas
RaiseEvent AfterFidelityPointSaveWithoutStep

'Si mode correction désactivé
If Not mvarCorrectionModeOfFidelityTest Then

    'Positionner à la prochaine mesure
    mvarNormalFidelityIndexOfMeasure = mvarNormalFidelityIndexOfMeasure + 1
    
    If mvarNormalFidelityIndexOfMeasure > currentFidelityPoint.Measures.Count Then
    
        mvarNormalFidelityIndexOfMeasure = 1
        mvarNormalFidelityIndexOfPoint = mvarNormalFidelityIndexOfPoint + 1
        
        'Si phase terminée
        If mvarNormalFidelityIndexOfPoint > mvarFidelityPoints.Count Then
            
            'Retour au début
            mvarNormalFidelityIndexOfMeasure = 1
            mvarNormalFidelityIndexOfPoint = 1
            mvarCurrentFidelityIndexOfMeasure = 1
            mvarCurrentFidelityIndexOfPoint = 1
            
            'Générer l'évènement
            mvarFidelityTestIsStart = False
            mvarFidelityTestIsEnd = True
    
            RaiseEvent FidelityPointsTestTerminated
            
        End If
        
    End If

Else

    'Mode correction activé - Retour à la normal
    mvarCorrectionModeOfFidelityTest = False
    RaiseEvent CorrectionModeOfFidelityTestChange
    
End If

'Libérer
Set currentFidelityPoint = Nothing

'Affecter
mvarCurrentFidelityIndexOfMeasure = mvarNormalFidelityIndexOfMeasure
mvarCurrentFidelityIndexOfPoint = mvarNormalFidelityIndexOfPoint

'Après la sauvegarde
RaiseEvent AfterFidelityPointSave

End Sub

Public Sub SetControlIsEnd()
'Signaler que le contrôle est terminé

'Le contrôle est terminé dans sa globalité
mvarControlIsEnd = True

'Evènement
RaiseEvent ControlIsEnd

End Sub
Public Sub SaveMeasure(sMeasure As Currency)
'Sauvegarder la mesure en cours pour les mesures en série

Dim currentSequences As Sequences
Dim currentMeasure As Measure

'Pointer sur la série de séquences correspondante
If mvarCurrentSequenceIsUp Then
    Set currentSequences = mvarUpSequences
Else
    Set currentSequences = mvarDownSequences
End If

'Pointer sur la mesure
Set currentMeasure = currentSequences(mvarCurrentIndexOfSequence).Measures(mvarCurrentIndexOfSerialValue)

'Sauvegarder
With currentMeasure
    .Enabled = True
    .Value = sMeasure
End With

'Avertir que la position va changer (Ou revenir à la position normale)
RaiseEvent LeavePosition

'Si le mode "Correction" n'est pas activé
If Not mvarCorrectionMode Then

    'Modifier la position selon le sens de la séquence
    If mvarNormalSequenceIsUp Then
        
        'Passer à la prochaine mesure
        mvarNormalIndexOfSerialValue = mvarNormalIndexOfSerialValue + 1
        
        'Si la séquence est terminée
        If mvarNormalIndexOfSerialValue > mvarSerialValues.Count Then
            
            'La séquence est descendante
            mvarNormalSequenceIsUp = False
            
            'La mesure passe à la dernière
            mvarNormalIndexOfSerialValue = mvarSerialValues.Count
            
        End If
        
    Else
    
        'Passer à la prochaine mesure
        mvarNormalIndexOfSerialValue = mvarNormalIndexOfSerialValue - 1
        
        'Si la séquence est terminée
        If mvarNormalIndexOfSerialValue = 0 Then
            
            'La séquence est montante
            mvarNormalSequenceIsUp = True
            
            'On recommence une nouvelle séquence
            mvarNormalIndexOfSerialValue = 1
            
            'On commence une nouvelle séquence
            mvarNormalIndexOfSequence = mvarNormalIndexOfSequence + 1
            
        End If
        
    End If
    
Else

    'Le mode "Correction" est activé, on réinitialise
    mvarCorrectionMode = False
    
    'Evènement
    RaiseEvent CorrectionModeChange
    
End If

'Copier les valeurs
mvarCurrentIndexOfSequence = mvarNormalIndexOfSequence
mvarCurrentIndexOfSerialValue = mvarNormalIndexOfSerialValue
mvarCurrentSequenceIsUp = mvarNormalSequenceIsUp

'Si le contrôle est terminé
If mvarCurrentIndexOfSequence > mvarSequencesCount Then
    
    'Enregistrer
    mvarSequencesAreEnd = True
    
    'Repositionner au début
    mvarNormalIndexOfSequence = 1
    mvarNormalIndexOfSerialValue = 1
    mvarNormalSequenceIsUp = True
    
    mvarCurrentIndexOfSequence = mvarNormalIndexOfSequence
    mvarCurrentIndexOfSerialValue = mvarNormalIndexOfSerialValue
    mvarCurrentSequenceIsUp = mvarNormalSequenceIsUp
    
    'Evènement
    RaiseEvent SerialValueTestTerminated
    
End If

'Avertir que la position a changée
RaiseEvent PositionChange

'Libérer
Set currentSequences = Nothing

End Sub

Public Property Let ControlIsStart(ByVal vData As Boolean)
'Accès en écriture sur l'état du contrôle

'Affecter
mvarControlIsStart = vData

End Property

Public Property Get SequencesCount() As Integer
'Accès en lecture sur le nombre de séries de mesure à effectuer

'Renvoyer
SequencesCount = mvarSequencesCount

End Property

Public Property Let SequencesCount(ByVal vData As Integer)
'Accès en écriture sur le nombre de séries de mesures à effectuer

'Affecter si le contrôle n'est pas commencé
If Not mvarControlIsStart Then

    'Affecter
    mvarSequencesCount = vData
    
End If

End Property

Public Property Set DownSequences(ByVal vData As Sequences)
'Accès en écriture sur la collection de séquences descendantes

'Affecter
Set mvarDownSequences = vData

End Property

Public Property Get DownSequences() As Sequences
'Accès en lecture sur la collection de séquences descendantes
    
'Renvoyer
Set DownSequences = mvarDownSequences

End Property

Public Property Set UpSequences(ByVal vData As Sequences)
'Accès en écriture sur la collection de séquences montantes

'Affecter
Set mvarUpSequences = vData

End Property

Public Property Get UpSequences() As Sequences
'Accès en lecture sur la collection de séquences montantes

'Affecter
Set UpSequences = mvarUpSequences

End Property

Public Property Let EquipmentReference(ByVal vData As String)
'Accès en écriture sur la référence de l'équipement

'Affecter
mvarEquipmentReference = vData

End Property

Public Property Get EquipmentReference() As String
'Accès en lecture sur la référence de l'équipement

'Renvoyer
EquipmentReference = mvarEquipmentReference

End Property

Public Property Let EquipmentManufacturer(ByVal vData As String)
'Accès en écriture sur le constructeur de l'équipement

'Affecter
mvarEquipmentManufacturer = vData

End Property

Public Property Get EquipmentManufacturer() As String
'Accès en lecture sur le constructeur de l'équipement

'Renvoyer
EquipmentManufacturer = mvarEquipmentManufacturer

End Property

Public Property Let EquipmentNumber(ByVal vData As String)
'Accès en écriture sur le numéro de référence de l'équipement

'Affecter
mvarEquipmentNumber = vData

End Property

Public Property Get EquipmentNumber() As String
'Accès en lecture sur le numéro de référence de l'équipement

'Renvoyer
EquipmentNumber = mvarEquipmentNumber

End Property

Public Property Let OperatorName(ByVal vData As String)
'Accès en écriture sur le nom de l'opérateur

'Affecter
mvarOperatorName = vData

End Property

Public Property Get OperatorName() As String
'Accès en lecture sur le nom de l'opérateur

'Renvoyer
OperatorName = mvarOperatorName

End Property

Public Property Let Humidity(ByVal vData As Integer)
'Accès en écriture sur le taux d'humidité

'Affecter
mvarHumidity = vData

End Property

Public Property Get Humidity() As Integer
'Accès en lecture sur le taux d'humidité

'Renvoyer
Humidity = mvarHumidity

End Property

Public Property Let Temperature(ByVal vData As Currency)
'Accès en écriture sur la température

'Affecter
mvarTemperature = vData

End Property

Public Property Get Temperature() As Currency
'Accès en lecture sur la température

'Renvoyer
Temperature = mvarTemperature

End Property

Public Property Let Detenteur(ByVal vData As String)
'Accès en écriture sur le détenteur

'Affecter
mvarDetenteur = vData

End Property

Public Property Get Detenteur() As String
'Accès en lecture sur le détenteur

'Renvoyer
Detenteur = mvarDetenteur

End Property


Public Property Let DateOfControl(ByVal vData As Date)
'Accès en écriture sur la date du controle

'Affecter
mvarDateOfControl = vData

End Property

Public Property Get DateOfControl() As Date
'Accès en lecture sur la date du contrôle

'Renvoyer
DateOfControl = mvarDateOfControl

End Property

Public Property Get SerialValues() As SerialValues
'Accès en lecture sur la collection de valeurs de la série

'Renvoyer
Set SerialValues = mvarSerialValues

End Property

Public Property Set SerialValues(vData As SerialValues)
'Accès en écriture sur la collection de valeurs de la série

'Affecter
Set mvarSerialValues = vData

End Property

Private Sub Class_Initialize()
'Constructeur

'Instancier
Set mvarSerialValues = New SerialValues
Set mvarUpSequences = New Sequences
Set mvarDownSequences = New Sequences
Set mvarFidelityPoints = New FidelityPoints

'Initialiser
mvarFidelityPoints.MeasuresCount = 5
InitializeData

End Sub

Private Sub Class_Terminate()
'Destructeur

'Libérer
Set mvarSerialValues = Nothing
Set mvarUpSequences = Nothing
Set mvarDownSequences = Nothing
Set mvarFidelityPoints = Nothing
Set mvarLevels = Nothing

End Sub

Public Property Let FidelityPoints(ByVal objData As FidelityPoints)
'Accès en écriture sur la collection de points de relevé d'écart de fidélité

'Affecter
Set mvarFidelityPoints = objData

End Property

Public Property Get FidelityPoints() As FidelityPoints
'Accès en lecture sur la collection de points de relevé d'écart de fidélité

'Renvoyer
Set FidelityPoints = mvarFidelityPoints

End Property

Public Property Get HysteresisError() As Currency
'Accès en lecture sur l'erreur d'Hystérésis
'Plus grande différence algébrique des ordonnées d'un même point sur les 2 courses / 2

Dim i As Integer

Dim cMaxValue As Currency
Dim cCurrentValue As Currency

'Sécurité
If mvarSerialValues.Count = 0 Then Exit Property

'Affecter une valeur par défaut
cMaxValue = Abs(mvarUpSequences.GetAVerage(1) - mvarDownSequences.GetAVerage(1))

'Stocker la valeur maxi
For i = 1 To mvarSerialValues.Count
    
    'Affecter
    cCurrentValue = Abs(mvarUpSequences.GetAVerage(i) - mvarDownSequences.GetAVerage(i))
    If cCurrentValue > cMaxValue Then cMaxValue = cCurrentValue
    
Next

'Renvoyer
HysteresisError = cMaxValue '/ 2

End Property

Public Property Get FidelityError() As Currency
'Accès en lecture sur l'erreur de fidélité (Valeur maximale)

Dim currentFidelityPoint As FidelityPoint

Dim cValMax As Currency
Dim cCurrent As Currency

'Enumérer les points de fidélités
For Each currentFidelityPoint In mvarFidelityPoints
    
    'Affecter
    cCurrent = currentFidelityPoint.FidelityError
    
    'Affecter si valeur max
    If cCurrent > cValMax Then cValMax = cCurrent
    
Next

'Libérer
Set currentFidelityPoint = Nothing

'Affecter
FidelityError = cValMax

End Property

Public Property Get ExactnessLocalErrorOld() As Currency
'Accès en lecture sur l'erreur de justesse locale
'Plus grande différence algébrique des ordonnées de 2 points successifs de la courbe

Dim currentSequences As Sequences
Dim cValue As Currency
Dim cLastValue As Currency

Dim s As Integer
Dim i As Integer
Dim iSValueCount As Integer

'Initialiser
iSValueCount = mvarSerialValues.Count
If iSValueCount < 2 Then Exit Property

'Stocker un résultat par défaut
cValue = Abs(mvarUpSequences.GetAVerage(1) - mvarUpSequences.GetAVerage(2))

'Enumérer les séquences
For s = 1 To 2

    'Affecter
    If s = 1 Then
        Set currentSequences = mvarUpSequences
    Else
        Set currentSequences = mvarDownSequences
    End If
    
    'Enumérer les valeurs de séries de mesure
    For i = 1 To iSValueCount - 1
    
        'Affecter la dernière valeur
        cLastValue = Abs(currentSequences.GetAVerage(i) - currentSequences.GetAVerage(i + 1))
        
        'Stocker si condition OK
        If cLastValue > cValue Then cValue = cLastValue
        
    Next
    
Next

'Libérer
Set currentSequences = Nothing

'Affecter
ExactnessLocalErrorOld = cValue

End Property

Public Property Get ExactnessLocalError() As Currency
'Accès en lecture sur l'erreur de justesse locale
'Plus grande différence algébrique des ordonnées de 2 points successifs de la courbe

Dim cValue As Currency
Dim cLastValue As Currency

Dim s As Integer
Dim i As Integer
Dim iSValueCount As Integer

'Initialiser
iSValueCount = mvarSerialValues.Count
If iSValueCount < 2 Then Exit Property

'Stocker un résultat par défaut
cValue = Abs(mvarUpSequences.GetAVerage(1) - mvarUpSequences.GetAVerage(2))
'Enumérer les valeurs de séries de mesure
For i = 1 To iSValueCount - 1
    
'Affecter la dernière valeur
    cLastValue = Abs(mvarUpSequences.GetAVerage(i) - mvarUpSequences.GetAVerage(i + 1))
    'Stocker si condition OK
    If cLastValue > cValue Then cValue = cLastValue
    cLastValue = Abs(mvarUpSequences.GetAVerage(i) - mvarDownSequences.GetAVerage(i + 1))
    'Stocker si condition OK
    If cLastValue > cValue Then cValue = cLastValue
    cLastValue = Abs(mvarDownSequences.GetAVerage(i) - mvarUpSequences.GetAVerage(i + 1))
    'Stocker si condition OK
    If cLastValue > cValue Then cValue = cLastValue
    cLastValue = Abs(mvarDownSequences.GetAVerage(i) - mvarDownSequences.GetAVerage(i + 1))
    'Stocker si condition OK
    If cLastValue > cValue Then cValue = cLastValue
Next
'Affecter
ExactnessLocalError = cValue

End Property

Public Property Get ExactnessTotalError() As Currency
'Accès en lecture sur l'erreur de justesse totale
'Différence algébrique des ordonnées maximale et minimal des courbes

ExactnessTotalError = GetMaxAverageOfAllSequences - GetMinAverageOfAllSequences

End Property

Public Function GetAVerage(IndexOfSerialMeasure As Integer, isUpSequence As Boolean) As Currency
'Renvoyer une moyenne d'écart obtenu pour une valeur de la série de mesures
'en phase montante ou descendante (La valeur est renvoyée en microns)


Dim localSequences As Sequences
Dim currentSequence As Sequence
Dim currentMeasure As Measure

Dim iNumber As Integer

Dim sSerialValue As Currency
Dim sTotal As Currency

'Affecter la séquence concernée
If mvarCurrentSequenceIsUp Then
    Set localSequences = mvarUpSequences
Else
    Set localSequences = mvarDownSequences
End If

'Stocker la valeur de la mesure dans la série
sSerialValue = mvarSerialValues(IndexOfSerialMeasure)

'Enumérer les séquences
For Each currentSequence In localSequences
    
    'Pointer sur la mesure correspondante
    Set currentMeasure = currentSequence.Measures(IndexOfSerialMeasure)
    
    'Si la mesure est validée
    If currentMeasure.Enabled Then
        
        sTotal = sTotal + ((sSerialValue - currentMeasure.Value) * 1000)
        iNumber = iNumber + 1
        
    End If
    
Next

'Renvoyer
If iNumber > 0 Then GetAVerage = Round(sTotal / iNumber, 2)

'Enregistrer
localSequences.SetAverage IndexOfSerialMeasure, GetAVerage

'Libérer
Set localSequences = Nothing
Set currentMeasure = Nothing
Set currentSequence = Nothing

End Function

Public Sub PrepareForANewControl()
'Méthode pour préparer à un nouveau contrôle

Dim varSerial As Variant
Dim i As Integer

'Travail sur la série de mesures
mvarSerialValues.Clear
varSerial = Array(0, 0.9, 2.1, 3, 4.1, 4.8, 6, 7.1, 7.8, 9, 10)
For i = LBound(varSerial) To UBound(varSerial)
    mvarSerialValues.Add varSerial(i)
Next

Set varSerial = Nothing

'Travail sur le nombre de séries
mvarSequencesCount = 2

'Initialiser
mvarFidelityPoints.Clear
mvarDownSequences.Clear
mvarUpSequences.Clear
InitializeData

End Sub

Public Function GetFidelityMicronsDifference(ByVal indexOfPoint As Integer, ByVal indexOfMeasure As Integer) As Currency
'Renvoie le calcul d'un écrat pour un point de fidélité

Dim currentFidelityPoint As FidelityPoint
Dim currentMeasure As Measure

'Affecter
Set currentFidelityPoint = mvarFidelityPoints(indexOfPoint)
Set currentMeasure = currentFidelityPoint.Measures(indexOfMeasure)

'Si mesure validée, renvoyer l'écart
If currentMeasure.Enabled Then GetFidelityMicronsDifference = (currentFidelityPoint.SerialValue - currentMeasure.Value) * 1000

'Libérer
Set currentMeasure = Nothing
Set currentFidelityPoint = Nothing

End Function

Public Function GetMicronsDifference(ByVal indexOfSerialValue As Integer, ByVal IndexOfSequence As Integer, _
ByVal SequenceIsUp As Boolean) As Currency
'Renvoie le calcul d'un écart

Dim currentSequences As Sequences
Dim currentMeasure As Measure

'Affecter la collection de séquence
If SequenceIsUp Then
    Set currentSequences = mvarUpSequences
Else
    Set currentSequences = mvarDownSequences
End If

'Affecter la mesure
Set currentMeasure = currentSequences(IndexOfSequence).Measures(indexOfSerialValue)

'Renvoyer l'écart
If currentMeasure.Enabled Then GetMicronsDifference = (mvarSerialValues(indexOfSerialValue) - currentMeasure.Value) * 1000

'Libérer
Set currentSequences = Nothing
Set currentMeasure = Nothing

End Function

Public Function GetMinAverageOfAllSequences() As Currency
'Renvoie la moyenne d'écart la plus basse des 2 collections de séquences

Dim i As Integer
Dim iCount As Integer

Dim cLastValue As Currency

'Initialisation
iCount = mvarSerialValues.Count
GetMinAverageOfAllSequences = mvarUpSequences.GetAVerage(iCount)

'Enumérer les moyennes de séquences montantes
For i = 1 To iCount
    cLastValue = mvarUpSequences.GetAVerage(i)
    If cLastValue < GetMinAverageOfAllSequences Then GetMinAverageOfAllSequences = cLastValue
Next

'Enumérer les moyennes de séquences descendantes
For i = 1 To iCount
    cLastValue = mvarDownSequences.GetAVerage(i)
    If cLastValue < GetMinAverageOfAllSequences Then GetMinAverageOfAllSequences = cLastValue
Next

End Function

Public Function GetMaxValueOfSerialValues() As Currency
'Renvoie la valeur maximale de la série de mesure

Dim currentSValue As SerialValue

'Sécurité
If mvarSerialValues.Count = 0 Then Exit Function

'Affecter par défaut
GetMaxValueOfSerialValues = mvarSerialValues(mvarSerialValues.Count).Value

'Enumérer
For Each currentSValue In mvarSerialValues
    
    'Affecter
    If currentSValue.Value > GetMaxValueOfSerialValues Then GetMaxValueOfSerialValues = currentSValue.Value
    
Next

'Libérer
Set currentSValue = Nothing

End Function

Public Function GetMaxAverageOfAllSequences() As Currency
'Renvoie la moyenne d'écart la plus haute des 2 collections de séquences

Dim i As Integer
Dim iCount As Integer

Dim cLastValue As Currency

'Initialisation
iCount = mvarSerialValues.Count
GetMaxAverageOfAllSequences = mvarUpSequences.GetAVerage(iCount)

'Enumérer les moyennes de séquences montantes
For i = 1 To iCount
    cLastValue = mvarUpSequences.GetAVerage(i)
    If cLastValue > GetMaxAverageOfAllSequences Then GetMaxAverageOfAllSequences = cLastValue
Next

'Enumérer les moyennes de séquences descendantes
For i = 1 To iCount
    cLastValue = mvarDownSequences.GetAVerage(i)
    If cLastValue > GetMaxAverageOfAllSequences Then GetMaxAverageOfAllSequences = cLastValue
Next

End Function

Private Sub PrintFidelityResult(objTarget As Object, objConfig As Configuration)
'Impression du résultat de relevé d'écarts de fidélité

Dim currentFidelityPoint As FidelityPoint
Dim strArray As Variant

Dim iPoint  As Integer
Dim iMeasure As Integer
Dim iCount As Integer

Dim cValue As Currency

'Titre
strArray = Array("Relevé des écarts de fidélité en " + CHAR_MICRON + "mMB")
PrintArrayOfString strArray, objTarget, objConfig

'Initialisation
iCount = mvarFidelityPoints.MeasuresCount
ReDim strArray(mvarFidelityPoints.MeasuresCount + 1)

'Enumérer les points de fidélité
For iPoint = 1 To mvarFidelityPoints.Count
    
    'Affecter
    Set currentFidelityPoint = mvarFidelityPoints(iPoint)
    
    'Construire le tableau de données
    strArray(0) = Format(currentFidelityPoint.SerialValue, "0.00") + IIf(currentFidelityPoint.IsUpDirection, " M", " D") + "LB"

    For iMeasure = 1 To iCount
    
        'Affecter la valeur de différence
        cValue = GetFidelityMicronsDifference(iPoint, iMeasure)
        
        'Affecter la chaine
        strArray(iMeasure) = IIf(cValue < 0, Format(cValue), "+" + Format(cValue)) + "MN"
        
    Next
    strArray(iMeasure) = "Moy: " + Format(currentFidelityPoint.GetAVerage, "0.00") + "LR"
        
    'Imprimer
    PrintArrayOfString strArray, objTarget, objConfig
    
Next

'Libérer
Set currentFidelityPoint = Nothing

End Sub

Public Sub PrintOn(objTarget As Object, objConfiguration As Configuration)
'Imprimer dans la cible spécifiée

'********************************************************************************
'Variables
'********************************************************************************

Dim i As Integer
Dim strArray As Variant

'********************************************************************************
'CODE
'********************************************************************************

'Initialisation graphique
Screen.MousePointer = vbHourglass
If Not (objTarget Is Printer) Then objTarget.Visible = False

'Initialisation de l'imprimante
With Printer
    .TrackDefault = True
    .ScaleMode = vbMillimeters
    .PaperSize = vbPRPSA4
    .Orientation = vbPRORPortrait
End With

'********************************************************************************
'Titre
'********************************************************************************
With objTarget
    .Font.Name = objConfiguration.Print_Font
    .CurrentY = ((297 - .ScaleHeight) / 2) + VERTICAL_MARGIN
    .FontBold = True
    .FontSize = 16
    .FontUnderline = True
    .CurrentX = (210 - .TextWidth("PROCES VERBAL DE CONTROLE")) / 2
    .DrawWidth = 1
End With

objTarget.Print "PROCES VERBAL DE CONTROLE"
objTarget.FontUnderline = False
objTarget.Print

'********************************************************************************
'Corps de texte
'********************************************************************************
strArray = Array("2ème RMAT BRUZLB", "Atelier: " + objConfiguration.Print_Service + "LL", _
"Date: " + Format(mvarDateOfControl, "dd/MM/yy") + "RR")
PrintArrayOfString strArray, objTarget, objConfiguration

strArray = Array("Désignation: " + objConfiguration.Print_Designation + "LL")
PrintArrayOfString strArray, objTarget, objConfiguration

strArray = Array("Position: " + objConfiguration.Print_Position + "LL")
PrintArrayOfString strArray, objTarget, objConfiguration

strArray = Array("Document de référence: " + objConfiguration.Print_DocReference + "LL")
PrintArrayOfString strArray, objTarget, objConfiguration

strArray = Array("Outillage: " + objConfiguration.Print_Tool + "LL")
PrintArrayOfString strArray, objTarget, objConfiguration

strArray = Array("Opérateur: " + mvarOperatorName + "LL", "Température:" + Str(mvarTemperature) + " °CLL", _
"Hygrométrie:" + Str(mvarHumidity) + " % HRLL")
PrintArrayOfString strArray, objTarget, objConfiguration
objTarget.Print

strArray = Array("Comparateur n°: " + mvarEquipmentNumber + "LR", "Référence: " + mvarEquipmentReference + "LR", _
"Fabriquant: " + mvarEquipmentManufacturer + "LR")
PrintArrayOfString strArray, objTarget, objConfiguration
objTarget.Print

'********************************************************************************
'Relevé d'écarts de fidélité
'********************************************************************************
PrintFidelityResult objTarget, objConfiguration
objTarget.Print

'********************************************************************************
'Résultat
'********************************************************************************
PrintResult objTarget, objConfiguration

'********************************************************************************
'Graphique
'********************************************************************************
objTarget.Print
strArray = Array("Courbe d'étalonnageMB")
PrintArrayOfString strArray, objTarget, objConfiguration
objTarget.Print
PrintGraph objTarget

'Mise en forme
With objTarget
    .FontBold = False
    .FontUnderline = False
    .FontSize = 12
End With
objTarget.Print

'********************************************************************************
'Commentaire
'********************************************************************************
mvarComment = Trim(mvarComment)
If mvarComment <> "" Then
    i = objConfiguration.Print_Lateral_Margin
    If Not (objTarget Is Printer) Then i = i + ((210 - Printer.ScaleWidth) / 2)
    PrintComment objTarget, i
End If

'********************************************************************************
'Fin
'********************************************************************************
strArray = Array("Page 1/1")
With objTarget
    .CurrentX = (210 - .TextWidth(strArray(0))) / 2
    .CurrentY = .ScaleHeight - VERTICAL_MARGIN - .TextHeight(strArray(0))
End With
objTarget.Print strArray(0)

'Terminer la tâche si l'objet est une imprimante
If objTarget Is Printer Then objTarget.EndDoc

'Libérer
Set strArray = Nothing

'Restauration graphique
If Not (objTarget Is Printer) Then objTarget.Visible = True
Screen.MousePointer = vbNormal

End Sub

Private Sub PrintArrayOfString(strArray As Variant, objTarget As Object, objConfig As Configuration)
'Imprime un tableau de chaine avec un encadrement
'Symbologie:
'L: Aligné à gauche
'M: Centré
'R: Aligné à droite

'N: Normal
'B: Mettre en gras
'L: Mettre en gras à gauche du ":"
'R: Mettre en gras à droite du ":"

'Chaine
Dim RealString As String
Dim strAlign As String * 1
Dim strFont As String * 1

'Tableau des tailles de colonnes
Dim iNormalWidth As Integer
Dim varData() As PrintData
Dim iLBound As Integer
Dim iUBound As Integer

'Compteur
Dim i As Integer

'Cotes de la cible
Dim iLeft As Integer
Dim iRight As Integer
Dim iWidth As Integer

'Impression
Dim iCurrentX As Integer
Dim iYSave As Integer
Dim iXStartOfCol As Integer
Dim iYBigest As Integer

'Calcul des espacements
Dim iTotalSize As Integer
Dim iSize As Currency

'********************************************************************************
'Dimensionnement dynamique
'********************************************************************************
'Initialisation
iLBound = LBound(strArray)
iUBound = UBound(strArray)
ReDim varData(iUBound - iLBound)

'********************************************************************************
'Calcul des tailles de colonnes
'********************************************************************************
'Travail sur l'objet cible
With objTarget
    
    'Affectation des variables
    iWidth = Printer.ScaleWidth - (2 * objConfig.Print_Lateral_Margin) - 1
    iLeft = objConfig.Print_Lateral_Margin
    If Not (objTarget Is Printer) Then iLeft = iLeft + ((210 - Printer.ScaleWidth) / 2)
    iRight = iLeft + iWidth
    
    'Condition initiale
    .FontSize = 12
    .FontBold = True
    .FontUnderline = False
    iYBigest = .TextHeight("|")
    
    'Calcul de l'espacement de chaque chaine
    For i = iLBound To iUBound
            
        'Mise en forme
        strArray(i) = Trim(strArray(i))
        
        'Sécurité
        If strArray(i) <> "" Then
        
            'Affectation
            RealString = strArray(i)
            strFont = Right(RealString, 1)
            RealString = Left(RealString, Len(RealString) - 2)
                                
            'Action selon les flags
            Select Case strFont
                
                'Chaine en gras
                Case "B"
                    
                    'Calcul
                    .FontBold = True
                    iSize = .TextWidth(RealString)
                    
                'Chaine en gras à gauche du ":"
                Case "L"
                
                    'Calcul
                    .FontBold = True
                    iSize = .TextWidth(RealString)
                    iSize = .TextWidth(Left(RealString, InStr(1, RealString, ":")))
                    .FontBold = False
                    iSize = iSize + .TextWidth(Right(RealString, Len(RealString) - InStr(1, RealString, ":")))

                'Chaine en gras à droite du ":"
                Case "R"
                
                    'Calcul
                    .FontBold = False
                    iSize = .TextWidth(Left(RealString, InStr(1, RealString, ":")))
                    .FontBold = True
                    iSize = iSize + .TextWidth(Right(RealString, Len(RealString) - InStr(1, RealString, ":")))
                    
                'Autre cas
                Case Else
                
                    'Calcul
                    .FontBold = False
                    iSize = .TextWidth(RealString)
                    
            End Select
            
        Else
            
            'La chaine est vide
            iSize = 0
            
        End If
        
        'Enregistrer la taille
        varData(i).TextWidth = iSize

        'Ajouter
        iTotalSize = iTotalSize + iSize
        
    Next

End With

'Sécurité
If iTotalSize = 0 Then

    'Libérer
    Erase varData
    
    'Sortir
    Exit Sub
    
End If

'Calculer la taille moyenne par défaut
If iLBound = 0 Then
    iNormalWidth = iWidth / (iUBound - iLBound + 1)
Else
    iNormalWidth = iWidth / (iUBound - iLBound)
End If
iNormalWidth = iNormalWidth - 1

'Calculer la nouvelle taille moyenne à partager
iTotalSize = iWidth
iSize = 0

'Enumérer les valeurs de colonnes
For i = iLBound To iUBound

    'Si taille supérieure à la moyenne
    If varData(i).TextWidth > iNormalWidth Then
        varData(i).ColWidth = varData(i).TextWidth
        iTotalSize = iTotalSize - varData(i).ColWidth
        iSize = iSize + 1
    End If
    
Next

'Si au moins une taille est supérieure à la moyenne
If iSize > 1 Then iNormalWidth = iTotalSize / iSize

'Appliquer les tailles
iTotalSize = 0
For i = iLBound To iUBound - 1
    
    'Appliquer la taille par défaut
    If varData(i).ColWidth < iNormalWidth Then varData(i).ColWidth = iNormalWidth
    iTotalSize = iTotalSize + varData(i).ColWidth
    
Next

'Egaliser
varData(iUBound).ColWidth = iWidth - iTotalSize

'********************************************************************************
'Impression dans l'objet cible
'********************************************************************************
'Initialisation de l'impression
With objTarget
    .CurrentX = iLeft
    iYSave = .CurrentY + 1
    .CurrentY = iYSave
End With
iXStartOfCol = iLeft

'Imprimer les textes
For i = iLBound To iUBound
    
    'Affectation
    RealString = strArray(i)
    strFont = Right(RealString, 1)
    strAlign = Mid(RealString, Len(RealString) - 1, 1)
    RealString = Left(RealString, Len(RealString) - 2)
    
    'Calculer la position de départ
    Select Case strAlign
        
        'Alignement à gauche
        Case "L"
        
            iCurrentX = iXStartOfCol + 1
            
        'Centré
        Case "M"
        
            iCurrentX = iXStartOfCol + (varData(i).ColWidth - varData(i).TextWidth) / 2
            
        'Alignement à droite
        Case "R"
        
            iCurrentX = (iXStartOfCol + varData(i).ColWidth) - varData(i).TextWidth - 1
            
    End Select
    
    'Positionner
    objTarget.CurrentX = iCurrentX
    
    'Afficher le texte selont la fonte
    Select Case strFont
    
        'Gras
        Case "B"
        
            objTarget.FontBold = True
            objTarget.Print RealString
            
        'En gras à droite du ":"
        Case "R"
            
            objTarget.FontBold = False
            objTarget.Print Left(RealString, InStr(1, RealString, ":"));
            objTarget.FontBold = True
            objTarget.Print Right(RealString, Len(RealString) - InStr(1, RealString, ":"))
            
        'En gras à gauche du ":"
        Case "L"
        
            objTarget.FontBold = True
            objTarget.Print Left(RealString, InStr(1, RealString, ":"));
            objTarget.FontBold = False
            objTarget.Print Right(RealString, Len(RealString) - InStr(1, RealString, ":"))
            
        'Normal
        Case Else
        
            objTarget.FontBold = False
            objTarget.Print RealString
            
    End Select
   
    'Pointer sur la prochaine position de départ
    iXStartOfCol = iXStartOfCol + varData(i).ColWidth
    objTarget.CurrentY = iYSave
    
Next

'********************************************************************************
'Encadrer les textes
'********************************************************************************
iXStartOfCol = iLeft
With objTarget
    .DrawWidth = 1
    .FontBold = False
End With

For i = iLBound To iUBound
    
    'Créer le cadre
    objTarget.Line (iXStartOfCol, iYSave - 1)-(iXStartOfCol + varData(i).ColWidth, iYSave + iYBigest), , B
    
    'Pointer sur la prochaine position de départ
    iXStartOfCol = iXStartOfCol + varData(i).ColWidth
    
Next

'Passer à la ligne suivante
objTarget.FontBold = False
objTarget.CurrentY = iYSave + iYBigest

'Libérer
Erase varData

End Sub

Private Sub PrintGraph(objTarget As Object)
'Sous-routine de création du graphique

'Constantes de taille à approcher
Const GRAPH_HEIGHT = 80
Const GRAPH_WIDTH = 155

'Objets
Dim currentSequences As Sequences

'Coefficient
Dim iDrawWidthCoeff As Integer

'Variables de propriétés du graph
Dim iGraph_Height As Integer
Dim iGraph_Width As Integer
Dim iGraphXLeftPosition As Integer
Dim iGraphXMax As Integer
Dim iGraphXStep As Integer
Dim iGraphYMin As Integer
Dim iGraphYMax As Integer
Dim iGraphYNullPosition As Integer
Dim iGraphYStep As Integer
Dim iGraphYTotal As Integer
Dim iGraphYTopPosition As Integer
Dim iYAxisPosition As Integer

'Variables de stockage temporaire
Dim iTmp1 As Integer
Dim iTmp2 As Integer

'Texte à imprimer
Dim strPrint As String

'********************************************************************************
'Affectation des variables
'********************************************************************************
'Coefficicent DrawWidth
If objTarget Is Printer Then
    iDrawWidthCoeff = 3
Else
    iDrawWidthCoeff = 1
End If

'Longueur d'un pas sur l'axe X
iGraphXMax = Int(mvarSerialValues(mvarSerialValues.Count)) + 1
iGraphXStep = GRAPH_WIDTH / iGraphXMax

'Nouvelle largeur du graph
iGraph_Width = iGraphXStep * iGraphXMax

'Position gauche de départ du graph
iGraphXLeftPosition = ((210 - iGraph_Width) / 2)

'Longueur d'un pas sur l'axe Y
iGraphYMin = Abs(Int(GetMinAverageOfAllSequences))
iGraphYMax = Abs(Int(GetMaxAverageOfAllSequences)) + 1
iGraphYTotal = Abs(iGraphYMin) + Abs(iGraphYMax)
iGraphYStep = GRAPH_HEIGHT / iGraphYTotal
iGraph_Height = iGraphYStep * iGraphYTotal

'********************************************************************************
'Affichage de la légende des courbes
'********************************************************************************
With objTarget
    iTmp1 = .TextHeight("|")
    iYAxisPosition = .CurrentY
    .CurrentX = iGraphXLeftPosition
    .DrawWidth = 2 * iDrawWidthCoeff
End With
strPrint = "Courbe montante: "
objTarget.Print strPrint;
objTarget.Line Step(0, (iTmp1 + 1) / 2)-Step(10, 0)
strPrint = "Courbe descendante: "
With objTarget
    .CurrentX = iGraphXLeftPosition + iGraph_Width - .TextWidth(strPrint) - 10
    .CurrentY = iYAxisPosition
End With
objTarget.Print strPrint;

With objTarget
    PrintSpecialLine .CurrentX, .CurrentY + (iTmp1 / 2), .CurrentX + 10, .CurrentY + (iTmp1 / 2), objTarget
    .CurrentY = .CurrentY - (iTmp1 / 2)
End With

'********************************************************************************
'Traçage des axes
'********************************************************************************
With objTarget
    iGraphYTopPosition = .CurrentY
    .CurrentX = iGraphXLeftPosition
End With

'Axes horizontaux
iTmp2 = iGraphYMax
For iTmp1 = 1 To iGraphYTotal + 1

    'Tracer la ligne
    objTarget.DrawWidth = IIf(iTmp2 <> 0, 1, 2 * iDrawWidthCoeff)
    objTarget.Line (iGraphXLeftPosition, iGraphYTopPosition + (iTmp1 * iGraphYStep))- _
    (iGraphXLeftPosition + iGraph_Width, iGraphYTopPosition + (iTmp1 * iGraphYStep))
    
    'Mémoriser la position Null
    If iTmp2 = 0 Then iGraphYNullPosition = iGraphYTopPosition + (iTmp1 * iGraphYStep)
    
    'Afficher la valeur de graduation
            
    'Positionner
    objTarget.CurrentY = iGraphYTopPosition + (iTmp1 * iGraphYStep) - (objTarget.TextHeight(Str(iTmp2)) / 2)
    objTarget.CurrentX = iGraphXLeftPosition - objTarget.TextWidth(Str(iTmp2)) - 2
    
    'Afficher
    objTarget.Print Str(iTmp2)
    
    'Décrémenter
    iTmp2 = iTmp2 - 1

Next

'Afficher la légende des ordonnées
strPrint = CHAR_MICRON + "m"
With objTarget
    .FontBold = True
    .CurrentX = iGraphXLeftPosition - .TextWidth(" 0") - .TextWidth(strPrint) - 3
    .CurrentY = iGraphYTopPosition + (iGraph_Height / 2)
End With
objTarget.Print strPrint
objTarget.FontBold = False

'Axes verticaux
For iTmp1 = 0 To iGraphXMax

    'Tracer la ligne
    objTarget.DrawWidth = IIf(iTmp1 = 0, 2 * iDrawWidthCoeff, 1)
    objTarget.Line (iGraphXLeftPosition + (iTmp1 * iGraphXStep), iGraphYTopPosition + iGraphYStep)- _
    (iGraphXLeftPosition + (iTmp1 * iGraphXStep), iGraphYTopPosition + (iGraphYStep * (iGraphYTotal + 1)))

    'Afficher la valeur de graduation
    objTarget.CurrentX = iGraphXLeftPosition + (iTmp1 * iGraphXStep) - (objTarget.TextWidth(Str(iTmp1)) / 2)
    objTarget.CurrentY = iGraphYTopPosition + (iGraphYStep * iGraphYTotal) + objTarget.TextHeight(Str(iTmp1)) + (iGraphYStep / 2)
    objTarget.Print Str(iTmp1)

Next
iYAxisPosition = objTarget.CurrentY

'********************************************************************************
'Traçage des courbes
'********************************************************************************
'Initialiser
objTarget.DrawWidth = 2 * iDrawWidthCoeff

'Enumérer les sens de séquences
For iTmp1 = 1 To 2
    
    'Affecter la collection de séquences en fonction du sens de la course
    If iTmp1 = 1 Then
        Set currentSequences = mvarUpSequences
    Else
        Set currentSequences = mvarDownSequences
    End If
    
    'Tracer les moyennes d'écart
    For iTmp2 = 1 To mvarSerialValues.Count - 1
        
        If iTmp1 = 1 Then
        
            'Tracer la courbe
            objTarget.Line ((mvarSerialValues(iTmp2) * iGraphXStep) + iGraphXLeftPosition, _
            iGraphYNullPosition - (currentSequences.GetAVerage(iTmp2) * iGraphYStep))- _
            ((mvarSerialValues(iTmp2 + 1) * iGraphXStep) + iGraphXLeftPosition, iGraphYNullPosition - _
            (currentSequences.GetAVerage(iTmp2 + 1) * iGraphYStep))
            
        Else
            
            'Tracer la courbe en mode spécial
            PrintSpecialLine (mvarSerialValues(iTmp2) * iGraphXStep) + iGraphXLeftPosition, _
            iGraphYNullPosition - (currentSequences.GetAVerage(iTmp2) * iGraphYStep), _
            (mvarSerialValues(iTmp2 + 1) * iGraphXStep) + iGraphXLeftPosition, _
            iGraphYNullPosition - (currentSequences.GetAVerage(iTmp2 + 1) * iGraphYStep), _
            objTarget
            
        End If
        
    Next
    
Next

'Libérer
Set currentSequences = Nothing

'********************************************************************************
'Afficher la légende des abscisses
'********************************************************************************
'Afficher la légende des abscisses
strPrint = "Course (mm)"
With objTarget
    .FontBold = True
    .ForeColor = QBColor(QBCOLOR_BLACK)
    .CurrentY = iYAxisPosition
    .CurrentX = iGraphXLeftPosition + (iGraph_Width - .TextWidth(strPrint)) / 2
End With
objTarget.Print (strPrint)

End Sub

Private Sub PrintSpecialLine(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, objTarget As Object)
'Trace une ligne en pointillé de taille constante avec un DrawWidth > 1

Const DASH_LENGTH = 1

Dim iLength As Currency                         'Longueur de la droite
Dim iDashCount As Currency                      'Nombre de pointillé à tracer
Dim iXStep As Currency                          'Pas moyen sur l'axe des X
Dim iCurrentX As Currency                       'X courant
Dim iNextX As Currency                          'Prochain X
Dim iNextY As Currency                          'Prochain Y

Dim bTrace As Boolean

'Equation de la droite
Dim a As Currency
Dim b As Currency

'Calcul de l'équation de droite
a = (y2 - y1) / (x2 - x1)
b = y1 - (a * x1)

'Longueur de la droite
iLength = Sqr(((x2 - x1) ^ 2) + ((IIf(y2 > y1, y2 - y1, y1 - y2)) ^ 2))

'Nombre de pointillés
iDashCount = iLength / DASH_LENGTH

'Pas moyen sur X
iXStep = (x2 - x1) / iDashCount

'Traçage
For iCurrentX = x1 To x2 Step iXStep
    
    'Permutation
    bTrace = Not bTrace
    
    If bTrace Then
    
        'Calcul
        iNextX = iCurrentX + iXStep
        If iNextX > x2 Then iNextX = x2
        iNextY = (a * iNextX) + b
        
        If y1 > y2 Then
            If iNextY < y2 Then iNextY = y2
        Else
            If iNextY > y2 Then iNextY = y2
        End If
        
        'Traçage
        objTarget.Line (iCurrentX, (a * iCurrentX) + b)-(iNextX, iNextY)
        
    End If
    
Next

'Tracer le point de fin si masqué
If Not bTrace And ((x2 - iNextX) > 1) Or Abs((y2 - iNextY)) > 1 Then
    
    iNextX = x2 - iXStep
    objTarget.Line (x2, y2)-(iNextX, (a * iNextX) + b)
    
End If

End Sub

Public Sub CreateASimulation()
'Créer une simulation

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim iCount As Integer

Dim bPlus As Boolean
Dim sDifference As Currency

'Initialiser
PrepareForANewControl
iCount = mvarSerialValues.Count
mvarControlIsStart = True

'Enumérer les séquences
For i = 1 To mvarSequencesCount
        
    'Enumérer les sens de séquence
    For k = 1 To 2
                        
        'Initialiser l'horloge
        Randomize
        
        'Enumérer les valeurs de la série de mesures
        For j = 1 To mvarSerialValues.Count
            
            'Déterminer si l'écart sera positif
            bPlus = ((Int((2 * Rnd) + 1) <> 2))
            
            'Si première mesure
            If i = 1 And j = 1 And k = 1 Then
                SaveMeasure 0
            Else
            
                'Enregistrer une mesure avec un écart aléatoire
                sDifference = Round(Int(Rnd * 100) / 10000, 3)
                If Not bPlus Then sDifference = -sDifference
                SaveMeasure SerialValues(IIf(k = 1, j, iCount - j + 1)).Value + sDifference
                
            End If
            
        Next
        
    Next
    
Next

End Sub

Public Property Get RealClass() As Integer
'Accès en lecture sur la classe réelle

'Renvoyer
RealClass = mvarRealClass

End Property

Public Property Let RealClass(ByVal vData As Integer)
'Accès en écriture sur la classe réelle

'Affecter
mvarRealClass = vData

End Property

Public Function GetTheoreticalClass(ByVal indxOfLevel As Integer, ByVal indxOfCapacity As Integer, _
ByVal isRealClass As Boolean) As Integer
'Renvoie la classe du comparateur (0 ou 1 ou 2 pour "A REBUTER")

Dim currentClass0 As Classification
'Dim currentClass1 As Classification

Dim currentValue As Currency
Dim iResult As Integer

'Sécurité
If mvarLevels Is Nothing Then Set mvarLevels = New Levels

'Affectation
mvarIndexOfCapacity = indxOfCapacity
mvarIndexOfLevel = indxOfLevel
Set currentClass0 = mvarLevels(indxOfLevel).Capacitys(indxOfCapacity).Classifications(1)
'Set currentClass1 = mvarLevels(indxOfLevel).Capacitys(indxOfCapacity).Classifications(2)

'********************************************************************************
'Travail sur l'erreur de justesse totale
'********************************************************************************
currentValue = ExactnessTotalError
'If currentValue > currentClass1.TotalExactness Then
'
'    GetTheoreticalClass = 2
'    If isRealClass Then mvarRealClass = 2
'
'    Exit Function
'
If currentValue >= currentClass0.TotalExactness Then
    iResult = 1
End If
'
''********************************************************************************
''Travail sur l'erreur de justesse locale
''********************************************************************************
currentValue = ExactnessLocalError
'If currentValue > currentClass1.LocalExactness Then
'
'    GetTheoreticalClass = 2
'    If isRealClass Then mvarRealClass = 2
'
'    Exit Function
'
If currentValue >= currentClass0.LocalExactness Then
    iResult = 1
End If
'
''********************************************************************************
''Travail sur l'erreur de fidélité
''********************************************************************************
currentValue = FidelityError
'If currentValue > currentClass1.Fidelity Then
'
'    GetTheoreticalClass = 2
'    If isRealClass Then mvarRealClass = 2
'
'    Exit Function
'
If currentValue >= currentClass0.Fidelity Then
    iResult = 1
End If
'
''********************************************************************************
''Travail sur l'erreur d'hystérésis
''********************************************************************************
currentValue = HysteresisError
'If currentValue > currentClass1.Hysteresis Then
'
'    GetTheoreticalClass = 2
'    If isRealClass Then mvarRealClass = 2
'
'    Exit Function
'
If currentValue >= currentClass0.Hysteresis Then
    iResult = 1
End If

'********************************************************************************
'Travail de fin
'********************************************************************************

'Affecter
GetTheoreticalClass = iResult
If isRealClass Then mvarRealClass = iResult

'Libérer
Set currentClass0 = Nothing
'Set currentClass1 = Nothing

End Function

Public Property Get Levels() As Levels
'Accès en lecture sur la collection d'échelons

'Initialiser au besoin
If mvarLevels Is Nothing Then Set mvarLevels = New Levels

'Renvoyer
Set Levels = mvarLevels

End Property

Public Property Get CorrectionModeOfFidelityTest() As Boolean
'Accès en lecture sur le mode de correction du test de fidélité

'Renvoyer
CorrectionModeOfFidelityTest = mvarCorrectionModeOfFidelityTest

End Property

Public Property Let CorrectionModeOfFidelityTest(ByVal bData As Boolean)
'Accès en écriture sur le mode de corretcion du test de fidélité

'Affecter
If bData <> mvarCorrectionModeOfFidelityTest Then
    
    'Affecter
    mvarCorrectionModeOfFidelityTest = bData
    
    'Evènement
    RaiseEvent CorrectionModeOfFidelityTestChange
    
End If

End Property

Private Sub PrintComment(objTarget As Object, iCurrentX As Integer)
'Impression du commentaire avec prise en compte des sauts de ligne

Dim iLeft As Integer
Dim iRight As Integer
Dim iLen As Integer

'Initialisation
iLen = Len(mvarComment)
iLeft = 1
iRight = InStr(1, mvarComment, vbCrLf)
If iRight = 0 Then iRight = iLen + 1

'Impression du titre
With objTarget
    .FontSize = 12
    .FontBold = True
    .FontUnderline = True
    .CurrentX = iCurrentX
End With
objTarget.Print "Observations:"

'Mise en forme
With objTarget
    .FontBold = False
    .FontUnderline = False
End With

'Boucle d'impression des lignes
Do
    
    'Impression de la ligne
    objTarget.CurrentX = iCurrentX
    objTarget.Print Mid(mvarComment, iLeft, iRight - iLeft)
    
    'Positionnement
    iLeft = iRight + 2
    If iLeft > iLen Then Exit Do
    
    iRight = InStr(iLeft, mvarComment, vbCrLf)
    If iRight = 0 Then iRight = iLen + 1
    
Loop

End Sub

Private Sub PrintResult(objTarget As Object, objConfiguration As Configuration)
'Impression du résultat du contrôle

Dim currentClass0 As Classification
Dim currentClass1 As Classification
Dim strArray As Variant

'Affectation
Set currentClass0 = mvarLevels(mvarIndexOfLevel).Capacitys(mvarIndexOfCapacity).Classifications(1)
'Set currentClass1 = mvarLevels(mvarIndexOfLevel).Capacitys(mvarIndexOfCapacity).Classifications(2)

'Impression
strArray = Array("Limites d'usure en " + CHAR_MICRON + "m d'après NFE 11-200MB")
PrintArrayOfString strArray, objTarget, objConfiguration

strArray = Array(" MB", "Justesse totaleMN", "Justesse localeMN", "FidélitéMN", "HystérésisMN")
PrintArrayOfString strArray, objTarget, objConfiguration

strArray = Array("Classe 0LN", Format(currentClass0.TotalExactness) + "MN", Format(currentClass0.LocalExactness) + "MN", _
CHAR_PLUS_MINUS + Format(currentClass0.Fidelity) + "MN", CHAR_PLUS_MINUS + Format(currentClass0.Hysteresis) + "MN")
PrintArrayOfString strArray, objTarget, objConfiguration

'strArray = Array("Classe 1LN", Format(currentClass1.TotalExactness) + "MN", Format(currentClass1.LocalExactness) + "MN", _
CHAR_PLUS_MINUS + Format(currentClass1.Fidelity) + "MN", CHAR_PLUS_MINUS + Format(currentClass1.Hysteresis) + "MN")
'PrintArrayOfString strArray, objTarget, objConfiguration

strArray = Array("MesuresLB", Format(ExactnessTotalError, "0.00") + "MB", Format(ExactnessLocalError, "0.00") + "MB", _
Format(FidelityError, "0.00") + "MB", Format(HysteresisError, "0.00") + "MB")
PrintArrayOfString strArray, objTarget, objConfiguration

ReDim strArray(0)
Select Case RealClass
    Case 0
        strArray(0) = "CLASSE 0"
    Case 1
        strArray(0) = "CLASSE 1"
    Case 2
        strArray(0) = "A REBUTER"
End Select
strArray(0) = strArray(0) + "MB"
PrintArrayOfString strArray, objTarget, objConfiguration

'Libérer
Set currentClass0 = Nothing
Set currentClass1 = Nothing
Set strArray = Nothing

End Sub


Public Sub Export_Excel(objConfiguration As Configuration, currentClassification As Classification)
Dim objExcel As Object, i As Integer, j As Integer, sTmp As String
Dim Sauve_Constats As String, iNbSeq As Integer, iNbMes As Integer
Const yMesures = 4
Const yConstat = 4
Sauve_Constats = App.Path & "\Constats\"
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Open Sauve_Constats & "Constat.xlt"
objExcel.Visible = False
objConfiguration.SaveConfig
objConfiguration.LoadConfig


'Constat
With objExcel.Worksheets("Constat")

    'Entete
    .cells(yConstat + 1, 4).Value = "'" & (mvarEquipmentNumber)
                               
    .cells(yConstat + 2, 4).Value = CStr(mvarEquipmentManufacturer)
    .cells(yConstat + 3, 4).Value = Str(mvarTemperature) & " °C"
    .cells(yConstat + 4, 4).Value = "12 mois"
    .cells(yConstat + 5, 4).Value = objConfiguration.Print_Tool
    .cells(yConstat + 1, 8).Value = objConfiguration.Print_Designation
    .cells(yConstat + 2, 8).Value = CStr(Levels.Item(mvarIndexOfLevel).Capacitys(mvarIndexOfCapacity).Value) & " mm"
    .cells(yConstat + 3, 9).Value = mvarDetenteur
    .cells(yConstat + 4, 10).Value = UCase(Format(DateAdd("m", 12, mvarDateOfControl), "mmmm yyyy"))
    .cells(yConstat + 5, 9).Value = CStr(Levels.Item(mvarIndexOfLevel).Value) & " mm"
        
   'Limites d'usure
   .cells(yConstat + 35, 1).Value = CStr(Levels.Item(mvarIndexOfLevel).Value)
   .cells(yConstat + 35, 3).Value = CStr(Levels.Item(mvarIndexOfLevel).Capacitys(mvarIndexOfCapacity).Value)
   .cells(yConstat + 37, 5).Value = Format(currentClassification.TotalExactness, "0.00")
   .cells(yConstat + 37, 9).Value = Format(currentClassification.Fidelity, "0.00")
   .cells(yConstat + 37, 11).Value = Format(currentClassification.Hysteresis, "0.00")
   .cells(yConstat + 35, 5).Value = Format(ExactnessTotalError, "0.00")
   If currentClassification.LocalExactness <= 100 Then
        .cells(yConstat + 37, 7).Value = Format(currentClassification.LocalExactness, "0.00")
        .cells(yConstat + 35, 7).Value = Format(ExactnessLocalError, "0.00")
   Else
        .cells(yConstat + 37, 7).Value = "Néant"
        .cells(yConstat + 35, 7).Value = "Néant"
   End If
   .cells(yConstat + 35, 9).Value = Format(FidelityError, "0.000")
   .cells(yConstat + 35, 11).Value = Format(HysteresisError, "0.00")
   '.cells(yConstat + 21, 8).Value = IIf(mvarRealClass = 1, "A REBUTER", "BON ETAT")
   If mvarRealClass = 1 Then
        .cells(45, 8).Font.Strikethrough = True
    
   Else
        .cells(45, 12).Font.Strikethrough = True
   End If
   
   .cells(yConstat + 43, 3).Value = mvarOperatorName
   .cells(yConstat + 44, 4).Value = Format(mvarDateOfControl, "dd mmm yyyy")
End With
    
'Mesures
With objExcel.Worksheets("Mesures")
    .cells(yMesures, 1).Value = "SÉRIES de MESURES"
    'Valeurs des séries
    iNbMes = mvarSerialValues.Count
    For i = 1 To iNbMes 'mvarSerialValues.Count
         '.cells(yMesures, i + 1).Value = Format(Round(mvarSerialValues(i).Value + 0.5), "0")
         .cells(yMesures + 1, i + 1).Value = Format(mvarSerialValues(i).Value, "0.00") '& " mm"
    Next i
    ' Mesures montantes
    iNbSeq = mvarUpSequences.Count
    For i = 1 To mvarUpSequences.Count
       .cells(yMesures + i + 1, 1).Value = "M" & i
        For j = 1 To iNbMes 'mvarUpSequences(i).Measures.Count
            sTmp = mvarUpSequences(i).Measures(j)
             sTmp = GetMicronsDifference(j, i, True)
            .cells(yMesures + i + 1, j + 1).Value = sTmp '* 1000
            sTmp = Format(mvarUpSequences.GetAVerage(j) + (mvarSerialValues(j).Value * 1000), "0.00")
            sTmp = mvarUpSequences.GetAVerage(j)
            .cells(2 + yMesures + mvarUpSequences.Count, j + 1).Value = sTmp
        Next j
    Next i
   .cells(2 + yMesures + mvarUpSequences.Count, 1).Value = "Moyenne"
    'Mesures descendantes
    For i = 1 To mvarDownSequences.Count
        .cells(2 + yMesures + i + mvarUpSequences.Count, 1).Value = "D" & i
        For j = 1 To iNbMes 'mvarDownSequences(i).Measures.Count
            sTmp = mvarDownSequences(i).Measures(j)
            sTmp = GetMicronsDifference(j, i, False)
            .cells(2 + yMesures + i + mvarUpSequences.Count, j + 1).Value = sTmp '* 1000
            sTmp = Format(mvarDownSequences.GetAVerage(j) + (mvarSerialValues(j).Value * 1000), "0.00")
            sTmp = mvarDownSequences.GetAVerage(j)
            .cells(3 + yMesures + mvarUpSequences.Count + mvarDownSequences.Count, j + 1).Value = sTmp
        Next j
    Next i
    .cells(3 + yMesures + mvarUpSequences.Count + mvarDownSequences.Count, 1).Value = "Moyenne"
    'Mise en forme
    '.Range("A" & yMesures & ":M" & yMesures).Font.Bold = True
    .Range(.cells(yMesures, 1), .cells(yMesures, iNbMes + 1)).merge
    .Rows(yMesures).RowHeight = 20
    .cells(yMesures, 1).HorizontalAlignment = -4108
    .cells(yMesures, 1).VerticalAlignment = -4108
    .Range(.cells(yMesures, 1), .cells(yMesures + (iNbSeq * 2) + 3, iNbMes + 1)).Font.Bold = True
    .Range(.cells(yMesures, 1), .cells(yMesures + (iNbSeq * 2) + 3, iNbMes + 1)).Borders.LineStyle = 1
    '.Columns("A:M").Columnwidth = 12
 
'Ecarts de fidélité
    .cells(yMesures + 26, 1).Value = "ECARTS de FIDÉLITÉ"
    For i = 1 To mvarFidelityPoints.Count
        .cells(yMesures + 27 + i, 1).Value = Format(mvarFidelityPoints(i).SerialValue, "0.00") + IIf(mvarFidelityPoints(i).IsUpDirection, " M", " D")
        For j = 1 To mvarFidelityPoints.MeasuresCount
            If (i = 1) Then
                .cells(yMesures + 26 + i, j + 1).Value = "N° " & j
            End If
            .cells(yMesures + 27 + i, j + 1).Value = Format(GetFidelityMicronsDifference(i, j), "0.00")
        Next j
        If (i = 1) Then
            .cells(yMesures + 26 + i, j + 1).Value = "Moyenne"
            .cells(yMesures + 26 + i, j + 2).Value = "Fidélité"
        End If
        .cells(yMesures + 27 + i, j + 1) = Format(mvarFidelityPoints(i).GetAVerage, "0.00")
        .cells(yMesures + 27 + i, j + 2) = Format(mvarFidelityPoints(i).FidelityError, "0.00")
    Next i
    
 'Mise en forme
     '.Range("A" & yMesures + 25 & ":M" & yMesures).Font.Bold = True
    .Range(.cells(yMesures + 26, 1), .cells(yMesures + 26, mvarFidelityPoints.MeasuresCount + 3)).merge
    .Rows(yMesures + 26).RowHeight = 20
    .cells(yMesures + 26, 1).HorizontalAlignment = -4108
    .cells(yMesures + 26, 1).VerticalAlignment = -4108
    .Range(.cells(yMesures + 26, 1), .cells(yMesures + (mvarFidelityPoints.Count) + 30, mvarFidelityPoints.MeasuresCount + 3)).Font.Bold = True
    .Range(.cells(yMesures + 26, 1), .cells(yMesures + (mvarFidelityPoints.Count) + 30, mvarFidelityPoints.MeasuresCount + 3)).Borders.LineStyle = 1
 End With
 
 'Courbe d'étalonnage
   ' la courbe est créer à partir des mesures ci-dessous
With objExcel.Worksheets("Constat").ChartObjects(1).Chart
    Dim iColFin As Integer, iLig As Integer
    .Type = 1
    iColFin = iNbMes + 1
    objExcel.Visible = True
    iLig = yMesures + 1
    .SeriesCollection(1).XValues = "=Mesures!R" & iLig & "C2:R" & iLig & "C" & iColFin
    iLig = iLig + iNbSeq + 1
    .SeriesCollection(1).Values = "=Mesures!R" & iLig & "C2:R" & iLig & "C" & iColFin
    iLig = iLig + iNbSeq + 1
    .SeriesCollection(2).Values = "=Mesures!R" & iLig & "C2:R" & iLig & "C" & iColFin
    .Type = 4
End With
'objExcel.Worksheets("Constat").cells(1, 1).Select
objExcel.ActiveWorkbook.SaveAs Sauve_Constats & mvarEquipmentNumber & "_" & Format(mvarDateOfControl, "dd-MM-yy") & ".xls"
objExcel.Visible = True
objExcel.Worksheets(1).PrintPreview
Set objExcel = Nothing
End Sub





