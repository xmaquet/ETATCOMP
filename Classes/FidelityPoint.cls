VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FidelityPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SerialValue"
Attribute VB_Ext_KEY = "Member1" ,"Measures"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UN POINT DE RELEVE D'ECART DE FIDELITE
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mvarMeasures As Measures
Private mvarSerialValue As Currency
Private mvarIsUpDirection As Boolean

Public Property Let IsUpDirection(ByVal bData As Boolean)
'Accès en écriture sur le sens de la séquence correspondante

'Affecter
mvarIsUpDirection = bData

End Property

Public Property Get IsUpDirection() As Boolean
'Accès en lecture sur le sens de la séquence correspondante

'Renvoyer
IsUpDirection = mvarIsUpDirection

End Property

Public Property Get Measures() As Measures
'Accès en lecture sur la collection de mesures associées

'Renvoyer
Set Measures = mvarMeasures

End Property

Public Property Let Measures(ByVal vData As Measures)
'Accès en écriture sur la valeur du point de mesure

'Affecter
Set mvarMeasures = vData

End Property

Public Property Let SerialValue(ByVal vData As Currency)
'Accès en écriture sur la valeur du point de mesure

'Affecter
mvarSerialValue = vData

End Property

Public Property Get SerialValue() As Currency
Attribute SerialValue.VB_UserMemId = 0
'Accès en lecture sur la valeur du point de mesure

'Renvoyer
SerialValue = mvarSerialValue

End Property

Private Sub Class_Initialize()
'Constructeur

'Initialiser
Set mvarMeasures = New Measures

End Sub

Private Sub Class_Terminate()
'Destructeur

'Libérer
Set mvarMeasures = Nothing

End Sub

Public Function GetAVerage() As Currency
'Renvoie la moyenne des écarts mesurés

Dim currentMeasure As Measure

Dim iValueCount As Integer
Dim iTotal As Long

'Enumérer les mesures
For Each currentMeasure In mvarMeasures
        
    'Prendre en compte si validée
    If currentMeasure.Enabled Then
        
        iTotal = iTotal + ((mvarSerialValue - currentMeasure.Value) * 1000)
        iValueCount = iValueCount + 1
        
    End If
    
Next

'Renvoyer la moyenne
If iValueCount > 0 Then GetAVerage = iTotal / iValueCount

End Function

Public Property Get FidelityError() As Currency 'modifier pour nouvelle norme
'Accès en lecture sur l'erreur de fidélité

Dim currentMeasure As Measure

Dim cMaxValue As Currency
Dim cAverage As Currency
Dim currentValue As Currency

'Affecter
cAverage = (GetAVerage)

'Enumérer les mesures
For Each currentMeasure In mvarMeasures
    
    'Si mesure validée
    If currentMeasure.Enabled Then
        
        'Calculer l'écart absolu
        currentValue = (((mvarSerialValue - currentMeasure.Value) * 1000) - cAverage)
        
        'Affecter si valeur maxi
        currentValue = currentValue * currentValue
        cMaxValue = (cMaxValue + currentValue)
    End If
    
Next
cMaxValue = cMaxValue / (mvarMeasures.Count - 1)
cMaxValue = Sqr(cMaxValue)

'Affecter
FidelityError = cMaxValue

'Libérer
Set currentMeasure = Nothing

End Property



'Public Property Get FidelityError() As Currency
'Accès en lecture sur l'erreur de fidélité

'Dim currentMeasure As Measure

'Dim cMaxValue As Currency
'Dim cAverage As Currency
'Dim currentValue As Currency

'Affecter
'cAverage = Abs(GetAVerage)

'Enumérer les mesures
'For Each currentMeasure In mvarMeasures
    
    'Si mesure validée
'    If currentMeasure.Enabled Then
        
        'Calculer l'écart absolu
'        currentValue = Abs(cAverage - (Abs(mvarSerialValue - currentMeasure.Value) * 1000))
        
        'Affecter si valeur maxi
'        If currentValue > cMaxValue Then cMaxValue = currentValue
        
'    End If
    
'Next

'Affecter
'FidelityError = cMaxValue

'Libérer
'Set currentMeasure = Nothing

'End Property
