VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Classification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UNE CLASSIFICATION (CLASSE 0 OU 1)
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mvarTotalExactness As Integer
Private mvarLocalExactness As Integer
Private mvarFidelity As Currency
Private mvarHysteresis As Currency

Public Property Let Hysteresis(ByVal vData As Currency)
'Accès en écriture sur la valeur d'hystérésis

'Affecter
mvarHysteresis = vData

End Property

Public Property Get Hysteresis() As Currency
'Accès en lecture sur la valeur d'hystérésis

'Renvoyer
Hysteresis = mvarHysteresis

End Property

Public Property Let Fidelity(ByVal vData As Currency)
'Accès en écriture sur la valeur de fidélité

'Affecter
mvarFidelity = vData

End Property

Public Property Get Fidelity() As Currency
'Accès en lecture sur la valeur de fidélité

'Renvoyer
Fidelity = mvarFidelity

End Property

Public Property Let LocalExactness(ByVal vData As Integer)
'Accès en écriture sur la valeur d'erreur locale

'Affecter
mvarLocalExactness = vData

End Property

Public Property Get LocalExactness() As Integer
'Accès en lecture sur la valeur d'erreur locale

'Renvoyer
LocalExactness = mvarLocalExactness

End Property

Public Property Let TotalExactness(ByVal vData As Integer)
'Accès en écriture sur la valeur totale d'erreur

'Affecter
mvarTotalExactness = vData

End Property

Public Property Get TotalExactness() As Integer
'Accès en lecture sur la valeur totale d'erreur

'Renvoyer
TotalExactness = mvarTotalExactness

End Property
