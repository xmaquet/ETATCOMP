VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Measure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UNE MESURE
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mvarEnabled As Currency
Private mvarValue As Currency

Public Property Let Enabled(ByVal bData As Boolean)
'Accès en écriture sur la validité d'une mesure

'Affecter
mvarEnabled = bData

End Property

Public Property Get Enabled() As Boolean
'Acès en lecture sur la validité d'une mesure

'Renvoyer
Enabled = mvarEnabled

End Property

Public Property Let Value(ByVal vData As Currency)
'Accès en écriture sur la valeur de la mesure

'Affecter
mvarValue = vData

End Property

Public Property Get Value() As Currency
Attribute Value.VB_UserMemId = 0
'Accès en lecture sur la valeur de la mesure

'Renvoyer
Value = mvarValue

End Property
