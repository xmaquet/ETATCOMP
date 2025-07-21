VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Measures"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UNE SEQUENCE DE MESURES
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mvarMeasures As Measures

Public Property Set Measures(ByVal vData As Measures)
'Accès en écriture sur la collection de mesures

'Affecter
Set mvarMeasures = vData

End Property

Public Property Get Measures() As Measures
'Accès en lecture sur la collection de mesures

'Renvoyer
Set Measures = mvarMeasures

End Property

Private Sub Class_Initialize()
'Constructeur

'Initaliser
Set mvarMeasures = New Measures

End Sub

Private Sub Class_Terminate()
'Destructeur

'Libérer
Set mvarMeasures = Nothing

End Sub
