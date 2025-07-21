VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Level"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Capacity"
Attribute VB_Ext_KEY = "Member1" ,"Capacitys"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UN ECHELON
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mvarCapacitys As Capacitys
Private mvarValue As Currency

Public Property Let Value(ByVal vData As Currency)
'Accès en écriture sur la valeur de l'échelon

'Affecter
mvarValue = vData

End Property

Public Property Get Value() As Currency
Attribute Value.VB_UserMemId = 0
'Accès en lecture sur la valeur de l'échelon

'Renvoyer
Value = mvarValue

End Property

Public Property Get Capacitys() As Capacitys
'Accès en lecture sur la collection de courses

'Renvoyer
Set Capacitys = mvarCapacitys

End Property

Public Property Set Capacitys(vData As Capacitys)
'Accès en écriture sur la collection de courses

'Affecter
Set mvarCapacitys = vData

End Property

Private Sub Class_Initialize()
'Constructeur

'Initialiser
Set mvarCapacitys = New Capacitys

End Sub

Private Sub Class_Terminate()
'Destructeur

'Libérer
Set mvarCapacitys = Nothing

End Sub
