VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Capacity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Classification"
Attribute VB_Ext_KEY = "Member1" ,"Classifications"
Option Explicit

'********************************************************************************
'IMPLEMENTATION D'UNE COURSE
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private mvarClassifications As Classifications
Private mvarValue As Currency

Public Property Let Value(ByVal vData As Currency)
'Accès en écriture sur la valeur de course

'Affecter
mvarValue = vData

End Property

Public Property Get Value() As Currency
'Accès en lecture sur la valeur de course

'Renvoyer
Value = mvarValue

End Property

Public Property Get Classifications() As Classifications
'Accès en lecture sur la collection de classification

'Renvoyer
Set Classifications = mvarClassifications

End Property

Public Property Set Classifications(vData As Classifications)
'Accès en écriture sur la collection de classification

'Affecter
Set mvarClassifications = vData

End Property

Private Sub Class_Initialize()
'Constructeur

'Initialiser
Set mvarClassifications = New Classifications

End Sub

Private Sub Class_Terminate()
'Destructeur

'Libérer
Set mvarClassifications = Nothing

End Sub
