VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'********************************************************************************
'IMPLEMENTATION DE LA CONFIGURATION DE L'APPLICATION
'********************************************************************************

'********************************************************************************
'Constantes
'********************************************************************************

'Constantes liées à l'application
Private Const APP_NAME = "EtaComp"

'Constantes liées à la base de registre
Private Const REG_KEY_COMM_BAUDS = "Bauds"
Private Const REG_KEY_COMM_DATA_BITS = "DataBits"
Private Const REG_KEY_COMM_PARITY = "Parity"
Private Const REG_KEY_COMM_PORT = "CommPort"
Private Const REG_KEY_COMM_STOP_BITS = "StopBits"

Private Const REG_KEY_CONTROL_BEEP = "BeepAfterMeasure"

Private Const REG_KEY_GEN_CONFIRM_EXIT = "ConfirmExit"

Private Const REG_KEY_LAST_OPERATOR_NAME = "LastOperatorName"
Private Const REG_KEY_LAST_EQUIPMENT_MANUFACTURER = "LastEquipmentManufacturer"

Private Const REG_KEY_PRINT_HOUSE = "House"
Private Const REG_KEY_PRINT_SERVICE = "Service"
Private Const REG_KEY_PRINT_DESIGNATION = "Designation"
Private Const REG_KEY_PRINT_DOCUMENT = "Document"
Private Const REG_KEY_PRINT_FONT = "Font"
Private Const REG_KEY_PRINT_LATERAL_MARGIN = "LateralMargin"
Private Const REG_KEY_PRINT_TOOL = "Tool"
Private Const REG_KEY_PRINT_POSITION = "Position"

Private Const REG_KEY_RESULT_LAST_LEVEL = "LastLevel"
Private Const REG_KEY_RESULT_LAST_CAPACITY = "LastCapacity"

Private Const REG_KEY_TIPS_CURRENT_TIP = "CurrentTip"
Private Const REG_KEY_TIPS_SHOW_TIPS = "ShowTips"

Private Const REG_SECTION_SETTINGS = "Settings"
Private Const REG_SECTION_COMM_SETTINGS = REG_SECTION_SETTINGS + "\Comm"
Private Const REG_SECTION_CONTROL_SETTINGS = REG_SECTION_SETTINGS + "\Control"
Private Const REG_SECTION_PRINT_SETTINGS = REG_SECTION_SETTINGS + "\Print"

'********************************************************************************
'Données membres
'********************************************************************************
Public Comm_Bauds As Integer
Public Comm_CommPort As Integer
Public Comm_DataBits As Integer
Public Comm_Parity As String
Public Comm_StopBits As Integer

Public Control_BeepAfterMeasure As Boolean

Public Gen_ConfirmExit As Boolean

Public Print_House As String
Public Print_Service As String
Public Print_Lateral_Margin As Integer
Public Print_Designation As String
Public Print_DocReference As String
Public Print_Font As String
Public Print_Tool As String
Public Print_Position As String

Public Result_LastLevel As Currency
Public Result_LastCapacity As Integer

Public Tips_CurrentTip As Integer
Public Tips_ShowTips As Boolean

Public Last_EquipmentManufacturer As String
Public Last_OperatorName As String

Public Sub LoadConfig()
'Charger la configuration depuis la base de registre

Dim strDefault As String

'Communication
Comm_Bauds = CInt(GetSetting(APP_NAME, REG_SECTION_COMM_SETTINGS, REG_KEY_COMM_BAUDS, "1200"))
Comm_CommPort = CInt(GetSetting(APP_NAME, REG_SECTION_COMM_SETTINGS, REG_KEY_COMM_PORT, "1"))
Comm_DataBits = CInt(GetSetting(APP_NAME, REG_SECTION_COMM_SETTINGS, REG_KEY_COMM_DATA_BITS, "7"))
Comm_Parity = GetSetting(APP_NAME, REG_SECTION_COMM_SETTINGS, REG_KEY_COMM_PARITY, "E")
Comm_StopBits = CInt(GetSetting(APP_NAME, REG_SECTION_COMM_SETTINGS, REG_KEY_COMM_STOP_BITS, "1"))

'Contrôle
Control_BeepAfterMeasure = CBool(GetSetting(APP_NAME, REG_SECTION_CONTROL_SETTINGS, REG_KEY_CONTROL_BEEP, "1"))

'Général
Gen_ConfirmExit = CBool(GetSetting(APP_NAME, REG_SECTION_SETTINGS, REG_KEY_GEN_CONFIRM_EXIT, "1"))

'Impression
strDefault = "COMPARATEUR mécanique à cadran à tige rentrante radiale"
Print_Designation = GetSetting(APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_DESIGNATION, strDefault)

strDefault = "Procédure VI-05 NFE 11-200/NFE 11-050"
Print_DocReference = GetSetting(APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_DOCUMENT, strDefault)

Print_Font = GetSetting(APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_FONT, Printer.FontName)

strDefault = "2ème RMAT de BRUZ"
Print_House = GetSetting(APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_HOUSE, strDefault)

strDefault = "CONTROLE DU COMPARATEUR EN POSITION HORIZONTALE"
Print_Position = GetSetting(APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_POSITION, strDefault)

Print_Service = GetSetting(APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_SERVICE, "MET. AMMP")

strDefault = "BANC de CONTROLE n°324 MR 0024 (Capacité: 0 à 35 mm, Précision: ± 2 " + CHAR_MICRON + "m)"
Print_Tool = GetSetting(APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_TOOL, strDefault)

Print_Lateral_Margin = CInt(GetSetting(APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_LATERAL_MARGIN, 0))

'Résultat
Result_LastCapacity = CInt(GetSetting(APP_NAME, REG_SECTION_CONTROL_SETTINGS, REG_KEY_RESULT_LAST_CAPACITY, "5"))
Result_LastLevel = CCur(GetSetting(APP_NAME, REG_SECTION_CONTROL_SETTINGS, REG_KEY_RESULT_LAST_LEVEL, 0.01))

'Tips
Tips_CurrentTip = CInt(GetSetting(APP_NAME, REG_SECTION_SETTINGS, REG_KEY_TIPS_CURRENT_TIP, "0"))
Tips_ShowTips = CBool(GetSetting(APP_NAME, REG_SECTION_SETTINGS, REG_KEY_TIPS_SHOW_TIPS, "1"))

'Last
Last_EquipmentManufacturer = GetSetting(APP_NAME, REG_SECTION_SETTINGS, REG_KEY_LAST_EQUIPMENT_MANUFACTURER)
Last_OperatorName = GetSetting(APP_NAME, REG_SECTION_SETTINGS, REG_KEY_LAST_OPERATOR_NAME)

End Sub

Public Sub SaveConfig()
'Enregistrer la configuration dans la base de registre

'Communication
SaveSetting APP_NAME, REG_SECTION_COMM_SETTINGS, REG_KEY_COMM_BAUDS, Comm_Bauds
SaveSetting APP_NAME, REG_SECTION_COMM_SETTINGS, REG_KEY_COMM_PORT, Comm_CommPort
SaveSetting APP_NAME, REG_SECTION_COMM_SETTINGS, REG_KEY_COMM_DATA_BITS, Comm_DataBits
SaveSetting APP_NAME, REG_SECTION_COMM_SETTINGS, REG_KEY_COMM_PARITY, Comm_Parity
SaveSetting APP_NAME, REG_SECTION_COMM_SETTINGS, REG_KEY_COMM_STOP_BITS, Comm_StopBits

'Contrôle
SaveSetting APP_NAME, REG_SECTION_CONTROL_SETTINGS, REG_KEY_CONTROL_BEEP, Control_BeepAfterMeasure

'Général
SaveSetting APP_NAME, REG_SECTION_SETTINGS, REG_KEY_GEN_CONFIRM_EXIT, Gen_ConfirmExit

'Impression
SaveSetting APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_DESIGNATION, Print_Designation
SaveSetting APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_DOCUMENT, Print_DocReference
SaveSetting APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_FONT, Print_Font
SaveSetting APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_HOUSE, Print_House
SaveSetting APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_LATERAL_MARGIN, Print_Lateral_Margin
SaveSetting APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_POSITION, Print_Position
SaveSetting APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_SERVICE, Print_Service
SaveSetting APP_NAME, REG_SECTION_PRINT_SETTINGS, REG_KEY_PRINT_TOOL, Print_Tool

'Résultat
SaveSetting APP_NAME, REG_SECTION_CONTROL_SETTINGS, REG_KEY_RESULT_LAST_CAPACITY, Result_LastCapacity
SaveSetting APP_NAME, REG_SECTION_CONTROL_SETTINGS, REG_KEY_RESULT_LAST_LEVEL, Result_LastLevel

'Tips
SaveSetting APP_NAME, REG_SECTION_SETTINGS, REG_KEY_TIPS_CURRENT_TIP, Tips_CurrentTip
SaveSetting APP_NAME, REG_SECTION_SETTINGS, REG_KEY_TIPS_SHOW_TIPS, Tips_ShowTips

'Last
SaveSetting APP_NAME, REG_SECTION_SETTINGS, REG_KEY_LAST_EQUIPMENT_MANUFACTURER, Last_EquipmentManufacturer
SaveSetting APP_NAME, REG_SECTION_SETTINGS, REG_KEY_LAST_OPERATOR_NAME, Last_OperatorName

End Sub

Private Sub Class_Initialize()
'Constructeur

'Charger la configuration
LoadConfig

End Sub

Public Sub Copy(objSource As Configuration)
'Copier un objet

'Affecter
With objSource
    
    'Communication
    Comm_Bauds = .Comm_Bauds
    Comm_CommPort = .Comm_CommPort
    Comm_DataBits = .Comm_DataBits
    Comm_Parity = .Comm_Parity
    Comm_StopBits = .Comm_StopBits
    
    'Contrôle
    Control_BeepAfterMeasure = .Control_BeepAfterMeasure
    
    'Général
    Gen_ConfirmExit = .Gen_ConfirmExit
    
    'Impression
    Print_Designation = .Print_Designation
    Print_DocReference = .Print_DocReference
    Print_Font = .Print_Font
    Print_House = .Print_House
    Print_Lateral_Margin = .Print_Lateral_Margin
    Print_Position = .Print_Position
    Print_Service = .Print_Service
    Print_Tool = .Print_Tool
    
    'Resultat
    Result_LastCapacity = .Result_LastCapacity
    Result_LastLevel = .Result_LastLevel
    
    'Tips
    Tips_CurrentTip = .Tips_CurrentTip
    Tips_ShowTips = .Tips_ShowTips
    
    'Last
    Last_EquipmentManufacturer = .Last_EquipmentManufacturer
    Last_OperatorName = .Last_OperatorName
    
End With

End Sub
