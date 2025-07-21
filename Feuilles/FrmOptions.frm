VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "FrmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   5360
      Begin VB.Frame FrameTips 
         Caption         =   "Astuces du jour"
         Height          =   735
         Left            =   0
         TabIndex        =   39
         Top             =   1680
         Width           =   5360
         Begin VB.CheckBox ChkTips 
            Caption         =   "Afficher les astuces du jour au démarrage"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame FrameGeneral 
         Caption         =   "Valeurs par défaut"
         Height          =   735
         Left            =   0
         TabIndex        =   3
         Top             =   840
         Width           =   5360
         Begin VB.CheckBox ChkConfirmExit 
            Caption         =   "Demander confirmation avant d'arrêter le programme"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Frame FrameMeasure 
         Caption         =   "Mesure"
         Height          =   735
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   5360
         Begin VB.CheckBox ChkBeep 
            Caption         =   "Emettre un beep après chaque mesure"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   3135
         End
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Liaison série RS232"
      Height          =   3855
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   420
      Visible         =   0   'False
      Width           =   5360
      Begin VB.ComboBox ComboPort 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   604
         Width           =   5175
      End
      Begin VB.ComboBox ComboBauds 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1212
         Width           =   5175
      End
      Begin VB.ComboBox ComboParity 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1820
         Width           =   5175
      End
      Begin VB.ComboBox ComboDataBit 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2428
         Width           =   5175
      End
      Begin VB.ComboBox ComboStopBit 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3036
         Width           =   5175
      End
      Begin VB.CommandButton CmdNowApply 
         Caption         =   "&Appliquer les paramètres maintenant"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3400
         Width           =   5175
      End
      Begin VB.Label LblPort 
         AutoSize        =   -1  'True
         Caption         =   "Port série:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   705
      End
      Begin VB.Label LblBauds 
         AutoSize        =   -1  'True
         Caption         =   "Vitesse de transmission (BAUDS):"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   968
         Width           =   2370
      End
      Begin VB.Label LblParity 
         AutoSize        =   -1  'True
         Caption         =   "Valeur de parité:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1576
         Width           =   1155
      End
      Begin VB.Label LblDataBit 
         AutoSize        =   -1  'True
         Caption         =   "Bits de données valides:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   2184
         Width           =   1725
      End
      Begin VB.Label LblStopBits 
         AutoSize        =   -1  'True
         Caption         =   "Bits d'arrêt:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2792
         Width           =   780
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Complément d'information"
      Height          =   3855
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   5360
      Begin VB.TextBox TxtHouse 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   602
         Width           =   5175
      End
      Begin VB.TextBox TxtService 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1176
         Width           =   5175
      End
      Begin VB.TextBox TxtDesignation 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1750
         Width           =   5175
      End
      Begin VB.TextBox TxtRefDoc 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2324
         Width           =   5175
      End
      Begin VB.TextBox TxtPosition 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   5175
      End
      Begin VB.TextBox TxtTool 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   2898
         Width           =   5175
      End
      Begin VB.Label LblHouse 
         AutoSize        =   -1  'True
         Caption         =   "Etablissement:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label LblService 
         AutoSize        =   -1  'True
         Caption         =   "Atelier :"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   934
         Width           =   525
      End
      Begin VB.Label LblDesignation 
         AutoSize        =   -1  'True
         Caption         =   "Désignation du comparateur:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1508
         Width           =   2040
      End
      Begin VB.Label LblDocReference 
         AutoSize        =   -1  'True
         Caption         =   "Document de référence:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2082
         Width           =   1725
      End
      Begin VB.Label LblPosition 
         AutoSize        =   -1  'True
         Caption         =   "Position du comparateur lors du contrôle:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   3230
         Width           =   2880
      End
      Begin VB.Label LblTool 
         AutoSize        =   -1  'True
         Caption         =   "Outillage:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2656
         Width           =   660
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   4
      Left            =   120
      TabIndex        =   33
      Top             =   420
      Width           =   5360
      Begin VB.Frame FrameMargin 
         Caption         =   "Marge horizontale en mm"
         Height          =   855
         Left            =   0
         TabIndex        =   36
         Top             =   3000
         Width           =   5340
         Begin MSComctlLib.Slider SliderMargin 
            Height          =   495
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   873
            _Version        =   393216
            Max             =   20
            TickStyle       =   2
            TextPosition    =   1
         End
         Begin VB.Label LblMargin 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4680
            TabIndex        =   38
            Top             =   405
            Width           =   525
         End
      End
      Begin VB.Frame FrameFont 
         Caption         =   "Police"
         Height          =   2895
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   5340
         Begin VB.ListBox ListFont 
            Height          =   2400
            ItemData        =   "FrmOptions.frx":0442
            Left            =   80
            List            =   "FrmOptions.frx":0444
            TabIndex        =   35
            Top             =   360
            Width           =   5175
         End
      End
   End
   Begin VB.CommandButton CmdValid 
      Caption         =   "Valider"
      Height          =   375
      Left            =   3000
      TabIndex        =   32
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   4320
      TabIndex        =   31
      Top             =   4440
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   4335
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Général"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Communication"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Procès verbaux"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Impression"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************
'FEUILLE DE CONTROLE DES OPTIONS DU PROGRAMME
'********************************************************************************

'********************************************************************************
'Données membres
'********************************************************************************
Private m_CopyOfConfig As Configuration
Private m_RefOfConfig As Configuration

Private m_IndxTab As Integer

Private Sub ChkBeep_Click()
'Clic sur la case à cocher "... Beep ..."

'Affecter
m_CopyOfConfig.Control_BeepAfterMeasure = CBool(ChkBeep)

End Sub

Private Sub ChkConfirmExit_Click()
'Clic sur la case à cocher "Demander confirmation ..."

'Affecter
m_CopyOfConfig.Gen_ConfirmExit = CBool(ChkConfirmExit)

End Sub

Private Sub ChkTips_Click()
'Clic sur la case à cocher "Astuces du jour"

m_CopyOfConfig.Tips_ShowTips = CBool(ChkTips.Value)

End Sub

Private Sub CmdCancel_Click()
'Clic sur le bouton "Annuler"

'Décharger la feuille
Unload Me

End Sub

Private Sub CmdNowApply_Click()
'Clic sur le bouton "Appliquer maintenant"

Dim strSave As String

Dim iSave As Integer
Dim iComboIndex(1 To 5) As Integer

'Les erreurs sont gérées
On Error GoTo GestErr

With FrmGen.MsComm

    'Sauvegarder les paramètres actuellement en vigueur
    strSave = .Settings
    iSave = .CommPort
    iComboIndex(1) = ComboPort.ListIndex
    iComboIndex(2) = ComboBauds.ListIndex
    iComboIndex(3) = ComboParity.ListIndex
    iComboIndex(4) = ComboDataBit.ListIndex
    iComboIndex(5) = ComboStopBit.ListIndex

    'Fermer le port si ouvert
    If .PortOpen Then
        
        .PortOpen = False
        FrmGen.StatusBar.Panels(1).Text = "RS232 KO"
        
    End If
    
    'Préparer un nouveau Settings
    .Settings = ComboBauds.Text + "," + ComboParity.Text + "," + ComboDataBit.Text + "," + ComboStopBit
    .CommPort = ComboPort.ListIndex + 1
    
    'Appliquer les nouveaux paramètres
    .PortOpen = True
    
    'Mettre à jour
    FrmGen.StatusBar.Panels(1).Text = "RS232 OK"
    
End With

'Appliquer la nouvelle configuration
With m_CopyOfConfig
    .Comm_Bauds = CInt(ComboBauds.Text)
    .Comm_CommPort = ComboPort.ListIndex + 1
    .Comm_DataBits = CInt(ComboDataBit.Text)
    .Comm_Parity = ComboParity.Text
    .Comm_StopBits = CInt(ComboStopBit.Text)
End With

'Afficher un message
strSave = "La nouvelle configuration du port série RS232 est supportée par le système."
iSave = MsgBox(strSave, vbInformation, "Communication")

'Sortir normalement
Exit Sub

'Gestion des erreurs
GestErr:

'Restaurer les anciens paramètres
With FrmGen.MsComm
    .Settings = IIf(strSave <> "", strSave, "1200,E,7,1")
    .CommPort = IIf(iSave > 0, iSave, 1)
    .PortOpen = True
End With

'Mettre à jour l'interface
ComboPort.ListIndex = iComboIndex(1)
ComboBauds.ListIndex = iComboIndex(2)
ComboParity.ListIndex = iComboIndex(3)
ComboDataBit.ListIndex = iComboIndex(4)
ComboStopBit.ListIndex = iComboIndex(5)
    
'Afficher un message
strSave = "La nouvelle configuration du port série RS232 n'est pas supportée par le système." + vbCrLf + _
          "La restauration des anciens paramètres a été effectuée." + vbCrLf + vbCrLf + _
          "Description de l'erreur: " + Err.Description
iSave = MsgBox(strSave, vbCritical, "Communication - Erreur")

End Sub

Private Sub CmdValid_Click()
'Clic sur le bouton "Valider"

'Enregistrer la configuration du port série
With m_CopyOfConfig
    'Enregistrer la configuration du port série
    .Comm_Bauds = CInt(ComboBauds.Text)
    .Comm_CommPort = ComboPort.ListIndex + 1
    .Comm_DataBits = CInt(ComboDataBit.Text)
    .Comm_Parity = ComboParity.Text
    .Comm_StopBits = CInt(ComboStopBit.Text)
    'Enregistrer la configuration de l'impression
    .Print_Designation = TxtDesignation.Text
    .Print_DocReference = TxtRefDoc.Text
    .Print_Tool = TxtTool.Text
    .Print_Service = TxtService.Text
    .Print_House = TxtHouse.Text
    .Print_Position = TxtPosition.Text
End With

'Copier
m_RefOfConfig.Copy m_CopyOfConfig

'Décharger la feuille
Unload Me

End Sub

Private Sub Form_Load()
'Chargement de la feuille

Dim varArray As Variant
Dim i As Integer

'Initialisation graphique
Screen.MousePointer = vbHourglass

'Initialiser les données membres
m_IndxTab = 1

'Charger les valeurs dans les zones de liste

'Liste des ports
For i = 1 To 4
    ComboPort.AddItem "COM" + Format(i)
Next

'Liste des BAUDS
varArray = Array("110", "300", "600", "1200", "2400", "4800", "9600", "14400", "19200", "28800", "38400", _
"56000", "128000", "256000")
For i = LBound(varArray) To UBound(varArray)
    ComboBauds.AddItem varArray(i)
Next

'Liste des parités
varArray = Array("E", "M", "N", "O", "S")
For i = LBound(varArray) To UBound(varArray)
    ComboParity.AddItem varArray(i)
Next

'Liste des bits de données
For i = 4 To 8
    ComboDataBit.AddItem Format(i)
Next

'Liste des bits de stop
For i = 1 To 2
    ComboStopBit.AddItem Format(i)
Next

'Liste des polices
For i = 0 To Printer.FontCount - 1
    
    'Ajouter
    ListFont.AddItem Printer.Fonts(i)
    
Next

'Libérer
Set varArray = Nothing

'Restauration graphique
Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Déchargement de la feuille

'Libérer
Set m_CopyOfConfig = Nothing
Set m_RefOfConfig = Nothing
Set FrmOptions = Nothing

End Sub

Public Sub SetReferenceToConfig(objConf As Configuration)
'Affecter la configuration courante à la feuille

Dim i As Integer

'Copier la référence
Set m_RefOfConfig = objConf

'Copier l'objet
Set m_CopyOfConfig = New Configuration
m_CopyOfConfig.Copy m_RefOfConfig

With m_CopyOfConfig

    'Initialiser l'interface graphique
    ChkBeep = IIf(.Control_BeepAfterMeasure, 1, 0)
    ChkConfirmExit = IIf(.Gen_ConfirmExit, 1, 0)
    ChkTips = IIf(.Tips_ShowTips, 1, 0)
    
    ComboBauds.Text = Format(.Comm_Bauds)
    ComboDataBit.Text = Format(.Comm_DataBits)
    ComboParity.Text = .Comm_Parity
    ComboPort.Text = "COM" + Format(.Comm_CommPort)
    ComboStopBit.Text = Format(.Comm_StopBits)

    TxtDesignation = .Print_Designation
    TxtHouse = .Print_House
    TxtPosition = .Print_Position
    TxtRefDoc = .Print_DocReference
    TxtService = .Print_Service
    TxtTool = .Print_Tool
    
    SliderMargin.Value = .Print_Lateral_Margin
    
End With

'Pointer sur la police
With ListFont
    For i = 0 To .ListCount - 1
        
        'On sélectionne si la police correspond
        If .List(i) = m_CopyOfConfig.Print_Font Then
            
            .ListIndex = i
            Exit For
            
        End If
        
    Next
End With

End Sub

Private Sub ListFont_Click()
'Clic sur la liste des polices

'Affecter
m_CopyOfConfig.Print_Font = ListFont.List(ListFont.ListIndex)

End Sub

Private Sub SliderMargin_Change()
'Modification de la marge

'Mise à jour
m_CopyOfConfig.Print_Lateral_Margin = CInt(SliderMargin.Value)
LblMargin = Format(m_CopyOfConfig.Print_Lateral_Margin) + " mm"

End Sub

Private Sub TabStrip_Click()
'Clic sur un onglet du TabStrip

'Optimisation
If TabStrip.SelectedItem.Index = m_IndxTab Then Exit Sub

'Basculer
Frame(m_IndxTab).Visible = False
m_IndxTab = TabStrip.SelectedItem.Index
Frame(m_IndxTab).Visible = True

End Sub

