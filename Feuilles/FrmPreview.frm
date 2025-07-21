VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPreview 
   Caption         =   "Aperçu avant impression"
   ClientHeight    =   7290
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7740
   Icon            =   "FrmPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   128.588
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   136.525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PctCont 
      Height          =   5775
      Left            =   600
      ScaleHeight     =   100.806
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   117.74
      TabIndex        =   1
      Top             =   720
      Width           =   6735
      Begin VB.PictureBox PctBack 
         BorderStyle     =   0  'None
         Height          =   283
         Left            =   480
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   5
         Top             =   240
         Width           =   283
      End
      Begin MSComCtl2.FlatScrollBar FlatHScrollBar 
         Height          =   225
         Left            =   0
         TabIndex        =   4
         Top             =   4200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         LargeChange     =   50
         Max             =   150
         Orientation     =   1179649
         SmallChange     =   20
      End
      Begin MSComCtl2.FlatScrollBar FlatVScrollBar 
         Height          =   1455
         Left            =   960
         TabIndex        =   3
         Top             =   1080
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   2566
         _Version        =   393216
         Appearance      =   0
         LargeChange     =   100
         Max             =   220
         Orientation     =   1179648
         SmallChange     =   10
      End
      Begin VB.PictureBox PctView 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   16838
         Left            =   567
         ScaleHeight     =   296.598
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   209.55
         TabIndex        =   2
         Top             =   567
         Width           =   11906
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimer"
            Key             =   "PRINT"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fermer"
            Key             =   "QUIT"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList32 
      Left            =   6720
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPreview.frx":014A
            Key             =   "QUIT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPreview.frx":02A4
            Key             =   "PRINT"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************
'FEUILLE D'APERCU AVANT IMPRESSION
'********************************************************************************

'********************************************************************************
'Constantes
'********************************************************************************
Private Const BUTTON_PRINT = "PRINT"
Private Const BUTTON_QUIT = "QUIT"

Private Const MARGIN = 10

'********************************************************************************
'Données membres
'********************************************************************************
Private mRefToControlObject As ControlObject
Private mRefToConfigObject As Configuration

Private Sub FlatHScrollBar_Change()
'Défilement horizontal

'Positionner
PctView.Left = -(FlatHScrollBar.Value) + MARGIN

End Sub

Private Sub FlatHScrollBar_Scroll()
'Défilement horizontal

'Positionner
PctView.Left = -(FlatHScrollBar.Value) + MARGIN

End Sub

Private Sub FlatVScrollBar_Change()
'Défilement vertical

'Positionner
PctView.Top = -(FlatVScrollBar.Value) + MARGIN

End Sub

Private Sub FlatVScrollBar_Scroll()
'Défilement vertical

'Positionner
PctView.Top = -(FlatVScrollBar.Value) + MARGIN

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Une touche est appuyée au-dessus de la feuille

'Action selon la touche
Select Case KeyCode
    Case vbKeyDown
        If FlatVScrollBar.Value < FlatVScrollBar.Max - FlatVScrollBar.SmallChange Then
            FlatVScrollBar.Value = FlatVScrollBar.Value + FlatVScrollBar.SmallChange
            FlatVScrollBar_Change
        Else
            FlatVScrollBar.Value = FlatVScrollBar.Max
            FlatVScrollBar_Change
        End If
    Case vbKeyUp
        If FlatVScrollBar.Value > FlatVScrollBar.SmallChange Then
            FlatVScrollBar.Value = FlatVScrollBar.Value - FlatVScrollBar.SmallChange
            FlatVScrollBar_Change
        Else
            FlatVScrollBar.Value = FlatVScrollBar.Min
            FlatVScrollBar_Change
        End If
    Case vbKeyPageDown, vbKeyEnd
        FlatVScrollBar.Value = FlatVScrollBar.Max
        FlatVScrollBar_Change
    Case vbKeyPageUp, vbKeyHome
        FlatVScrollBar.Value = FlatVScrollBar.Min
        FlatVScrollBar_Change
    Case vbKeyRight
        If FlatHScrollBar.Value < FlatHScrollBar.Max - FlatHScrollBar.SmallChange Then
            FlatHScrollBar.Value = FlatHScrollBar.Value + FlatHScrollBar.SmallChange
            FlatHScrollBar_Change
        Else
            FlatHScrollBar.Value = FlatHScrollBar.Max
            FlatHScrollBar_Change
        End If
    Case vbKeyLeft
        If FlatHScrollBar.Value > FlatHScrollBar.SmallChange Then
            FlatHScrollBar.Value = FlatHScrollBar.Value - FlatHScrollBar.SmallChange
            FlatHScrollBar_Change
        Else
            FlatHScrollBar.Value = FlatHScrollBar.Min
            FlatHScrollBar_Change
        End If
End Select

End Sub

Private Sub Form_Load()
'Chargement de la feuille

Dim currentButton As MSComctlLib.Button
Dim valKey As String

'Travail sur la barre d'outils
With Toolbar

    'Imagelist
    .ImageList = ImageList32
    
    'Enumérer les bouttons
    For Each currentButton In .Buttons
        
        'Affeter l'image
        valKey = currentButton.Key
        If valKey <> "" Then currentButton.Image = valKey
        
    Next
    
End With

'Libérer
Set currentButton = Nothing

'Positionnement de la vue
With PctView
    .Left = MARGIN
    .Top = MARGIN
End With

'Initialisation des barre de défielement
With FlatHScrollBar
    .Max = 150
End With
With FlatVScrollBar
    .Max = 220
End With

End Sub

Private Sub Form_Resize()
'Redimensionnement de la feuille

Dim bVisible As Boolean

Const MIN_HEIGHT = 1500
Const MIN_WIDTH = 6000

'Sécurité
If WindowState = vbMinimized Then Exit Sub

'Tailles mini
If Height < MIN_HEIGHT Or Width < MIN_WIDTH Then
    Move Left, Top, IIf(Width < MIN_WIDTH, MIN_WIDTH, Width), IIf(Height < MIN_HEIGHT, MIN_HEIGHT, Height)
    Exit Sub
End If

'Redimensionnement du conteneur
PctCont.Move 0, Toolbar.Height, ScaleWidth, ScaleHeight - Toolbar.Height

'Travail sur la barre de défilement horizontale
 If PctCont.Width > 210 + (2 * MARGIN) Then
    bVisible = False
    PctView.Left = (PctCont.Width - 210) / 2
Else
    bVisible = True
End If
FlatHScrollBar.Visible = bVisible
PctBack.Visible = bVisible

'Redimensionnement des objets
FlatVScrollBar.Move PctCont.Width - FlatVScrollBar.Width - 1, 0, FlatVScrollBar.Width, PctCont.Height - IIf(bVisible, PctBack.Height, 0)
If bVisible Then FlatHScrollBar.Move 0, PctCont.Height - FlatHScrollBar.Height - 1, PctCont.Width - FlatVScrollBar.Width - 1
PctBack.Move FlatHScrollBar.Width, FlatVScrollBar.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Déchargement de la feuille

'Libérer
Set mRefToControlObject = Nothing
Set mRefToConfigObject = Nothing
Set FrmPreview = Nothing

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'Clic sur un boutton de la barre d'outils

'Action selon le bouton
Select Case Button.Key
    Case BUTTON_PRINT
        FrmGen.ConnectToMnuFile_Print_Slot
    Case BUTTON_QUIT
        Unload Me
End Select

End Sub


Public Sub SetRefToObjects(objControl As ControlObject, objConfiguration As Configuration)
'Affecter la référence à l'objet principal

'Affecter
Set mRefToControlObject = objControl
Set mRefToConfigObject = objConfiguration

'Dessiner
mRefToControlObject.PrintOn PctView, mRefToConfigObject

End Sub
