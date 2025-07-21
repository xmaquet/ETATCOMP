VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FrmHelp 
   Caption         =   "Aide en ligne de EtaComp"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13035
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   13035
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   1005
      ButtonWidth     =   1085
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fermer"
            Key             =   "QUIT"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList32 
      Left            =   8400
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHelp.frx":058A
            Key             =   "QUIT"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************************
'FEUILLE D'AIDE EN LIGNE
'********************************************************************************

'********************************************************************************
'Constantes
'********************************************************************************
Private Const BUTTON_QUIT = "QUIT"

Private Sub Form_Load()
'Chargement de la feuille

Dim currentButton As MSComctlLib.Button
Dim valKey As String

'Initialiser la barre d'outils
With Toolbar
    .ImageList = ImageList32
    For Each currentButton In .Buttons
        valKey = currentButton.Key
        If valKey <> "" Then currentButton.Image = valKey
    Next
End With

'Libérer
Set currentButton = Nothing

'Charger la page Web
WebBrowser.Navigate2 App.Path + APP_WEB_PAGE

End Sub

Private Sub Form_Resize()
'Redimensionnement de la feuille

Const MIN_HEIGHT = 2000
Const MIN_WIDTH = 3000

Dim iToolBarHeight As Integer

'Sécurité
If WindowState = vbMinimized Then Exit Sub

'Tailles mini
If Height < MIN_HEIGHT Or Width < MIN_WIDTH Then

    'Redimensionner
    Move Left, Top, IIf(Width < MIN_WIDTH, MIN_WIDTH, Width), IIf(Height < MIN_HEIGHT, MIN_HEIGHT, Height)
    
    'Sortir pour éviter le double-appel
    Exit Sub
    
End If

'Redimensionnement
iToolBarHeight = Toolbar.Height
WebBrowser.Move 0, iToolBarHeight + 40, ScaleWidth, ScaleHeight - iToolBarHeight - 40

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Déchargement de la feuille

'Libérer
Set FrmHelp = Nothing

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'Clic sur un bouton de la barre d'outils

'Action selon le bouton
Select Case Button.Key
    Case BUTTON_QUIT
        Unload Me
End Select

End Sub
