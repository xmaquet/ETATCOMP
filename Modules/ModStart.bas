Attribute VB_Name = "ModStart"
Option Explicit

' Ajout suite � �volution du programme en 2013
'librairie port.dll utilis� pour fonctionnemnt avec adaptatateur USB-Serie

Declare Function OPENCOM Lib "c:\windows\system32\Port.dll" (ByVal a$) As Integer
Declare Sub CLOSECOM Lib "Port" ()
Declare Sub SENDBYTE Lib "Port" (ByVal b%)
Declare Function READBYTE Lib "Port" () As Integer
Declare Sub DTR Lib "Port" (ByVal b%)
Declare Sub RTS Lib "Port" (ByVal b%)
Declare Sub TXD Lib "Port" (ByVal b%)
Declare Function CTS Lib "Port" () As Integer
Declare Function DSR Lib "Port" () As Integer
Declare Function RI Lib "Port" () As Integer
Declare Function DCD Lib "Port" () As Integer
Declare Sub DELAY Lib "Port" (ByVal b%)
Declare Sub DELAYUS Lib "Port" (ByVal l As Long)
Declare Sub TIMEINIT Lib "Port" ()
Declare Sub TIMEINITUS Lib "Port" ()
Declare Function TIMEREAD Lib "Port" () As Long
Declare Function TIMEREADUS Lib "Port" () As Long
Declare Sub REALTIME Lib "Port" (ByVal i As Boolean)
Declare Sub OUTPORT Lib "Port" (ByVal adr%, b%)
Declare Function INPORT Lib "Port" () As Integer

Public Sub Ouvrir_Port(vCom As String)
Dim P As Integer
On Error Resume Next
CLOSECOM
DELAYUS 500
P = OPENCOM(vCom)
DELAYUS 500
If P = 0 Then
    MsgBox "Ouverture port impossible"
    Exit Sub
End If
End Sub

Public Function Reception() As String
Dim vAscii As String, vChaine As String, R As Integer
R = READBYTE
Do While (R <> 13) 'And Not (Arret)
    DoEvents
    If R > -1 Then
        vAscii = Chr(R)
        vChaine = vChaine & vAscii
    End If
    R = READBYTE
Loop
Reception = vChaine
End Function

Public Sub Fermer_Port()
CLOSECOM
End Sub

' fin ajout

'********************************************************************************
'MODULE DE DEMARRAGE DE L'APPLICATION
'********************************************************************************

Private Sub Main()
'D�marrage du programme

Dim txt As String
Dim i As Integer

'Si le programme est d�j� en cours d'ex�cution
If App.PrevInstance Then

    'Afficher un message
    txt = App.EXEName + " est d�j� en cours d'ex�cution sur le syst�me." + vbCrLf + vbCrLf + _
          "D�marrage annul� !"
    i = MsgBox(txt, vbCritical, "D�marrage de " + App.EXEName)
    
    'Quitter
    End
    
End If

'S'il n'y a pas d'imprimante install�e sur le syst�me
If Printers.Count = 0 Then

    'Afficher un message
    txt = App.EXEName + " ne peut pas �tre ex�cut� correctement car il n'y a pas" + vbCrLf + _
         "d'imprimante install�e sur le syst�me."
    i = MsgBox(txt, vbCritical, "D�marrage de " + App.EXEName)
    
    'Quitter
    End
    
End If

'Afficher le Splash
#If NOSPLASH Then

    With FrmGen
        .Enabled = True
        .Show vbModeless
    End With

#Else

    FrmSplash.Show vbModeless

#End If

End Sub
