VERSION 5.00
Begin VB.Form F_SysTray 
   Caption         =   "Système Tray"
   ClientHeight    =   1140
   ClientLeft      =   9435
   ClientTop       =   435
   ClientWidth     =   2190
   Icon            =   "F_SysTray.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1140
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   720
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1560
      Picture         =   "F_SysTray.frx":000C
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   960
      Picture         =   "F_SysTray.frx":044E
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "F_SysTray.frx":0890
      Top             =   0
      Width           =   480
   End
   Begin VB.Menu menu 
      Caption         =   "Menu déroulant"
      Begin VB.Menu M_Action 
         Caption         =   "Activer"
      End
      Begin VB.Menu Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu Quitter 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "F_SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type IconeTray
    cbSize As Long      'Taille de l'icône (en octets)
    hWnd As Long        'Handle de la fenêtre chargée de recevoir les messages envoyés lors des évènements sur l'icône (clics, doubles-clics...)
    uID As Long         'Identificateur de l'icône
    uFlags As Long
    uCallbackMessage As Long    'Messages à renvoyer
    hIcon As Long               'Handle de l'icône
    szTip As String * 64        'Texte à mettre dans la bulle d'aide
End Type
Dim IconeT As IconeTray

'Constantes nécessaires
Private Const AJOUT = &H0
Private Const MODIF = &H1
Private Const SUPPRIME = &H2
Private Const MOUSEMOVE = &H200
Private Const MESSAGE = &H1
Private Const Icone = &H2
Private Const TIP = &H4

Private Const DOUBLE_CLICK_GAUCHE = &H203
Private Const BOUTON_GAUCHE_POUSSE = &H201
Private Const BOUTON_GAUCHE_LEVE = &H202
Private Const DOUBLE_CLICK_DROIT = &H206
Private Const BOUTON_DROIT_POUSSE = &H204
Private Const BOUTON_DROIT_LEVE = &H205

'API nécessaire
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As IconeTray) As Boolean

Private Sub Form_Load()
'Préparation de la variable IconeT
IconeT.cbSize = Len(IconeT) 'Taille de l'icône en octet
IconeT.hWnd = Me.hWnd       'Handle de l'application (pour qu'elle reçoive les messages envoyés lors d'un clic, double-clic...
IconeT.uID = 1&             'Identificateur de l'icône
IconeT.uFlags = Icone Or TIP Or MESSAGE
IconeT.uCallbackMessage = MOUSEMOVE     'Renvoyer les messages concernant l'action de la souris
IconeT.hIcon = Image1.Picture   'Mettre en icône l'image qui est dans le contrôle "Image1"
IconeT.szTip = "Systeme tray" & Chr$(0) 'Texte de la bulle d'aide
'Appel de la fonction pour mettre l'icône dans le système tray
Shell_NotifyIcon AJOUT, IconeT
Me.Hide     'Cache la fenêtre
App.TaskVisible = False     'Retire le bouton de l'application de la barre                          'des tâches
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static rec As Boolean, msg As Long
'Se produit lorsque l'utilisateur agit avec la souris sur
'l'icône placée dans le système tray
msg = X / Screen.TwipsPerPixelX
If rec = False Then
    rec = True
    Select Case msg     'Différentes possibilité d'action
        Case DOUBLE_CLICK_GAUCHE:   'mettez
              If M_Action.Enabled Then
                M_Action_Click 'ici
              End If
        Case BOUTON_GAUCHE_POUSSE:  'ce
        Case BOUTON_GAUCHE_LEVE:    'que
        Case DOUBLE_CLICK_DROIT:    'vous
        Case BOUTON_DROIT_POUSSE:   'voudrez
        Case BOUTON_DROIT_LEVE:     'qu'il se passe
            PopupMenu Menu, , , , M_Action     'fait apparaitre le menu
            '"A propos de" apparaitra en gras
    End Select
    rec = False
End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Refait appel à l'API pour retirer l'icône du système tray
'lorsque le programme se ferme, en utilisant cette fois la constante SUPPRIME
'au lieu de AJOUT
IconeT.cbSize = Len(IconeT)
IconeT.hWnd = Me.hWnd
IconeT.uID = 1&
Shell_NotifyIcon SUPPRIME, IconeT
End Sub

Private Sub M_Action_Click()
If M_Action.Caption = "Activer" Then
    M_Action.Caption = "Pause"
    Timer1.Enabled = True
Else
    M_Action.Caption = "Activer"
    Timer1.Enabled = False
    IconeT.hIcon = Image1.Picture
    Shell_NotifyIcon MODIF, IconeT
End If
End Sub

Private Sub Quitter_Click()
Unload Me   'retire la fenêtre
End Sub

Private Sub Timer1_Timer()
If IconeT.hIcon = Image2.Picture Then
    IconeT.hIcon = Image3.Picture
    Shell_NotifyIcon MODIF, IconeT
Else
    IconeT.hIcon = Image2.Picture
    Shell_NotifyIcon MODIF, IconeT
End If
End Sub
