VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   " C++ Buttons"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command3 
      Caption         =   "Deaktivierter C++ Button"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "C++ Button"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Normaler VB button"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________________________________________
'|                                                                                                             |
'|Code Name.........: C++ Buttons                                                                              |
'|Code Autor........: Zpage                                                                                    |
'|E-Mail vom mir....: Zpage@gmx.net                                                                            |
'|HQ von mir........: h**t://www.GTApalace.de.vu                                                               |
'|Relase Team.......: LegalAccessNavy (LAN)                                                                    |
'|Relase Team's HQ..: h**p://www.LegalAccessNavy.de.vu                                                         |
'|Beschreibung......: Dieser code ändert einfache VB Buttons in hübsche C++ Buttons um.                        |
'|Copyright.........: Wenn du den code editierst, dann lass auf jeden Fall die Copyright-Zeile stehen.         |
'|You find this in English on http://www.planet-source-code.com/.Just enter in the Search C++ Buttons [English]|                                                                                  |
'|_____________________________________________________________________________________________________________|
'
'The folloing code MUST be in every App. Just set it in the code on the top of you app.

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const WS_DISABLED = &H58018000

Private Sub Form_Load()
'For every Button you must made 'btnFlat [ButtonName]'
'This defines the First button
btnFlat Command1
'This the second
btnFlat Command2
'And this the third
btnFlat Command3
End Sub

Function btnFlat(Button As CommandButton)
'Here also: For every Button you must made this code:
    SetWindowLong Command1.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    Command1.Visible = True
    SetWindowLong Command2.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    Command2.Visible = True
    SetWindowLong Command3.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT Or WS_DISABLED
    Command3.Visible = True
'If you want that the Button is Disabled, then add after 'BT_FLAT' the folloing: 'Or WS_DISABLED'
End Function
