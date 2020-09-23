VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adding to the system menu"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4260
      TabIndex        =   0
      Top             =   2940
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   2190
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "You could do just about anything when the menu is clicked. Showing this form is just an example."
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   2340
      TabIndex        =   1
      Top             =   1440
      Width           =   3045
   End
   Begin VB.Label lblTitle 
      Caption         =   "Adding a menu to the system menu"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   2625
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":2C10
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   255
      TabIndex        =   2
      Top             =   2700
      Width           =   3690
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
Option Explicit
Private Sub cmdOK_Click()
  Unload Me
End Sub
