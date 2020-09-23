VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Right click the titlebar"
      Height          =   255
      Left            =   1260
      TabIndex        =   0
      Top             =   1080
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
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
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function CreateMenu Lib "user32" () As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_SEPARATOR = &H800&
Private Sub Form_Load()
    Dim hSysMenu As Long
    Dim nCnt As Long
    Dim NewMenuSeparator As Long
    hSysMenu = GetSystemMenu(Me.hwnd, False)
    If hSysMenu Then
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            NewMenu = CreateMenu
            InsertMenu hSysMenu, nCnt, MF_BYPOSITION Or MF_SEPARATOR, NewMenuSeparator, ""
            InsertMenu hSysMenu, nCnt + 1, MF_BYPOSITION, NewMenu, "&About"
        End If
    End If
    SubClass hwnd
End Sub

