VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "OK not the real error msgbox but the one on my machine"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9795
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgAbout 
      Height          =   2085
      Left            =   0
      Picture         =   "frmabout.frx":0000
      Top             =   0
      Width           =   9765
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgAbout_Click()

  Unload Me

End Sub

