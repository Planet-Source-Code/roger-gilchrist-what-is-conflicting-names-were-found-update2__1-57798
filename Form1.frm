VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "This is an almost useless program just to show the bug"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read about the problem"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "See the error message again"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()

  frmabout.Show

End Sub

Private Sub cmdRead_Click()

  Shell "explorer.exe What is Conflicting names were found.html", vbNormalFocus

End Sub
