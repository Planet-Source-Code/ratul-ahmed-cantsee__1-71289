VERSION 5.00
Begin VB.Form frmabut 
   BorderStyle     =   0  'None
   Caption         =   "frmabut"
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   Icon            =   "frmabut.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmabut.frx":628A
   ScaleHeight     =   2220
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmabut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload frmabut

End Sub

Private Sub Form_Load()
StayOnTop frmabut
End Sub
