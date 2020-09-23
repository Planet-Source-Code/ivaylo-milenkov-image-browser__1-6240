VERSION 5.00
Begin VB.Form frmImage 
   BorderStyle     =   0  'None
   Caption         =   "    "
   ClientHeight    =   1980
   ClientLeft      =   90
   ClientTop       =   -90
   ClientWidth     =   1935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Browse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   1935
   Begin VB.Image imgTwo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Move 0, 0
End Sub

Private Sub imgTwo_Click()
    frmBrowser.Show
    frmImage.Hide
End Sub
