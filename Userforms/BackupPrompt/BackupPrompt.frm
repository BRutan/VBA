VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BackupPrompt 
   Caption         =   "Backup Current Page"
   ClientHeight    =   1716
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   2952
   OleObjectBlob   =   "BackupPrompt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BackupPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Result As Boolean
Private Sub Userform_Initialize()
    Result = False
End Sub
Private Sub NoButton_Click()
    Result = False
    Me.Hide
End Sub

Private Sub YesButton_Click()
    Result = True
    Me.Hide
End Sub


