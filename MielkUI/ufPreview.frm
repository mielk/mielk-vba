VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPreview 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "ufPreview.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Click()
    Debug.Print ComboBox1.text
    Debug.Print ComboBox1.value
End Sub

Private Sub UserForm_Initialize()
    With ComboBox1
        Call .addItem("a")
        Call .addItem("b")
        Call .addItem("c")
    End With
End Sub
