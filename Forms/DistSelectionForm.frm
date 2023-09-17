VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DistSelectionForm 
   Caption         =   "Distribution Selection"
   ClientHeight    =   1352
   ClientLeft      =   88
   ClientTop       =   408
   ClientWidth     =   4304
   OleObjectBlob   =   "DistSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DistSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub okbutton_Click()
    dist = DistSelectionForm.DistSelector.Value
    
    If dist = "" Then
        MsgBox "No Items Selected!", vbCritical, "Error"
        Exit Sub
    End If
    
    Select Case dist
    
    Case "Normal"
        Unload DistSelectionForm
        NormLognormInputsForm.Show
    Case "Lognormal"
        Unload DistSelectionForm
        NormLognormInputsForm.Show
    Case "Inverse Lognormal"
        Unload DistSelectionForm
        NormLognormInputsForm.Show
    Case "Uniform"
        Unload DistSelectionForm
        UserForm4.Show
    Case "PERT"
        Unload DistSelectionForm
        UserForm4.Show
    Case "Binomial"
        Unload DistSelectionForm
        UserForm5.Show
    Case Else
        MsgBox "Error", vbCritical, "Error"
        Exit Sub
    End Select
End Sub

Private Sub cancelbutton_Click()
    Unload DistSelectionForm
End Sub

Private Sub UserForm_Initialize()
    With DistSelectionForm.DistSelector
        .AddItem "Normal"
        .AddItem "Lognormal"
        .AddItem "Inverse Lognormal"
        .AddItem "Uniform"
        .AddItem "PERT"
        .AddItem "Binomial"
    End With
End Sub

