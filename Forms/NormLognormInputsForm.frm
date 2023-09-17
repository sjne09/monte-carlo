VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NormLognormInputsForm 
   Caption         =   "Inputs"
   ClientHeight    =   3432
   ClientLeft      =   88
   ClientTop       =   408
   ClientWidth     =   4560
   OleObjectBlob   =   "NormLognormInputsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NormLognormInputsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub another_Click()
    iterations = NormLognormInputsForm.inputIterations.Value
    mean = Range(NormLognormInputsForm.inputMean.Value).Value
    stddev = NormLognormInputsForm.inputStddev.Value
    varName = NormLognormInputsForm.inputName.Value
    refcell = Range(NormLognormInputsForm.inputMean).Address(External:=True)
    
    Application.Run "SampleAndRun.sample"
    
    Unload NormLognormInputsForm
    
    DistSelectionForm.Show
End Sub

Private Sub okbutton_Click()
    iterations = NormLognormInputsForm.inputIterations.Value
    mean = Range(NormLognormInputsForm.inputMean.Value).Value
    stddev = NormLognormInputsForm.inputStddev.Value
    varName = NormLognormInputsForm.inputName.Value
    refcell = Range(NormLognormInputsForm.inputMean).Address(External:=True)
    
    Application.Run "SampleAndRun.sample"
    
    Unload NormLognormInputsForm
End Sub

Private Sub cancelbutton_Click()
    Unload NormLognormInputsForm
End Sub

Private Sub UserForm_Initialize()
    inputName.SetFocus
End Sub

