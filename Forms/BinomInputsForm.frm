VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BinomInputsForm 
   Caption         =   "Inputs"
   ClientHeight    =   3432
   ClientLeft      =   88
   ClientTop       =   408
   ClientWidth     =   4304
   OleObjectBlob   =   "BinomInputsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BinomInputsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub another_Click()
    iterations = BinomInputsForm.inputIterations.Value
    mean = Range(BinomInputsForm.inputMean.Value).Value
    prob = BinomInputsForm.inputProb.Value
    varName = BinomInputsForm.inputName.Value
    refcell = Range(BinomInputsForm.inputMean).Address(External:=True)
    
    Application.Run "SampleAndRun.sample"
    
    Unload BinomInputsForm
    
    DistSelectionForm.Show
End Sub

Private Sub okbutton_Click()
    iterations = BinomInputsForm.inputIterations.Value
    mean = Range(BinomInputsForm.inputMean.Value).Value
    prob = BinomInputsForm.inputProb.Value
    varName = BinomInputsForm.inputName.Value
    refcell = Range(BinomInputsForm.inputMean).Address(External:=True)
    
    Application.Run "SampleAndRun.sample"
    
    Unload BinomInputsForm
End Sub

Private Sub cancelbutton_Click()
    Unload BinomInputsForm
End Sub

Private Sub UserForm_Initialize()
    inputName.SetFocus
End Sub

