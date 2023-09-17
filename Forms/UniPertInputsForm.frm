VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UniPertInputsForm 
   Caption         =   "Inputs"
   ClientHeight    =   4016
   ClientLeft      =   88
   ClientTop       =   408
   ClientWidth     =   4304
   OleObjectBlob   =   "UniPertInputsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UniPertInputsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub another_Click()
    iterations = UniPertInputsForm.inputIterations.Value
    mean = Range(UniPertInputsForm.inputMean.Value).Value
    max = UniPertInputsForm.inputMax.Value
    min = UniPertInputsForm.inputMin.Value
    varName = UniPertInputsForm.inputName.Value
    refcell = Range(UniPertInputsForm.inputMean).Address(External:=True)
    
    Application.Run "SampleAndRun.sample"
    
    Unload UniPertInputsForm
    
    DistSelectionForm.Show
End Sub

Private Sub okbutton_Click()
    iterations = CLng(UniPertInputsForm.inputIterations.Value)
    mean = Range(UniPertInputsForm.inputMean.Value).Value
    max = UniPertInputsForm.inputMax.Value
    min = UniPertInputsForm.inputMin.Value
    varName = UniPertInputsForm.inputName.Value
    refcell = Range(UniPertInputsForm.inputMean).Address(External:=True)
    
    Application.Run "SampleAndRun.sample"
    
    Unload UniPertInputsForm
End Sub

Private Sub cancelbutton_Click()
    Unload UniPertInputsForm
End Sub

Private Sub UserForm_Initialize()
    inputName.SetFocus
End Sub

