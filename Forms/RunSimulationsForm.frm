VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RunSimulationsForm 
   Caption         =   "MCS"
   ClientHeight    =   1816
   ClientLeft      =   88
   ClientTop       =   408
   ClientWidth     =   4304
   OleObjectBlob   =   "RunSimulationsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RunSimulationsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelbutton_Click()
    Unload RunSimulationsForm
End Sub

Private Sub okbutton_Click()
    Set inputhead = Range(RunSimulationsForm.inputHeader)
    dependent = Range(RunSimulationsForm.depvar).Address(External:=True)

    Application.Run "SampleAndRun.run_mcs"
    Unload RunSimulationsForm
End Sub

